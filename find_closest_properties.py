"""
Find the closest rental properties to an input address.

This script reads addresses from "Housing Database.xlsx", geocodes them with a
free provider, and ranks the nearest properties by straight-line distance.

Default provider:
    OpenStreetMap Nominatim

Fallback provider:
    US Census Geocoder and ArcGIS World Geocoder, used automatically for
    workbook rows that Nominatim cannot resolve. These fallbacks do not require
    API keys.

Optional provider:
    OpenCage Geocoder, if OPENCAGE_API_KEY is set and --provider opencage is used.

Usage:
    python find_closest_properties.py
    python find_closest_properties.py --address "2020 S Hacienda Blvd Hacienda Heights, CA 91745"
    python find_closest_properties.py --top 5 --provider nominatim

When no --address is provided, the script asks for one. Press Enter to use the
default address: 2020 S Hacienda Blvd Hacienda Heights, CA 91745.

Notes:
    Nominatim is free but usage-policy limited. The script includes a cache and
    a one-second delay between uncached requests to be gentle with the service.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import pandas as pd


DEFAULT_ADDRESS = "2020 S Hacienda Blvd Hacienda Heights, CA 91745"
DEFAULT_EXCEL_FILE = "Housing Database.xlsx"
DEFAULT_CACHE_FILE = "geocode_cache.json"
DEFAULT_TOP_N = 3

REQUIRED_COLUMNS = [
    "Rental Property Address",
    "Street Address Line 2",
    "City",
    "State",
    "Zip Code",
]


@dataclass(frozen=True)
class Coordinates:
    latitude: float
    longitude: float


def load_addresses(filepath: Path) -> pd.DataFrame:
    """Load the spreadsheet and validate the address columns."""
    if not filepath.exists():
        raise FileNotFoundError(f"Could not find Excel file: {filepath}")

    df = pd.read_excel(filepath, dtype=str).fillna("")
    missing = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in Excel file: {missing}")

    return df


def build_full_address(row: pd.Series) -> str:
    """Build a geocodable address from the workbook's address columns."""
    parts = [clean_part(row["Rental Property Address"])]
    line_2 = clean_part(row["Street Address Line 2"])

    if line_2:
        parts.append(line_2)

    city = clean_part(row["City"])
    state = clean_part(row["State"])
    zip_code = clean_part(row["Zip Code"])
    city_state_zip = " ".join(part for part in [state, zip_code] if part)

    if city and city_state_zip:
        parts.append(f"{city}, {city_state_zip}")
    elif city:
        parts.append(city)
    elif city_state_zip:
        parts.append(city_state_zip)

    return ", ".join(part for part in parts if part)


def clean_part(value: Any) -> str:
    return str(value).strip()


def load_cache(cache_path: Path) -> dict[str, Any]:
    if not cache_path.exists():
        return {}

    with cache_path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def save_cache(cache_path: Path, cache: dict[str, Any]) -> None:
    with cache_path.open("w", encoding="utf-8") as handle:
        json.dump(cache, handle, indent=2, sort_keys=True)
        handle.write("\n")


def cache_key(provider: str, address: str) -> str:
    return f"{provider}:{address.strip().lower()}"


def cached_coordinates(
    cache: dict[str, Any],
    provider: str,
    address: str,
) -> Coordinates | None:
    cached = cache.get(cache_key(provider, address))
    if not cached or cached.get("missing"):
        return None

    return Coordinates(cached["latitude"], cached["longitude"])


def has_cached_answer(cache: dict[str, Any], provider: str, address: str) -> bool:
    return cache_key(provider, address) in cache


def has_any_cached_answer(
    cache: dict[str, Any],
    providers: list[str],
    address: str,
) -> bool:
    return any(has_cached_answer(cache, provider, address) for provider in providers)


def cache_coordinates(
    cache: dict[str, Any],
    provider: str,
    address: str,
    coords: Coordinates | None,
) -> None:
    key = cache_key(provider, address)
    if coords:
        cache[key] = {
            "latitude": coords.latitude,
            "longitude": coords.longitude,
        }
    else:
        cache[key] = {"missing": True}


def geocode_address(
    address: str,
    *,
    provider: str,
    cache: dict[str, Any],
    user_agent: str,
    opencage_api_key: str | None,
    timeout: int,
) -> Coordinates | None:
    """Geocode an address, using the local cache when possible."""
    if has_cached_answer(cache, provider, address):
        return cached_coordinates(cache, provider, address)

    if provider == "nominatim":
        coords = geocode_with_nominatim(address, user_agent=user_agent, timeout=timeout)
    elif provider == "opencage":
        if not opencage_api_key:
            raise EnvironmentError(
                "Set OPENCAGE_API_KEY or use --provider nominatim."
            )
        coords = geocode_with_opencage(
            address, api_key=opencage_api_key, timeout=timeout
        )
    else:
        raise ValueError(f"Unsupported geocoder provider: {provider}")

    cache_coordinates(cache, provider, address, coords)

    return coords


def geocode_property_row(
    row: pd.Series,
    *,
    provider: str,
    cache: dict[str, Any],
    user_agent: str,
    opencage_api_key: str | None,
    timeout: int,
) -> Coordinates | None:
    """Geocode a workbook row, with Census fallback for US address columns."""
    address = row["full_address"]

    for cached_provider in [provider, "census", "arcgis"]:
        if has_cached_answer(cache, cached_provider, address):
            coords = cached_coordinates(cache, cached_provider, address)
            if coords:
                return coords

    coords = geocode_address(
        address,
        provider=provider,
        cache=cache,
        user_agent=user_agent,
        opencage_api_key=opencage_api_key,
        timeout=timeout,
    )
    if coords or provider != "nominatim":
        return coords

    coords = geocode_with_census(
        street=clean_part(row["Rental Property Address"]),
        city=clean_part(row["City"]),
        state=clean_part(row["State"]),
        zip_code=clean_part(row["Zip Code"]),
        timeout=timeout,
    )
    cache_coordinates(cache, "census", address, coords)
    if coords:
        return coords

    coords = geocode_with_arcgis(address, timeout=timeout)
    cache_coordinates(cache, "arcgis", address, coords)

    return coords


def geocode_with_nominatim(
    address: str,
    *,
    user_agent: str,
    timeout: int,
) -> Coordinates | None:
    results = get_json(
        "https://nominatim.openstreetmap.org/search",
        params={
            "q": address,
            "format": "jsonv2",
            "limit": 1,
            "countrycodes": "us",
        },
        headers={"User-Agent": user_agent},
        timeout=timeout,
    )

    if not results:
        return None

    first = results[0]
    return Coordinates(float(first["lat"]), float(first["lon"]))


def geocode_with_opencage(
    address: str,
    *,
    api_key: str,
    timeout: int,
) -> Coordinates | None:
    data = get_json(
        "https://api.opencagedata.com/geocode/v1/json",
        params={
            "q": address,
            "key": api_key,
            "limit": 1,
            "countrycode": "us",
            "no_annotations": 1,
        },
        timeout=timeout,
    )
    results = data.get("results", [])

    if not results:
        return None

    geometry = results[0]["geometry"]
    return Coordinates(float(geometry["lat"]), float(geometry["lng"]))


def geocode_with_census(
    *,
    street: str,
    city: str,
    state: str,
    zip_code: str,
    timeout: int,
) -> Coordinates | None:
    if not street or not city or not state:
        return None

    data = get_json(
        "https://geocoding.geo.census.gov/geocoder/locations/address",
        params={
            "street": street,
            "city": city,
            "state": state,
            "zip": zip_code,
            "benchmark": "Public_AR_Current",
            "format": "json",
        },
        timeout=timeout,
    )
    matches = data.get("result", {}).get("addressMatches", [])

    if not matches:
        return None

    coordinates = matches[0]["coordinates"]
    return Coordinates(float(coordinates["y"]), float(coordinates["x"]))


def geocode_with_arcgis(address: str, *, timeout: int) -> Coordinates | None:
    data = get_json(
        "https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/findAddressCandidates",
        params={
            "SingleLine": address,
            "f": "json",
            "maxLocations": 1,
            "countryCode": "USA",
        },
        timeout=timeout,
    )
    candidates = data.get("candidates", [])

    if not candidates or candidates[0].get("score", 0) < 80:
        return None

    location = candidates[0]["location"]
    return Coordinates(float(location["y"]), float(location["x"]))


def get_json(
    url: str,
    *,
    params: dict[str, Any],
    timeout: int,
    headers: dict[str, str] | None = None,
) -> Any:
    query_url = f"{url}?{urlencode(params)}"
    request = Request(query_url, headers=headers or {})

    try:
        with urlopen(request, timeout=timeout) as response:
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as error:
        body = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {error.code} from geocoder: {body[:300]}") from error
    except URLError as error:
        raise RuntimeError(f"Could not reach geocoder: {error.reason}") from error


def haversine_miles(origin: Coordinates, destination: Coordinates) -> float:
    """Calculate straight-line distance in miles."""
    earth_radius_miles = 3958.7613
    lat1 = math.radians(origin.latitude)
    lon1 = math.radians(origin.longitude)
    lat2 = math.radians(destination.latitude)
    lon2 = math.radians(destination.longitude)

    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = (
        math.sin(dlat / 2) ** 2
        + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    )
    return earth_radius_miles * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def find_closest_properties(args: argparse.Namespace) -> pd.DataFrame:
    df = load_addresses(Path(args.excel_file))
    df["full_address"] = df.apply(build_full_address, axis=1)
    df = df[df["full_address"].str.strip().ne("")].copy()

    cache_path = Path(args.cache_file)
    cache = load_cache(cache_path)
    cache_changed = False

    origin = geocode_address(
        args.address,
        provider=args.provider,
        cache=cache,
        user_agent=args.user_agent,
        opencage_api_key=args.opencage_api_key,
        timeout=args.timeout,
    )
    if origin is None:
        raise RuntimeError(f"Could not geocode input address: {args.address}")
    cache_changed = True

    rows: list[dict[str, Any]] = []
    uncached_requests = 0

    for index, row in df.iterrows():
        address = row["full_address"]
        was_cached = has_any_cached_answer(
            cache,
            [args.provider, "census", "arcgis"],
            address,
        )
        coords = geocode_property_row(
            row,
            provider=args.provider,
            cache=cache,
            user_agent=args.user_agent,
            opencage_api_key=args.opencage_api_key,
            timeout=args.timeout,
        )

        if not was_cached:
            cache_changed = True
            uncached_requests += 1
            if args.provider == "nominatim" and uncached_requests:
                time.sleep(args.rate_limit_seconds)

        if coords is None:
            print(f"Skipping ungeocoded address: {address}", file=sys.stderr)
            continue

        record = row.to_dict()
        record.update(
            {
                "source_row": int(index) + 2,
                "latitude": coords.latitude,
                "longitude": coords.longitude,
                "distance_miles": haversine_miles(origin, coords),
            }
        )
        rows.append(record)

    if cache_changed:
        save_cache(cache_path, cache)

    if not rows:
        raise RuntimeError("No rental property addresses could be geocoded.")

    ranked = pd.DataFrame(rows).sort_values("distance_miles", kind="mergesort")
    if args.unique_addresses:
        ranked = ranked.drop_duplicates(
            subset=["full_address"],
            keep="first",
        )
    return ranked.head(args.top).reset_index(drop=True)


def print_results(results: pd.DataFrame, input_address: str) -> None:
    print(f"\nTop {len(results)} closest properties to:")
    print(f"  {input_address}")
    print("=" * 72)

    for rank, row in results.iterrows():
        print(f"\n#{rank + 1}  {row['full_address']}")
        print(f"    Distance: {row['distance_miles']:.2f} miles straight-line")
        print(f"    City: {row['City']}, {row['State']} {row['Zip Code']}")
        print(f"    Spreadsheet row: {row['source_row']}")

        listing_name = clean_part(row.get("Rental Property Listing Name", ""))
        if listing_name:
            print(f"    Listing: {listing_name}")

    print("\n" + "=" * 72)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Find the nearest rental properties in an Excel database."
    )
    parser.add_argument(
        "--address",
        default=None,
        help=(
            "Input address to compare against. If omitted, you will be prompted; "
            f"blank input uses {DEFAULT_ADDRESS}."
        ),
    )
    parser.add_argument("--excel-file", default=DEFAULT_EXCEL_FILE)
    parser.add_argument("--top", type=int, default=DEFAULT_TOP_N)
    parser.add_argument(
        "--provider",
        choices=["nominatim", "opencage"],
        default="nominatim",
        help="Geocoding service to use. Default: nominatim.",
    )
    parser.add_argument("--cache-file", default=DEFAULT_CACHE_FILE)
    parser.add_argument(
        "--unique-addresses",
        action="store_true",
        help="Return only one result for each distinct full address.",
    )
    parser.add_argument(
        "--user-agent",
        default="housing-distance-ranker/1.0 (local script)",
        help="Required by Nominatim usage policy.",
    )
    parser.add_argument(
        "--rate-limit-seconds",
        type=float,
        default=1.0,
        help="Delay after uncached Nominatim requests.",
    )
    parser.add_argument("--timeout", type=int, default=20)
    parser.add_argument(
        "--opencage-api-key",
        default=os.getenv("OPENCAGE_API_KEY"),
        help="Optional. Defaults to OPENCAGE_API_KEY environment variable.",
    )
    return parser.parse_args()


def get_input_address(address_arg: str | None) -> str:
    if address_arg and address_arg.strip():
        return address_arg.strip()

    try:
        entered = input(f"Enter input address (default: {DEFAULT_ADDRESS}): ").strip()
    except EOFError:
        entered = ""

    return entered or DEFAULT_ADDRESS


def main() -> None:
    args = parse_args()
    args.address = get_input_address(args.address)
    results = find_closest_properties(args)
    print_results(results, args.address)


if __name__ == "__main__":
    main()
