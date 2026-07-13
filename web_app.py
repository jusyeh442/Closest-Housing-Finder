from __future__ import annotations

import argparse
import html
import json
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from types import SimpleNamespace
from urllib.parse import parse_qs, urlparse

from find_closest_properties import (
    Coordinates,
    DEFAULT_ADDRESS,
    DEFAULT_CACHE_FILE,
    DEFAULT_EXCEL_FILE,
    DEFAULT_TOP_N,
    find_closest_properties,
    geocode_address,
    get_json,
)


DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8000
SUGGESTION_CACHE: dict[str, list[str]] = {}


PAGE = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Closest Housing Finder</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #17212b;
      --muted: #5c6875;
      --line: #d9e1e8;
      --panel: #ffffff;
      --field: #f7f9fb;
      --accent: #176b5d;
      --accent-dark: #105247;
      --warn: #9a4d12;
      --bg: #eef3f6;
    }

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      min-height: 100vh;
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      color: var(--ink);
      background: linear-gradient(180deg, #f8fbfc 0%, var(--bg) 100%);
    }

    main {
      width: min(980px, calc(100% - 32px));
      margin: 0 auto;
      padding: 48px 0;
    }

    header {
      margin-bottom: 24px;
    }

    h1 {
      margin: 0 0 8px;
      font-size: clamp(2rem, 5vw, 4rem);
      line-height: 1;
      letter-spacing: 0;
    }

    .subtitle {
      margin: 0;
      max-width: 680px;
      color: var(--muted);
      font-size: 1rem;
      line-height: 1.5;
    }

    .search-panel {
      padding: 18px;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel);
      box-shadow: 0 12px 32px rgba(23, 33, 43, 0.08);
    }

    form {
      display: grid;
      gap: 14px;
    }

    label {
      display: block;
      margin-bottom: 7px;
      color: #32404d;
      font-size: 0.88rem;
      font-weight: 650;
    }

    input[type="text"],
    input[type="number"] {
      width: 100%;
      min-height: 46px;
      padding: 10px 12px;
      border: 1px solid var(--line);
      border-radius: 6px;
      background: var(--field);
      color: var(--ink);
      font: inherit;
    }

    input:focus {
      outline: 3px solid rgba(23, 107, 93, 0.18);
      border-color: var(--accent);
      background: #ffffff;
    }

    .address-field {
      position: relative;
    }

    .suggestions {
      position: absolute;
      z-index: 10;
      top: calc(100% + 6px);
      right: 0;
      left: 0;
      overflow: hidden;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #ffffff;
      box-shadow: 0 18px 40px rgba(23, 33, 43, 0.14);
    }

    .suggestions[hidden] {
      display: none;
    }

    .suggestion {
      width: 100%;
      min-height: 42px;
      padding: 10px 12px;
      border: 0;
      border-bottom: 1px solid #edf1f4;
      border-radius: 0;
      background: #ffffff;
      color: var(--ink);
      text-align: left;
      font-weight: 560;
    }

    .suggestion:last-child {
      border-bottom: 0;
    }

    .suggestion:hover,
    .suggestion:focus {
      background: #eef7f4;
      color: var(--accent-dark);
      outline: none;
    }

    .controls {
      display: grid;
      grid-template-columns: minmax(96px, 128px) minmax(180px, 1fr) auto;
      gap: 12px;
      align-items: end;
    }

    .advanced {
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fafcfd;
    }

    .advanced summary {
      cursor: pointer;
      list-style: none;
      padding: 12px 14px;
      font-size: 0.95rem;
      font-weight: 700;
      color: #2a3a48;
      user-select: none;
    }

    .advanced summary::-webkit-details-marker {
      display: none;
    }

    .advanced summary::after {
      content: "Open";
      float: right;
      color: var(--muted);
      font-weight: 600;
      font-size: 0.84rem;
    }

    .advanced[open] summary::after {
      content: "Close";
    }

    .advanced-fields {
      padding: 0 14px 14px;
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 12px;
    }

    .hint {
      margin: 0;
      grid-column: 1 / -1;
      color: var(--muted);
      font-size: 0.84rem;
      line-height: 1.4;
    }

    .check {
      min-height: 46px;
      display: flex;
      align-items: center;
      gap: 9px;
      color: #32404d;
      font-size: 0.92rem;
    }

    .check input {
      width: 18px;
      height: 18px;
      accent-color: var(--accent);
    }

    button {
      min-height: 46px;
      border: 0;
      border-radius: 6px;
      padding: 0 18px;
      background: var(--accent);
      color: #ffffff;
      font: inherit;
      font-weight: 750;
      cursor: pointer;
    }

    button:hover {
      background: var(--accent-dark);
    }

    button:disabled {
      cursor: wait;
      opacity: 0.72;
    }

    .status {
      min-height: 24px;
      color: var(--muted);
      font-size: 0.93rem;
    }

    .status.error {
      color: #9f2424;
    }

    .results {
      margin-top: 22px;
      display: grid;
      gap: 12px;
    }

    .result {
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel);
      padding: 16px;
      display: grid;
      grid-template-columns: auto 1fr auto;
      gap: 14px;
      align-items: start;
    }

    .rank {
      width: 34px;
      height: 34px;
      border-radius: 999px;
      display: grid;
      place-items: center;
      background: #e2f0ec;
      color: var(--accent-dark);
      font-weight: 800;
    }

    .address {
      margin: 0 0 6px;
      font-size: 1.02rem;
      font-weight: 760;
      line-height: 1.3;
    }

    .meta {
      margin: 0;
      color: var(--muted);
      font-size: 0.9rem;
      line-height: 1.45;
    }

    .listing {
      margin-top: 8px;
      color: #32404d;
      font-size: 0.92rem;
      line-height: 1.45;
    }

    .distance {
      white-space: nowrap;
      text-align: right;
      font-size: 1.15rem;
      font-weight: 800;
      color: var(--accent-dark);
    }

    .distance span {
      display: block;
      margin-top: 2px;
      color: var(--muted);
      font-size: 0.76rem;
      font-weight: 650;
    }

    @media (max-width: 700px) {
      main {
        width: min(100% - 24px, 980px);
        padding: 28px 0;
      }

      .controls {
        grid-template-columns: 1fr;
      }

      .advanced-fields {
        grid-template-columns: 1fr;
      }

      button {
        width: 100%;
      }

      .result {
        grid-template-columns: auto 1fr;
      }

      .distance {
        grid-column: 2;
        text-align: left;
      }
    }
  </style>
</head>
<body>
  <main>
    <header>
      <h1>Closest Housing Finder</h1>
      <p class="subtitle">Search the housing workbook by distance from one address.</p>
    </header>

    <section class="search-panel" aria-label="Address search">
      <form id="search-form">
        <div class="address-field">
          <label for="address">Input address</label>
          <input id="address" name="address" type="text" autocomplete="street-address" value="{{ default_address }}">
          <div id="suggestions" class="suggestions" role="listbox" aria-label="Address suggestions" hidden></div>
        </div>
        <div class="controls">
          <div>
            <label for="top">Results</label>
            <input id="top" name="top" type="number" min="1" max="25" value="{{ default_top }}">
          </div>
          <label class="check" for="unique">
            <input id="unique" name="unique" type="checkbox">
            Unique addresses only
          </label>
          <button id="submit" type="submit">Find closest</button>
        </div>
        <details class="advanced">
          <summary>Advanced Search</summary>
          <div class="advanced-fields">
            <div>
              <label for="bedrooms">Number of Bedrooms</label>
              <input id="bedrooms" name="bedrooms" type="number" min="0" step="any" placeholder="e.g. 2">
            </div>
            <div>
              <label for="bathrooms">Number of Bathrooms</label>
              <input id="bathrooms" name="bathrooms" type="number" min="0" step="0.5" placeholder="e.g. 1.5">
            </div>
            <div>
              <label for="max_monthly_rate">Max Monthly Rate</label>
              <input id="max_monthly_rate" name="max_monthly_rate" type="number" min="0" step="1" placeholder="e.g. 2500">
            </div>
            <p class="hint">Leave any field blank to ignore that filter.</p>
          </div>
        </details>
        <div id="status" class="status" role="status"></div>
      </form>
    </section>

    <section id="results" class="results" aria-live="polite"></section>
  </main>

  <script>
    const form = document.querySelector("#search-form");
    const statusEl = document.querySelector("#status");
    const resultsEl = document.querySelector("#results");
    const submitBtn = document.querySelector("#submit");
    const addressInput = document.querySelector("#address");
    const suggestionsEl = document.querySelector("#suggestions");
    let suggestTimer;
    let activeSuggestController;

    function escapeHtml(value) {
      return String(value ?? "").replace(/[&<>"']/g, (char) => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;"
      }[char]));
    }

    function renderResults(items) {
      resultsEl.innerHTML = items.map((item, index) => `
        <article class="result">
          <div class="rank">${index + 1}</div>
          <div>
            <p class="address">${escapeHtml(item.full_address)}</p>
            <p class="meta">${escapeHtml(item.city)}, ${escapeHtml(item.state)} ${escapeHtml(item.zip_code)} · Spreadsheet row ${escapeHtml(item.source_row)}</p>
            ${item.listing_name ? `<div class="listing">${escapeHtml(item.listing_name)}</div>` : ""}
          </div>
          <div class="distance">
            ${Number(item.distance_miles).toFixed(2)}<span>miles straight-line</span>
            ${item.quickest_route_miles == null ? "" : `${Number(item.quickest_route_miles).toFixed(2)}<span>miles quickest route</span>`}
          </div>
        </article>
      `).join("");
    }

    function hideSuggestions() {
      suggestionsEl.hidden = true;
      suggestionsEl.innerHTML = "";
    }

    function renderSuggestions(items) {
      if (!items.length) {
        hideSuggestions();
        return;
      }

      suggestionsEl.innerHTML = items.map((item) => `
        <button class="suggestion" type="button" role="option" data-address="${escapeHtml(item)}">${escapeHtml(item)}</button>
      `).join("");
      suggestionsEl.hidden = false;
    }

    suggestionsEl.addEventListener("click", (event) => {
      const option = event.target.closest(".suggestion");
      if (!option) {
        return;
      }
      addressInput.value = option.dataset.address;
      hideSuggestions();
      addressInput.focus();
    });

    addressInput.addEventListener("input", () => {
      const query = addressInput.value.trim();
      clearTimeout(suggestTimer);

      if (activeSuggestController) {
        activeSuggestController.abort();
      }

      if (query.length < 2) {
        hideSuggestions();
        return;
      }

      suggestTimer = setTimeout(async () => {
        activeSuggestController = new AbortController();
        try {
          const response = await fetch(`/api/suggest?q=${encodeURIComponent(query)}`, {
            signal: activeSuggestController.signal
          });
          const data = await response.json();
          if (response.ok && addressInput.value.trim() === query) {
            renderSuggestions(data.suggestions || []);
          }
        } catch (error) {
          if (error.name !== "AbortError") {
            hideSuggestions();
          }
        }
      }, 350);
    });

    addressInput.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        hideSuggestions();
      }
    });

    document.addEventListener("click", (event) => {
      if (!event.target.closest(".address-field")) {
        hideSuggestions();
      }
    });

    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      const formData = new FormData(form);
      const params = new URLSearchParams({
        address: formData.get("address") || "",
        top: formData.get("top") || "{{ default_top }}",
        unique: formData.get("unique") === "on" ? "1" : "0",
        bedrooms: (formData.get("bedrooms") || "").toString().trim(),
        bathrooms: (formData.get("bathrooms") || "").toString().trim(),
        max_monthly_rate: (formData.get("max_monthly_rate") || "").toString().trim()
      });

      submitBtn.disabled = true;
      hideSuggestions();
      statusEl.className = "status";
      statusEl.textContent = "Finding closest properties...";
      resultsEl.innerHTML = "";

      try {
        const response = await fetch(`/api/search?${params}`);
        const data = await response.json();
        if (!response.ok) {
          throw new Error(data.error || "Search failed.");
        }
        statusEl.textContent = `Showing ${data.results.length} closest result${data.results.length === 1 ? "" : "s"} for ${data.address}.`;
        renderResults(data.results);
      } catch (error) {
        statusEl.className = "status error";
        statusEl.textContent = error.message;
      } finally {
        submitBtn.disabled = false;
      }
    });
  </script>
</body>
</html>
"""


def make_search_args(
  address: str,
  top: int,
  unique: bool,
  bedrooms: str | None,
  bathrooms: str | None,
  max_monthly_rate: str | None,
) -> SimpleNamespace:
    return SimpleNamespace(
        address=address or DEFAULT_ADDRESS,
        excel_file=DEFAULT_EXCEL_FILE,
        top=top,
    bedrooms=bedrooms,
    bathrooms=bathrooms,
    max_monthly_rate=max_monthly_rate,
        provider="nominatim",
        cache_file=DEFAULT_CACHE_FILE,
        unique_addresses=unique,
        user_agent="housing-distance-ranker/1.0 (local web app)",
        rate_limit_seconds=1.0,
        timeout=20,
        opencage_api_key=None,
    )


def results_to_json(results) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for _, row in results.iterrows():
        rows.append(
            {
                "full_address": str(row["full_address"]),
                "distance_miles": float(row["distance_miles"]),
                "latitude": float(row["latitude"]),
                "longitude": float(row["longitude"]),
                "city": str(row["City"]),
                "state": str(row["State"]),
                "zip_code": str(row["Zip Code"]),
                "source_row": int(row["source_row"]),
                "listing_name": str(row.get("Rental Property Listing Name", "")),
            }
        )
    return rows


def quickest_route_miles(
    origin: Coordinates,
    destination: Coordinates,
) -> float | None:
    """Return quickest driving-route distance in miles using OSRM."""
    data = get_json(
        "http://router.project-osrm.org/route/v1/driving/"
        f"{origin.longitude},{origin.latitude};{destination.longitude},{destination.latitude}",
        params={
            "overview": "false",
            "alternatives": "false",
            "steps": "false",
        },
        timeout=10,
    )

    routes = data.get("routes", [])
    if not routes:
        return None

    distance_meters = routes[0].get("distance")
    if distance_meters is None:
        return None

    return float(distance_meters) * 0.000621371


def suggest_addresses(query: str) -> list[str]:
    normalized_query = " ".join(query.lower().split())
    if len(normalized_query) < 2:
        return []

    cached = SUGGESTION_CACHE.get(normalized_query)
    if cached is not None:
        return cached

    suggestions = suggest_with_arcgis(query)
    if suggestions:
        SUGGESTION_CACHE[normalized_query] = suggestions
        return suggestions

    suggestions = suggest_with_nominatim(query)
    SUGGESTION_CACHE[normalized_query] = suggestions
    return suggestions


def suggest_with_arcgis(query: str) -> list[str]:
    data = get_json(
        "https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/suggest",
        params={
            "text": query,
            "f": "json",
            "maxSuggestions": 5,
            "countryCode": "USA",
        },
        timeout=8,
    )

    suggestions: list[str] = []
    seen: set[str] = set()
    for result in data.get("suggestions", []):
        suggestion = str(result.get("text", "")).strip()
        key = suggestion.lower()
        if suggestion and key not in seen:
            seen.add(key)
            suggestions.append(suggestion)

    return suggestions


def suggest_with_nominatim(query: str) -> list[str]:
    results = get_json(
        "https://nominatim.openstreetmap.org/search",
        params={
            "q": query,
            "format": "jsonv2",
            "addressdetails": 1,
            "limit": 5,
            "countrycodes": "us",
        },
        headers={"User-Agent": "housing-distance-ranker/1.0 (local web app)"},
        timeout=8,
    )

    suggestions: list[str] = []
    seen: set[str] = set()
    for result in results:
        display_name = str(result.get("display_name", "")).strip()
        if not display_name:
            continue

        suggestion = shorten_display_name(display_name)
        key = suggestion.lower()
        if key not in seen:
            seen.add(key)
            suggestions.append(suggestion)

    return suggestions


def shorten_display_name(display_name: str) -> str:
    parts = [part.strip() for part in display_name.split(",") if part.strip()]
    if len(parts) <= 6:
        return ", ".join(parts)

    return ", ".join(parts[:6])


class HousingRequestHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_html(PAGE.replace("{{ default_address }}", html.escape(DEFAULT_ADDRESS)).replace("{{ default_top }}", str(DEFAULT_TOP_N)))
            return

        if parsed.path == "/api/search":
            self.handle_search(parsed.query)
            return

        if parsed.path == "/api/suggest":
            self.handle_suggest(parsed.query)
            return

        self.send_error(HTTPStatus.NOT_FOUND)

    def handle_search(self, query: str) -> None:
        params = parse_qs(query)
        address = params.get("address", [DEFAULT_ADDRESS])[0].strip() or DEFAULT_ADDRESS
        unique = params.get("unique", ["0"])[0] == "1"
        bedrooms = params.get("bedrooms", [""])[0].strip() or None
        bathrooms = params.get("bathrooms", [""])[0].strip() or None
        max_monthly_rate = params.get("max_monthly_rate", [""])[0].strip() or None

        try:
            top = int(params.get("top", [str(DEFAULT_TOP_N)])[0])
        except ValueError:
            top = DEFAULT_TOP_N
        top = max(1, min(top, 25))

        try:
            search_args = make_search_args(
                address,
                top,
                unique,
                bedrooms,
                bathrooms,
                max_monthly_rate,
            )
            results = find_closest_properties(search_args)
            origin = geocode_address(
                address,
                provider=search_args.provider,
                cache={},
                user_agent=search_args.user_agent,
                opencage_api_key=search_args.opencage_api_key,
                timeout=search_args.timeout,
            )
            payload = {
                "address": address,
                "results": results_to_json(results),
            }

            if origin is not None:
                for result in payload["results"]:
                    destination = Coordinates(
                        latitude=float(result["latitude"]),
                        longitude=float(result["longitude"]),
                    )
                    try:
                        result["quickest_route_miles"] = quickest_route_miles(
                            origin,
                            destination,
                        )
                    except Exception:
                        result["quickest_route_miles"] = None
            else:
                for result in payload["results"]:
                    result["quickest_route_miles"] = None

            self.send_json(payload)
        except Exception as error:
            self.send_json({"error": str(error)}, status=HTTPStatus.INTERNAL_SERVER_ERROR)

    def handle_suggest(self, query: str) -> None:
        params = parse_qs(query)
        search_text = params.get("q", [""])[0].strip()

        try:
            self.send_json({"suggestions": suggest_addresses(search_text)})
        except Exception:
            self.send_json({"suggestions": []})

    def send_html(self, body: str) -> None:
        encoded = body.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def send_json(self, payload: dict[str, object], status: HTTPStatus = HTTPStatus.OK) -> None:
        encoded = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def log_message(self, format: str, *args: object) -> None:
        return


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the closest housing web app.")
    parser.add_argument("--host", default=DEFAULT_HOST)
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    server = ThreadingHTTPServer((args.host, args.port), HousingRequestHandler)
    print(f"Serving closest housing finder at http://{args.host}:{args.port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
