# Executive Dashboard - CLAUDE.md

For PCMA context, brand guidelines, workstreams, and preferences, see `../CLAUDE.md` in the parent PCMA Dashboards folder.

## What This Is

A single-page executive dashboard for PCMA's 2026 Enterprise Goals. It is a self-contained `index.html` (HTML + CSS + JS, no build step) that displays KPI progress across five weighted sections with live data polling. Hosted on GitHub Pages at stevenrieger.github.io/PCMA-Executive-Dashboard.

## Running

Open `index.html` directly in a browser, or serve it with any static server (e.g. `npx serve .`). Live data polling requires the file to be served over HTTP so `fetch()` works.

## Architecture

**Single file:** All markup, styles, and logic live in `index.html` (~809 lines). There is no build system, bundler, or framework.

**Data flow:**
1. Hardcoded fallback data lives in the `DATA` constant (line ~252) with five sections: `financial`, `membership`, `endowment`, `education`, `internal`.
2. On load, `showLoadingState()` runs, then `tryLiveData()` fetches `dashboard_data.json` (generated externally by Power Automate from a SharePoint Excel list).
3. `parseExcelRows()` (line ~639) maps SharePoint row names to section/metric indices via `NAME_MAP` - this is the critical data-binding layer. When SharePoint column names change, `NAME_MAP` fragments must be updated.
4. `buildDashboard()` and `buildComposite()` re-render the entire UI from the `DATA` object.
5. Polls every 30 seconds via `setInterval(tryLiveData, 30000)`.

**Key functions:**
- `parseExcelRows()` - Converts raw SharePoint rows into section updates. Uses name-fragment matching, not positional indexing.
- `buildDashboard()` - Generates all section blocks, cubes, detail panels, and chart canvases.
- `buildDetailCharts()` - Creates Chart.js bar/progress charts when a section is expanded.
- `secAvg()` / `compositeScore()` - Calculate weighted progress scores displayed in the header and thermometers.
- `pColor()` / `chipClass()` - Determine RAG (red/amber/green) status colors based on progress percentage.

**External dependency:** Chart.js loaded via CDN (`<script>` tag, line ~242).

**Metric properties:** Each metric object can have: `goal`, `ytd`, `gl` (goal label), `yl` (YTD label), `fmt` (currency/number/percent), `src` (data source name), `inverse` (lower-is-better, e.g. turnover), `threshold` (target-based), `exceeds` (exceeded goal flag), `status` (text status like "Quiet Phase"), `checklist` (array of {label, done} for the Data & Technology initiatives), `budget`, `variance`.

## Key Conventions

- Section weights (in the `DATA` config) must sum to 100 - they drive the composite score calculation.
- The `NAME_MAP` in `parseExcelRows()` uses substring matching (`includes()`) against the SharePoint "Financial Results" column. Order matters: first match wins.
- Financial metrics use a special pattern: the goal row maps to the annual goal, but YTD/budget/variance come from separate named rows (e.g. "PCMA: Consolidated February 2026 YTD Net Operating Income Actual").
- Internal metrics with decimal YTD values < 1 are auto-converted to percentages (multiplied by 100).
- The `dashboard_data.json` file is machine-generated - do not hand-edit it. To change displayed data, modify the `DATA` constant in `index.html` or update the upstream SharePoint list.
- CORS is a hard architectural constraint for dashboards hosted on external domains - JSON must be served from the same origin (GitHub Pages) rather than SharePoint to avoid cross-origin fetch failures.
- Excel dashboard data should be parsed by name, not row position, so adding rows does not break the pipeline.
