# CSO Insights — General Insights Page Design

**Date:** 2026-05-09
**Status:** Approved
**Owner:** PRISM

---

## Overview

A single-page HTML dashboard (`index.html`) for the CSO Insights GitHub repo (`FlomaticAuto/olympic-paints-cso-insights`). Provides an executive cross-functional view via a card grid that drills into full insight detail pages — all within one file, no page reloads.

---

## Architecture

**Approach:** Single HTML file, JS-driven SPA. Two views (grid / detail) swap via `display` toggling. All data injected at build time by `build_cso_insights.py` as a JS `const INSIGHTS = [...]` block. Browser back button supported via `popstate`.

**Build script:** `1.Projects/AWS Data/build_cso_insights.py`
**Output:** `1.Projects/AWS Data/cso_insights/index.html` + `logo.jpg`
**Repo:** `FlomaticAuto/olympic-paints-cso-insights` (public, GitHub Pages not yet enabled)
**Push pattern:** `gh auth token --user FlomaticAuto` — same as all other Olympic Paints dashboards

---

## Design System

Fully compliant with the Olympic Paints HTML design standard (DESIGN_SYSTEM.md):
- Four themes: Light / Dark / Brand / Navy — default `theme-dark`
- Theme class on `<html>`, persisted via `localStorage` key `oly-theme`
- Fonts: Barlow Condensed (display/headings) + Barlow (body/UI) — Google Fonts
- All colours via `--color-*` CSS custom property tokens — no hardcoded hex in components
- Logo: `Olympic Paints Logo Digital.jpg` in `border-radius:50%;overflow:hidden` wrapper
- No CSS frameworks, no external JS except Chart.js from cdnjs
- `barLabels` Chart.js plugin registered for all bar charts

---

## Page Structure

### Header (fixed, both views)
- Olympic Paints logo (48px circular)
- Title: "CSO INSIGHTS" (Barlow Condensed 900, uppercase)
- Today's date (subtitle)
- 4-button theme toggle (Light / Dark / Brand / Navy)

### Grid View (default)
- Eyebrow label: "INSIGHTS"
- 3-column responsive card grid (→ 2-col tablet → 1-col mobile)
- Each card contains:
  - Icon (insight-specific SVG inline)
  - Title (Barlow Condensed 700)
  - One-line summary (Barlow 400, 14px)
  - Last updated date badge (`--color-neutral-*` tokens)
  - Brand-yellow left border (`--color-border-brand`)
  - Hover lift: `translateY(-3px)` + `box-shadow` increase

### Detail View (on card click)
- `← Back to Insights` button (top-left, below header)
- Insight title: H1 Barlow Condensed 900 uppercase
- Three content blocks in order:
  1. **Written analysis** — plain-English paragraph (`--color-text-secondary`)
  2. **Chart** — Chart.js, insight-specific type, Olympic Paints colour sequence
  3. **Data table** — vanilla JS filterable, zebra-striped (`--color-surface-sunken`), sticky header

---

## Data Model

Each insight is one object in the `INSIGHTS` array:

```js
{
  id: String,          // kebab-case, used as URL hash
  title: String,
  summary: String,     // one-line card description
  updated: String,     // ISO date, e.g. "2026-05-09"
  icon: String,        // inline SVG string
  analysis: String,    // plain-English HTML paragraph
  chart: {
    type: String,      // "bar" | "line" (horizontal bars use indexAxis:"y" option)
    labels: Array,
    datasets: Array    // Chart.js dataset objects
  },
  table: {
    columns: Array,    // { key, label } objects
    rows: Array        // plain objects matching column keys
  }
}
```

---

## The Three Insights

### 1. Store Buying Frequency
- **Source:** `Sales_Invoices_All.parquet` — computed by Python at build time
- **Chart:** Horizontal bar (`indexAxis: 'y'`) — stores on Y-axis, order count on X-axis (top 20 stores by frequency)
- **Table columns:** Store Name · Account Ref · Orders (last 12m) · Last Order Date · Avg Days Between Orders
- **Colour:** Yellow (`#F5C400`) bars

### 2. Rep Performance
- **Source:** Manual — `REPS` constant imported from `build_kpi_dashboard.py` (single source of truth)
- **Chart:** Vertical grouped bar — actual (yellow) vs target (navy) per rep
- **Table columns:** Rep · MTD Sales · Target · % vs Target · Status badge
- **Colour:** Yellow actual, Navy target

### 3. Product Mix
- **Source:** `Sales_Invoices_All.parquet` — `ivnett` by product group, computed by Python at build time
- **Chart:** Stacked bar — product groups as segments per rep
- **Table columns:** Product Group · Revenue (NET) · % of Mix · YoY Change
- **Colour:** Multi-series order from design system

---

## Build Script Behaviour

1. Read `Sales_Invoices_All.parquet` → compute Store Buying Frequency and Product Mix tables
2. Import `REPS` from `build_kpi_dashboard.py` for Rep Performance
3. Copy `LOGO_SRC` (`3.Resources/9. Brand Assets & Images/Misc Pictures/Olympic Paints Logo Digital.jpg`) to `cso_insights/logo.jpg`
4. Inject all computed data into the HTML template as `const INSIGHTS = [...]`
5. Write `cso_insights/index.html`
6. Push `cso_insights/` contents to `FlomaticAuto/olympic-paints-cso-insights` via embedded token URL

**Run command:** `python build_cso_insights.py`

---

## Navigation

- Clicking a card: sets `window.location.hash` to `#<insight-id>`, shows detail view
- Back button click: clears hash, shows grid view
- Browser back: `popstate` listener restores grid view
- Direct URL with hash (e.g. `index.html#store-buying-frequency`): opens detail view directly on load

---

## Out of Scope (this iteration)

- Category/tag filtering on the grid
- Search across insights
- Print/export per insight
- GitHub Pages enablement (manual step, done separately)
