"""
Step 2 of the build pipeline: JSON -> HTML dashboard.

Loads data/sales.json plus the two bundled libraries (build/chart.umd.js,
build/xlsx.core.min.js), injects all three into an HTML/CSS/JS template, and
writes the result to dashboard/Sales Dashboard.html.

There's no server and no build tooling beyond plain Python — the "template"
below is one large string with placeholders (__SALES_JSON__, __CHARTJS_SRC__,
__XLSXLIB_SRC__) substituted at the end of the script. The resulting HTML
file is fully self-contained: no internet connection needed to open it.

Category and Sale Location are NOT hardcoded here (see documentation/
Documentation.md §5) — they're derived entirely in the browser from
whatever's actually in the data, ranked by revenue and coloured by rank
against a fixed 8-slot palette. This file has no opinion on your taxonomy.

Run from inside build/:
    python3 build_dashboard.py
"""

import json
import os

SALES_JSON_PATH = "../data/sales.json"
CHARTJS_PATH = "chart.umd.js"
XLSXLIB_PATH = "xlsx.core.min.js"
OUTPUT_PATH = "../dashboard/Sales Dashboard.html"
# Also written to docs/index.html so GitHub Pages (serving main branch,
# /docs folder) can host this at https://<user>.github.io/<repo>/ — see
# documentation/Documentation.md's note on GitHub Pages for the important
# caveat about that URL being effectively public on a GitHub Free plan.
DOCS_OUTPUT_PATH = "../docs/index.html"

TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Sales Dashboard</title>
<script>__CHARTJS_SRC__</script>
<script>__XLSXLIB_SRC__</script>
<style>
  :root {
    color-scheme: light;
    --surface-1:      #fcfcfb;
    --page-plane:     #f9f9f7;
    --text-primary:   #0b0b0b;
    --text-secondary: #52514e;
    --muted:          #898781;
    --grid:           #e1e0d9;
    --baseline:       #c3c2b7;
    --border:         rgba(11,11,11,0.10);
    --chip-bg:        #f0efec;
    --chip-bg-active: #0b0b0b;
    --chip-fg-active: #fcfcfb;

    --series-primary: #2a78d6; /* blue - generic single-metric colour (trend line, top-products bar) */
    --series-profit:  #1baf7a; /* aqua - profit line */

    /* Fixed 8-slot categorical palette (references/palette.md, validated
       order) used POSITIONALLY: Category and Sale Location are ranked by
       revenue at load time and assigned slots 1..8 in that order, not by
       name. A 9th+ distinct value falls back to --muted ("Other"). */
    --series-1: #2a78d6; /* blue */
    --series-2: #eb6834; /* orange */
    --series-3: #1baf7a; /* aqua */
    --series-4: #eda100; /* yellow */
    --series-5: #e87ba4; /* magenta */
    --series-6: #008300; /* green */
    --series-7: #4a3aa7; /* violet */
    --series-8: #e34948; /* red */

    --good: #006300;
    --critical: #d03b3b;
  }
  @media (prefers-color-scheme: dark) {
    :root:where(:not([data-theme="light"])) {
      color-scheme: dark;
      --surface-1:      #1a1a19;
      --page-plane:     #0d0d0d;
      --text-primary:   #ffffff;
      --text-secondary: #c3c2b7;
      --muted:          #898781;
      --grid:           #2c2c2a;
      --baseline:       #383835;
      --border:         rgba(255,255,255,0.10);
      --chip-bg:        #262624;
      --chip-bg-active: #ffffff;
      --chip-fg-active: #1a1a19;

      --series-primary: #3987e5;
      --series-profit:  #199e70;

      --series-1: #3987e5;
      --series-2: #d95926;
      --series-3: #199e70;
      --series-4: #c98500;
      --series-5: #d55181;
      --series-6: #008300;
      --series-7: #9085e9;
      --series-8: #e66767;

      --good: #0ca30c;
    }
  }
  :root[data-theme="dark"] {
    color-scheme: dark;
    --surface-1:      #1a1a19;
    --page-plane:     #0d0d0d;
    --text-primary:   #ffffff;
    --text-secondary: #c3c2b7;
    --muted:          #898781;
    --grid:           #2c2c2a;
    --baseline:       #383835;
    --border:         rgba(255,255,255,0.10);
    --chip-bg:        #262624;
    --chip-bg-active: #ffffff;
    --chip-fg-active: #1a1a19;

    --series-primary: #3987e5;
    --series-profit:  #199e70;

    --series-1: #3987e5;
    --series-2: #d95926;
    --series-3: #199e70;
    --series-4: #c98500;
    --series-5: #d55181;
    --series-6: #008300;
    --series-7: #9085e9;
    --series-8: #e66767;

    --good: #0ca30c;
  }

  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
    background: var(--page-plane);
    color: var(--text-primary);
  }
  header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 16px;
    padding: 18px 28px;
    background: var(--surface-1);
    border-bottom: 1px solid var(--border);
    flex-wrap: wrap;
  }
  header h1 { font-size: 20px; margin: 0; }
  header .meta { color: var(--text-secondary); font-size: 13px; margin-top: 2px; }
  header .actions { display: flex; align-items: center; gap: 10px; }

  .tabs {
    display: flex;
    gap: 4px;
    padding: 0 28px;
    background: var(--surface-1);
    border-bottom: 1px solid var(--border);
  }
  .tab-btn {
    background: none;
    border: none;
    border-bottom: 2px solid transparent;
    border-radius: 0;
    padding: 12px 6px;
    margin: 0 14px 0 0;
    font-size: 13px;
    color: var(--text-secondary);
    cursor: pointer;
  }
  .tab-btn:hover { background: none; color: var(--text-primary); }
  .tab-btn.active { color: var(--text-primary); border-bottom-color: var(--series-primary); font-weight: 600; }
  .tab-panel { display: none; }
  .tab-panel.active { display: block; }
  button {
    font-family: inherit;
    font-size: 13px;
    border: 1px solid var(--border);
    background: var(--surface-1);
    color: var(--text-primary);
    padding: 8px 14px;
    border-radius: 8px;
    cursor: pointer;
  }
  button:hover { background: var(--chip-bg); }
  button.primary { background: var(--series-primary); color: #fff; border-color: transparent; }

  main { padding: 24px 28px 60px; max-width: 1400px; margin: 0 auto; }

  section { margin-bottom: 32px; }
  section > h2 {
    font-size: 14px;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    color: var(--text-secondary);
    margin: 0 0 12px;
  }

  .kpi-row {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 12px;
  }
  .kpi-tile {
    background: var(--surface-1);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 14px 16px;
  }
  .kpi-tile .label { font-size: 12px; color: var(--text-secondary); }
  .kpi-tile .value {
    font-size: 24px;
    font-weight: 600;
    margin-top: 4px;
    font-variant-numeric: proportional-nums;
  }
  .kpi-tile .sub { font-size: 12px; color: var(--muted); margin-top: 2px; }

  .filters {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 18px;
    background: var(--surface-1);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 12px 16px;
  }
  .filter-group { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
  .filter-group .fg-label { font-size: 12px; color: var(--text-secondary); margin-right: 2px; }
  .chip {
    font-size: 12px;
    padding: 5px 11px;
    border-radius: 999px;
    background: var(--chip-bg);
    color: var(--text-primary);
    cursor: pointer;
    user-select: none;
    border: 1px solid transparent;
  }
  .chip.active { background: var(--chip-bg-active); color: var(--chip-fg-active); }
  .chip .dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%; margin-right: 6px; }
  select, input[type="search"] {
    font-family: inherit;
    font-size: 13px;
    padding: 6px 10px;
    border-radius: 8px;
    border: 1px solid var(--border);
    background: var(--surface-1);
    color: var(--text-primary);
  }
  input[type="search"] { min-width: 220px; }

  .chart-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
  }
  .chart-card {
    background: var(--surface-1);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 16px;
  }
  .chart-card h3 { font-size: 13px; margin: 0 0 10px; font-weight: 600; }
  .chart-card .canvas-wrap { position: relative; height: 260px; }

  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  th, td { text-align: left; padding: 7px 10px; border-bottom: 1px solid var(--grid); }
  th { color: var(--text-secondary); font-weight: 600; cursor: pointer; white-space: nowrap; }
  th.sorted-asc::after { content: " \2191"; }
  th.sorted-desc::after { content: " \2193"; }
  td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
  tr:hover td { background: var(--chip-bg); }
  .table-card {
    background: var(--surface-1);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 16px;
    overflow-x: auto;
  }
  .table-toolbar {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 10px;
    gap: 12px;
    flex-wrap: wrap;
  }
  .pivot-row { cursor: pointer; }
  .pivot-row:hover td { background: var(--chip-bg); }
  .swatch { display: inline-block; width: 9px; height: 9px; border-radius: 2px; margin-right: 7px; }

  .editable-cell { cursor: pointer; }
  .editable-cell:hover { background: var(--chip-bg); }
  .editable-cell::after { content: "\270E"; opacity: 0; font-size: 10px; margin-left: 6px; color: var(--muted); }
  .editable-cell:hover::after { opacity: 1; }
  #tx-table td input, #tx-table td select {
    font-family: inherit;
    font-size: 13px;
    padding: 3px 6px;
    border-radius: 4px;
    border: 1px solid var(--series-primary);
    background: var(--surface-1);
    color: var(--text-primary);
    width: 100%;
    box-sizing: border-box;
  }

  .pagination { display: flex; align-items: center; gap: 10px; font-size: 12px; color: var(--text-secondary); }
  .pagination button { padding: 4px 9px; }
  .viewall-label { display: flex; align-items: center; gap: 6px; font-size: 12px; color: var(--text-secondary); }

  .legend-row { display: flex; flex-wrap: wrap; gap: 10px 16px; margin-top: 10px; font-size: 12px; color: var(--text-secondary); }
  .legend-row .item { display: flex; align-items: center; }

  .form-card {
    background: var(--surface-1);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 18px;
  }
  .form-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
    gap: 14px;
  }
  .form-grid label {
    display: flex;
    flex-direction: column;
    gap: 5px;
    font-size: 12px;
    color: var(--text-secondary);
  }
  .form-grid input, .form-grid select {
    font-family: inherit;
    font-size: 13px;
    padding: 7px 10px;
    border-radius: 8px;
    border: 1px solid var(--border);
    background: var(--surface-1);
    color: var(--text-primary);
  }
  .link-btn {
    background: none;
    border: none;
    padding: 0;
    margin-left: 6px;
    font-size: 11px;
    font-weight: 400;
    color: var(--series-primary);
    cursor: pointer;
    text-decoration: underline;
  }
  .link-btn:hover { background: none; }
  .form-preview {
    display: flex;
    gap: 24px;
    margin-top: 16px;
    padding-top: 14px;
    border-top: 1px solid var(--grid);
    font-size: 13px;
    color: var(--text-secondary);
  }
  .form-preview strong { color: var(--text-primary); font-variant-numeric: tabular-nums; }
  .form-actions { display: flex; align-items: center; gap: 14px; margin-top: 16px; }
  .form-message { font-size: 12px; }
  .form-message.success { color: var(--good); }
  .form-message.error { color: var(--critical); }
  .hint { font-size: 12px; color: var(--muted); padding: 4px 0 0; }

  footer { padding: 24px 28px 40px; color: var(--muted); font-size: 12px; text-align: center; }

  @media (max-width: 900px) {
    .chart-grid { grid-template-columns: 1fr; }
  }
</style>
</head>
<body>

<header>
  <div>
    <h1>Sales Dashboard</h1>
    <div class="meta" id="hdr-meta"></div>
    <div class="meta" id="load-status"></div>
  </div>
  <div class="actions">
    <button id="load-excel-btn">Load from Excel</button>
    <input type="file" id="load-excel-input" accept=".xlsx,.xls" hidden>
    <button id="export-btn">Export data (.xlsx)</button>
    <button id="theme-toggle">Dark mode</button>
  </div>
</header>

<nav class="tabs">
  <button class="tab-btn active" data-panel="panel-dashboard">Dashboard</button>
  <button class="tab-btn" data-panel="panel-add">Add Sale</button>
</nav>

<main>
<div id="panel-dashboard" class="tab-panel active">

  <section>
    <div class="kpi-row" id="kpi-row"></div>
  </section>

  <section>
    <h2>Filters</h2>
    <div class="filters">
      <div class="filter-group">
        <span class="fg-label">Category</span>
        <div id="slicer-category"></div>
      </div>
      <div class="filter-group">
        <span class="fg-label">Sale Location</span>
        <div id="slicer-location"></div>
      </div>
      <div class="filter-group">
        <span class="fg-label">Year</span>
        <select id="year-select"></select>
      </div>
      <div class="filter-group" style="margin-left:auto;">
        <input type="search" id="search-box" placeholder="Search product, order ID, salesperson...">
      </div>
    </div>
  </section>

  <section>
    <h2>Trends</h2>
    <div class="chart-grid">
      <div class="chart-card" style="grid-column: 1 / -1;">
        <h3>Revenue &amp; Profit by Month</h3>
        <div class="canvas-wrap"><canvas id="chart-trend"></canvas></div>
      </div>
      <div class="chart-card">
        <h3>Revenue by Category</h3>
        <div class="canvas-wrap"><canvas id="chart-category"></canvas></div>
      </div>
      <div class="chart-card">
        <h3>Top 10 Products by Revenue</h3>
        <div class="canvas-wrap"><canvas id="chart-products"></canvas></div>
      </div>
      <div class="chart-card">
        <h3>Revenue by Sale Location</h3>
        <div class="canvas-wrap"><canvas id="chart-location"></canvas></div>
      </div>
    </div>
  </section>

  <section>
    <h2>Category &times; Year Summary</h2>
    <div class="table-card">
      <table id="pivot-table"></table>
    </div>
  </section>

  <section>
    <h2>Transactions</h2>
    <p class="hint">
      Click any cell except Order ID, Revenue, or Profit to edit it &mdash; Revenue and Profit recalculate
      automatically from Units/Unit Price/Cost. Press Enter or click away to save, Escape to cancel.
      Edits are included in <strong>Export data</strong> and kept in this browser tab until then (see the
      Add Sale tab's note on session-only changes).
    </p>
    <div class="table-card">
      <div class="table-toolbar">
        <label class="viewall-label"><input type="checkbox" id="viewall" checked> Show all rows</label>
        <div class="pagination" id="pagination"></div>
      </div>
      <div style="overflow-x:auto;">
        <table id="tx-table"></table>
      </div>
    </div>
  </section>

</div>

<div id="panel-add" class="tab-panel">

  <section>
    <h2>Add a Sale</h2>
    <div class="form-card">
      <form id="add-sale-form">
        <div class="form-grid">
          <label>Purchase Date
            <input type="date" id="add-purchase-date" required>
          </label>
          <label>Sale Date
            <input type="date" id="add-sale-date" required>
          </label>
          <label>Category
            <input type="text" id="add-category" list="category-options" required placeholder="e.g. Uncategorized">
            <datalist id="category-options"></datalist>
          </label>
          <label>Product
            <input type="text" id="add-product" list="product-options" required placeholder="e.g. Wireless Earbuds">
            <datalist id="product-options"></datalist>
          </label>
          <label>Sale Location
            <input type="text" id="add-location" list="location-options" required placeholder="e.g. Online">
            <datalist id="location-options"></datalist>
          </label>
          <label>Customer Segment
            <select id="add-segment" required></select>
          </label>
          <label>Salesperson
            <input type="text" id="add-salesperson" list="salesperson-options" required placeholder="e.g. Priya Shah">
            <datalist id="salesperson-options"></datalist>
          </label>
          <label>Units
            <input type="number" id="add-units" min="1" step="1" value="1" required>
          </label>
          <label>Unit Price
            <input type="number" id="add-price" min="0" step="0.01" required>
          </label>
          <label>Cost <button type="button" id="cost-reset-btn" class="link-btn">use estimate</button>
            <input type="number" id="add-cost" min="0" step="0.01" required>
          </label>
        </div>
        <div class="form-preview">
          <span>Revenue: <strong id="add-revenue-preview">£0.00</strong></span>
          <span>Profit: <strong id="add-profit-preview">£0.00</strong></span>
        </div>
        <div class="form-actions">
          <button type="submit" class="primary">Add sale</button>
          <span id="add-form-message" class="form-message"></span>
        </div>
      </form>
    </div>
    <p class="hint">
      Category and Sale Location are free text &mdash; type an existing one (autocomplete will suggest what's
      already in use) or a brand new one; new values get their own filter chip and chart colour automatically.
      Cost is auto-estimated from each category's historical cost ratio &mdash; edit it directly to override,
      or click "use estimate" to go back to the auto value. New sales are kept in this browser tab only
      (see the note below); use <strong>Export data</strong> to save everything, including what you just
      added, to an Excel file.
    </p>
  </section>

  <section>
    <h2>Added this session (<span id="added-count">0</span>) <button type="button" id="clear-added-btn" class="link-btn" style="display:none;">clear all</button></h2>
    <div class="table-card">
      <table id="added-table"></table>
      <p class="hint" id="added-empty-hint">Nothing added yet in this browser tab.</p>
    </div>
  </section>

</div>
</main>

<footer>
  See documentation/Documentation.md for how this dashboard is built and how data flows into it.
</footer>

<script>
let SALES = __SALES_JSON__;

// ---------------------------------------------------------------------
// Sales added via the "Add Sale" tab this session, persisted best-effort
// to localStorage (see documentation/Documentation.md, "Adding new sales")
// ---------------------------------------------------------------------
let addedRows = [];
const ADDED_ROWS_KEY = "sales-dashboard-added-rows-v1";

(function restoreAddedRows() {
  try {
    const raw = localStorage.getItem(ADDED_ROWS_KEY);
    if (!raw) return;
    const rows = JSON.parse(raw);
    if (Array.isArray(rows)) {
      rows.forEach(r => { SALES.push(r); addedRows.push(r); });
    }
  } catch (e) {}
})();

function persistAddedRows() {
  try { localStorage.setItem(ADDED_ROWS_KEY, JSON.stringify(addedRows)); } catch (e) {}
}

// ---------------------------------------------------------------------
// In-place edits made via the transactions table this session, persisted
// best-effort to localStorage the same way addedRows is (see §7.9 in
// documentation/Documentation.md). Stored as a diff — { orderId: { field:
// value } } — not full rows, and reapplied onto SALES by orderId on load.
// ---------------------------------------------------------------------
let editedFields = {};
const EDITED_FIELDS_KEY = "sales-dashboard-edited-fields-v1";

function persistEditedFields() {
  try { localStorage.setItem(EDITED_FIELDS_KEY, JSON.stringify(editedFields)); } catch (e) {}
}

// Applies one field's value to a row IN PLACE, keeping derived fields
// (revenue/profit/year/month/monthLabel) consistent with it. Shared by
// live cell edits and by restoring persisted edits on load, so both paths
// recompute derived fields identically.
function applyFieldValue(row, key, val) {
  row[key] = val;
  if (key === "units" || key === "unitPrice") {
    row.revenue = Math.round(row.units * row.unitPrice * 100) / 100;
    row.profit = Math.round((row.revenue - row.cost) * 100) / 100;
  } else if (key === "cost") {
    row.profit = Math.round((row.revenue - row.cost) * 100) / 100;
  } else if (key === "saleDate") {
    const [y, m] = String(row.saleDate).split("-");
    row.year = parseInt(y, 10);
    row.month = parseInt(m, 10);
    row.monthLabel = new Date(row.saleDate + "T00:00:00").toLocaleDateString("en-GB", { month: "short", year: "numeric" });
  }
}

(function restoreEditedFields() {
  try {
    const raw = localStorage.getItem(EDITED_FIELDS_KEY);
    if (!raw) return;
    const patches = JSON.parse(raw);
    if (!patches || typeof patches !== "object") return;
    editedFields = patches;
    Object.entries(patches).forEach(([orderId, fields]) => {
      const row = SALES.find(r => r.orderId === orderId);
      if (!row) return; // row no longer present — orphaned patch, harmless
      Object.entries(fields).forEach(([key, val]) => applyFieldValue(row, key, val));
    });
  } catch (e) {}
})();

function cssVar(name) {
  return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
}
function fmtMoney(n) {
  return "£" + Math.round(n).toLocaleString("en-GB");
}
function fmtMoney2(n) {
  return "£" + n.toLocaleString("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function fmtInt(n) {
  return Math.round(n).toLocaleString("en-GB");
}

// ---------------------------------------------------------------------
// Theme
// ---------------------------------------------------------------------
function setTheme(mode) {
  if (mode === "dark") {
    document.documentElement.setAttribute("data-theme", "dark");
    document.getElementById("theme-toggle").textContent = "Light mode";
  } else {
    document.documentElement.setAttribute("data-theme", "light");
    document.getElementById("theme-toggle").textContent = "Dark mode";
  }
  try { localStorage.setItem("sales-dashboard-theme", mode); } catch (e) {}
  renderAll();
}
document.getElementById("theme-toggle").addEventListener("click", () => {
  const current = document.documentElement.getAttribute("data-theme") === "dark" ? "dark" : "light";
  setTheme(current === "dark" ? "light" : "dark");
});
(function initTheme() {
  let saved = null;
  try { saved = localStorage.getItem("sales-dashboard-theme"); } catch (e) {}
  if (saved === "dark" || saved === "light") {
    document.documentElement.setAttribute("data-theme", saved);
    document.getElementById("theme-toggle").textContent = saved === "dark" ? "Light mode" : "Dark mode";
  }
})();

// ---------------------------------------------------------------------
// Category / Sale Location: ranked by revenue, coloured by rank
// ---------------------------------------------------------------------
// Fixed 8-slot categorical palette, used POSITIONALLY (see the CSS above
// and references/palette.md) — Category and Sale Location each get their
// own independent ranking, assigned to these same 8 slots. They never
// appear in the same chart/legend together, so reusing the slots is safe.
const COLOR_SLOTS = ["--series-1", "--series-2", "--series-3", "--series-4",
                      "--series-5", "--series-6", "--series-7", "--series-8"];
const OTHER_COLOR_VAR = "--muted"; // 9th+ distinct value in one field

// Ranks the distinct values of `field` by total revenue (desc) and assigns
// each the next colour slot, up to 8 — an order/colour pairing that's
// fixed once computed, not re-ranked on every filter change (colour
// follows the entity, not its rank). Recomputed on full data reloads only
// (initial load, Load from Excel, or a brand-new value via Add Sale).
function computeOrderAndColors(field) {
  const revenueByValue = new Map();
  SALES.forEach(r => {
    const v = r[field];
    revenueByValue.set(v, (revenueByValue.get(v) || 0) + r.revenue);
  });
  const order = [...revenueByValue.entries()].sort((a, b) => b[1] - a[1]).map(e => e[0]);
  const colorVar = {};
  order.forEach((v, i) => { colorVar[v] = i < COLOR_SLOTS.length ? COLOR_SLOTS[i] : OTHER_COLOR_VAR; });
  return { order, colorVar };
}

let CATEGORY_ORDER, CATEGORY_COLOR_VAR, SALE_LOCATION_ORDER, SALE_LOCATION_COLOR_VAR;
function recomputeCategoryAndLocation() {
  ({ order: CATEGORY_ORDER, colorVar: CATEGORY_COLOR_VAR } = computeOrderAndColors("category"));
  ({ order: SALE_LOCATION_ORDER, colorVar: SALE_LOCATION_COLOR_VAR } = computeOrderAndColors("saleLocation"));
}
recomputeCategoryAndLocation();

// ---------------------------------------------------------------------
// Filter state
// ---------------------------------------------------------------------
const state = {
  categories: new Set(CATEGORY_ORDER),
  locations: new Set(SALE_LOCATION_ORDER),
  year: "All",
  search: "",
  sortCol: "saleDate",
  sortDir: "desc",
  page: 1,
  pageSize: 50,
};

// Reference data derived from whatever's currently in SALES. Each is wrapped
// in a compute function, not just computed once, because "Load from Excel"
// (§ below) replaces SALES wholesale and needs to redo all of this.
// Time-based grouping (year/month/monthLabel) is always derived from
// SaleDate, not PurchaseDate — see extract_data.py for the same choice.
function computeAllYears() { return [...new Set(SALES.map(r => r.year))].sort(); }
function computeAllSegments() { return [...new Set(SALES.map(r => r.segment))].sort(); }
function computeAllProducts() { return [...new Set(SALES.map(r => r.product))].sort(); }
function computeSalespeople() { return [...new Set(SALES.map(r => r.salesperson))].sort(); }

// Average cost-to-revenue ratio per category, used to auto-suggest a Cost
// value on the Add Sale form (the user can always override it).
function computeCostRatios() {
  const ratios = {};
  CATEGORY_ORDER.forEach(cat => {
    const rows = SALES.filter(r => r.category === cat);
    const rev = rows.reduce((s, r) => s + r.revenue, 0);
    const cost = rows.reduce((s, r) => s + r.cost, 0);
    ratios[cat] = rev > 0 ? cost / rev : 0.5;
  });
  return ratios;
}

function computeNextOrderSeq() {
  // Grabs trailing digits regardless of prefix style — real order IDs
  // won't necessarily match this project's own "SO-100001" convention
  // (e.g. "S-00001" also parses fine here).
  return SALES.reduce((max, r) => {
    const m = /(\d+)$/.exec(String(r.orderId));
    return m ? Math.max(max, parseInt(m[1], 10)) : max;
  }, 0);
}

let ALL_YEARS = computeAllYears();
let ALL_SEGMENTS = computeAllSegments();
let ALL_PRODUCTS = computeAllProducts();
let ALL_SALESPEOPLE = computeSalespeople();
let COST_RATIO_BY_CATEGORY = computeCostRatios();
let nextOrderSeq = computeNextOrderSeq();

function filteredRows() {
  const q = state.search.trim().toLowerCase();
  return SALES.filter(r => {
    if (!state.categories.has(r.category)) return false;
    if (!state.locations.has(r.saleLocation)) return false;
    if (state.year !== "All" && r.year !== state.year) return false;
    if (q) {
      const hay = (r.orderId + " " + r.product + " " + r.salesperson).toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

// ---------------------------------------------------------------------
// Slicers (chip groups)
// ---------------------------------------------------------------------
function createSlicer(containerId, items, selectedSet, colorVarMap, onChange) {
  const el = document.getElementById(containerId);
  function render() {
    el.innerHTML = "";
    const allChip = document.createElement("span");
    allChip.className = "chip" + (selectedSet.size === items.length ? " active" : "");
    allChip.textContent = "All";
    allChip.onclick = () => {
      if (selectedSet.size === items.length) { selectedSet.clear(); }
      else { items.forEach(i => selectedSet.add(i)); }
      render(); onChange();
    };
    el.appendChild(allChip);
    items.forEach(item => {
      const chip = document.createElement("span");
      chip.className = "chip" + (selectedSet.has(item) ? " active" : "");
      if (colorVarMap && colorVarMap[item]) {
        const dot = document.createElement("span");
        dot.className = "dot";
        dot.style.background = cssVar(colorVarMap[item]);
        chip.appendChild(dot);
      }
      chip.appendChild(document.createTextNode(item));
      chip.onclick = () => {
        if (selectedSet.has(item)) selectedSet.delete(item); else selectedSet.add(item);
        render(); onChange();
      };
      el.appendChild(chip);
    });
  }
  render();
  return { render, setSelection(list) { selectedSet.clear(); list.forEach(i => selectedSet.add(i)); render(); onChange(); } };
}

let catSlicer, locationSlicer;

// (Re)builds the Category and Sale Location slicers against the current
// CATEGORY_ORDER/SALE_LOCATION_ORDER, resetting both filters to "all
// selected". Needs to fully recreate the slicer (not just call .render())
// because CATEGORY_ORDER/state.categories are reassigned, not mutated —
// the old slicer's closures would otherwise still point at stale objects.
function rebuildCategoryAndLocationFilters() {
  recomputeCategoryAndLocation();
  state.categories = new Set(CATEGORY_ORDER);
  state.locations = new Set(SALE_LOCATION_ORDER);
  catSlicer = createSlicer("slicer-category", CATEGORY_ORDER, state.categories, CATEGORY_COLOR_VAR, () => { state.page = 1; renderAll(); });
  locationSlicer = createSlicer("slicer-location", SALE_LOCATION_ORDER, state.locations, SALE_LOCATION_COLOR_VAR, () => { state.page = 1; renderAll(); });
}

function initFilters() {
  rebuildCategoryAndLocationFilters();

  buildYearSelectOptions();
  document.getElementById("year-select").addEventListener("change", (evt) => {
    state.year = evt.target.value === "All" ? "All" : Number(evt.target.value);
    state.page = 1;
    renderAll();
  });

  const searchBox = document.getElementById("search-box");
  searchBox.addEventListener("input", () => { state.search = searchBox.value; state.page = 1; renderAll(); });

  document.getElementById("viewall").addEventListener("change", () => { state.page = 1; renderAll(); });
}

function buildYearSelectOptions() {
  const yearSel = document.getElementById("year-select");
  const current = yearSel.value;
  yearSel.innerHTML = '<option value="All">All years</option>' +
    ALL_YEARS.map(y => `<option value="${y}">${y}</option>`).join("");
  if ([...yearSel.options].some(o => o.value === current)) yearSel.value = current;
}

// Recomputes ALL_YEARS from SALES and rebuilds the Year filter's options.
// Pass the year of a just-added row to skip the recompute when it's
// already a known year (the common case).
function refreshYearSelectIfNeeded(newYear) {
  if (newYear !== undefined && ALL_YEARS.includes(newYear)) return;
  ALL_YEARS = [...new Set(SALES.map(r => r.year))].sort();
  buildYearSelectOptions();
}

// ---------------------------------------------------------------------
// Tabs
// ---------------------------------------------------------------------
function initTabs() {
  document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".tab-panel").forEach(p => p.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(btn.dataset.panel).classList.add("active");
    });
  });
}

// ---------------------------------------------------------------------
// KPIs
// ---------------------------------------------------------------------
function renderKPIs(rows) {
  const revenue = rows.reduce((s, r) => s + r.revenue, 0);
  const profit = rows.reduce((s, r) => s + r.profit, 0);
  const units = rows.reduce((s, r) => s + r.units, 0);
  const orders = rows.length;
  const margin = revenue > 0 ? (profit / revenue) * 100 : 0;
  const aov = orders > 0 ? revenue / orders : 0;

  const tiles = [
    { label: "Total Revenue", value: fmtMoney(revenue) },
    { label: "Total Profit", value: fmtMoney(profit), sub: margin.toFixed(1) + "% margin" },
    { label: "Units Sold", value: fmtInt(units) },
    { label: "Orders", value: fmtInt(orders) },
    { label: "Avg Order Value", value: fmtMoney2(aov) },
  ];
  document.getElementById("kpi-row").innerHTML = tiles.map(t => `
    <div class="kpi-tile">
      <div class="label">${t.label}</div>
      <div class="value">${t.value}</div>
      ${t.sub ? `<div class="sub">${t.sub}</div>` : ""}
    </div>
  `).join("");
}

// ---------------------------------------------------------------------
// Charts
// ---------------------------------------------------------------------
let chartTrend, chartCategory, chartProducts, chartLocation;

function baseGridOptions() {
  return {
    grid: { color: cssVar("--grid"), drawTicks: false },
    ticks: { color: cssVar("--muted"), font: { size: 11 } },
    border: { color: cssVar("--baseline") },
  };
}

function renderTrendChart(rows) {
  const byMonth = new Map();
  rows.forEach(r => {
    const key = r.year + "-" + String(r.month).padStart(2, "0");
    if (!byMonth.has(key)) byMonth.set(key, { label: r.monthLabel, revenue: 0, profit: 0, sortKey: key });
    const m = byMonth.get(key);
    m.revenue += r.revenue;
    m.profit += r.profit;
  });
  const months = [...byMonth.values()].sort((a, b) => a.sortKey.localeCompare(b.sortKey));

  const ctx = document.getElementById("chart-trend").getContext("2d");
  if (chartTrend) chartTrend.destroy();
  try {
    chartTrend = new Chart(ctx, {
      type: "line",
      data: {
        labels: months.map(m => m.label),
        datasets: [
          {
            label: "Revenue",
            data: months.map(m => m.revenue),
            borderColor: cssVar("--series-primary"),
            backgroundColor: cssVar("--series-primary"),
            borderWidth: 2,
            pointRadius: 4,
            pointHoverRadius: 6,
            tension: 0.25,
          },
          {
            label: "Profit",
            data: months.map(m => m.profit),
            borderColor: cssVar("--series-profit"),
            backgroundColor: cssVar("--series-profit"),
            borderWidth: 2,
            pointRadius: 4,
            pointHoverRadius: 6,
            tension: 0.25,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false },
        plugins: {
          legend: { position: "bottom", labels: { color: cssVar("--text-secondary"), boxWidth: 10, boxHeight: 10 } },
          tooltip: {
            callbacks: { label: (c) => c.dataset.label + ": " + fmtMoney(c.parsed.y) },
          },
        },
        scales: {
          x: baseGridOptions(),
          y: { ...baseGridOptions(), ticks: { ...baseGridOptions().ticks, callback: (v) => fmtMoney(v) } },
        },
      },
    });
  } catch (e) { console.error("chart-trend failed", e); }
}

function renderCategoryChart(rows) {
  const byCat = new Map(CATEGORY_ORDER.map(c => [c, 0]));
  rows.forEach(r => byCat.set(r.category, (byCat.get(r.category) || 0) + r.revenue));
  const entries = [...byCat.entries()].sort((a, b) => b[1] - a[1]);

  const ctx = document.getElementById("chart-category").getContext("2d");
  if (chartCategory) chartCategory.destroy();
  try {
    chartCategory = new Chart(ctx, {
      type: "bar",
      data: {
        labels: entries.map(e => e[0]),
        datasets: [{
          label: "Revenue",
          data: entries.map(e => e[1]),
          backgroundColor: entries.map(e => cssVar(CATEGORY_COLOR_VAR[e[0]])),
          borderRadius: 4,
          borderSkipped: false,
        }],
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: (c) => fmtMoney(c.parsed.x) } },
        },
        onClick: (evt, elements) => {
          if (elements.length) {
            const cat = entries[elements[0].index][0];
            catSlicer.setSelection([cat]);
          }
        },
        scales: {
          x: { ...baseGridOptions(), ticks: { ...baseGridOptions().ticks, callback: (v) => fmtMoney(v) } },
          y: baseGridOptions(),
        },
      },
    });
  } catch (e) { console.error("chart-category failed", e); }
}

function renderProductsChart(rows) {
  const byProduct = new Map();
  rows.forEach(r => byProduct.set(r.product, (byProduct.get(r.product) || 0) + r.revenue));
  const entries = [...byProduct.entries()].sort((a, b) => b[1] - a[1]).slice(0, 10).reverse();

  const ctx = document.getElementById("chart-products").getContext("2d");
  if (chartProducts) chartProducts.destroy();
  try {
    chartProducts = new Chart(ctx, {
      type: "bar",
      data: {
        labels: entries.map(e => e[0]),
        datasets: [{
          label: "Revenue",
          data: entries.map(e => e[1]),
          backgroundColor: cssVar("--series-primary"),
          borderRadius: 4,
          borderSkipped: false,
        }],
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: (c) => fmtMoney(c.parsed.x) } },
        },
        scales: {
          x: { ...baseGridOptions(), ticks: { ...baseGridOptions().ticks, callback: (v) => fmtMoney(v) } },
          y: baseGridOptions(),
        },
      },
    });
  } catch (e) { console.error("chart-products failed", e); }
}

function renderLocationChart(rows) {
  const byLocation = new Map(SALE_LOCATION_ORDER.map(c => [c, 0]));
  rows.forEach(r => byLocation.set(r.saleLocation, (byLocation.get(r.saleLocation) || 0) + r.revenue));
  const entries = SALE_LOCATION_ORDER.map(loc => [loc, byLocation.get(loc) || 0]);

  const ctx = document.getElementById("chart-location").getContext("2d");
  if (chartLocation) chartLocation.destroy();
  try {
    chartLocation = new Chart(ctx, {
      type: "bar",
      data: {
        labels: entries.map(e => e[0]),
        datasets: [{
          label: "Revenue",
          data: entries.map(e => e[1]),
          backgroundColor: entries.map(e => cssVar(SALE_LOCATION_COLOR_VAR[e[0]])),
          borderRadius: 4,
          borderSkipped: false,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { callbacks: { label: (c) => fmtMoney(c.parsed.y) } },
        },
        onClick: (evt, elements) => {
          if (elements.length) {
            const loc = entries[elements[0].index][0];
            locationSlicer.setSelection([loc]);
          }
        },
        scales: {
          x: baseGridOptions(),
          y: { ...baseGridOptions(), ticks: { ...baseGridOptions().ticks, callback: (v) => fmtMoney(v) } },
        },
      },
    });
  } catch (e) { console.error("chart-location failed", e); }
}

// ---------------------------------------------------------------------
// Category x Year pivot table
// ---------------------------------------------------------------------
function renderPivotTable(rows) {
  const years = ALL_YEARS;
  const matrix = new Map(CATEGORY_ORDER.map(c => [c, new Map(years.map(y => [y, 0]))]));
  rows.forEach(r => {
    if (matrix.has(r.category) && matrix.get(r.category).has(r.year)) {
      matrix.get(r.category).set(r.year, matrix.get(r.category).get(r.year) + r.revenue);
    }
  });

  let html = "<thead><tr><th>Category</th>" + years.map(y => `<th class="num">${y}</th>`).join("") + `<th class="num">Total</th></tr></thead><tbody>`;
  CATEGORY_ORDER.forEach(cat => {
    const rowMap = matrix.get(cat);
    const total = [...rowMap.values()].reduce((s, v) => s + v, 0);
    html += `<tr class="pivot-row" data-cat="${cat}">
      <td><span class="swatch" style="background:${cssVar(CATEGORY_COLOR_VAR[cat])}"></span>${cat}</td>
      ${years.map(y => `<td class="num">${fmtMoney(rowMap.get(y))}</td>`).join("")}
      <td class="num"><strong>${fmtMoney(total)}</strong></td>
    </tr>`;
  });
  html += "</tbody>";
  const table = document.getElementById("pivot-table");
  table.innerHTML = html;
  table.querySelectorAll(".pivot-row").forEach(tr => {
    tr.addEventListener("click", () => catSlicer.setSelection([tr.dataset.cat]));
  });
}

// ---------------------------------------------------------------------
// Transactions table
// ---------------------------------------------------------------------
// `revenue`/`profit` are deliberately NOT editable — they're always
// recomputed from Units/UnitPrice and Cost (applyFieldValue()) so they
// can't drift out of sync with the values they're derived from.
const TX_COLUMNS = [
  { key: "purchaseDate", label: "Purchase Date", editable: true, editType: "date" },
  { key: "saleDate", label: "Sale Date", editable: true, editType: "date" },
  { key: "orderId", label: "Order ID" },
  { key: "category", label: "Category", editable: true, editType: "text", datalistId: "category-options" },
  { key: "product", label: "Product", editable: true, editType: "text", datalistId: "product-options" },
  { key: "saleLocation", label: "Sale Location", editable: true, editType: "text", datalistId: "location-options" },
  { key: "segment", label: "Segment", editable: true, editType: "select", optionsFn: () => ALL_SEGMENTS },
  { key: "salesperson", label: "Salesperson", editable: true, editType: "text", datalistId: "salesperson-options" },
  { key: "units", label: "Units", num: true, editable: true, editType: "number", step: "1" },
  { key: "unitPrice", label: "Unit Price", num: true, money: true, editable: true, editType: "number", step: "0.01" },
  { key: "cost", label: "Cost", num: true, money: true, editable: true, editType: "number", step: "0.01" },
  { key: "revenue", label: "Revenue", num: true, money: true },
  { key: "profit", label: "Profit", num: true, money: true },
];

function renderTransactionsTable(rows) {
  const sorted = [...rows].sort((a, b) => {
    const av = a[state.sortCol], bv = b[state.sortCol];
    let cmp;
    if (typeof av === "number" && typeof bv === "number") cmp = av - bv;
    else cmp = String(av).localeCompare(String(bv));
    return state.sortDir === "asc" ? cmp : -cmp;
  });

  const viewAll = document.getElementById("viewall").checked;
  const pageSize = viewAll ? sorted.length || 1 : state.pageSize;
  const totalPages = Math.max(1, Math.ceil(sorted.length / pageSize));
  if (state.page > totalPages) state.page = totalPages;
  const start = viewAll ? 0 : (state.page - 1) * pageSize;
  const pageRows = sorted.slice(start, start + pageSize);

  let thead = "<thead><tr>" + TX_COLUMNS.map(c => {
    const sortCls = state.sortCol === c.key ? (state.sortDir === "asc" ? "sorted-asc" : "sorted-desc") : "";
    return `<th class="${c.num ? "num " : ""}${sortCls}" data-key="${c.key}">${c.label}</th>`;
  }).join("") + "</tr></thead>";

  let tbody = "<tbody>" + pageRows.map((r, i) => "<tr>" + TX_COLUMNS.map(c => {
    let v = r[c.key];
    if (c.money) v = fmtMoney2(v);
    else if (c.num) v = fmtInt(v);
    const cls = (c.num ? "num " : "") + (c.editable ? "editable-cell" : "");
    const editAttrs = c.editable ? ` data-editable="true" data-row-idx="${i}" data-key="${c.key}"` : "";
    return `<td class="${cls}"${editAttrs}>${v}</td>`;
  }).join("") + "</tr>").join("") + "</tbody>";

  const table = document.getElementById("tx-table");
  table.innerHTML = thead + tbody;
  table.querySelectorAll("th[data-key]").forEach(th => {
    th.addEventListener("click", () => {
      const key = th.dataset.key;
      if (state.sortCol === key) state.sortDir = state.sortDir === "asc" ? "desc" : "asc";
      else { state.sortCol = key; state.sortDir = "asc"; }
      renderTransactionsTable(filteredRows());
    });
  });
  table.querySelectorAll('td[data-editable="true"]').forEach(td => {
    td.addEventListener("click", () => {
      const row = pageRows[parseInt(td.dataset.rowIdx, 10)];
      const colDef = TX_COLUMNS.find(c => c.key === td.dataset.key);
      startEditingCell(td, row, colDef);
    });
  });

  const pag = document.getElementById("pagination");
  if (viewAll) {
    pag.innerHTML = `Showing all ${sorted.length.toLocaleString()} rows`;
  } else {
    pag.innerHTML = `
      <button id="pg-prev" ${state.page <= 1 ? "disabled" : ""}>&larr; Prev</button>
      <span>Page ${state.page} of ${totalPages} &middot; ${sorted.length.toLocaleString()} rows</span>
      <button id="pg-next" ${state.page >= totalPages ? "disabled" : ""}>Next &rarr;</button>
    `;
    const prevBtn = document.getElementById("pg-prev");
    const nextBtn = document.getElementById("pg-next");
    if (prevBtn) prevBtn.addEventListener("click", () => { state.page--; renderTransactionsTable(filteredRows()); });
    if (nextBtn) nextBtn.addEventListener("click", () => { state.page++; renderTransactionsTable(filteredRows()); });
  }
}

// ---------------------------------------------------------------------
// Editing a transaction cell in place
// ---------------------------------------------------------------------
// Replaces one <td>'s content with an input (or select) matching its
// column's editType, reusing the same <datalist> elements the Add Sale
// form uses (§7.7) for autocomplete on Category/Product/Sale
// Location/Salesperson. Committing (blur/Enter) or cancelling (Escape)
// both end by re-rendering the table, so there's never more than one
// cell being edited at a time — a click-to-edit table, not a live grid
// of hundreds of permanently-mounted inputs.
function startEditingCell(td, row, colDef) {
  if (td.querySelector("input, select")) return; // already editing this cell
  const key = colDef.key;
  const original = row[key];
  let inputEl;

  if (colDef.editType === "select") {
    inputEl = document.createElement("select");
    colDef.optionsFn().forEach(opt => {
      const o = document.createElement("option");
      o.value = opt;
      o.textContent = opt;
      if (opt === original) o.selected = true;
      inputEl.appendChild(o);
    });
  } else {
    inputEl = document.createElement("input");
    inputEl.type = colDef.editType; // "date" | "text" | "number"
    inputEl.value = original;
    if (colDef.editType === "number") { inputEl.min = "0"; inputEl.step = colDef.step || "0.01"; }
    if (colDef.datalistId) inputEl.setAttribute("list", colDef.datalistId);
  }

  td.innerHTML = "";
  td.appendChild(inputEl);
  inputEl.focus();
  if (inputEl.select) inputEl.select();

  function commit() {
    let newVal = inputEl.value;
    if (colDef.editType === "text") newVal = newVal.trim();
    if (colDef.editType === "date" && !newVal) { renderTransactionsTable(filteredRows()); return; } // don't allow clearing a date
    if (colDef.editType === "number") {
      newVal = parseFloat(newVal);
      if (!Number.isFinite(newVal) || newVal < 0) { renderTransactionsTable(filteredRows()); return; }
    }
    if (colDef.editType !== "number" && !newVal) { renderTransactionsTable(filteredRows()); return; } // don't allow blanking required text
    applyCellEdit(row, key, newVal);
  }
  inputEl.addEventListener("blur", commit);
  inputEl.addEventListener("keydown", (evt) => {
    if (evt.key === "Enter") { evt.preventDefault(); inputEl.blur(); }
    else if (evt.key === "Escape") {
      evt.preventDefault();
      inputEl.removeEventListener("blur", commit); // don't also commit on the blur this triggers
      renderTransactionsTable(filteredRows());
    }
  });
}

// Applies one committed cell edit: updates the row (and any derived
// fields — applyFieldValue()), records it in editedFields for export/
// persistence, refreshes whatever reference data the edit affects (a
// renamed category/location, a new product/salesperson, a moved sale
// date), and re-renders everything so KPIs/charts/pivot/table all reflect
// it immediately.
function applyCellEdit(row, key, newVal) {
  if (row[key] === newVal) { renderTransactionsTable(filteredRows()); return; } // no actual change

  const isNewCategory = key === "category" && !CATEGORY_ORDER.includes(newVal);
  const isNewLocation = key === "saleLocation" && !SALE_LOCATION_ORDER.includes(newVal);

  applyFieldValue(row, key, newVal);

  if (!editedFields[row.orderId]) editedFields[row.orderId] = {};
  editedFields[row.orderId][key] = newVal;
  persistEditedFields();

  if (isNewCategory || isNewLocation) {
    rebuildCategoryAndLocationFilters();
  }
  if (key === "category" || key === "units" || key === "unitPrice" || key === "cost") {
    COST_RATIO_BY_CATEGORY = computeCostRatios();
  }
  if (key === "product" && !ALL_PRODUCTS.includes(newVal)) {
    ALL_PRODUCTS.push(newVal); ALL_PRODUCTS.sort(); refreshProductOptions();
  }
  if (key === "salesperson" && !ALL_SALESPEOPLE.includes(newVal)) {
    ALL_SALESPEOPLE.push(newVal); ALL_SALESPEOPLE.sort(); refreshSalespersonOptions();
  }
  if (key === "saleDate") refreshYearSelectIfNeeded(row.year);
  if (key === "category" || key === "saleLocation") { refreshCategoryOptions(); refreshLocationOptions(); }

  renderAddedTable(); // no-op on the added-rows list itself, but refreshes the export badge
  renderHeaderMeta();
  renderAll();
}

// ---------------------------------------------------------------------
// Add Sale form
// ---------------------------------------------------------------------
let costManuallyEdited = false;

// Fills in the form's option lists from current reference data. Safe to
// call repeatedly (e.g. after "Load from Excel" replaces SALES) since it
// only touches innerHTML/value, never attaches listeners.
function populateAddFormOptions() {
  document.getElementById("add-segment").innerHTML =
    ALL_SEGMENTS.map(s => `<option value="${s}">${s}</option>`).join("");
  const today = new Date().toISOString().slice(0, 10);
  document.getElementById("add-purchase-date").value = today;
  document.getElementById("add-sale-date").value = today;

  refreshCategoryOptions();
  refreshLocationOptions();
  refreshProductOptions();
  refreshSalespersonOptions();
}

// One-time setup: populates the form and wires its event listeners. Only
// ever called once, at page load — re-calling would double-attach listeners.
function populateAddForm() {
  populateAddFormOptions();

  document.getElementById("add-category").addEventListener("input", recomputeAddFormDerived);
  document.getElementById("add-units").addEventListener("input", recomputeAddFormDerived);
  document.getElementById("add-price").addEventListener("input", recomputeAddFormDerived);
  document.getElementById("add-cost").addEventListener("input", () => {
    costManuallyEdited = true;
    recomputeAddFormDerived();
  });
  document.getElementById("cost-reset-btn").addEventListener("click", () => {
    costManuallyEdited = false;
    recomputeAddFormDerived();
  });
  document.getElementById("clear-added-btn").addEventListener("click", clearAllAddedRows);
  document.getElementById("add-sale-form").addEventListener("submit", handleAddSaleSubmit);
}

function refreshCategoryOptions() {
  document.getElementById("category-options").innerHTML =
    CATEGORY_ORDER.map(c => `<option value="${c}">`).join("");
}
function refreshLocationOptions() {
  document.getElementById("location-options").innerHTML =
    SALE_LOCATION_ORDER.map(c => `<option value="${c}">`).join("");
}
function refreshProductOptions() {
  document.getElementById("product-options").innerHTML =
    ALL_PRODUCTS.map(p => `<option value="${p}">`).join("");
}
function refreshSalespersonOptions() {
  document.getElementById("salesperson-options").innerHTML =
    ALL_SALESPEOPLE.map(p => `<option value="${p}">`).join("");
}

function recomputeAddFormDerived() {
  const units = parseFloat(document.getElementById("add-units").value) || 0;
  const price = parseFloat(document.getElementById("add-price").value) || 0;
  const revenue = units * price;
  document.getElementById("add-revenue-preview").textContent = fmtMoney2(revenue);

  const costInput = document.getElementById("add-cost");
  if (!costManuallyEdited) {
    const cat = document.getElementById("add-category").value.trim();
    const ratio = COST_RATIO_BY_CATEGORY[cat] ?? 0.5;
    costInput.value = (revenue * ratio).toFixed(2);
  }
  const cost = parseFloat(costInput.value) || 0;
  document.getElementById("add-profit-preview").textContent = fmtMoney2(revenue - cost);
}

function handleAddSaleSubmit(evt) {
  evt.preventDefault();
  const msgEl = document.getElementById("add-form-message");

  const purchaseDateStr = document.getElementById("add-purchase-date").value;
  const saleDateStr = document.getElementById("add-sale-date").value;
  const category = document.getElementById("add-category").value.trim();
  const product = document.getElementById("add-product").value.trim();
  const saleLocation = document.getElementById("add-location").value.trim();
  const segment = document.getElementById("add-segment").value;
  const salesperson = document.getElementById("add-salesperson").value.trim();
  const units = parseInt(document.getElementById("add-units").value, 10);
  const unitPrice = parseFloat(document.getElementById("add-price").value);
  const cost = parseFloat(document.getElementById("add-cost").value);

  if (!purchaseDateStr || !saleDateStr || !category || !product || !saleLocation || !segment || !salesperson
      || !Number.isFinite(units) || units <= 0
      || !Number.isFinite(unitPrice) || unitPrice < 0
      || !Number.isFinite(cost) || cost < 0) {
    msgEl.textContent = "Please fill in every field with a valid value.";
    msgEl.className = "form-message error";
    return;
  }

  // year/month/monthLabel are derived from Sale Date, not Purchase Date —
  // this is the dashboard's primary time dimension (see computeAllYears()).
  const [y, m] = saleDateStr.split("-");
  const revenue = Math.round(units * unitPrice * 100) / 100;
  const roundedCost = Math.round(cost * 100) / 100;
  const profit = Math.round((revenue - roundedCost) * 100) / 100;
  nextOrderSeq += 1;

  const row = {
    orderId: "SO-" + nextOrderSeq,
    purchaseDate: purchaseDateStr,
    saleDate: saleDateStr,
    year: parseInt(y, 10),
    month: parseInt(m, 10),
    monthLabel: new Date(saleDateStr + "T00:00:00").toLocaleDateString("en-GB", { month: "short", year: "numeric" }),
    category, product, saleLocation, segment, salesperson,
    units, unitPrice: Math.round(unitPrice * 100) / 100,
    revenue, cost: roundedCost, profit,
  };

  const isNewCategory = !CATEGORY_ORDER.includes(category);
  const isNewLocation = !SALE_LOCATION_ORDER.includes(saleLocation);

  SALES.push(row);
  addedRows.push(row);
  persistAddedRows();

  if (!ALL_PRODUCTS.includes(product)) {
    ALL_PRODUCTS.push(product);
    ALL_PRODUCTS.sort();
  }
  if (!ALL_SALESPEOPLE.includes(salesperson)) {
    ALL_SALESPEOPLE.push(salesperson);
    ALL_SALESPEOPLE.sort();
  }
  COST_RATIO_BY_CATEGORY = computeCostRatios();

  // Only rebuild the category/location filter chips (and reset their
  // selection to "all") when a genuinely new value was introduced — a
  // plain revenue-rank shuffle among existing values isn't worth
  // disrupting the user's current filter selection for.
  if (isNewCategory || isNewLocation) {
    rebuildCategoryAndLocationFilters();
  }

  refreshYearSelectIfNeeded(row.year);
  refreshCategoryOptions();
  refreshLocationOptions();
  refreshProductOptions();
  refreshSalespersonOptions();

  msgEl.textContent = `Added ${row.orderId}.`;
  msgEl.className = "form-message success";

  renderAddedTable();
  renderHeaderMeta();
  renderAll();

  // Reset fields that vary order-to-order; leave Category/Sale Location/
  // Segment as-is since consecutive entries often share them.
  document.getElementById("add-product").value = "";
  document.getElementById("add-salesperson").value = "";
  document.getElementById("add-units").value = "1";
  document.getElementById("add-price").value = "";
  costManuallyEdited = false;
  recomputeAddFormDerived();
}

function renderAddedTable() {
  const table = document.getElementById("added-table");
  const emptyHint = document.getElementById("added-empty-hint");
  document.getElementById("added-count").textContent = addedRows.length;
  document.getElementById("clear-added-btn").style.display = addedRows.length ? "inline" : "none";
  updateExportBadge();

  if (!addedRows.length) {
    table.innerHTML = "";
    emptyHint.style.display = "block";
    return;
  }
  emptyHint.style.display = "none";

  const rows = [...addedRows].reverse(); // most recently added first
  let thead = `<thead><tr>
    <th>Purchase Date</th><th>Sale Date</th><th>Order ID</th><th>Category</th><th>Product</th>
    <th>Sale Location</th><th>Salesperson</th><th class="num">Units</th><th class="num">Revenue</th>
    <th class="num">Profit</th><th></th>
  </tr></thead>`;
  let tbody = "<tbody>" + rows.map(r => `
    <tr>
      <td>${r.purchaseDate}</td><td>${r.saleDate}</td><td>${r.orderId}</td><td>${r.category}</td><td>${r.product}</td>
      <td>${r.saleLocation}</td><td>${r.salesperson}</td>
      <td class="num">${fmtInt(r.units)}</td><td class="num">${fmtMoney2(r.revenue)}</td>
      <td class="num">${fmtMoney2(r.profit)}</td>
      <td><button type="button" class="link-btn" data-order="${r.orderId}">Remove</button></td>
    </tr>`).join("") + "</tbody>";
  table.innerHTML = thead + tbody;
  table.querySelectorAll("button[data-order]").forEach(btn => {
    btn.addEventListener("click", () => removeAddedRow(btn.dataset.order));
  });
}

function removeAddedRow(orderId) {
  addedRows = addedRows.filter(r => r.orderId !== orderId);
  SALES = SALES.filter(r => r.orderId !== orderId);
  persistAddedRows();
  if (editedFields[orderId]) { delete editedFields[orderId]; persistEditedFields(); }
  refreshYearSelectIfNeeded();
  rebuildCategoryAndLocationFilters();
  renderAddedTable();
  renderHeaderMeta();
  renderAll();
}

function clearAllAddedRows() {
  if (!addedRows.length) return;
  if (!confirm("Remove all rows added in this session? This cannot be undone.")) return;
  const ids = new Set(addedRows.map(r => r.orderId));
  SALES = SALES.filter(r => !ids.has(r.orderId));
  addedRows = [];
  persistAddedRows();
  refreshYearSelectIfNeeded();
  rebuildCategoryAndLocationFilters();
  renderAddedTable();
  renderHeaderMeta();
  renderAll();
}

function updateExportBadge() {
  const btn = document.getElementById("export-btn");
  const parts = [];
  if (addedRows.length) parts.push(`${addedRows.length} new`);
  const editedCount = Object.keys(editedFields).length;
  if (editedCount) parts.push(`${editedCount} edited`);
  btn.textContent = parts.length ? `Export data (.xlsx) · ${parts.join(", ")}` : "Export data (.xlsx)";
}

// ---------------------------------------------------------------------
// Export
// ---------------------------------------------------------------------
document.getElementById("export-btn").addEventListener("click", () => {
  try {
    const rows = filteredRows();
    const sheetData = rows.map(r => ({
      "OrderID": r.orderId, "PurchaseDate": r.purchaseDate, "SaleDate": r.saleDate,
      "Category": r.category, "Product": r.product, "SaleLocation": r.saleLocation,
      "CustomerSegment": r.segment, "Salesperson": r.salesperson,
      "Units": r.units, "UnitPrice": r.unitPrice, "Revenue": r.revenue, "Cost": r.cost, "Profit": r.profit,
    }));
    const ws = XLSX.utils.json_to_sheet(sheetData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Sales");
    XLSX.writeFile(wb, "Sales Dashboard Export.xlsx");
  } catch (e) {
    console.error("export failed", e);
    alert("Export failed - see console for details.");
  }
});

// ---------------------------------------------------------------------
// Load from Excel — re-reads a workbook client-side (via the bundled
// SheetJS library) and replaces the dataset in place, no rebuild needed.
// This is a JS re-implementation of extract_data.py's column mapping;
// keep the two in sync if data/Sales Data.xlsx's schema ever changes.
// ---------------------------------------------------------------------
const REQUIRED_COLUMNS = [
  "OrderID", "PurchaseDate", "SaleDate", "Category", "Product", "SaleLocation",
  "CustomerSegment", "Salesperson", "Units", "UnitPrice", "Revenue", "Cost", "Profit",
];

function excelSerialToDate(serial) {
  // Excel's day 0 is 1899-12-30 (this accounts for its 1900 leap-year bug).
  // Only used as a fallback — cellDates:true on XLSX.read() normally means
  // date cells already arrive as JS Date objects, not raw serial numbers.
  const utcDays = Math.floor(serial - 25569);
  return new Date(utcDays * 86400 * 1000);
}
function excelDateToIso(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}
// Converts one cell's raw value (Date object, serial number, or string) to
// an ISO "yyyy-mm-dd" string, or null if it isn't a recognisable date.
function parseDateCell(val) {
  if (val instanceof Date) return excelDateToIso(val);
  if (typeof val === "number") return excelDateToIso(excelSerialToDate(val));
  if (typeof val === "string" && val.trim()) return val.slice(0, 10);
  return null;
}

// Mirrors extract_data.py: blank category/location/segment/salesperson are
// bucketed under a labelled placeholder rather than left blank/undefined,
// and a row missing one of the two dates falls back to whichever is
// present (only dropped if NEITHER date is available).
function rowFromSheetRecord(rec) {
  if (!rec["OrderID"]) return null;

  let purchaseDate = parseDateCell(rec["PurchaseDate"]);
  let saleDate = parseDateCell(rec["SaleDate"]);
  if (!saleDate) saleDate = purchaseDate;
  if (!purchaseDate) purchaseDate = saleDate;
  if (!saleDate) return null;

  const [y, m] = saleDate.split("-");
  return {
    orderId: String(rec["OrderID"]),
    purchaseDate,
    saleDate,
    year: parseInt(y, 10),
    month: parseInt(m, 10),
    monthLabel: new Date(saleDate + "T00:00:00").toLocaleDateString("en-GB", { month: "short", year: "numeric" }),
    category: rec["Category"] || "Uncategorized",
    product: rec["Product"] || "Unspecified",
    saleLocation: rec["SaleLocation"] || "Unspecified",
    segment: rec["CustomerSegment"] || "Unspecified",
    salesperson: rec["Salesperson"] || "Unassigned",
    units: Number(rec["Units"]) || 1,
    unitPrice: Math.round((Number(rec["UnitPrice"]) || 0) * 100) / 100,
    revenue: Math.round((Number(rec["Revenue"]) || 0) * 100) / 100,
    cost: Math.round((Number(rec["Cost"]) || 0) * 100) / 100,
    profit: Math.round((Number(rec["Profit"]) || 0) * 100) / 100,
  };
}

function rowsFromWorkbook(wb) {
  const sheetName = wb.SheetNames.includes("Sales") ? "Sales" : wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { defval: null });
  if (!raw.length) throw new Error(`Sheet "${sheetName}" has no data rows.`);

  const missing = REQUIRED_COLUMNS.filter(c => !(c in raw[0]));
  if (missing.length) throw new Error(`Sheet "${sheetName}" is missing column(s): ${missing.join(", ")}.`);

  const rows = raw.map(rowFromSheetRecord).filter(r => r);
  if (!rows.length) throw new Error(`No valid rows found in sheet "${sheetName}" (every row is missing an Order ID, or both dates).`);
  return rows;
}

function showLoadStatus(text, kind) {
  const el = document.getElementById("load-status");
  el.textContent = text;
  el.style.color = kind === "success" ? cssVar("--good")
    : kind === "error" ? cssVar("--critical")
    : cssVar("--text-secondary");
}

async function handleLoadExcelFile(file) {
  const editedCount = Object.keys(editedFields).length;
  if (addedRows.length || editedCount) {
    const bits = [];
    if (addedRows.length) bits.push(`${addedRows.length} sale(s) added`);
    if (editedCount) bits.push(`${editedCount} row(s) edited`);
    const proceed = confirm(
      `You have ${bits.join(" and ")} this session that haven't been exported. ` +
      `Loading "${file.name}" will replace the current dataset with what's in that file. ` +
      `Export first if you want to keep them. Continue?`
    );
    if (!proceed) return;
  }

  showLoadStatus(`Reading ${file.name}...`, null);
  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    const rows = rowsFromWorkbook(wb);

    SALES = rows;
    addedRows = [];
    persistAddedRows();
    editedFields = {};
    persistEditedFields();

    ALL_YEARS = computeAllYears();
    ALL_SEGMENTS = computeAllSegments();
    ALL_PRODUCTS = computeAllProducts();
    ALL_SALESPEOPLE = computeSalespeople();
    rebuildCategoryAndLocationFilters(); // also recomputes CATEGORY/LOCATION order+colours
    COST_RATIO_BY_CATEGORY = computeCostRatios();
    nextOrderSeq = computeNextOrderSeq();

    // Reset the rest of the filter/sort/paging state to defaults so
    // nothing hides the fresh data (categories/locations were already
    // reset to "all" by rebuildCategoryAndLocationFilters() above).
    state.year = "All";
    state.search = "";
    state.sortCol = "saleDate";
    state.sortDir = "desc";
    state.page = 1;
    document.getElementById("search-box").value = "";
    buildYearSelectOptions();
    document.getElementById("year-select").value = "All";

    populateAddFormOptions();
    costManuallyEdited = false;
    recomputeAddFormDerived();
    renderAddedTable();
    renderHeaderMeta();
    renderAll();

    showLoadStatus(`Loaded ${rows.length.toLocaleString()} rows from "${file.name}".`, "success");
    alert(`Loaded ${rows.length.toLocaleString()} rows from "${file.name}".\n\nFilters have been reset to show everything.`);
  } catch (e) {
    console.error("Load from Excel failed", e);
    showLoadStatus(`Couldn't load "${file.name}": ${e.message}`, "error");
    alert(`Couldn't load "${file.name}":\n\n${e.message}`);
  }
}

document.getElementById("load-excel-btn").addEventListener("click", () => {
  document.getElementById("load-excel-input").click();
});
document.getElementById("load-excel-input").addEventListener("change", (evt) => {
  const file = evt.target.files[0];
  evt.target.value = ""; // allow re-selecting the same file next time
  if (file) handleLoadExcelFile(file);
});

// ---------------------------------------------------------------------
// Header meta
// ---------------------------------------------------------------------
function renderHeaderMeta() {
  const dates = SALES.map(r => r.saleDate).sort();
  const addedNote = addedRows.length ? ` · ${addedRows.length} added this session` : "";
  document.getElementById("hdr-meta").textContent =
    `${SALES.length.toLocaleString()} transactions · ${dates[0]} to ${dates[dates.length - 1]}${addedNote}`;
}

// ---------------------------------------------------------------------
// Render orchestration
// ---------------------------------------------------------------------
function renderAll() {
  const rows = filteredRows();
  renderKPIs(rows);
  renderTrendChart(rows);
  renderCategoryChart(rows);
  renderProductsChart(rows);
  renderLocationChart(rows);
  renderPivotTable(rows);
  renderTransactionsTable(rows);
}

renderHeaderMeta();
initFilters();
initTabs();
populateAddForm();
recomputeAddFormDerived();
renderAddedTable();
renderAll();
</script>
</body>
</html>
"""


def main():
    with open(SALES_JSON_PATH) as f:
        sales = json.load(f)
    with open(CHARTJS_PATH) as f:
        chartjs_src = f.read()
    with open(XLSXLIB_PATH) as f:
        xlsxlib_src = f.read()

    html = TEMPLATE
    html = html.replace("__SALES_JSON__", json.dumps(sales, separators=(",", ":")))
    html = html.replace("__CHARTJS_SRC__", chartjs_src)
    html = html.replace("__XLSXLIB_SRC__", xlsxlib_src)

    with open(OUTPUT_PATH, "w") as f:
        f.write(html)

    os.makedirs(os.path.dirname(DOCS_OUTPUT_PATH), exist_ok=True)
    with open(DOCS_OUTPUT_PATH, "w") as f:
        f.write(html)

    print(f"Wrote {OUTPUT_PATH} and {DOCS_OUTPUT_PATH} "
          f"({len(html):,} bytes) from {len(sales):,} rows")


if __name__ == "__main__":
    main()
