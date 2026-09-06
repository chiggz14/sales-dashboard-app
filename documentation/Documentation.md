# Sales Dashboard — Documentation

This document explains how the dashboard is built, how data flows from a
spreadsheet into it, what the dashboard's own code does once it's open in a
browser, and how to move it onto a database later.

## 1. Folder structure

```
1. Dashboard - Sales/
├── documentation/
│   └── Documentation.md          this file
├── data/
│   ├── Sales Data.xlsx           source spreadsheet (tblSalesTransactions)
│   └── sales.json                extracted, flattened copy of the same data
├── build/
│   ├── generate_sample_data.py   one-off: creates the sample Sales Data.xlsx
│   ├── extract_data.py           step 1: spreadsheet -> JSON
│   ├── build_dashboard.py        step 2: JSON -> HTML dashboard
│   ├── chart.umd.js              bundled charting library (Chart.js v4)
│   └── xlsx.core.min.js          bundled spreadsheet library (SheetJS), used
│                                  by the dashboard's "Export data" button
├── dashboard/
│   └── Sales Dashboard.html      the finished, self-contained dashboard
└── docs/
    └── index.html                 identical copy, for GitHub Pages (§10)
```

The dashboard HTML file is **fully self-contained** — the two library files
in `build/` are copied *into* it at build time, so the file in `dashboard/`
has no external dependencies and no internet connection is required to open
it. `chart.umd.js` and `xlsx.core.min.js` only need to exist in `build/` at
*build* time, not when you open the dashboard. `build_dashboard.py` writes
the exact same content to both `dashboard/Sales Dashboard.html` (the nicely-
named local copy) and `docs/index.html` (which GitHub Pages requires be
named `index.html` — see §10) — there's no meaningful difference between
the two files, just where they live and what serves them.

This mirrors the structure used by the Cashflows dashboard elsewhere in this
project — same three-stage pipeline (source data → JSON → self-contained
HTML), so the two are easy to maintain side by side.

## 2. Where the data comes from right now

**`data/Sales Data.xlsx` now holds real data**, not the sample set — it was
replaced with an actual sales export. The dashboard and pipeline handle
that transparently (see §5's note on dynamic categories/locations, which
exists specifically because real data rarely matches a hardcoded taxonomy).

For reference, `build/generate_sample_data.py` is what originally bootstrapped
this project with ~1,700 synthetic transactions (1 Jan–31 Aug 2026, GBP, 6
categories, 26 products, 3 sale locations, 3 customer segments, 10
salespeople, with seasonality and category growth trends baked in). It's
no longer needed now that real data is in place, but it's still there —
re-run it if you ever want to reset `data/Sales Data.xlsx` back to a clean
synthetic dataset (e.g. for testing). `START`/`END` in that script control
the date range it generates.

**Swapping in a different real export** works the same way each time:
replace `data/Sales Data.xlsx` with the new workbook, keeping the same 13
columns (see §4), then either re-run the two build steps (§3) or use
**Load from Excel** in the dashboard itself (§7.8).

## 3. The build pipeline

Two scripts turn the spreadsheet into the dashboard. Run them in order, from
inside the `build/` folder:

```
cd build
python3 extract_data.py       # spreadsheet -> data/sales.json
python3 build_dashboard.py    # JSON -> dashboard/Sales Dashboard.html
```

(`generate_sample_data.py` is a third script, run separately and only when
you want to regenerate or expand the *sample* dataset — see §2.)

**Step 1 — `extract_data.py`**
Opens `data/Sales Data.xlsx` with `openpyxl` and reads the Excel Table
`tblSalesTransactions` on the `Sales` sheet. Each row is converted into a
plain JSON object with a few derived fields added (`year`, `month`,
`monthLabel` — derived from `SaleDate`, not `PurchaseDate`; see §4). Output:
`data/sales.json`.

**Step 2 — `build_dashboard.py`**
Loads `data/sales.json` plus the two bundled libraries, injects all three
into an HTML/CSS/JS template (as a literal JSON array and two `<script>`
blocks), and writes the result to `dashboard/Sales Dashboard.html`. There's
no server and no build tooling beyond plain Python — the "template" is just
a large Python string with placeholders (`__SALES_JSON__`,
`__CATEGORY_ORDER_JSON__`, `__SALE_LOCATION_ORDER_JSON__`, `__CHARTJS_SRC__`,
`__XLSXLIB_SRC__`) substituted at the end of the script.

**When to re-run this:** whenever `Sales Data.xlsx` changes (new
transactions, edited rows, etc.). Replace the file in `data/`, then re-run
both scripts in order — **or** skip this whole pipeline and use the
**Load from Excel** button in the dashboard itself (§7.8), which reads an
updated workbook straight into the open page with no terminal involved.
The two stay in sync with each other on purpose (see §7.8) but are
independent: rebuilding via Python updates the file every future visitor
opens; Load from Excel only updates what's on screen in that browser tab
right now, and is lost on refresh unless you also rebuild.

## 4. Data field reference

### `data/Sales Data.xlsx` — one row per order, in Excel Table `tblSalesTransactions` (sheet `Sales`)

| Column | Type | Notes |
|---|---|---|
| `OrderID` | text | e.g. `SO-100001`, unique per row |
| `PurchaseDate` | date | when the customer placed the order |
| `SaleDate` | date | when the sale was completed/recognised — **this is the date the dashboard's charts, Year filter, and pivot table group by**, not `PurchaseDate` (see §5) |
| `Category` | text | any value you use — **not a fixed list**, see §5 |
| `Product` | text | specific product name |
| `SaleLocation` | text | any value you use — **not a fixed list**, see §5 |
| `CustomerSegment` | text | any value you use |
| `Salesperson` | text | account owner for the order |
| `Units` | number | quantity sold |
| `UnitPrice` | currency (£) | price per unit actually charged |
| `Revenue` | currency (£) | `Units × UnitPrice` |
| `Cost` | currency (£) | cost of goods sold for the order |
| `Profit` | currency (£) | `Revenue − Cost` |

Blank cells in `Category`, `SaleLocation`, `CustomerSegment`, and
`Salesperson` aren't left blank in `sales.json` — they're bucketed under a
labelled placeholder (`"Uncategorized"`, `"Unspecified"`, `"Unspecified"`,
`"Unassigned"` respectively) so nothing is unfilterable or uncoloured; see
§5. A row missing `PurchaseDate` or `SaleDate` (but not both) has the
missing one filled in from the one that's present; a row missing **both**
is dropped (there's no date to place it on the trend chart).

All currency columns are plain numbers — GBP is a display convention (the
Excel number format and the dashboard's `fmtMoney`/`fmtMoney2` both prefix
`£`), not a stored unit, so no conversion happens anywhere in the pipeline.

### `data/sales.json` — the same rows, flattened

Same fields as above (camelCase: `orderId`, `purchaseDate`, `saleDate`,
`category`, `product`, `saleLocation`, `segment`, `salesperson`, `units`,
`unitPrice`, `revenue`, `cost`, `profit`), plus three fields derived from
**`saleDate`** and used for grouping by time: `year`, `month` (1–12),
`monthLabel` (e.g. `"Mar 2026"`).

## 5. Design decisions worth knowing about

- **Category and Sale Location are NOT a fixed, hardcoded list** —
  `build_dashboard.py` has no opinion on your taxonomy at all. Instead,
  `computeOrderAndColors()` (in the dashboard's own JavaScript) ranks
  whatever distinct values actually appear in the loaded data by total
  revenue (descending) and assigns each one a colour slot in that order.
  This changed from an earlier, hardcoded-list design specifically because
  real sales data (§2) rarely matches a taxonomy invented for sample data —
  a business might use two categories, twenty, or none at all.
- **Colours come from a validated, colour-vision-deficiency-safe 8-slot
  palette** (see the `dataviz` skill's `references/palette.md`), assigned
  *positionally* by revenue rank rather than by name — the highest-revenue
  category gets slot 1 (blue), the next gets slot 2 (orange), and so on.
  Category and Sale Location are ranked independently but share the same 8
  CSS variables (`--series-1` … `--series-8`) since they never appear
  together in the same chart or legend. A **9th+ distinct value in either
  field falls back to a shared neutral grey** (`--muted`) — still its own
  filter chip and pivot row, just not individually colour-coded (the
  palette's own validated slot count tops out at 8; see `color-formula.md`'s
  "snap-to-passing" rule). Once assigned, a colour/rank pairing is **not**
  reshuffled by filtering (only by a full data reload — see the next
  bullet) — "colour follows the entity, never its rank."
- **The colour/order pairing is recomputed on a full reload, not on every
  edit**: initial page load, **Load from Excel** (§7.8), and adding a sale
  whose Category or Sale Location has never been seen before (§7.7) all
  call `rebuildCategoryAndLocationFilters()`, which re-ranks everything and
  resets both filters to "all selected". Adding a sale that reuses an
  *existing* category/location does not re-rank or reset filters, even
  though the revenue totals technically shifted slightly — re-ranking on
  every single row would be more disruptive (colours/positions shuffling
  constantly) than useful.
- **Blank Category/SaleLocation/CustomerSegment/Salesperson become labelled
  placeholders**, not nulls (see §4) — `"Uncategorized"`, `"Unspecified"`,
  `"Unspecified"`, `"Unassigned"`. This matters because a `null` category
  can't be matched against a filter chip or given a colour; a labelled
  placeholder can, and shows up as an honest, visible bucket instead of
  silently making rows disappear (which is exactly what was happening
  before this was added — real data is often incomplete on these fields).
- **One measure per axis, no dual-axis charts.** The "Revenue & Profit by
  Month" chart plots two lines on a single, shared currency axis rather than
  giving Profit its own scale — a dual-axis chart is a well-known way to
  visually mislead (mismatched scales make unrelated trends look
  correlated), so this project avoids it everywhere.
- **All charts, the pivot table, and the transactions table respond to the
  same filter state** (Category, Sale Location, Year, search) — the
  Dashboard tab is a single page rather than further sub-tabs, since
  (unlike the Cashflows dashboard) there's only one dataset here, not two.
- **`SaleDate`, not `PurchaseDate`, drives every time-based view** — the
  Year filter, the trend chart, and the Category × Year pivot all group by
  `SaleDate`. `PurchaseDate` is carried through as a plain informational
  column (visible in the transactions table and the Add Sale form) but
  isn't used for any grouping or filtering. This was a deliberate choice
  (not inferred from the data) since "when was this recognised as a sale"
  is usually what a sales trend chart is meant to answer; swap the field
  used in `computeAllYears()`, `extract_data.py`'s date-derivation, and
  `handleAddSaleSubmit()`/`rowFromSheetRecord()` if you'd rather group by
  `PurchaseDate` instead.
- **Clicking a category** — in either the "Revenue by Category" chart or a
  row of the Category × Year Summary table — sets the Category filter to
  just that one category, as a quick way to drill in.
- **New sales added via the Add Sale tab live in the browser only** (see
  §7.7) — they're never written back to `Sales Data.xlsx`. This is a
  deliberate, browser-imposed limitation (no web page can write to an
  arbitrary local file for security reasons), the same one the Cashflows
  dashboard documents for its own editing feature.

## 6. Dashboard structure (HTML)

The HTML file has four main pieces, top to bottom:

1. **`<head>`** — two embedded `<script>` blocks (Chart.js, then SheetJS)
   and one `<style>` block using CSS custom properties (`--surface-1`,
   `--text-primary`, `--cat-electronics`, etc.) so the whole colour scheme
   — chart colours included — can flip between light and dark by swapping
   one `data-theme` attribute, following the OS theme by default.
2. **`<header>`** — title, transaction count/date range, a status line for
   the last Load/Export action, the **Load from Excel** button (§7.8), the
   **Export data** button (with a badge once new rows exist), and the
   dark-mode toggle.
3. **`<nav class="tabs">`** — switches between the two panels below by
   toggling an `active` CSS class; no page reloads.
4. **`<main>`** — two tab panels:
   - **`#panel-dashboard`** — five sections:
     - **KPI row** — Total Revenue, Total Profit (+ margin %), Units Sold,
       Orders, Avg Order Value.
     - **Filters** — Category / Sale Location chip slicers, a Year
       dropdown, and a free-text search box.
     - **Trends** — four charts: Revenue & Profit by Month (line), Revenue
       by Category (bar), Top 10 Products by Revenue (bar), Revenue by
       Sale Location (bar).
     - **Category × Year Summary** — a pivot table, one row per category,
       one column per year, click a row to filter to that category.
     - **Transactions** — a sortable, searchable, paginated (or "show all"),
       **click-to-edit** (see §7.9) table of every order matching the
       current filters.
   - **`#panel-add`** — the **Add Sale** form (see §7.7) and a table of
     everything added so far this session, each row removable.

## 7. Dashboard behaviour (JavaScript)

Everything lives in the third `<script>` block, at the bottom of the file.
There's no framework — plain DOM APIs, rebuilding `innerHTML` on state
changes, same approach as the Cashflows dashboard.

### 7.1 Data loading
`SALES` is the JSON array, embedded directly as a JS literal (no `fetch()` —
this is what makes the file work fully offline). It's declared `let` rather
than `const` because the Add Sale tab (§7.7) appends to it at runtime.

### 7.2 Filter state
A single `state` object holds the current Category/Sale Location selections
(each a `Set`), the selected Year, the search text, and the transactions
table's current sort column/direction/page. `filteredRows()` applies all of
these to `SALES` in one pass; every chart, the pivot table, and the
transactions table are re-derived from its output whenever anything changes
(`renderAll()`).

### 7.3 Slicers (the chip filters)
`createSlicer(containerId, items, selectedSet, colorVarMap, onChange)` is a
small reusable function that renders one chip per item plus an "All" chip,
tracks the current multi-selection in a `Set`, and calls `onChange()`
whenever the user clicks a chip. It returns a `setSelection(list)` method,
which is what "click a category bar / pivot row to filter" calls into.

### 7.4 Charts
Built with Chart.js. Colours are read from the page's CSS custom properties
at render time (`cssVar()`), so switching themes just re-runs the same
render functions rather than needing separate light/dark chart configs.
Every chart-building call is wrapped in `try/catch` so that if Chart.js
were ever unavailable, the rest of the page still renders instead of the
whole script halting partway through.

### 7.5 The transactions table
`renderTransactionsTable()` handles sorting (click any column header),
free-text search (across Order ID, Product, Salesperson), and pagination —
a "Show all rows" checkbox (checked by default) swaps the 50-row page size
for the full filtered row count and replaces the Prev/Next controls with a
plain row count.

### 7.6 Export
The **Export data** button uses the bundled SheetJS library to build an
in-memory workbook containing exactly the rows currently visible under the
active filters, as a single `Sales Export` sheet, then triggers a normal
browser download (`XLSX.writeFile`). This includes any rows added via the
Add Sale tab (§7.7), since `filteredRows()` reads from the same `SALES`
array they're pushed onto. This is a deliberate, browser-imposed bridge: **a
static HTML file cannot write back to `Sales Data.xlsx` directly** (no
browser allows arbitrary filesystem writes from a web page, for security
reasons) — download the export, then send it back (or fold it in yourself)
to update the real spreadsheet.

### 7.7 Add Sale tab
The form (`#add-sale-form`) collects one new order at a time:

- **Purchase Date and Sale Date are two separate date inputs**, both
  defaulting to today. Sale Date, not Purchase Date, is what determines the
  row's `year`/`month`/`monthLabel` (see §5) — so entering a Sale Date in a
  different month to Purchase Date changes which month the sale shows up
  under in every chart, the Year filter, and the pivot table.
- **Category, Product, Sale Location, and Salesperson are all free-text
  inputs** backed by a `<datalist>` (autocomplete from existing values), so
  a brand-new category or sale location can be typed straight in — it gets
  its own colour slot and filter chip automatically via
  `rebuildCategoryAndLocationFilters()` (§5). **Customer Segment** is still
  a `<select>` (its list is dynamic too, via `computeAllSegments()`, just
  presented as a dropdown rather than free text since typos there are less
  self-evident from context than a mistyped category name).
- **Product and Salesperson's datalists** refresh after every add so
  newly-typed values are suggested next time; **Category and Sale
  Location's datalists** only refresh (via `rebuildCategoryAndLocationFilters()`)
  when the *submitted* value was genuinely new — see §5.
- **Cost is auto-estimated**: `computeCostRatios()` gives each category's
  average cost ÷ revenue ratio across whatever's currently in `SALES` (run
  once at page load, and again after Load from Excel — §7.8 — replaces the
  dataset). As Units/Unit Price/Category change, the Cost field is kept in
  sync at `Revenue × ratio` — until the user types into it directly
  (`costManuallyEdited` flips to `true` and auto-updates stop), at which
  point "use estimate" resets to the computed value. Revenue and Profit are
  pure derived previews, not editable fields.
- **`handleAddSaleSubmit()`** validates every field, assigns the next
  `SO-#####` order ID (`nextOrderSeq` tracks the highest trailing number
  found across every order ID currently in `SALES`, regardless of prefix
  style — so it keeps counting up correctly even against real IDs like
  `S-00221` rather than this project's own `SO-100001` convention), derives
  `year`/`month`/`monthLabel` from Sale Date
  the same way `extract_data.py` does, pushes the row onto `SALES` and onto
  `addedRows`, then calls `renderAll()` — the new row is reflected in every
  KPI, chart, the pivot table, and the transactions table immediately, and
  is included the next time **Export data** is clicked (its badge shows a
  running count of session additions).
- **Persistence**: `addedRows` is saved to `localStorage`
  (`sales-dashboard-added-rows-v1`) after every add/remove, wrapped in
  `try/catch` like the dark-mode preference (§7.2), and restored on page
  load — so an accidental refresh of the *same* file in the *same* browser
  doesn't lose what was typed in. This is best-effort, not a database: see
  §9.
- **"Added this session" table** lists everything currently in `addedRows`
  (newest first) with a per-row **Remove** button, plus a **clear all**
  link once anything has been added. Both call back into `renderAll()` so
  removing a mistaken entry immediately un-does its effect on every chart.

### 7.8 Load from Excel
The **Load from Excel** button in the header lets you refresh the dashboard
from an updated spreadsheet **without running the Python build pipeline at
all** — everything happens client-side, using the same bundled SheetJS
library the Export button uses:

- Clicking it opens a normal file picker (`#load-excel-input`, a hidden
  `<input type="file">`); picking a `.xlsx` reads it with
  `file.arrayBuffer()` and `XLSX.read(buf, { type: "array", cellDates: true })`.
- **`rowsFromWorkbook()`** looks for a sheet named `Sales` (falling back to
  the first sheet if there isn't one), converts it to row objects with
  `XLSX.utils.sheet_to_json()`, and checks it has all 13 required columns
  (`REQUIRED_COLUMNS`) — if any are missing, it throws a specific error
  naming them rather than silently misreading the file. Each row then goes
  through `rowFromSheetRecord()`, which is **a JS re-implementation of
  `extract_data.py`'s column mapping** — the two need to be kept in sync by
  hand if the schema ever changes, since nothing enforces that
  automatically. Both `PurchaseDate` and `SaleDate` go through the same
  `parseDateCell()` helper; date cells normally arrive as JS `Date` objects
  (because of `cellDates: true`), with a raw Excel serial-number fallback
  (`excelSerialToDate()`) for the rare case where a cell isn't a
  properly-typed date.
- **On success**, `SALES` is replaced outright with the freshly parsed
  rows, `addedRows` is cleared (see below), every derived reference
  structure (`ALL_YEARS`, `ALL_SEGMENTS`, `ALL_PRODUCTS`, `ALL_SALESPEOPLE`,
  `CATEGORY_ORDER`/`SALE_LOCATION_ORDER` and their colours,
  `COST_RATIO_BY_CATEGORY`, `nextOrderSeq`) is recomputed from it,
  filters/search/sort/paging reset to their defaults so nothing hides the
  new data, and everything re-renders. Both outcomes are reported two ways:
  a status line under the title (green for success, red for an error), and
  a blocking `alert()` with the same message — added deliberately so a
  failure (or a load that happens to look identical to what was already
  showing) can't go unnoticed the way the status line alone could.
- **If you have unexported session additions** (§7.7) when you click Load
  from Excel, you're asked to confirm first — loading a file replaces the
  whole dataset, so anything only sitting in `addedRows`/`localStorage` and
  not yet folded into a spreadsheet would otherwise be silently lost.
- **What it can't do**: this is still a one-click action, not automatic —
  nothing watches the file for changes (browsers can't do that from a
  static page), so you re-click the button each time you want to pick up
  edits. It also only updates what's on screen in *this* browser tab; unlike
  a rebuild via `build_dashboard.py`, it doesn't touch
  `dashboard/Sales Dashboard.html` on disk, so refreshing the page or
  reopening the file reverts to whatever was last built (plus any
  `localStorage`-persisted Add Sale rows — see §7.7).
- **Expects the same 13 column headers as `Sales Data.xlsx`** (§4) — but
  *not* the same Category/Sale Location vocabulary, since neither is a
  fixed list (§5); a workbook using an entirely different taxonomy loads
  fine and just re-ranks/re-colours against whatever it finds. It's meant
  for reloading an updated version of the same kind of workbook (e.g. the
  real spreadsheet after folding in an export, per §7.6), not for pointing
  the dashboard at a structurally unrelated file.

### 7.9 Editing the transactions table
Every cell in the Transactions table is click-to-edit, except `Order ID`
(the row's identity) and `Revenue`/`Profit` (always derived, never stored
independently — see below).

- **Clicking an editable cell** (`startEditingCell()`) swaps its text for
  an input matching the column: a date picker for Purchase/Sale Date, a
  plain number field for Units/Unit Price/Cost, a `<select>` for Segment,
  and a free-text field with autocomplete for Category/Product/Sale
  Location/Salesperson — reusing the **same `<datalist>` elements the Add
  Sale form uses** (§7.7), so a category typed in one place is suggested in
  the other. Only one cell is ever mid-edit at a time; there's no
  permanently-mounted grid of inputs, which is what keeps this fast even
  with hundreds of rows on screen ("Show all" — §7.5).
- **Enter or clicking away commits the edit**; **Escape cancels it**
  (the cell reverts to its previous value — the Escape handler explicitly
  detaches the commit-on-blur listener first, so cancelling doesn't
  double-fire a commit from the blur that removing the input triggers).
  Blanking a required text field or clearing a date is rejected (the edit
  is simply discarded, same as Escape) rather than accepted as empty.
- **`applyFieldValue()`** is the single place that writes a field onto a
  row and keeps everything derived from it consistent: editing Units or
  Unit Price recalculates Revenue (and then Profit from the new Revenue);
  editing Cost recalculates Profit; editing Sale Date recalculates
  `year`/`month`/`monthLabel` (§5) so the row moves to the right place on
  the trend chart and pivot table. This same function is used both for a
  live edit and for reapplying persisted edits on page load, so the two
  paths can't drift out of sync with each other.
- **`applyCellEdit()`** wraps that with everything an edit might need to
  ripple outward: a brand-new Category or Sale Location triggers
  `rebuildCategoryAndLocationFilters()` (§5) exactly like the Add Sale tab
  does; a new Product or Salesperson gets added to the relevant datalist;
  editing Sale Date into an unseen year rebuilds the Year filter; and
  `COST_RATIO_BY_CATEGORY` is recomputed whenever Category, Units, Unit
  Price, or Cost changes, so the Add Sale tab's cost estimate stays
  accurate. It finishes by calling `renderAll()` — every KPI, chart, the
  pivot table, and the transactions table itself reflect the edit
  immediately.
- **Persistence and export mirror the Add Sale tab exactly** (§7.7): edits
  are recorded as a diff — `{ orderId: { field: value } }` — in
  `editedFields`, saved to `localStorage`
  (`sales-dashboard-edited-fields-v1`) after every change, and reapplied
  (via `applyFieldValue()`) onto matching rows the next time the page
  loads. Nothing special was needed to make **Export data** (§7.6) include
  edits — it already reads live rows out of `SALES` via `filteredRows()`,
  and edits mutate those same row objects in place, so an edited value is
  simply what's there by the time Export runs. The **Export data** button's
  badge shows both counts together when relevant, e.g. "3 new, 5 edited".
- **Loading a different file (§7.8) clears `editedFields`** the same way
  it clears `addedRows` — old edits are keyed by order ID against the
  dataset that was loaded when they were made, so reapplying them onto a
  freshly-loaded, potentially unrelated file isn't safe. The confirmation
  dialog before a replacing load mentions both counts if either is
  nonzero.

## 8. Moving to a database eventually

The pipeline was deliberately split into two stages so that only the first
one needs to change:

```
[ data source ]  --(extract_data.py)-->  data/sales.json  --(build_dashboard.py)-->  dashboard/*.html
     ^ this is the only piece that changes
```

To move off Excel and onto a database:

1. **Replace the body of `extract_data.py`** — swap the `openpyxl` workbook
   read for a database query (e.g. via `psycopg2`, `pyodbc`, `sqlite3`, or an
   ORM), keeping the output shape identical: a list of dicts with exactly
   the fields listed in §4's `sales.json` table, written to
   `data/sales.json`.
2. **`build_dashboard.py` needs no changes at all** — it only ever reads
   `data/sales.json`, and has no idea whether that file came from a
   spreadsheet or a database.
3. **If you want the dashboard to always show live data** rather than a
   point-in-time export, the next step beyond this project's current scope
   would be replacing the "embed JSON directly in the HTML" approach with a
   small local web server that serves `/api/sales` from the database and has
   the dashboard `fetch()` it — at that point the dashboard stops being a
   single offline file and becomes a small web app. Ask if/when you want
   this; it's a bigger structural change than swapping the extract step.
4. **The Add Sale tab (§7.7) would move from `localStorage` to a real
   write** at the same point — `handleAddSaleSubmit()` would `POST` to
   `/api/sales` instead of pushing onto the in-memory `SALES` array, and the
   "Added this session" list would become unnecessary once every viewer
   sees the same live data.

## 9. Known limitations

- **No automatic refresh** — a browser can't watch an arbitrary local file
  for changes (the same security restriction that stops it writing to one).
  The HTML file itself is a point-in-time snapshot from whenever it was
  last built; **Load from Excel (§7.8)** removes the need to re-run the
  Python pipeline to see updated data, but it's still a manual, one-click
  action each time, only affects the current browser tab, and expects the
  same 13-column headers as `Sales Data.xlsx` (Category/Sale Location
  values can differ freely — see §5 — but a file missing a required
  *column* produces an error).
- **Only 8 categories and 8 sale locations get an individually distinct
  colour** (§5) — the 9th+ distinct value (ranked by revenue, ascending)
  shares a neutral grey. They're still separate, correctly-filterable
  entries everywhere else (filter chips, pivot rows, chart bars); only the
  colour-coding stops distinguishing them individually past 8.
- **Sales added or edited in the dashboard never reach `Sales Data.xlsx`
  on their own** and aren't shared between browsers, devices, or people —
  both `addedRows` (§7.7) and `editedFields` (§7.9) live in that one
  browser tab's `localStorage` until exported (§7.6) and folded into the
  source spreadsheet by hand, or until the dashboard is rebuilt (§3),
  which starts fresh from `sales.json` and leaves them behind unless they
  were exported and re-incorporated first. If `localStorage` is cleared,
  private-browsing is used, or a different browser/computer opens the
  file, session changes won't be there.
- **No undo for edits** (unlike Add Sale's "remove"/"clear all" for new
  rows — §7.7) — an edited cell's *previous* value isn't kept anywhere
  once a new one is committed, so reverting a mistaken edit means typing
  the original value back in by hand, or reloading the source file fresh
  via Load from Excel (§7.8, which discards session edits after
  confirmation).
- **`localStorage` (dark mode preference) may not persist** in some
  browsers when opening the file directly via `file://` rather than through
  a web server — it fails silently rather than erroring.
- **"Show all" on very large filtered views** renders every matching row as
  real DOM elements. With a few thousand rows this is still fast in modern
  browsers; a much larger dataset (tens of thousands of rows) may notice a
  delay when toggling it on.

## 10. Hosting on GitHub Pages

The dashboard is published at **https://chiggz14.github.io/sales-dashboard-app/**,
served by GitHub Pages from the `sales-dashboard-app` repo's `main` branch,
`/docs` folder (`docs/index.html` — see §1).

**This repo, and therefore this URL, is public.** That was a deliberate,
explicit choice, made after two rounds of confirmation — not a default:

- On a **GitHub Free** plan (which this account is on), GitHub Pages simply
  doesn't work on a private repository at all — not "works but public,"
  genuinely unavailable. Getting *any* Pages site required either making
  the repo public, or upgrading to GitHub Pro/Team/Enterprise (which
  supports a real access-restricted Pages visibility on a private repo).
- The repo was made public specifically to get this link working. That
  means the source code, full commit history, and the dashboard itself —
  **with the real sales data baked into `docs/index.html` and
  `dashboard/Sales Dashboard.html`, per §2's data-inclusion choice** — are
  all publicly visible to anyone, not just people who happen to find the
  Pages URL.
- **If this ever needs to be undone**: `gh repo edit chiggz14/sales-dashboard-app
  --visibility private` (this alone will also break the Pages site, per the
  limitation above, unless paired with a Pro/Team/Enterprise upgrade first).

**Keeping the published site up to date**: since `build_dashboard.py`
writes to `docs/index.html` automatically (§1, §3), a normal rebuild +
commit + push (§3, and the `sales-dashboard-git-workflow` house rule this
project follows — commit/push after changes rather than leaving them
local) is all that's needed — GitHub rebuilds the Pages site from
`docs/index.html` on every push to `main` automatically, typically live
within a minute or two.
