# Clinic Streak

A web dashboard that reproduces the **Clinic Streak** pivot from the Metabase
export. It shows order `subtotal` totals broken down by **Rep → Clinic → Drug**
across monthly columns, with interactive filters for **rep**, **clinic name**, and
**drug name**.

The app reads the raw `Query result` data (one row per prescription line item),
buckets each order by calendar month, and renders an expandable matrix in the
shape of the `Clinic_Streak` sheet.

---

## What it does

- **Pivot view** — rows grouped Rep ▸ Clinic ▸ Drug, columns = order months
  (e.g. Sep 2025, Oct 2025, …), values = sum of `subtotal`, with a Grand Total column.
- **Filters** — multi-select by **Sales Rep**, **Clinic Name**, and **Drug Name**.
  Filters combine (AND) and instantly recompute the matrix.
- **Monthly columns** — each order is bucketed by the calendar month of its
  `order_date`.
- **Recency sort** — rows are ordered by subtotal in the most recent month first,
  then the prior month, and so on, so the longer it has been since a row's last
  order, the lower it sinks (a quick read on "months since last order").
- **Whole-number values** — amounts are shown rounded, no decimals.
- **Expand / collapse** — drill from Rep down to Clinic down to Drug.
- **Totals** — per-row monthly subtotals plus a Grand Total per row and per month.

---

## Repository index

| Path | Description |
|------|-------------|
| [`README.md`](README.md) | This index. |
| [`docs/DATA_DICTIONARY.md`](docs/DATA_DICTIONARY.md) | Every source column, its meaning, and which ones the app uses. |
| [`docs/MONTH_BUCKETING.md`](docs/MONTH_BUCKETING.md) | How orders are grouped into month columns. |
| [`docs/SETUP.md`](docs/SETUP.md) | Install, run, and deploy instructions. |
| [`clinic-streak-standalone.html`](clinic-streak-standalone.html) | **Self-contained single file** — double-click it, drop your `.xlsx`, done. No server, no internet, no install. |
| [`src/index.html`](src/index.html) | Multi-file dashboard (served over HTTP). Same UI; loads the modules below. |
| [`src/pivot.js`](src/pivot.js) | Pivot/aggregation engine, month bucketing, and recency sort. |
| [`src/xlsx-lite.js`](src/xlsx-lite.js) | Zero-dependency in-browser `.xlsx` reader (unzips + parses with built-in browser APIs — no SheetJS, no CDN). |
| [`scripts/convert_xlsx_to_json.py`](scripts/convert_xlsx_to_json.py) | Optional: converts the Metabase `.xlsx` export into a committed `data.json`. |
| [`data/data.json`](data/data.json) | Generated dataset (run the script to (re)create it). |
| [`.gitignore`](.gitignore) | Excludes raw exports and local artifacts. |
| [`LICENSE`](LICENSE) | MIT license. |

---

## Quick start

**Easiest — no tooling at all:** double-click **`clinic-streak-standalone.html`**,
then drag a `Metabase_*.xlsx` export onto the drop zone. It unzips and parses the
file in your browser (no server, no internet, nothing uploaded) and builds the
pivot in a couple of seconds.

**Multi-file version** (if you'd rather serve the `src/` folder):

```bash
git clone https://github.com/<your-org>/clinic-streak.git
cd clinic-streak
python -m http.server 8000   # or: npx serve  /  VS Code Live Server
# open http://localhost:8000/src/index.html  -> then drop your .xlsx
```

Prefer to pre-bake the data with Python? See
[`docs/SETUP.md`](docs/SETUP.md) Option B.

---

## Data source

The source file is the Metabase export `Metabase_YYYY-MM-DD.xlsx`, which contains:

- **`Query result`** — the raw line-item table (one row per order line). This is
  what the app consumes.
- **`Clinic_Streak`** — the Excel pivot the app reproduces. Used only as a
  reference for matching layout and totals; the app does not read it.

The export is **not committed** to the repo (it can contain customer data) — see
[`.gitignore`](.gitignore). Generate `data/data.json` locally and commit that, or
keep it out of version control per your data-handling policy.

---

## How the pivot is built

1. Load every row of `Query result`.
2. Bucket each row by the calendar month of its `order_date`.
3. Apply the active Rep / Clinic / Drug filters.
4. Group surviving rows by Rep, then Clinic, then Drug.
5. Sum `subtotal` into each (group, month) cell and into Grand Total.
6. Render the nested, expandable matrix.

The aggregation is in [`src/pivot.js`](src/pivot.js) and is intentionally
filter-agnostic, so adding more dimensions later is straightforward.
