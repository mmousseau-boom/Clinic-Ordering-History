# Clinic Streak

A web dashboard that reproduces the **Clinic Streak** pivot from the Metabase
export. It shows order `subtotal` totals broken down by **Rep → Clinic → Drug**
across weekly columns, with interactive filters for **rep**, **clinic name**, and
**drug name**.

The app reads the raw `Query result` data (one row per prescription line item),
computes the same week numbering used in the Excel pivot, and renders an
expandable weekly matrix identical in shape to the `Clinic_Streak` sheet.

---

## What it does

- **Pivot view** — rows grouped Rep ▸ Clinic ▸ Drug, columns = order weeks
  (Week 1, 2, 3, …), values = sum of `subtotal`, with a Grand Total column.
- **Filters** — multi-select by **Sales Rep**, **Clinic Name**, and **Drug Name**.
  Filters combine (AND) and instantly recompute the matrix.
- **Weekly columns** — same week index as Metabase: `INT((order_date - 45894) / 7)`,
  i.e. Week 1 begins **2025-09-01** (Excel serial base `45894` = 2025-08-25).
- **Expand / collapse** — drill from Rep down to Clinic down to Drug.
- **Totals** — per-row weekly subtotals plus a Grand Total per row and per week.

---

## Repository index

| Path | Description |
|------|-------------|
| [`README.md`](README.md) | This index. |
| [`docs/DATA_DICTIONARY.md`](docs/DATA_DICTIONARY.md) | Every source column, its meaning, and which ones the app uses. |
| [`docs/WEEK_NUMBERING.md`](docs/WEEK_NUMBERING.md) | How week numbers are derived to match Metabase exactly. |
| [`docs/SETUP.md`](docs/SETUP.md) | Install, run, and deploy instructions. |
| [`src/index.html`](src/index.html) | Single-page dashboard (HTML + CSS + JS, no build step). |
| [`src/pivot.js`](src/pivot.js) | Pivot/aggregation engine and week-numbering logic. |
| [`scripts/convert_xlsx_to_json.py`](scripts/convert_xlsx_to_json.py) | Converts the Metabase `.xlsx` export into the `data.json` the app loads. |
| [`data/data.json`](data/data.json) | Generated dataset (run the script to (re)create it). |
| [`.gitignore`](.gitignore) | Excludes raw exports and local artifacts. |
| [`LICENSE`](LICENSE) | MIT license. |

---

## Quick start

```bash
# 1. Clone
git clone https://github.com/<your-org>/clinic-streak.git
cd clinic-streak

# 2. Generate data.json from a fresh Metabase export
python scripts/convert_xlsx_to_json.py path/to/Metabase_YYYY-MM-DD.xlsx data/data.json

# 3. Serve the app (any static server works)
python -m http.server 8000 --directory src
# open http://localhost:8000
```

See [`docs/SETUP.md`](docs/SETUP.md) for full details, including refreshing the
data each week and deploying to GitHub Pages.

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
2. Compute each row's week index from `order_date`.
3. Apply the active Rep / Clinic / Drug filters.
4. Group surviving rows by Rep, then Clinic, then Drug.
5. Sum `subtotal` into each (group, week) cell and into Grand Total.
6. Render the nested, expandable matrix.

The aggregation is in [`src/pivot.js`](src/pivot.js) and is intentionally
filter-agnostic, so adding more dimensions later is straightforward.
