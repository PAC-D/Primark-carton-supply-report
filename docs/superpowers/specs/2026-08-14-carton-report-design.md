# Carton Supply Report Tool — Design Spec

Date: 2026-08-14
Status: Approved (2026-08-14)

## Purpose

A local desktop tool that reads the Primark carton supply workbook (`Sales Record V1.xlsx`), lets the user filter the data by packaging supplier, supplier, factory, and a month range, shows the result as a single flattened table with one column per month (like the source Excel layout), and exports the filtered view to a multi-page landscape A4 PDF.

The user runs this on their own PC. No hosting, no accounts, no servers.

## Source data model

The workbook contains sheets named `<Packaging Supplier> - <Year>` (e.g. `M&U - 2026`, `Uniglory Re - 2023`), plus summary sheets `2024` and `Uniglory Sales Analy` which are ignored.

Each data sheet has:

- Row 1: a title (ignored)
- Row 2: headers — `Sl`, `Suppliers`, `Factory `, then one column per month as `datetime` values (first day of month)
- Rows 3+: one row per factory; `Suppliers` (col B) and `Factory ` (col C) are names, month columns are numeric carton quantities (may be empty)

### Parsing rules

1. **Sheet selection:** every sheet is parsed except `2024` and `Uniglory Sales Analy`. A sheet is skipped (with a visible warning) if its header row contains no month columns.
2. **Re replacement:** when both `<Supplier> - <Year>` and `<Supplier> Re - <Year>` exist, only the `Re` sheet is used. No double counting.
3. **Record shape:** each data row becomes one record `(packaging_supplier, supplier, factory, year, month, cartons)`. Rows are flattened to month-level records so the duration filter can sum over arbitrary ranges.
4. **Exclusions:** rows whose factory/supplier name contains "Total" are excluded. Records with zero cartons are kept during parsing (they may become visible when a different range is selected) but contribute nothing to sums.
5. **New sheets** are picked up automatically on every launch and on manual refresh — no code change needed.

## Output table

One row per factory, one column per month in the selected range, plus a trailing `Total` column (row sum across the range).

| Packaging Supplier | Supplier | Factory | Jan-24 | Feb-24 | ... | Total |

- Rows with zero cartons across the whole selected range are hidden.
- Columns are in chronological order; all-zero months still appear as columns.
- One row per (packaging supplier, supplier, factory) triple.

## Filters (4, above the table)

1. **Packaging Supplier** — dropdown: All / each packaging supplier present in data.
2. **Supplier** — dropdown: All / trading suppliers; choices filtered by selected packaging supplier.
3. **Factory** — dropdown: All / factories; choices filtered by packaging supplier + supplier.
4. **Duration** — From / To month pickers.

### Duration rules

- **Maximum range: 24 months.** The To picker is clamped so a selection can never span more than 24 months.
- **Default (no selection):** the latest 24 months — from (latest month with data − 23 months) to (latest month with data).
- "Latest month with data" = the newest (year, month) across all parsed records.

## PDF export

- Button exports the currently filtered table.
- **A4 landscape**; header row repeated on every page; alternating row shading; grid lines.
- Title line with supplier context (e.g. `M&U — Jan-25 to Jul-26`) or "All suppliers"; a `Generated on <date>` line.
- `Total` row at the end of the table (sum of each month column + grand total).
- File name: `carton-report.pdf`, saved via the browser download.
- Disabled when the filtered result is empty.
- ReportLab platypus (handles arbitrary page counts; no system dependencies on Windows).

## Error handling

| Situation | Behaviour |
|---|---|
| Workbook missing or locked (open in Excel) | Error message on screen; app stays usable; Refresh retries |
| Sheet with unparseable shape | Skip it, list it in a warning area; app never crashes |
| Filters produce zero rows | "No data for the selected filters" message; PDF disabled |

## Implementation notes

- Stack: Python 3.14, Streamlit (UI), pandas/openpyxl (Excel read), ReportLab (PDF). pytest for logic tests.
- Single app file `app.py`, `requirements.txt`, `start.bat` (first run installs deps, then opens the app in the default browser), short `README.md`.
- Data is reloaded on app launch and via a Refresh button.

## Testing

- Unit tests (pytest) on a small fixture workbook mimicking the real one:
  - parsing (normal sheet, Re replacement, Total-row exclusion, empty cells, unparseable sheet)
  - duration default (latest 24 months) and 24-month clamp
  - filtering and row hiding
- Manual verification against the real workbook: totals of a few factories match raw Excel; PDF opens with repeated headers across pages.

## Out of scope

- Editing/adding data in the app.
- Multi-user access, hosting, authentication.
- Charts or summary dashboards.
