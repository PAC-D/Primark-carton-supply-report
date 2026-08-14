# Carton Report: Excel Export, No PDF, No Month Cap

Date: 2026-08-15

## Problem

The tool currently exports to PDF (landscape pages sized to fit, whole numbers) and caps the month range at 24 months. The user wants the PDF gone, the cap gone (any month range), and an Excel export instead.

## Requirements

1. **No 24-month cap (filtering.py)** — `default_duration` returns the full range (earliest period with data to latest period with data). `clamp_duration` and `MAX_MONTHS` are removed. The app's date pickers allow any From/To within the data (From min = earliest data month, To max = latest data month). The "Range capped at 24 months" caption is removed. All other filtering behavior (multi-select, SL column, zero-total row hiding) is unchanged.
2. **PDF removed** — delete `pdf.py` and `tests/test_pdf.py`; remove `reportlab` and `pypdf` from `requirements.txt`; the app's "Export PDF" button and its error handling are removed.
3. **Excel export (new `excel_export.py`)** — exports the current filtered table as a formatted `.xlsx` workbook via openpyxl (already a dependency):
   - Single sheet named "Report"
   - Row 1: merged title cell, bold, e.g. `M&U, Union — Jan-25 to Dec-26` (same title string the PDF used)
   - Row 2: column headers (bold, blue fill) — `SL | Packaging Supplier | Supplier | Factory | <months> | Total`
   - Rows 1-2 frozen (`freeze_panes = "A3"`)
   - Data rows: whole numbers only (no decimals)
   - Final row: totals (bold, green fill) — per-month sums and the grand total; "Total" label in the Packaging Supplier cell, blank in the other label cells
   - Column widths sized to content (capped to keep long factory names readable)
   - `render_excel(df, title, out)` — accepts a file path or file-like object (mirrors the old `render_pdf` signature); uses `os.fspath` for path-like inputs
4. **App (app.py)** — "Export Excel" download button (`carton-report.xlsx`, `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`); the SL column and table rendering are unchanged.
5. **Docs** — README updated: no PDF mention, Excel export wording, no 24-month cap wording.

## Behavior Details

- Whole numbers: values are formatted as integers via `int(round(v))`; label cells stay strings.
- The default view on opening the app is the full range (earliest to latest data month).
- With the cap gone, the web table can be wide (~84 columns on current data); Streamlit's horizontal scrolling handles it.
- Excel export file is a real .xlsx readable by Excel, WPS, or openpyxl.

## Files Touched

- `filtering.py` — remove cap; `default_duration` = full range; delete `clamp_duration`, `MAX_MONTHS`
- `app.py` — picker bounds, remove clamp + caption, replace PDF button with Excel button
- Delete: `pdf.py`, `tests/test_pdf.py`
- Create: `excel_export.py`, `tests/test_excel_export.py`
- `requirements.txt` — drop `reportlab`, `pypdf`
- `README.md` — wording updates
- Tests: update `tests/test_filtering.py` (default_duration tests, remove clamp tests)

## Out of Scope

- Keeping any PDF code path
- Multiple sheets in the export (single "Report" sheet)
- Changing the web table's layout or filters
- Month cap of any other size