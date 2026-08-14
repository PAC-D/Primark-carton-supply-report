# Carton Report: Column A Row Selection + SL Column

Date: 2026-08-14

## Problem

The current loader decides which workbook rows are data using two heuristics: a row is skipped when both supplier and factory names are blank, or when the names contain "total". The user wants a single, deterministic rule: a row is data iff its column A cell holds a numeric value; otherwise the row is dropped. Additionally, the output table (web and PDF) gains a fresh 1..N "SL" index column.

## Requirements

1. **Row selection (loader.py)** — a row from a data sheet is kept iff column A (column 1) contains a numeric value. Rows with blank or non-numeric column A (headers, grand-total rows like "Total"/"Total Sales Unit (Qty of cartons)") are dropped. The existing blank-name and "total"-name checks are removed.
   - Known and approved consequence: 11 "Total Fashion PJT" subtotal rows carry a numeric column A and are now included, so per-supplier subtotal figures (e.g., PARKSCENE) appear and double-count alongside their factories. This is intentional per the user's explicit choice.
2. **SL column (filtering.py)** — the output table's first column is `SL`, numbered 1..N in the final sorted/filtered table (after zero-total row hiding). SL is a fresh output index, not taken from the workbook.
3. **PDF (pdf.py)** — SL renders like the other label columns: left-aligned, blank in the totals row. Page sizing, whole-number formatting, repeating headers, and the totals row are unchanged.
4. **App (app.py)** — no code change; the SL column flows through via the DataFrame.

## Behavior Details

- Column A values seen in the workbook: header text "Sl", integers 1..N, blank. "Numeric" means Python int/float (excluding bool). Row 1-2 headers are non-numeric and therefore never selected.
- Sheet selection, Re-sheet replacement, summary-sheet skip, month columns from row 2 date cells, supplier/factory columns 2/3 with the Epyllion label scan, and per-month record creation all stay as-is.
- Column order becomes: `SL | Packaging Supplier | Supplier | Factory | <month cols> | Total`.
- Zero-total rows remain hidden; multi-select filters, 24-month cap, and latest-24 default are unchanged.
- Record count on the real workbook rises from 16,756 (the 11 subtotal rows add ~12 records each).

## Files Touched

- `loader.py` — replace row-selection heuristics with the numeric column A check
- `filtering.py` — SL column in BASE_COLUMNS and numbering
- `pdf.py` — SL in the label-column handling
- `tests/` — loader fixtures use numeric column A; expected column lists updated; new SL and PDF-SL tests

## Out of Scope

- Keeping the "total" name check (explicitly rejected)
- SL from workbook serials (explicitly rejected — fresh 1..N chosen)
- Changing filtering, PDF layout, or app UI behavior