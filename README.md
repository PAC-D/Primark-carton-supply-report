# Primark Carton Supply Report

Local tool to view and export the carton supply workbook as a filterable month-column table.

## Run

Double-click `start.bat` (first run installs dependencies). The app opens in your browser at http://localhost:8501.

## Use

- Place `Sales Record V1.xlsx` in the same folder as `app.py`.
- Filter by Packaging Supplier, Supplier, Factory (multi-select; empty = All), and From/To months (max 24 months; default = latest 24).
- The table starts with an SL (serial) column, numbered 1, 2, 3... in the current view.
- Click "Export PDF" to download a multi-page report. The page is sized to fit the table (no cut-off cells), with repeating headers and a totals row; values are whole numbers.

## Notes

- A row is only read if its column A holds a number; rows with blank column A are skipped. (Subtotal rows such as "Total Fashion PJT" do have numbers and therefore appear in the report.)
