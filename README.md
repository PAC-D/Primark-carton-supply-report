# Primark Carton Supply Report

Local tool to view and export the carton supply workbook as a filterable month-column table.

## Run

Double-click `start.bat` (first run installs dependencies). The app opens in your browser at http://localhost:8501.

## Use

- Place `Sales Record V1.xlsx` in the same folder as `app.py`.
- Filter by Packaging Supplier, Supplier, Factory, and From/To months (max 24 months; default = latest 24).
- Click "Export PDF" to download a landscape A4 report with repeating headers and a totals row.

## Notes

- New sheets added to the workbook are picked up on app start or after clicking "Refresh data".
- `Re` sheets replace the original sheet for the same supplier and year.
- Sheets `2024` and `Uniglory Sales Analy` are ignored.
