# Primark Carton Supply Report

Local tool to view and export the carton supply workbook as a filterable month-column table.

## Run

Double-click `start.bat` (first run installs dependencies). The app opens in your browser at http://localhost:8501.

## Use

- Place `Sales Record V1.xlsx` in the same folder as `app.py`.
- Filter by Packaging Supplier, Supplier, Factory (multi-select; empty = All), and any From/To month range (default: full range of the data).
- The table starts with an SL (serial) column, numbered 1, 2, 3... in the current view.
- Click "Export Excel" to download the filtered table as a formatted .xlsx (header row, totals row, whole numbers).

## Notes

- A row is only read if its column A holds a number; rows with blank column A are skipped. (Subtotal rows such as "Total Fashion PJT" do have numbers and therefore appear in the report.)

## Publish to GitHub Pages

### One-time setup

1. Rename branch and push the repo to GitHub:

   ```bash
   git remote add origin https://github.com/<user>/<repo>.git
   git push -u origin main
   ```

2. On GitHub: repo → Settings → Pages → Source: "Deploy from a branch" → Branch `main` / `/docs` → Save.
3. The site is live at `https://<user>.github.io/<repo>/`.

### Publish (every time data changes)

1. Open the app (`start.bat`), click **Publish site**. This regenerates `docs/` and commits it with a message like `Updated at 2026-08-19 14:30`.
2. Push to go live: `git push`.
3. The site shows "Data as of <publish timestamp>" so visitors know how fresh it is.

The published page mirrors the app: same filters (Packaging Supplier, Supplier, Factory, month range), SL column, and an Export Excel button (full styled .xlsx built in the browser via ExcelJS). Anyone with the link can view and download the data.
