# Carton Report: Static Site Publish (GitHub Pages)

Date: 2026-08-19

## Problem

The tool is a local Streamlit app. The user wants colleagues at PacD and external partners/suppliers to view the report (same filters, same Excel export) on a published static page on GitHub Pages. Anyone with the link may view/download the data — no access control needed. The workbook changes monthly or twice a week, so publishing must be a near-effortless step inside the existing app.

## Requirements

1. **"Publish site" button in the app (app.py)** — next to "Refresh data". On click:
   - Reads the workbook (existing `load_records` path)
   - Builds the full unfiltered table via existing `build_table`
   - Writes the static site files to `docs/site/`
   - Runs `git add docs/site` then `git commit -m "Updated at <YYYY-MM-DD HH:MM>"` (no push — user pushes manually)
   - Shows success or error message inline; git failure shows the git error text, never crashes
2. **New module `publish.py`** — `publish_site(records, out_dir, now)` writes:
   - `index.html` — static shell: filter controls, table container, Export button, caption with "Data as of <publish timestamp>"
   - `data.json` — all table rows as JSON (from `build_table` with no filters)
   - `app.js` — client-side filtering + SheetJS Excel generation (CDN)
   - `site-manifest.json` — publish timestamp, row count
3. **Static page behavior (visitors)** — mirrors the app:
   - Filters: Packaging Supplier, Supplier, Factory (multi-select; empty = All), From/To month pickers (From > To swapped, same rule as app.py:86)
   - Table with SL column numbered 1, 2, 3... in current view
   - Caption: `<row count> rows · <month count> months · <from> to <to>`
   - "Export Excel" builds the .xlsx in-browser with SheetJS: header row bold, totals row, whole-number formatting — matching `render_excel` output as closely as SheetJS allows
   - Footer: "Data as of <publish timestamp>" from `site-manifest.json`
   - Empty data: "No data" message like the app
   - SheetJS CDN unavailable: clear "Excel export unavailable" message, no silent failure
4. **Branch** — `main` (rename local `master` with `git branch -m master main` during setup)
5. **README** — document: one-time setup (create repo, push `main`, enable Pages → "Deploy from a branch" → `main` / `/docs`), and the publish flow

## Behavior Details

- Site URL: `https://<user>.github.io/<repo>/site/` (Pages serves the whole `/docs` folder; existing superpowers docs are served as raw markdown, harmless)
- `data.json` is committed to git (public data is accepted); `Sales Record V1.xlsx` stays gitignored
- Workbook is ~267 KB; JSON of the flattened table is small (a few hundred KB max)
- Publish button is disabled with explanatory text if the workbook is missing/unreadable

## Files Touched

- Create: `publish.py`, `tests/test_publish.py`
- Edit: `app.py` (Publish button + handler), `README.md` (setup + publish docs)
- Tests: `tests/test_publish.py` covers — files written; `data.json` exactly matches `build_table` output for a fixture workbook; manifest timestamp/row count; git commit stages `docs/site` when run from a temp repo (existing `tests/fixtures.py` patterns)

## Out of Scope

- Auto-push to GitHub (user pushes manually)
- Access control / auth on the published page (anyone may view)
- Re-implementing workbook parsing in JS (Python stays the single source of truth)
- JavaScript test framework (minimal Node script, run manually, no new toolchain)
- Changing the existing app's table, filters, or Excel export behavior