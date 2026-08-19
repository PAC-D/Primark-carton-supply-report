# PACD Restyle — Design

Date: 2026-08-19
Status: Approved

## Goal

Restyle the Primark Carton Supply Report static site and its Excel export to match the Carton Price Calculator project's PACD website theme (colors, header logos, footer, favicon). The reference project is `C:\Users\Shoaib\OneDrive - PacD\Projects\Primark\Primark Carton SQM Analysis\Carton-Price-Calculator`, specifically its `primark-pricing-data-check` subpage, which has the same shape (filters + table + export).

## Scope

- Static site: `site_template/` (source of truth) and the generated `docs/site/` (regenerated on publish).
- Excel export palette: `excel_export.py` `render_excel` (Python, used by the Streamlit app's download button) and the mirrored JS `toWorkbookData` in `site_template/app.js` (site export button). Both must stay identical.
- Out of scope: the Streamlit app page look (unchanged), the welcome/step wizard page structure (ours is a report, not a wizard), lucide icons.

## Decisions

- Approach A: mirror the data-check subpage — separate `css/style.css`, copied binary logo assets, navbar/footer/bg-decoration shell.
- Footer: text only — "© 2026 PACD. All rights reserved. | Developed by EV1" (no cross-link to the calculator).
- Zero JS logic changes for the site; all changes are HTML/CSS/asset + Excel palette.

## Site structure & assets

`site_template/index.html` (rewritten shell, same IDs and script order `logic.js` → ExcelJS CDN → `app.js`):

- Head: title "Primark Carton Supply Report", `<link rel="icon" href="favicon.png">`, Google Fonts link (Outfit 300–800 + Inter 400–700, same URL as reference).
- `div.bg-decoration` (fixed, radial gradients, `z-index: -1`).
- `nav.navbar`: `.logo-section` with `img.pacd.png` (40px, `alt="PACD Logo"`), `h1` "Primark Carton Supply Report" (navy, bold), direct navbar `img.primark-logo` (assets/primark-logo.png, 30px, `alt="Primark Logo"`).
- `div.main-container` → `div.calculator-card` (glass-card shell) containing: filters card (section label "Filters"), export button, caption, table container, error divs — existing element IDs preserved exactly (`pkg`, `sup`, `fac`, `from`, `to`, `export`, `caption`, `table`, `export-error`, `data-error`); `data-as-of` replaces the former `footer` id (publish timestamp line).
- `footer.app-footer`: text only, exact reference styles.

New files in `site_template/` (binary copies from the reference project):
- `css/style.css`
- `favicon.png`
- `pacd.png`
- `assets/primark-logo.png`

`publish.py`: add `ASSETS = ["css/style.css", "favicon.png", "pacd.png", "assets/primark-logo.png"]`; `publish_site` copies each from `site_template/` to `docs/site/` (byte-identical). The 5 generated files stay generated as today.

## CSS

Port of `primark-pricing-data-check/styles.css` with the same `:root` variables:

- `--primary: #00205b`, `--primary-hover: #001540`, `--secondary: #e31837`, `--bg-dark: #f8fafc`, `--bg-card: #ffffff`, `--bg-glass: rgba(255,255,255,0.9)`, `--border: #e2e8f0`, `--text-main: #0f172a`, `--text-muted: #64748b`, shadows as reference.
- Body: `font-family: 'Outfit', sans-serif`, `--bg-dark` background, column flex, min-height 100vh.
- Navbar: height 70px, `--bg-glass` + `backdrop-filter: blur(12px)`, bottom border, `--shadow-sm`, space-between padding 0 2rem.
- `.calculator-card`: white, `--border`, `border-radius: 1.5rem`, `box-shadow: var(--shadow-lg)`, `slideUp` animation, **max-width 1400px** (adaptation: report table is ~78 columns wide; horizontal scroll below that).
- Filters: grid `repeat(auto-fit, minmax(180px, 1fr))`, gap 1rem; `.section-label` uppercase red; `.custom-select` single selects (chevron background, focus ring); multi-selects styled to match (border/radius/focus) with `size` kept; `.select-label` per reference.
- Table: `#table` (max-height 500px, overflow auto, border, radius 8px), `#table thead` sticky; **th: navy `#00205B` fill, white, 13px uppercase, letter-spacing 0.5px** (adaptation); td 14px, `border-bottom #f1f5f9`, row hover `#eff6ff`; left-aligned label columns (SL/Packaging Supplier/Supplier/Factory), right-aligned numeric columns with navy weight-600 values.
- `.primary-btn`: pill (radius 99px), navy, white, shadow, hover `--primary-hover` + lift, disabled gray `#94a3b8`.
- `.error-box`: red-tinted (reference) for data/export errors.
- `.app-footer`: reference styles verbatim; `.divider`, `.developer-link` navy.
- `@media (max-width: 640px)`: navbar padding, logo 22px/16px, card padding 1.5rem, filters single column.

## Excel export palette

Applied identically in `render_excel` (Python) and `toWorkbookData` (JS):

- Title row: bold 14pt, merged, font color `#00205B` (was black).
- Header row: fill `#00205B`, bold white text (was `#D9E2F3` fill / black).
- Totals row: fill `#D9E2F3`, bold navy `#00205B` text (was `#E2EFDA` / black).
- Unchanged: borders `#BFBFBF`, widths cap 40, freeze pane A3, right-aligned numerics, whole-number format, `whole()` behavior, workbook structure (columns/months alignment).

## Tests

- `tests/test_publish.py`: new test — `publish_site` output contains the 4 asset files, each byte-identical to its `site_template/` source.
- `tests/test_logic.js`: currently asserts workbook structure only (no color assertions); add assertions for the new header/title/totals colors in `toWorkbookData`.
- `tests/test_excel_export.py` (lines ~46–50): update `A2` header fill assertion from `D9E2F3` → `00205B` and `B7` totals fill from `E2EFDA` → `D9E2F3`, plus the new font color assertions (title/header/totals).
- Existing 38 tests must stay green.

## Deployment

- Regenerate `docs/site` with real data (one-off publish invocation) and commit ("Updated at …" message per existing flow) so the deployed artifact matches.
- README: no changes (publish flow unchanged).

## Testing note

Run: `node tests/test_logic.js`, `python -m pytest -q`, and the headless/served page check via `python -m http.server` from `docs/` (or `docs/site` root) to confirm the site renders with the new shell and data loads.
