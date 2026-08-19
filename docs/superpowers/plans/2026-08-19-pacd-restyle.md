# PACD Restyle Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Restyle the static report site and the Excel export to the Carton Price Calculator project's PACD theme.

**Architecture:** `site_template/` stays the single source of truth: rewrite `index.html` as the PACD shell, add `css/style.css` plus binary logo/favicon assets, extend `publish.py` to copy the new assets to `docs/site/`, and recolor the shared Excel styling in both `excel_export.py` (Python) and `logic.js`/`app.js` (JS). No behavioral changes.

**Tech Stack:** Python 3 (openpyxl), Node (assert-based test runner), vanilla HTML/CSS/JS, ExcelJS 4.4.0 (CDN, unchanged).

## Global Constraints

- Palette (verbatim from spec): `--primary: #00205b`, `--primary-hover: #001540`, `--secondary: #e31837`, `--bg-dark: #f8fafc`, `--bg-card: #ffffff`, `--bg-glass: rgba(255,255,255,0.9)`, `--border: #e2e8f0`, `--text-main: #0f172a`, `--text-muted: #64748b`.
- Excel palette: title/header font navy `00205B`, header fill navy `00205B`, header text white `FFFFFF`, totals fill light navy `D9E2F3`, totals text navy `00205B`. Keep `BFBFBF` borders, widths cap 40, freeze A3, right-aligned numerics.
- Footer is text-only: `© 2026 PACD. All rights reserved. | Developed by EV1` (no cross-link).
- Existing element IDs preserved: `pkg`, `sup`, `fac`, `from`, `to`, `export`, `caption`, `table`, `export-error`, `data-error`. New ID added: `data-as-of`.
- Script order unchanged: `logic.js` → ExcelJS CDN (`https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js`) → `app.js`.
- Published `docs/site/` files must stay byte-identical to `site_template/` (generated files and copied assets alike).
- Reference project for binary asset copies: `C:\Users\Shoaib\OneDrive - PacD\Projects\Primark\Primark Carton SQM Analysis\Carton-Price-Calculator`.
- Commit messages follow repo style (`feat:`/`fix:`/`docs:`/`style:`); publish commits are auto-generated "Updated at <timestamp>".

---

### Task 1: PACD palette in Python Excel export

**Files:**
- Modify: `excel_export.py:9-12,49-71`
- Test: `tests/test_excel_export.py:41-54`

**Interfaces:**
- Consumes: nothing new — `render_excel(df, title, out)` as-is.
- Produces: same `render_excel(df, title, out)` signature; styled workbook with the new palette.

- [ ] **Step 1: Update the styling test to the new palette (make it fail first)**

Replace the assertions in `test_excel_export_styling` (`tests/test_excel_export.py:46-50`) with:

```python
    assert ws["A1"].font.bold
    assert ws["A1"].font.color.rgb.endswith("00205B")
    assert ws["A2"].font.bold
    assert ws["A2"].font.color.rgb.endswith("FFFFFF")
    assert ws["A2"].fill.start_color.rgb.endswith("00205B")
    assert ws["B7"].value == "Total"
    assert ws["B7"].font.bold
    assert ws["B7"].font.color.rgb.endswith("00205B")
    assert ws["B7"].fill.start_color.rgb.endswith("D9E2F3")
    assert ws["A3"].border.left.style == "thin"
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `python -m pytest tests/test_excel_export.py::test_excel_export_styling -v`
Expected: FAIL — color/fill assertions mismatch the current `D9E2F3`/`E2EFDA` styling.

- [ ] **Step 3: Update the constants and style application**

In `excel_export.py`, replace lines 9-10 with:

```python
HEADER_BG = "00205B"
TOTAL_BG = "D9E2F3"
TITLE_COLOR = "00205B"
HEADER_COLOR = "FFFFFF"
TOTAL_COLOR = "00205B"
```

Then update the three style sites:

```python
    title_cell.font = Font(bold=True, size=14, color=TITLE_COLOR)
```
(was `title_cell.font = Font(bold=True, size=14)`)

```python
        cell.font = Font(bold=True, color=HEADER_COLOR)
```
(was `cell.font = Font(bold=True)` in the header loop)

```python
            cell.font = Font(bold=True, color=TOTAL_COLOR)
```
(was `cell.font = Font(bold=True)` in the totals loop)

- [ ] **Step 4: Run the test to verify it passes**

Run: `python -m pytest tests/test_excel_export.py -v`
Expected: PASS — all 5 excel_export tests green.

- [ ] **Step 5: Run the full suite**

Run: `python -m pytest -q`
Expected: 38 passed.

- [ ] **Step 6: Commit**

```bash
git add excel_export.py tests/test_excel_export.py
git commit -m "style: PACD palette for Excel export"
```

---

### Task 2: PACD palette in JS Excel export

**Files:**
- Modify: `site_template/logic.js` (append constants block)
- Modify: `site_template/app.js:112-135` (consume constants)
- Test: `tests/test_logic.js` (append assertion)

**Interfaces:**
- Consumes: nothing — `CartLogic` namespace already exported via UMD.
- Produces: `CartLogic.WORKBOOK_COLORS` — an object with keys `headerBg`, `totalBg`, `titleColor`, `headerColor`, `totalColor`, values as 8-digit ARGB strings. `app.js` `exportExcel()` uses it for all workbook colors.

- [ ] **Step 1: Append the failing test**

Append to `tests/test_logic.js` (after the existing assertions, before the final success log):

```js
const COLORS = Logic.WORKBOOK_COLORS;
assert.deepStrictEqual(COLORS, {
  headerBg: "FF00205B",
  totalBg: "FFD9E2F3",
  titleColor: "FF00205B",
  headerColor: "FFFFFFFF",
  totalColor: "FF00205B",
});
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `node tests/test_logic.js`
Expected: FAIL — `Logic.WORKBOOK_COLORS` is `undefined`.

- [ ] **Step 3: Add the constants to logic.js**

In `site_template/logic.js`, add the member to the returned object literal (lines 133-142, next to `toWorkbookData`):

```js
    toWorkbookData: toWorkbookData,
    WORKBOOK_COLORS: {
      headerBg: "FF00205B",
      totalBg: "FFD9E2F3",
      titleColor: "FF00205B",
      headerColor: "FFFFFFFF",
      totalColor: "FF00205B"
    }
  };
```

- [ ] **Step 4: Consume the constants in app.js**

In `site_template/app.js` `exportExcel()`, replace line 112:

```js
    var HEADER_BG = { argb: "FFD9E2F3" }, TOTAL_BG = { argb: "FFE2EFDA" };
```

with:

```js
    var COLORS = CartLogic.WORKBOOK_COLORS;
    var HEADER_BG = { argb: COLORS.headerBg }, TOTAL_BG = { argb: COLORS.totalBg };
```

Update the title font (line 123):

```js
    titleCell.font = { bold: true, size: 14, color: { argb: COLORS.titleColor } };
```

Update the header font (line 127):

```js
      headerCell.font = { bold: true, color: { argb: COLORS.headerColor } };
```

Update the totals font (line 133):

```js
      totalCell.font = { bold: true, color: { argb: COLORS.totalColor } };
```

- [ ] **Step 5: Run the test to verify it passes**

Run: `node tests/test_logic.js`
Expected: PASS — "All logic tests passed" (including the new color assertions).

- [ ] **Step 6: Syntax-check app.js**

Run: `node --check site_template/app.js`
Expected: no output, exit 0.

- [ ] **Step 7: Commit**

```bash
git add site_template/logic.js site_template/app.js tests/test_logic.js
git commit -m "style: PACD palette for JS Excel export"
```

---

### Task 3: PACD site shell — assets, CSS, index.html

**Files:**
- Create: `site_template/css/style.css`
- Create: `site_template/favicon.png`, `site_template/pacd.png`, `site_template/assets/primark-logo.png` (binary copies)
- Modify: `site_template/index.html` (full rewrite)
- Modify: `site_template/app.js:204` (data-as-of target)

**Interfaces:**
- Consumes: existing `logic.js`/`app.js` IDs (`pkg`, `sup`, `fac`, `from`, `to`, `export`, `caption`, `table`, `export-error`, `data-error`) — preserved.
- Produces: `site_template/css/style.css`, `favicon.png`, `pacd.png`, `assets/primark-logo.png` — consumed by `publish.py` in Task 4.

- [ ] **Step 1: Copy the binary assets from the reference project**

```powershell
$ref = "C:\Users\Shoaib\OneDrive - PacD\Projects\Primark\Primark Carton SQM Analysis\Carton-Price-Calculator"
Copy-Item "$ref\favicon.png" "site_template\favicon.png"
Copy-Item "$ref\pacd.png" "site_template\pacd.png"
New-Item -ItemType Directory -Path "site_template\assets" -Force | Out-Null
Copy-Item "$ref\assets\primark-logo.png" "site_template\assets\primark-logo.png"
```

Verify: `Get-ChildItem site_template -Recurse -File | Select-Object FullName, Length` — the three images exist with non-zero sizes.

- [ ] **Step 2: Write `site_template/css/style.css`**

Complete content (PACD theme port with adaptations: 1400px card, navy table header, auto-fit filter grid, multi-select styling, error boxes hidden when empty):

```css
/* Primark Carton Supply Report - PACD Theme (reference: Carton Price Calculator) */
:root {
  --primary: #00205b;
  --primary-hover: #001540;
  --secondary: #e31837;
  --bg-dark: #f8fafc;
  --bg-card: #ffffff;
  --bg-glass: rgba(255, 255, 255, 0.9);
  --border: #e2e8f0;
  --text-main: #0f172a;
  --text-muted: #64748b;
  --shadow-sm: 0 1px 2px 0 rgb(0 0 0 / 0.05);
  --shadow-lg: 0 10px 15px -3px rgb(0 0 0 / 0.1), 0 4px 6px -4px rgb(0 0 0 / 0.05);
}

* { box-sizing: border-box; margin: 0; padding: 0; }

html { scroll-behavior: smooth; }

body {
  font-family: 'Outfit', sans-serif;
  background-color: var(--bg-dark);
  color: var(--text-main);
  min-height: 100vh;
  display: flex;
  flex-direction: column;
}

.bg-decoration {
  position: fixed;
  top: 0; left: 0;
  width: 100%; height: 100%;
  background: radial-gradient(circle at 10% 10%, rgba(227, 24, 55, 0.03) 0%, transparent 40%),
    radial-gradient(circle at 90% 90%, rgba(0, 32, 91, 0.03) 0%, transparent 40%);
  z-index: -1;
  pointer-events: none;
}

.navbar {
  height: 70px;
  background: var(--bg-glass);
  backdrop-filter: blur(12px);
  border-bottom: 1px solid var(--border);
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 2rem;
  z-index: 50;
  box-shadow: var(--shadow-sm);
  flex-shrink: 0;
}

.logo-section {
  display: flex;
  align-items: center;
  gap: 1rem;
}

.logo-img { height: 40px; width: auto; object-fit: contain; }

.navbar h1 {
  font-size: 1.25rem;
  font-weight: 700;
  letter-spacing: -0.01em;
  color: var(--primary);
}

.primark-logo { height: 30px; width: auto; object-fit: contain; }

.main-container {
  flex: 1;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: flex-start;
  padding: 40px 20px 80px;
}

.calculator-card {
  background: var(--bg-card);
  border: 1px solid var(--border);
  border-radius: 1.5rem;
  padding: 2.5rem;
  max-width: 1400px;
  width: 100%;
  box-shadow: var(--shadow-lg);
  animation: slideUp 0.5s cubic-bezier(0.16, 1, 0.3, 1);
}

@keyframes slideUp {
  from { opacity: 0; transform: translateY(20px); }
  to { opacity: 1; transform: translateY(0); }
}

.card-section { margin-bottom: 1.5rem; }

.section-label {
  display: block;
  font-size: 0.75rem;
  font-weight: 700;
  text-transform: uppercase;
  color: var(--secondary);
  margin-bottom: 0.75rem;
  letter-spacing: 0.05em;
}

.filter-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 1rem;
}

.form-group { margin-bottom: 0; position: relative; }

.select-label {
  display: block;
  margin-bottom: 0.5rem;
  font-size: 0.85rem;
  color: var(--text-main);
  font-weight: 500;
}

.custom-select {
  width: 100%;
  background: #ffffff;
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 0.75rem;
  color: var(--text-main);
  font-family: inherit;
  font-size: 0.95rem;
  cursor: pointer;
  box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.02);
  appearance: none;
  background-image: url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' fill='none' viewBox='0 0 20 20'%3e%3cpath stroke='%2364748b' stroke-linecap='round' stroke-linejoin='round' stroke-width='1.5' d='M6 8l4 4 4-4'/%3e%3c/svg%3e");
  background-position: right 0.75rem center;
  background-repeat: no-repeat;
  background-size: 1.5em 1.5em;
  padding-right: 2.5rem;
}

.custom-select:focus {
  outline: none;
  border-color: var(--primary);
  box-shadow: 0 0 0 2px rgba(0, 32, 91, 0.1);
}

.multi-select {
  width: 100%;
  background: #ffffff;
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 0.5rem;
  color: var(--text-main);
  font-family: inherit;
  font-size: 0.95rem;
  box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.02);
}

.multi-select:focus {
  outline: none;
  border-color: var(--primary);
  box-shadow: 0 0 0 2px rgba(0, 32, 91, 0.1);
}

.export-area {
  margin-top: 1.5rem;
  display: flex;
  justify-content: flex-end;
}

.primary-btn {
  background: var(--primary);
  color: white;
  border: none;
  padding: 0 1.5rem;
  height: 50px;
  border-radius: 99px;
  font-weight: 600;
  font-size: 1rem;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 0.5rem;
  transition: all 0.3s;
  box-shadow: 0 4px 12px rgba(0, 32, 91, 0.2);
  font-family: inherit;
}

.primary-btn:hover {
  background: var(--primary-hover);
  transform: translateY(-1px);
  box-shadow: 0 6px 16px rgba(0, 32, 91, 0.25);
}

.primary-btn:disabled {
  background: #94a3b8;
  box-shadow: none;
  cursor: not-allowed;
  transform: none;
}

.row-count {
  font-size: 0.85rem;
  color: var(--text-muted);
  margin-bottom: 0.75rem;
}

#table {
  max-height: 500px;
  overflow: auto;
  border: 1px solid var(--border);
  border-radius: 8px;
}

#table table { width: 100%; border-collapse: collapse; }

#table thead { position: sticky; top: 0; z-index: 1; }

#table th {
  background: var(--primary);
  color: #ffffff;
  padding: 12px 16px;
  text-align: right;
  font-weight: 600;
  font-size: 13px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  white-space: nowrap;
}

#table th:nth-child(-n+4) { text-align: left; }

#table td {
  padding: 10px 16px;
  border-bottom: 1px solid #f1f5f9;
  font-size: 14px;
  text-align: right;
  white-space: nowrap;
}

#table td:nth-child(-n+4) { text-align: left; }

#table td:nth-child(n+5) { font-weight: 600; color: var(--primary); }

#table tbody tr { transition: background-color 0.15s; }
#table tbody tr:hover { background-color: #eff6ff; }

#table p {
  text-align: center;
  color: var(--text-muted);
  padding: 1rem;
  font-style: italic;
}

.error-box {
  background: #fdecea;
  border: 1px solid #f5c6cb;
  color: #8b1e1e;
  padding: 12px 16px;
  border-radius: 8px;
  margin-bottom: 1rem;
  font-size: 0.9rem;
}

#export-error:empty, #data-error:empty { display: none; }

.app-footer {
  display: flex;
  justify-content: center;
  align-items: center;
  gap: 1rem;
  flex-wrap: wrap;
  padding: 1rem 2rem;
  color: var(--text-muted);
  font-size: 0.75rem;
  font-family: inherit;
}

.app-footer .divider { opacity: 0.5; margin: 0 0.5rem; }

.app-footer .developer-link {
  color: var(--primary);
  text-decoration: none;
  font-weight: 500;
  transition: opacity 0.2s;
}

.app-footer .developer-link:hover { opacity: 0.8; text-decoration: underline; }

@media (max-width: 640px) {
  .navbar { padding: 0 0.75rem; gap: 0.5rem; }
  .navbar h1 { font-size: 0.85rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; min-width: 0; }
  .logo-img { height: 22px; }
  .primark-logo { height: 16px; }
  .main-container { padding: 20px 1rem 80px; }
  .calculator-card { padding: 1.5rem; }
  .filter-grid { grid-template-columns: 1fr; }
}
```

- [ ] **Step 3: Rewrite `site_template/index.html`**

Complete content (same IDs and script order; shell + classes; `data-as-of` replaces the old `footer` div):

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Primark Carton Supply Report</title>
  <link rel="icon" type="image/png" href="favicon.png">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="css/style.css">
</head>
<body>
  <div class="bg-decoration"></div>

  <nav class="navbar">
    <div class="logo-section">
      <img src="pacd.png" alt="PACD Logo" class="logo-img">
    </div>
    <h1>Primark Carton Supply Report</h1>
    <img src="assets/primark-logo.png" alt="Primark Logo" class="primark-logo">
  </nav>

  <div class="main-container">
    <div class="calculator-card">
      <div class="card-section">
        <label class="section-label">Filters</label>
        <div class="filter-grid">
          <div class="form-group">
            <label class="select-label" for="pkg">Packaging Supplier</label>
            <select id="pkg" class="multi-select" multiple size="6"></select>
          </div>
          <div class="form-group">
            <label class="select-label" for="sup">Supplier</label>
            <select id="sup" class="multi-select" multiple size="6"></select>
          </div>
          <div class="form-group">
            <label class="select-label" for="fac">Factory</label>
            <select id="fac" class="multi-select" multiple size="6"></select>
          </div>
          <div class="form-group">
            <label class="select-label" for="from">From</label>
            <select id="from" class="custom-select"></select>
          </div>
          <div class="form-group">
            <label class="select-label" for="to">To</label>
            <select id="to" class="custom-select"></select>
          </div>
        </div>
      </div>

      <div class="export-area">
        <button id="export" class="primary-btn" disabled>Export Excel</button>
      </div>
      <p id="data-as-of" class="row-count"></p>
      <p id="caption" class="row-count"></p>
      <div id="table" class="table-container"></div>
      <p id="export-error" class="error-box"></p>
      <p id="data-error" class="error-box"></p>
    </div>
  </div>

  <footer class="app-footer">
    <p>&copy; 2026 PACD. All rights reserved. <span class="divider">|</span> Developed by <a href="https://ev1shoaib.netlify.app" target="_blank" class="developer-link">EV1</a></p>
  </footer>

  <script src="logic.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"></script>
  <script src="app.js"></script>
</body>
</html>
```

- [ ] **Step 4: Update app.js data-as-of target**

In `site_template/app.js`, line 204, change:

```js
      byId("footer").textContent = "Data as of " + manifest.published;
```

to:

```js
      byId("data-as-of").textContent = "Data as of " + manifest.published;
```

Verify no other `byId("footer")` references remain: `Select-String -Path site_template\app.js -Pattern 'footer'` — zero matches.

- [ ] **Step 5: Verify nothing else broke**

Run: `node --check site_template/app.js` and `node --check site_template/logic.js` and `node tests/test_logic.js`
Expected: no output (checks), "All logic tests passed".

- [ ] **Step 6: Serve and eyeball the shell**

```powershell
python -m http.server 8010 --directory site_template
```

Open `http://localhost:8010/` — the page must show: glass navbar (PACD logo left, title, Primark logo right), filters card, pill Export button (disabled), caption + "Data as of …" lines, empty table area, centered footer text, favicon in the tab. (Serve `site_template` here — `docs/site` is regenerated later in Task 5.) Stop the server afterwards.

- [ ] **Step 7: Commit**

```bash
git add site_template/css/style.css site_template/favicon.png site_template/pacd.png site_template/assets/primark-logo.png site_template/index.html site_template/app.js
git commit -m "style: PACD site shell with logos, favicon and theme CSS"
```

---

### Task 4: Publish assets via publish.py

**Files:**
- Modify: `publish.py:1-40`
- Test: `tests/test_publish.py`

**Interfaces:**
- Consumes: `TEMPLATE_DIR = Path(__file__).parent / "site_template"`; the assets created in Task 3.
- Produces: `publish_site(records, out_dir, now)` now also copies `css/style.css`, `favicon.png`, `pacd.png`, `assets/primark-logo.png` from `site_template/` to `out_dir/` (byte-identical). `commit_site` unchanged.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_publish.py`:

```python
def test_publish_copies_assets(tmp_path):
    out = publish_site(make_records(), tmp_path / "site", datetime.datetime(2026, 8, 19, 14, 30))
    for name in ("css/style.css", "favicon.png", "pacd.png", "assets/primark-logo.png"):
        assert (out / name).exists()
        assert (out / name).read_bytes() == (PROJECT_ROOT / "site_template" / name).read_bytes()
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `python -m pytest tests/test_publish.py::test_publish_copies_assets -v`
Expected: FAIL — asset files do not exist in the output dir.

- [ ] **Step 3: Implement asset copying**

In `publish.py`, add `import shutil` at the top and after the existing template loop (line 39):

```python
    for name in ("css/style.css", "favicon.png", "pacd.png", "assets/primark-logo.png"):
        shutil.copyfile(TEMPLATE_DIR / name, out_dir / name)
```

- [ ] **Step 4: Run the test to verify it passes**

Run: `python -m pytest tests/test_publish.py -v`
Expected: PASS — all 7 publish tests green.

- [ ] **Step 5: Run the full suite**

Run: `python -m pytest -q`
Expected: 39 passed.

- [ ] **Step 6: Commit**

```bash
git add publish.py tests/test_publish.py
git commit -m "feat: copy PACD assets when publishing the site"
```

---

### Task 5: Regenerate docs/site and verify

**Files:**
- Modify: `docs/site/*` (regenerated artifact — committed, no hand edits)

**Interfaces:**
- Consumes: `load_records` (loader.py), `publish_site`/`commit_site` (publish.py).
- Produces: published `docs/site/` in the new style, committed via the standard "Updated at …" flow.

- [ ] **Step 1: Regenerate the published site with real data**

```powershell
python -c "from pathlib import Path; import datetime; from loader import load_records; from publish import publish_site, commit_site; records = load_records('Sales Record V1.xlsx')[0]; now = datetime.datetime.now(); publish_site(records, Path('docs/site'), now); print(commit_site(Path('.'), 'docs/site', now))"
```

Expected: prints "Updated at <timestamp>"; `git log --oneline -1` shows that commit.

- [ ] **Step 2: Verify byte-identity of docs/site vs site_template**

```powershell
Get-ChildItem site_template -Recurse -File | ForEach-Object {
  $rel = $_.FullName.Substring((Resolve-Path site_template).Path.Length + 1)
  $other = "docs\site\$rel"
  $same = (Get-FileHash $_.FullName).Hash -eq (Get-FileHash $other).Hash
  "{0} {1}" -f ($(if ($same) { "OK " } else { "DIFF" }), $rel)
}
```

Expected: every line starts with `OK` — all 9 files (5 generated + 4 assets) byte-identical.

- [ ] **Step 3: Run the full test suites**

Run: `python -m pytest -q` then `node tests/test_logic.js`
Expected: 39 passed; "All logic tests passed".

- [ ] **Step 4: Final serve check**

```powershell
python -m http.server 8010 --directory docs
```

Open `http://localhost:8010/site/` — data table renders with navy headers, filters work (pkg → sup → fac), Export Excel produces a downloaded workbook. Stop the server.

- [ ] **Step 5: Commit the plan**

```bash
git add docs/superpowers/plans/2026-08-19-pacd-restyle.md
git commit -m "docs: PACD restyle implementation plan"
```
