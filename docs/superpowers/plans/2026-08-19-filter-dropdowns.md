# Dropdown Filters with Multi-Select and Select All — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the three native `<select multiple size="6">` list boxes on the static site with PACD-styled dropdown widgets that support multiple selection and a "Select all" option, mirroring the Streamlit app's empty-selection-means-All filtering model.

**Architecture:** Three units — (1) pure filter helpers in `site_template/logic.js` (Node-testable), (2) static markup + CSS in `site_template/index.html` and `site_template/css/style.css`, (3) a DOM-only widget factory in `site_template/app.js` that drives the three filters and chains them through the existing `CartLogic.options`/`buildView` pipeline. The widget keeps the selection as an array; empty array = "All" = no filter (exactly Streamlit's `placeholder="All"`). `publish.py` needs no changes; `docs/site/` is regenerated at the end.

**Tech Stack:** Vanilla ES5 JavaScript (IIFE, `var`, no build step), CSS, HTML. No new runtime dependencies (only existing ExcelJS 4.4.0 CDN). Tests: Node `assert` script `tests/test_logic.js`, pytest suite (unchanged, must stay green).

## Global Constraints

- ES5 only: `var`, `function` declarations, no arrow functions, no template literals, no `let`/`const` — match `app.js`/`logic.js` style exactly.
- No new files in `site_template/` other than edits to `index.html`, `app.js`, `logic.js`, `css/style.css`. No dependency additions.
- Empty selection array means "All" (no filter) — never pass `null` for pkg/sup/fac; `CartLogic.options`/`buildView` already accept empty arrays.
- All three filters get identical widgets; From/To selects are untouched.
- Line endings: commit text files with LF via git autocrlf (existing repo behavior); do not edit binary assets.
- Every commit must leave `python -m pytest -q` (39 passed) and `node tests/test_logic.js` ("All logic tests passed") green, except where a step explicitly runs them mid-task.
- PACD palette variables: `--primary` = `#00205B` navy, `--border` used by `.custom-select`.

---

### Task 1: Pure filter helpers in logic.js

**Files:**
- Modify: `site_template/logic.js` (add two functions before the `return` statement at line 133; add both to the returned object)
- Test: `tests/test_logic.js` (append assertions before the final `console.log` line 60)

**Interfaces:**
- Consumes: nothing new (existing `logic.js` internals only)
- Produces:
  - `CartLogic.selectionSummary(selection)` — takes an array of selected values; returns `"All"` when the array is empty (or falsy), otherwise the string `"N selected"` where N is the length.
  - `CartLogic.pruneSelection(selection, available)` — takes two string arrays; returns a new array of `selection`'s values (order preserved) that are present in `available`.

- [ ] **Step 1: Write the failing tests**

Append to `tests/test_logic.js` before the final `console.log`:

```js
assert.deepStrictEqual(Logic.selectionSummary([]), "All");
assert.deepStrictEqual(Logic.selectionSummary(null), "All");
assert.deepStrictEqual(Logic.selectionSummary(["S1"]), "1 selected");
assert.deepStrictEqual(Logic.selectionSummary(["S1", "S2"]), "2 selected");
assert.deepStrictEqual(Logic.pruneSelection(["S1", "S2", "S9"], ["S1", "S2", "S3"]), ["S1", "S2"]);
assert.deepStrictEqual(Logic.pruneSelection([], ["S1", "S2"]), []);
assert.deepStrictEqual(Logic.pruneSelection(["S1"], ["S2"]), []);
```

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/test_logic.js`
Expected: FAIL — `TypeError: Logic.selectionSummary is not a function`

- [ ] **Step 3: Implement the helpers**

Add above the `return` in `site_template/logic.js` (near `toWorkbookData`, after line 131):

```js
  function selectionSummary(selection) {
    return selection && selection.length ? selection.length + " selected" : "All";
  }

  function pruneSelection(selection, available) {
    var out = [];
    for (var i = 0; i < selection.length; i++) {
      if (available.indexOf(selection[i]) >= 0) {
        out.push(selection[i]);
      }
    }
    return out;
  }
```

Add both names to the returned object (after `toWorkbookData: toWorkbookData,`):

```js
    selectionSummary: selectionSummary,
    pruneSelection: pruneSelection,
```

- [ ] **Step 4: Run test to verify it passes**

Run: `node tests/test_logic.js`
Expected: `All logic tests passed`

- [ ] **Step 5: Commit**

```bash
git add site_template/logic.js tests/test_logic.js
git commit -m "feat: add filter selection summary and prune helpers to logic"
```

---

### Task 2: Widget markup and styling

**Files:**
- Modify: `site_template/index.html:29-40` (the three multi-select form-groups)
- Modify: `site_template/css/style.css` (replace `.multi-select` rules at lines 148-164 with widget styles; keep the other rules)

**Interfaces:**
- Consumes: nothing at runtime yet (widgets are inert until Task 3 wires them)
- Produces (DOM contract Task 3 depends on, per filter key `pkg` / `sup` / `fac`):
  - Container `<div class="form-group filter-widget" id="<key>-widget">`
  - `<button type="button" class="filter-trigger" aria-haspopup="listbox" aria-expanded="false">` containing `<span class="filter-summary">All</span>`
  - `<div class="filter-panel" role="listbox" hidden>` containing `<label class="filter-all"><input type="checkbox" checked> Select all</label>` and `<div class="filter-options"></div>`
  - The label element keeps class `select-label` with the same text ("Packaging Supplier", "Supplier", "Factory"), with no `for` attribute.

- [ ] **Step 1: Replace the three multi-select form-groups**

In `site_template/index.html`, replace the `pkg` form-group (lines 29-32) with:

```html
          <div class="form-group filter-widget" id="pkg-widget">
            <label class="select-label">Packaging Supplier</label>
            <button type="button" class="filter-trigger" aria-haspopup="listbox" aria-expanded="false">
              <span class="filter-summary">All</span>
            </button>
            <div class="filter-panel" role="listbox" hidden>
              <label class="filter-all"><input type="checkbox" checked> Select all</label>
              <div class="filter-options"></div>
            </div>
          </div>
```

Replace the `sup` form-group (lines 33-36) with the same block using `id="sup-widget"`, label "Supplier"; replace the `fac` form-group (lines 37-40) with `id="fac-widget"`, label "Factory". Leave the `from`/`to` form-groups (lines 41-48) untouched.

- [ ] **Step 2: Replace the .multi-select CSS with widget styles**

In `site_template/css/style.css`, delete the `.multi-select` and `.multi-select:focus` rules (lines 148-164) and insert in their place:

```css
.filter-widget { position: relative; }

.filter-trigger {
  width: 100%;
  background: #ffffff;
  border: 1px solid var(--border);
  border-radius: 8px;
  padding: 0.75rem 2.5rem 0.75rem 0.75rem;
  color: var(--text-main);
  font-family: inherit;
  font-size: 0.95rem;
  text-align: left;
  cursor: pointer;
  box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.02);
  background-image: url("data:image/svg+xml,%3csvg xmlns='http://www.w3.org/2000/svg' fill='none' viewBox='0 0 20 20'%3e%3cpath stroke='%2364748b' stroke-linecap='round' stroke-linejoin='round' stroke-width='1.5' d='M6 8l4 4 4-4'/%3e%3c/svg%3e");
  background-position: right 0.75rem center;
  background-repeat: no-repeat;
  background-size: 1.5em 1.5em;
}

.filter-trigger:focus,
.filter-trigger[aria-expanded="true"] {
  outline: none;
  border-color: var(--primary);
  box-shadow: 0 0 0 2px rgba(0, 32, 91, 0.1);
}

.filter-summary {
  display: block;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.filter-panel {
  position: absolute;
  top: calc(100% + 4px);
  left: 0;
  right: 0;
  background: #ffffff;
  border: 1px solid var(--border);
  border-radius: 8px;
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.12);
  max-height: 260px;
  overflow-y: auto;
  z-index: 100;
  padding: 0.25rem 0;
}

.filter-all {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  padding: 0.55rem 0.75rem;
  font-weight: 600;
  color: var(--primary);
  border-bottom: 1px solid var(--border);
  cursor: pointer;
}

.filter-all:hover { background: rgba(0, 32, 91, 0.05); }

.filter-option {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  padding: 0.55rem 0.75rem;
  font-size: 0.9rem;
  cursor: pointer;
}

.filter-option:hover { background: rgba(0, 32, 91, 0.05); }

.filter-all input,
.filter-option input { accent-color: var(--primary); cursor: pointer; }
```

- [ ] **Step 3: Serve-check the markup renders**

Start the server in a second terminal (the command blocks):
```powershell
python -m http.server 8010 --directory docs
```
Open `http://localhost:8010/site/` — expected (pre-Task 3, widgets are inert):
- Three buttons styled like the From/To selects with chevrons and "All" text
- Panels stay hidden (no interaction yet)
- No styling regressions on From/To/Export/table (the server is already running it; refresh)

- [ ] **Step 4: Verify suites still green**

Run: `python -m pytest -q` and `node tests/test_logic.js`
Expected: `39 passed`, `All logic tests passed`

- [ ] **Step 5: Commit**

```bash
git add site_template/index.html site_template/css/style.css
git commit -m "style: filter widget markup and dropdown panel styles"
```

---

### Task 3: Widget factory and app wiring

**Files:**
- Modify: `site_template/app.js` — replace `fillSelect` (lines 20-29), `readSelections` (lines 50-56), `refreshOptions` (lines 58-66), the pkg/sup/fac listener block (lines 178-184); add `makeFilterWidget`; add `var widgets = {};` near the `state` declaration (line 7-14); create widgets in `init` before the `Promise.all` (before line 199)

**Interfaces:**
- Consumes: DOM contract from Task 2 (`#<key>-widget` structure); `CartLogic.selectionSummary`, `CartLogic.pruneSelection` from Task 1; existing `CartLogic.options`, `CartLogic.buildView`
- Produces: nothing new for later tasks (Task 4 only consumes the committed files)

- [ ] **Step 1: Add the widget factory**

Insert `makeFilterWidget` after `byId` (after line 18) in `site_template/app.js`:

```js
  function makeFilterWidget(optionKey) {
    var root = byId(optionKey + "-widget");
    var trigger = root.querySelector(".filter-trigger");
    var summary = root.querySelector(".filter-summary");
    var panel = root.querySelector(".filter-panel");
    var optionsBox = root.querySelector(".filter-options");
    var allInput = root.querySelector(".filter-all input");
    var optionValues = [];
    var selection = [];
    var onChange = null;

    function updateSummary() {
      summary.textContent = CartLogic.selectionSummary(selection);
    }

    function syncAllCheckbox() {
      allInput.checked = selection.length === 0;
    }

    function rebuildRows() {
      optionsBox.innerHTML = "";
      optionValues.forEach(function (v) {
        var row = document.createElement("label");
        row.className = "filter-option";
        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.checked = selection.indexOf(v) >= 0;
        row.appendChild(cb);
        row.appendChild(document.createTextNode(v));
        cb.addEventListener("change", function () {
          var idx = selection.indexOf(v);
          if (cb.checked && idx < 0) { selection.push(v); }
          if (!cb.checked && idx >= 0) { selection.splice(idx, 1); }
          syncAllCheckbox();
          updateSummary();
          if (onChange) { onChange(selection.slice()); }
        });
        optionsBox.appendChild(row);
      });
    }

    function setOptions(values) {
      optionValues = values.slice();
      rebuildRows();
    }

    function setSelection(values) {
      selection = values.slice();
      syncAllCheckbox();
      updateSummary();
      rebuildRows();
    }

    function getSelection() {
      return selection.slice();
    }

    function setOpen(open) {
      panel.hidden = !open;
      trigger.setAttribute("aria-expanded", open ? "true" : "false");
    }

    trigger.addEventListener("click", function () {
      setOpen(panel.hidden);
    });

    allInput.addEventListener("change", function () {
      if (allInput.checked) {
        selection = [];
      } else {
        selection = optionValues.slice();
      }
      syncAllCheckbox();
      updateSummary();
      rebuildRows();
      if (onChange) { onChange(selection.slice()); }
    });

    document.addEventListener("click", function (e) {
      if (!root.contains(e.target)) { setOpen(false); }
    });

    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && !panel.hidden) { setOpen(false); }
    });

    return {
      setOptions: setOptions,
      setSelection: setSelection,
      getSelection: getSelection,
      onChange: function (fn) { onChange = fn; }
    };
  }
```

- [ ] **Step 2: Wire the widgets in**

Replace `fillSelect` (lines 20-29) with nothing (delete the function). Add `var widgets = {};` to the `state` declaration block (line 7-14), e.g. after `var state = { ... };` add a new line `var widgets = {};`.

Replace `readSelections` (lines 50-56) with:

```js
  function readSelections() {
    state.pkg = widgets.pkg.getSelection();
    state.sup = widgets.sup.getSelection();
    state.fac = widgets.fac.getSelection();
    state.fromKey = Number(byId("from").value);
    state.toKey = Number(byId("to").value);
  }
```

Replace `refreshOptions` (lines 58-66) with:

```js
  function refreshOptions() {
    var opts = CartLogic.options(state.data, { pkg: state.pkg, sup: state.sup, fac: state.fac });
    state.sup = CartLogic.pruneSelection(state.sup, opts.suppliers);
    opts = CartLogic.options(state.data, { pkg: state.pkg, sup: state.sup, fac: state.fac });
    state.fac = CartLogic.pruneSelection(state.fac, opts.factories);
    widgets.pkg.setOptions(opts.packagingSuppliers);
    widgets.sup.setOptions(opts.suppliers);
    widgets.fac.setOptions(opts.factories);
  }
```

Replace the pkg/sup/fac listener block (lines 178-184) with:

```js
    widgets.pkg = makeFilterWidget("pkg");
    widgets.sup = makeFilterWidget("sup");
    widgets.fac = makeFilterWidget("fac");
    ["pkg", "sup", "fac"].forEach(function (key) {
      widgets[key].onChange(function () {
        readSelections();
        refreshOptions();
        render();
      });
    });
```

Note: the widgets must be created inside `init()` before the `Promise.all` call (place the block where the old listener block was). The From/To listener block (lines 185-197) is untouched.

- [ ] **Step 3: Syntax-check both JS files**

Run: `node --check site_template/app.js` and `node --check site_template/logic.js`
Expected: no output, exit code 0

- [ ] **Step 4: Run both test suites**

Run: `node tests/test_logic.js` and `python -m pytest -q`
Expected: `All logic tests passed`, `39 passed`

- [ ] **Step 5: Manual behavior check**

With the server from Task 2 still running in its second terminal (`python -m http.server 8010 --directory docs`), open `http://localhost:8010/site/` and verify:
1. All three dropdowns show "All" when closed
2. Clicking a trigger opens its panel; opening another closes the first; click-outside and Esc close
3. Checking a specific option changes the summary to "N selected" and unchecks Select all
4. Unchecking the last option re-checks Select all and the summary returns to "All"
5. Select-all unchecked state selects every option (summary "N selected") and filters nothing (table rows unchanged from "All")
6. Selecting one Packaging Supplier narrows the Supplier options; a previously checked Supplier missing from the new list disappears and its count drops
7. Table, caption, export button, From/To swap all still work (compare with the current published behavior)

- [ ] **Step 6: Commit**

```bash
git add site_template/app.js
git commit -m "feat: dropdown filter widgets with select all for site filters"
```

---

### Task 4: Regenerate the published site

**Files:**
- Modify: `docs/site/` (regenerated output — `index.html`, `app.js`, `logic.js`, `css/style.css`; assets unchanged)

**Interfaces:**
- Consumes: final `site_template/` from Tasks 1-3; real data source `Sales Record V1.xlsx` (424 rows) via the existing `publish_site` path
- Produces: committed `docs/site/` matching `site_template/`

- [ ] **Step 1: Regenerate from real data**

Run one-off publish from the project root using the existing machinery (replicating `app.py`'s Publish site button). PowerShell has no heredoc, so write a temp script, run it, then delete it:

```powershell
Set-Content -LiteralPath "$env:TEMP\publish_now.py" -Value @'
from pathlib import Path
import datetime
from load_data import load_records
from publish import publish_site, commit_site

root = Path(".")
out_dir = root / "docs" / "site"
records = load_records("Sales Record V1.xlsx")[0]
now = datetime.datetime.now()
publish_site(records, out_dir, now)
commit_site(root, "docs/site", now)
print("published", len(records), "rows at", now.isoformat())
'@
python "$env:TEMP\publish_now.py"; if ($?) { Remove-Item -LiteralPath "$env:TEMP\publish_now.py" }
```

Expected: prints `published 424 rows at <now>`; creates commit `Updated at <now>`.

- [ ] **Step 2: Verify byte-identity of template files**

Run (PowerShell):

```powershell
$pairs = @("index.html","app.js","logic.js","css/style.css","favicon.png","pacd.png","assets/primark-logo.png")
$mismatch = @()
foreach ($f in $pairs) {
  $a = [System.IO.File]::ReadAllBytes("site_template\$f")
  $b = [System.IO.File]::ReadAllBytes("docs\site\$f")
  if (-not ($a.SequenceEqual($b))) { $mismatch += $f }
}
if ($mismatch.Count) { "MISMATCH: $($mismatch -join ', ')" } else { "ALL 7 BYTE-IDENTICAL" }
```

Expected: `ALL 7 BYTE-IDENTICAL` — if any mismatch, fix the source of drift (likely CRLF from a hand edit; write the file with LF) and re-run Step 1.

- [ ] **Step 3: Verify suites and serve check**

Run: `python -m pytest -q`, `node tests/test_logic.js`, `node --check site_template/app.js`, `node --check site_template/logic.js`
Expected: `39 passed`, `All logic tests passed`, no syntax output.
Refresh `http://localhost:8010/site/` — the filter widgets behave per Task 3 Step 5 list.

- [ ] **Step 4: Commit**

`commit_site` in Step 1 already created the `Updated at ...` commit. Verify with `git log --oneline -1`; if no such commit exists (e.g. nothing to commit), create it:

```powershell
git add docs/site
git commit -m "Updated at $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
```

- [ ] **Step 5: Final green check**

Run: `git status --short` (clean or only the two commits listed), `python -m pytest -q`, `node tests/test_logic.js`
Expected: `39 passed`, `All logic tests passed`
