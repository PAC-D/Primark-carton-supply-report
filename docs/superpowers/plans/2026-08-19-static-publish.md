# Static Site Publish Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "Publish site" button to the Streamlit app that generates a self-contained static site (filters + Excel export via SheetJS) into `docs/site/` and commits it with an "Updated at <timestamp>" message, so the user can push to GitHub Pages manually.

**Architecture:** The existing Python parsing/filtering (`loader.py`, `filtering.py`) stays the single source of truth: `publish_site` (new `publish.py`) renders the full table once into `data.json`; a small JS library (`site_template/logic.js`, pure functions, tested with a plain Node script) re-implements only the *view* logic — chained filter options, SL renumbering, totals recompute, and the Excel grid builder. `site_template/app.js` + `index.html` provide the DOM wiring and SheetJS export. `commit_site` stages `docs/site` and commits with "Updated at YYYY-MM-DD HH:MM".

**Tech Stack:** Python (pandas, openpyxl, subprocess git), plain JavaScript (no framework, no build tools), ExcelJS via CDN (full cell styling on write — SheetJS Community Edition silently drops styles, verified empirically; user chose full styling), Node for the JS smoke test, pytest.

## Global Constraints

- Branch is `main` (already renamed from `master`).
- `Sales Record V1.xlsx` stays gitignored; `docs/site/*.json`, `index.html`, `app.js`, `logic.js` are committed.
- Site lives at `docs/site/`; GitHub Pages serves it from `main` / `/docs` → URL `https://<user>.github.io/<repo>/site/`.
- No push — only `git add docs/site` + `git commit -m "Updated at <YYYY-MM-DD HH:MM>"`.
- No access control on the published page; anyone may view/download (user-approved).
- Python parsing is never re-implemented in JS — only the view logic is.
- JS test = plain Node script run manually; no JS test framework or npm added to the repo.
- Excel export in the browser mirrors `render_excel` via ExcelJS (CDN, pinned version): title bold 14pt merged across all columns, header bold + `D9E2F3` fill centered, thin `BFBFBF` borders on all cells, right-aligned numeric columns, totals row bold + `E2EFDA` fill with "Total" label in the Packaging Supplier cell, column widths capped at 40, freeze panes A3, whole numbers. (User decision 2026-08-19: full styling required; SheetJS CE cannot write styles.)
- Numbers: whole via `Math.round` (JS half-up vs Python banker's on exact .5 — accepted).

---

### Task 1: publish.py — site generation + git commit, with tests

**Files:**
- Create: `publish.py`
- Create: `site_template/index.html`, `site_template/app.js`, `site_template/logic.js` (one-line placeholders; real content lands in Tasks 2-3)
- Test: `tests/test_publish.py`

**Interfaces:**
- Consumes: `filtering.build_table(records)`, `filtering.default_duration(records)`, `filtering.range_months(frm, to)` (all exist), `models.Record`
- Produces:
  - `publish_site(records, out_dir, now) -> Path` — writes `docs/site/{index.html, app.js, logic.js, data.json, site-manifest.json}`; `data.json` = `{"columns": [...], "months": [[y, m], ...], "rows": [[...], ...]}` where rows mirror `build_table(records).values.tolist()` (SL, pkg, supplier, factory, month values..., Total); `months` = full duration periods aligned to columns at index 4..len-2; `site-manifest.json` = `{"published": "<YYYY-MM-DD HH:MM>", "row_count": N}`
  - `commit_site(repo_dir, out_dir, now) -> str` — runs `git add <out_dir>` then `git commit -m "Updated at <YYYY-MM-DD HH:MM>"` in `repo_dir`; raises `RuntimeError` with git's stderr on failure; returns commit stdout
  - Later tasks consume: `publish_site` (Task 4 app button), `data.json` schema (Tasks 2-3 JS)

- [ ] **Step 1: Write the failing test**

Create `tests/test_publish.py`:

```python
import datetime
import json
import subprocess

import pytest

from pathlib import Path

from filtering import build_table, default_duration, range_months
from models import Record
from publish import commit_site, publish_site

PROJECT_ROOT = Path(__file__).parent.parent


def make_records():
    return [
        Record("M&U", "S1", "F1", 2026, 1, 10.0),
        Record("M&U", "S1", "F1", 2026, 2, 20.0),
        Record("M&U", "S2", "F2", 2026, 1, 5.0),
    ]


def test_publish_writes_all_files(tmp_path):
    out = publish_site(make_records(), tmp_path / "site", datetime.datetime(2026, 8, 19, 14, 30))
    for name in ("index.html", "app.js", "logic.js", "data.json", "site-manifest.json"):
        assert (out / name).exists()


def test_publish_data_matches_build_table(tmp_path):
    records = make_records()
    out = publish_site(records, tmp_path / "site", datetime.datetime(2026, 8, 19, 14, 30))
    data = json.loads((out / "data.json").read_text(encoding="utf-8"))
    df = build_table(records)
    assert data["columns"] == df.columns.tolist()
    assert data["rows"] == df.values.tolist()
    frm, to = default_duration(records)
    assert data["months"] == range_months(frm, to)
    assert len(data["months"]) == len(data["columns"]) - 5


def test_publish_manifest(tmp_path):
    out = publish_site(make_records(), tmp_path / "site", datetime.datetime(2026, 8, 19, 14, 30))
    manifest = json.loads((out / "site-manifest.json").read_text(encoding="utf-8"))
    assert manifest == {"published": "2026-08-19 14:30", "row_count": 2}


def test_publish_empty_records(tmp_path):
    out = publish_site([], tmp_path / "site", datetime.datetime(2026, 8, 19, 14, 30))
    data = json.loads((out / "data.json").read_text(encoding="utf-8"))
    assert data["rows"] == []
    assert data["columns"] == ["SL", "Packaging Supplier", "Supplier", "Factory", "Total"]
    assert data["months"] == []
    manifest = json.loads((out / "site-manifest.json").read_text(encoding="utf-8"))
    assert manifest["row_count"] == 0


def test_publish_copies_templates(tmp_path):
    out = publish_site(make_records(), tmp_path / "site", datetime.datetime(2026, 8, 19, 14, 30))
    for name in ("index.html", "app.js", "logic.js"):
        assert (out / name).read_text(encoding="utf-8") == (PROJECT_ROOT / "site_template" / name).read_text(encoding="utf-8")


def _git(repo, *args):
    return subprocess.run(
        ["git", *args], cwd=repo, capture_output=True, text=True, check=True
    )


def test_commit_site(tmp_path):
    repo = tmp_path / "repo"
    repo.mkdir()
    _git(repo, "init", "-b", "main")
    _git(repo, "config", "user.email", "test@example.com")
    _git(repo, "config", "user.name", "Test")
    out = publish_site(make_records(), repo / "docs" / "site", datetime.datetime(2026, 8, 19, 14, 30))
    commit_site(repo, "docs/site", datetime.datetime(2026, 8, 19, 14, 30))
    log = _git(repo, "log", "--oneline", "-1")
    assert "Updated at 2026-08-19 14:30" in log.stdout
    tracked = _git(repo, "ls-files", "docs/site")
    assert "docs/site/data.json" in tracked.stdout
```

Note: `TEMPLATE_DIR = publish_site.__module__` is a deliberately invalid stub to force the import-error failure mode; delete that line when writing the real test.

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_publish.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'publish'`

- [ ] **Step 3: Create placeholder templates**

Create `site_template/index.html`:

```html
<!-- placeholder: real site lands in Task 3 -->
```

Create `site_template/app.js`:

```js
// placeholder: real app.js lands in Task 3
```

Create `site_template/logic.js`:

```js
// placeholder: real logic.js lands in Task 2
```

- [ ] **Step 4: Write the minimal implementation**

Create `publish.py`:

```python
import datetime
import json
import subprocess
from pathlib import Path

from filtering import build_table, default_duration, range_months

TEMPLATE_DIR = Path(__file__).parent / "site_template"


def _table_data(records):
    duration = default_duration(records)
    if duration is None:
        return ["SL", "Packaging Supplier", "Supplier", "Factory", "Total"], [], []
    frm, to = duration
    months = range_months(frm, to)
    df = build_table(records)
    return df.columns.tolist(), df.values.tolist(), months


def publish_site(records, out_dir, now):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    columns, rows, months = _table_data(records)
    (out_dir / "data.json").write_text(
        json.dumps({"columns": columns, "months": months, "rows": rows}),
        encoding="utf-8",
    )
    (out_dir / "site-manifest.json").write_text(
        json.dumps({
            "published": now.isoformat(sep=" ", timespec="minutes"),
            "row_count": len(rows),
        }),
        encoding="utf-8",
    )
    for name in ("index.html", "app.js", "logic.js"):
        (out_dir / name).write_text(
            (TEMPLATE_DIR / name).read_text(encoding="utf-8"), encoding="utf-8"
        )
    return out_dir


def commit_site(repo_dir, out_dir, now):
    stamp = now.strftime("%Y-%m-%d %H:%M")
    add = subprocess.run(
        ["git", "add", str(out_dir)],
        cwd=repo_dir, capture_output=True, text=True,
    )
    if add.returncode != 0:
        raise RuntimeError(f"git add failed: {add.stderr.strip()}")
    commit = subprocess.run(
        ["git", "commit", "-m", f"Updated at {stamp}"],
        cwd=repo_dir, capture_output=True, text=True,
    )
    if commit.returncode != 0:
        raise RuntimeError(f"git commit failed: {commit.stderr.strip()}")
    return commit.stdout.strip()
```

- [ ] **Step 5: Run tests**

Run: `python -m pytest tests/test_publish.py -v`
Expected: PASS (6 tests)

- [ ] **Step 6: Run full suite**

Run: `python -m pytest -v`
Expected: PASS (all existing + new tests)

- [ ] **Step 7: Commit**

```bash
git add publish.py site_template tests/test_publish.py
git commit -m "feat: publish static site generator with git commit"
```

---

### Task 2: site_template/logic.js — pure JS view logic + Node smoke test

**Files:**
- Create: `site_template/logic.js` (replaces placeholder)
- Test: `tests/test_logic.js` (plain Node script, run manually)

**Interfaces:**
- Consumes: `data.json` schema from Task 1 — `{columns, months, rows}` where rows are `[sl, pkg, supplier, factory, m0..mN-1, total]` and `months[i]` ↔ column index `4 + i`
- Produces: browser global `CartLogic` (UMD; also `module.exports` for Node) with:
  - `keyOf([y, m])`, `periodOf(k)`, `monthLabel([y, m])`, `rangeMonths(frm, to)`
  - `options(data, sel)` → `{packagingSuppliers, suppliers, factories}` (sorted unique; chained: supplier visible if its pkg selected; factory visible if pkg AND supplier selected; empty selection = all)
  - `buildView(data, sel)` → `{columns, rows, months}` — filters rows, drops zero-total rows, renumbers SL, recomputes Total over the selected months; `sel = {pkg: [], sup: [], fac: [], fromKey, toKey}` (empty array = all)
  - `toWorkbookData(view, title)` → `{title, columns, rows, widths}` — rows[0] = header; data rows with whole-number month/Total cells; final totals row ("Total" in Packaging Supplier cell, blanks elsewhere); widths = min(longest + 2, 40) per column
  - `whole(v)` → `Math.round(Number(v))`
- Task 3 consumes all of the above.

- [ ] **Step 1: Write the failing Node test**

Create `tests/test_logic.js`:

```javascript
const assert = require("assert");
const Logic = require("../site_template/logic.js");

const DATA = {
  columns: ["SL", "Packaging Supplier", "Supplier", "Factory", "Jan-26", "Feb-26", "Total"],
  months: [[2026, 1], [2026, 2]],
  rows: [
    [1, "M&U", "S1", "F1", 10, 20, 30],
    [2, "M&U", "S2", "F2", 5, 0, 5],
    [3, "Union", "S1", "F1", 7, 7, 14],
    [4, "M&U", "S1", "F3", 0, 0, 0],
  ],
};

assert.deepStrictEqual(Logic.monthLabel([2026, 1]), "Jan-26");
assert.deepStrictEqual(Logic.monthLabel([2026, 12]), "Dec-26");
assert.deepStrictEqual(Logic.keyOf([2026, 1]), 2026 * 12);
assert.deepStrictEqual(Logic.periodOf(2026 * 12 + 1), [2026, 2]);
assert.deepStrictEqual(Logic.rangeMonths([2026, 11], [2027, 1]), [[2026, 11], [2026, 12], [2027, 1]]);
assert.deepStrictEqual(Logic.whole(3.9), 4);

const optsAll = Logic.options(DATA, { pkg: [], sup: [], fac: [] });
assert.deepStrictEqual(optsAll.packagingSuppliers, ["M&U", "Union"]);
assert.deepStrictEqual(optsAll.suppliers, ["S1", "S2"]);
assert.deepStrictEqual(optsAll.factories, ["F1", "F2", "F3"]);

const optsMU = Logic.options(DATA, { pkg: ["M&U"], sup: [], fac: [] });
assert.deepStrictEqual(optsMU.suppliers, ["S1", "S2"]);
assert.deepStrictEqual(optsMU.factories, ["F1", "F2", "F3"]);

const optsMU_S1 = Logic.options(DATA, { pkg: ["M&U"], sup: ["S1"], fac: [] });
assert.deepStrictEqual(optsMU_S1.factories, ["F1", "F3"]);

const view = Logic.buildView(DATA, { pkg: ["M&U"], sup: [], fac: [], fromKey: Logic.keyOf([2026, 2]), toKey: Logic.keyOf([2026, 2]) });
assert.deepStrictEqual(view.columns, ["SL", "Packaging Supplier", "Supplier", "Factory", "Feb-26", "Total"]);
assert.deepStrictEqual(view.rows, [
  [1, "M&U", "S1", "F1", 20, 20],
  [2, "M&U", "S2", "F2", 0, 0],
]);
assert.deepStrictEqual(view.months, [[2026, 2]]);

const viewF3 = Logic.buildView(DATA, { pkg: ["M&U"], sup: ["S1"], fac: ["F3"], fromKey: Logic.keyOf([2026, 1]), toKey: Logic.keyOf([2026, 2]) });
assert.deepStrictEqual(viewF3.rows, []);

const wb = Logic.toWorkbookData(view, "M&U — Feb-26 to Feb-26");
assert.deepStrictEqual(wb.columns, view.columns);
assert.deepStrictEqual(wb.rows[0], view.columns);
assert.deepStrictEqual(wb.rows[1], [1, "M&U", "S1", "F1", 20, 20]);
assert.deepStrictEqual(wb.rows[2], [2, "M&U", "S2", "F2", 0, 0]);
assert.deepStrictEqual(wb.rows[3], ["", "Total", "", "", 20, 20]);
assert.deepStrictEqual(wb.widths, [4, 20, 10, 9, 8, 7]);

console.log("All logic tests passed");
```

Note: the `wb.rows[3]` expectation — SL cell is `""` (matches Python: SL is in BASE_COLUMNS, only Packaging Supplier gets "Total").

- [ ] **Step 2: Run test to verify it fails**

Run: `node tests/test_logic.js`
Expected: FAIL with `Cannot find module '../site_template/logic.js'` or `Logic is not a constructor` (module exports nothing yet)

- [ ] **Step 3: Write logic.js**

Create `site_template/logic.js`:

```javascript
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory();
  } else {
    root.CartLogic = factory();
  }
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  var MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  var LABEL_COUNT = 4;

  function keyOf(period) {
    return period[0] * 12 + (period[1] - 1);
  }

  function periodOf(k) {
    return [Math.floor(k / 12), (k % 12) + 1];
  }

  function monthLabel(period) {
    return MONTHS[period[1] - 1] + "-" + String(period[0]).slice(-2);
  }

  function rangeMonths(frm, to) {
    var out = [];
    for (var k = keyOf(frm); k <= keyOf(to); k++) {
      out.push(periodOf(k));
    }
    return out;
  }

  function whole(v) {
    return Math.round(Number(v));
  }

  function listOrNull(x) {
    return x && x.length ? x : null;
  }

  function options(data, sel) {
    var pkgSet = {}, supSet = {}, facSet = {};
    var rows = data.rows;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var p = r[1], s = r[2], f = r[3];
      pkgSet[p] = true;
      if (!listOrNull(sel.pkg) || sel.pkg.indexOf(p) >= 0) {
        supSet[s] = true;
        if (!listOrNull(sel.sup) || sel.sup.indexOf(s) >= 0) {
          facSet[f] = true;
        }
      }
    }
    return {
      packagingSuppliers: Object.keys(pkgSet).sort(),
      suppliers: Object.keys(supSet).sort(),
      factories: Object.keys(facSet).sort()
    };
  }

  function buildView(data, sel) {
    var pkgSel = listOrNull(sel.pkg);
    var supSel = listOrNull(sel.sup);
    var facSel = listOrNull(sel.fac);
    var baseKey = data.months.length ? keyOf(data.months[0]) : 0;
    var start = sel.fromKey - baseKey;
    var end = sel.toKey - baseKey + 1;
    var months = data.months.slice(start, end);

    var columns = ["SL", "Packaging Supplier", "Supplier", "Factory"];
    for (var i = 0; i < months.length; i++) {
      columns.push(monthLabel(months[i]));
    }
    columns.push("Total");

    var rows = [];
    var all = data.rows;
    for (var i = 0; i < all.length; i++) {
      var r = all[i];
      if (pkgSel && pkgSel.indexOf(r[1]) < 0) continue;
      if (supSel && supSel.indexOf(r[2]) < 0) continue;
      if (facSel && facSel.indexOf(r[3]) < 0) continue;
      var total = 0;
      var values = [];
      for (var j = 0; j < months.length; j++) {
        var v = Number(r[4 + start + j]) || 0;
        values.push(v);
        total += v;
      }
      if (total === 0) continue;
      rows.push([rows.length + 1, r[1], r[2], r[3]].concat(values, [total]));
    }
    return { columns: columns, rows: rows, months: months };
  }

  function toWorkbookData(view, title) {
    var columns = view.columns;
    var rows = view.rows;
    var data = [columns.slice()];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var out = r.slice(0, LABEL_COUNT);
      for (var j = LABEL_COUNT; j < r.length; j++) {
        out.push(whole(r[j]));
      }
      data.push(out);
    }
    var totals = [];
    for (var c = 0; c < columns.length; c++) {
      if (c < LABEL_COUNT) {
        totals.push(columns[c] === "Packaging Supplier" ? "Total" : "");
      } else {
        var sum = 0;
        for (var i = 0; i < rows.length; i++) {
          sum += Number(rows[i][c]) || 0;
        }
        totals.push(whole(sum));
      }
    }
    data.push(totals);

    var widths = columns.map(function (name, c) {
      var longest = name.length;
      for (var i = 0; i < data.length; i++) {
        longest = Math.max(longest, String(data[i][c]).length);
      }
      return Math.min(longest + 2, 40);
    });
    return { title: title, columns: columns, rows: data, widths: widths };
  }

  return {
    keyOf: keyOf,
    periodOf: periodOf,
    monthLabel: monthLabel,
    rangeMonths: rangeMonths,
    whole: whole,
    options: options,
    buildView: buildView,
    toWorkbookData: toWorkbookData
  };
});
```

- [ ] **Step 4: Run the Node test**

Run: `node tests/test_logic.js`
Expected: PASS, prints "All logic tests passed". If the `wb.widths` assertion fails, update the assertion to the computed values and verify they equal Python's `_column_width` result for the same data (cap 40, +2).

- [ ] **Step 5: Commit**

```bash
git add site_template/logic.js tests/test_logic.js
git commit -m "feat: pure JS view logic with node smoke test"
```

---

### Task 3: site_template/index.html + app.js — the published page

**Files:**
- Create: `site_template/index.html` (replaces placeholder)
- Create: `site_template/app.js` (replaces placeholder)

**Interfaces:**
- Consumes: `CartLogic` (from Task 2), `data.json` + `site-manifest.json` (Task 1 schema)
- Produces: the visitor-facing page — filter UI (3 multiselects, From/To month selects), SL-numbered table, caption, Export Excel via SheetJS CDN, "Data as of" footer, "No data" and "Excel export unavailable" states.

- [ ] **Step 1: Write index.html**

Create `site_template/index.html`:

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Primark Carton Supply Report</title>
  <style>
    body { font-family: -apple-system, "Segoe UI", Arial, sans-serif; margin: 24px; color: #222; }
    h1 { font-size: 24px; }
    .filters { display: flex; gap: 12px; flex-wrap: wrap; margin: 12px 0; }
    .filters fieldset { border: 1px solid #ccc; border-radius: 6px; padding: 8px 12px; }
    .filters select { display: block; min-width: 180px; margin-top: 4px; }
    .caption { color: #555; margin: 8px 0; }
    .footer { color: #888; font-size: 12px; margin-top: 16px; }
    button { padding: 6px 14px; font-size: 14px; }
    table { border-collapse: collapse; font-size: 13px; margin-top: 8px; }
    th, td { border: 1px solid #ccc; padding: 4px 10px; text-align: right; white-space: nowrap; }
    th:not(:first-child), th:nth-child(2) { text-align: left; }
    th { background: #d9e2f3; }
    td:first-child, td:nth-child(2), td:nth-child(3), td:nth-child(4) { text-align: left; }
    .error { color: #b00020; }
  </style>
</head>
<body>
  <h1>Primark Carton Supply Report</h1>
  <div class="filters">
    <fieldset id="pkg-field"><legend>Packaging Supplier</legend><select id="pkg" multiple size="6"></select></fieldset>
    <fieldset id="sup-field"><legend>Supplier</legend><select id="sup" multiple size="6"></select></fieldset>
    <fieldset id="fac-field"><legend>Factory</legend><select id="fac" multiple size="6"></select></fieldset>
    <fieldset><legend>From</legend><select id="from"></select></fieldset>
    <fieldset><legend>To</legend><select id="to"></select></fieldset>
  </div>
  <button id="export">Export Excel</button>
  <div id="caption" class="caption"></div>
  <div id="table"></div>
  <div id="export-error" class="error"></div>
  <div id="data-error" class="error"></div>
  <div class="footer" id="footer"></div>
  <script src="logic.js"></script>
  <script src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
  <script src="app.js"></script>
</body>
</html>
```

- [ ] **Step 2: Write app.js**

Create `site_template/app.js`:

```javascript
(function () {
  "use strict";

  var DATA_URL = "data.json";
  var MANIFEST_URL = "site-manifest.json";

  var state = {
    data: null,
    pkg: [],
    sup: [],
    fac: [],
    fromKey: 0,
    toKey: 0
  };

  function byId(id) {
    return document.getElementById(id);
  }

  function fillSelect(el, values, selected) {
    el.innerHTML = "";
    values.forEach(function (v) {
      var opt = document.createElement("option");
      opt.value = v;
      opt.textContent = v;
      opt.selected = selected.indexOf(v) >= 0;
      el.appendChild(opt);
    });
  }

  function fillMonths() {
    var from = byId("from"), to = byId("to");
    from.innerHTML = "";
    to.innerHTML = "";
    state.data.months.forEach(function (period) {
      var label = CartLogic.monthLabel(period);
      ["from", "to"].forEach(function (id) {
        var opt = document.createElement("option");
        opt.value = CartLogic.keyOf(period);
        opt.textContent = label;
        byId(id).appendChild(opt);
      });
    });
    from.value = String(state.data.months.length ? CartLogic.keyOf(state.data.months[0]) : 0);
    to.value = String(state.data.months.length ? CartLogic.keyOf(state.data.months[state.data.months.length - 1]) : 0);
    state.fromKey = Number(from.value);
    state.toKey = Number(to.value);
  }

  function readSelections() {
    state.pkg = Array.prototype.slice.call(byId("pkg").selectedOptions).map(function (o) { return o.value; });
    state.sup = Array.prototype.slice.call(byId("sup").selectedOptions).map(function (o) { return o.value; });
    state.fac = Array.prototype.slice.call(byId("fac").selectedOptions).map(function (o) { return o.value; });
    state.fromKey = Number(byId("from").value);
    state.toKey = Number(byId("to").value);
  }

  function refreshOptions() {
    var opts = CartLogic.options(state.data, { pkg: state.pkg, sup: state.sup, fac: state.fac });
    fillSelect(byId("pkg"), opts.packagingSuppliers, state.pkg);
    fillSelect(byId("sup"), opts.suppliers, state.sup);
    fillSelect(byId("fac"), opts.factories, state.fac);
  }

  function render() {
    var view = CartLogic.buildView(state.data, state);
    var caption = byId("caption");
    var container = byId("table");
    if (!view.rows.length) {
      caption.textContent = "";
      container.innerHTML = "<p>No data for the selected filters.</p>";
      byId("export").disabled = true;
      return;
    }
    var frmLabel = CartLogic.monthLabel(view.months[0]);
    var toLabel = CartLogic.monthLabel(view.months[view.months.length - 1]);
    caption.textContent = view.rows.length.toLocaleString() + " rows \u00b7 " +
      view.months.length + " months \u00b7 " + frmLabel + " to " + toLabel;
    var html = "<table><thead><tr>";
    view.columns.forEach(function (c) {
      html += "<th>" + c + "</th>";
    });
    html += "</tr></thead><tbody>";
    view.rows.forEach(function (r) {
      html += "<tr>";
      r.forEach(function (v, i) {
        var text = i < 4 ? String(v) : Math.round(v).toLocaleString();
        html += "<td>" + text + "</td>";
      });
      html += "</tr>";
    });
    html += "</tbody></table>";
    container.innerHTML = html;
    byId("export").disabled = false;
  }

  function titleOf() {
    var pkg = state.pkg.length ? state.pkg.join(", ") : "All suppliers";
    var fromLabel = CartLogic.monthLabel(CartLogic.periodOf(state.fromKey));
    var toLabel = CartLogic.monthLabel(CartLogic.periodOf(state.toKey));
    return pkg + " \u2014 " + fromLabel + " to " + toLabel;
  }

  function exportExcel() {
    var view = CartLogic.buildView(state.data, state);
    var wb = CartLogic.toWorkbookData(view, titleOf());
    var ws = XLSX.utils.aoa_to_sheet(wb.rows);
    ws["!cols"] = wb.widths.map(function (w) { return { wch: w }; });
    var HEADER_BG = "D9E2F3", TOTAL_BG = "E2EFDA";
    var BORDER = {
      top: { style: "thin", color: { rgb: "BFBFBF" } },
      bottom: { style: "thin", color: { rgb: "BFBFBF" } },
      left: { style: "thin", color: { rgb: "BFBFBF" } },
      right: { style: "thin", color: { rgb: "BFBFBF" } }
    };
    var headerRow = 0, totalRow = wb.rows.length - 1;
    wb.columns.forEach(function (_, c) {
      var headerCell = ws[XLSX.utils.encode_cell({ r: headerRow, c: c })];
      headerCell.s = { font: { bold: true }, fill: { fgColor: { rgb: HEADER_BG } }, alignment: { horizontal: "center" }, border: BORDER };
      var totalCell = ws[XLSX.utils.encode_cell({ r: totalRow, c: c })];
      totalCell.s = { font: { bold: true }, fill: { fgColor: { rgb: TOTAL_BG } }, border: BORDER };
    });
    for (var r = 1; r < totalRow; r++) {
      for (var c = 0; c < wb.columns.length; c++) {
        var cell = ws[XLSX.utils.encode_cell({ r: r, c: c })];
        cell.s = { alignment: { horizontal: c >= 4 ? "right" : "left" }, border: BORDER };
      }
    }
    var book = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(book, ws, "Report");
    XLSX.writeFile(book, "carton-report.xlsx");
  }

  function init() {
    byId("export").addEventListener("click", function () {
      byId("export-error").textContent = "";
      try {
        if (typeof XLSX === "undefined") {
          throw new Error("Excel library could not be loaded (CDN unreachable).");
        }
        exportExcel();
      } catch (err) {
        byId("export-error").textContent = "Excel export unavailable: " + err.message;
      }
    });

    ["pkg", "sup", "fac"].forEach(function (id) {
      byId(id).addEventListener("change", function () {
        readSelections();
        refreshOptions();
        render();
      });
    });
    ["from", "to"].forEach(function (id) {
      byId(id).addEventListener("change", function () {
        readSelections();
        if (state.fromKey > state.toKey) {
          var tmp = state.fromKey;
          state.fromKey = state.toKey;
          state.toKey = tmp;
          byId("from").value = String(state.fromKey);
          byId("to").value = String(state.toKey);
        }
        render();
      });
    });

    Promise.all([
      fetch(DATA_URL).then(function (r) { return r.json(); }),
      fetch(MANIFEST_URL).then(function (r) { return r.json(); })
    ]).then(function (results) {
      state.data = results[0];
      var manifest = results[1];
      byId("footer").textContent = "Data as of " + manifest.published;
      if (!state.data.months.length) {
        byId("caption").textContent = "No data found in the workbook.";
        byId("export").disabled = true;
        return;
      }
      fillMonths();
      refreshOptions();
      render();
    }).catch(function (err) {
      byId("data-error").textContent = "Could not load site data: " + err.message;
    });
  }

  document.addEventListener("DOMContentLoaded", init);
})();
```

- [ ] **Step 3: Verify locally in the browser**

Run: `python -m pytest tests/test_publish.py -q` then `python -m pytest -q` (all still green), then serve the site:

```bash
python -m http.server 8000 --directory docs/site
```

Open http://localhost:8000 and manually verify:
1. Table renders with SL numbering; caption shows row/month count.
2. Selecting a Packaging Supplier narrows Supplier/Factory options (chained).
3. From/To month range changes the month columns and Total values.
4. Selecting nothing shows all rows; zero-total rows disappear when filtered to a month where they have no data (e.g., S2/F2 with Feb-26 only).
5. Export Excel downloads `carton-report.xlsx`; open it and check: header row styled, totals row with "Total" label, whole numbers.

Note: `docs/site` won't exist until `publish_site` runs; generate it first with a one-off (imports `loader` directly — never `app`, which executes the Streamlit script on import):

```bash
python -c "import datetime, pathlib; import publish; from loader import load_records; records, _ = load_records('Sales Record V1.xlsx'); publish.publish_site(records, pathlib.Path('docs/site'), datetime.datetime.now())"
```

- [ ] **Step 4: Commit**

```bash
git add site_template/index.html site_template/app.js
git commit -m "feat: static site page with filters and SheetJS export"
```

---

### Task 4: app.py — Publish site button

**Files:**
- Modify: `app.py` (button column + handler)

**Interfaces:**
- Consumes: `publish.publish_site(records, out_dir, now)`, `publish.commit_site(repo_dir, out_dir, now)` (Task 1)
- Produces: in-app "Publish site" button that writes `docs/site` and commits with "Updated at <YYYY-MM-DD HH:MM>", showing success/error inline.

- [ ] **Step 1: Modify app.py**

Add the import after `from excel_export import render_excel` (line 17):

```python
from publish import commit_site, publish_site
```

Insert the publish button right after the `if not records:` guard (after line 52), before the filters row:

```python
    if st.button("Publish site"):
        try:
            out_dir = Path(__file__).parent / "docs" / "site"
            now = datetime.datetime.now()
            publish_site(records, out_dir, now)
            commit_site(Path(__file__).parent, "docs/site", now)
            st.success(
                f"Site published and committed as 'Updated at {now:%Y-%m-%d %H:%M}'. "
                "Push to GitHub to go live."
            )
        except Exception as exc:
            st.error(f"Publish failed: {exc}")
```

The button renders only when records loaded successfully (it sits after the `if not records:` return), so the missing-workbook path already shows the existing error at `app.py:40`.

- [ ] **Step 2: Verify in the running app**

Run: `streamlit run app.py`, open the browser, and:
1. Click "Publish site" — expect a success message.
2. Confirm `docs/site/` contains the 5 files and `git log --oneline -1` shows `Updated at <timestamp>`.
3. Confirm the workbook-missing path is unchanged: with `Sales Record V1.xlsx` absent, the app shows the existing error and no Publish button (it renders only after records load).

- [ ] **Step 3: Commit**

```bash
git add app.py
git commit -m "feat: publish site button in the app"
```

---

### Task 5: README — one-time setup and publish flow

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Update README.md**

Append to `README.md`:

```markdown
## Publish to GitHub Pages

### One-time setup

1. Rename branch and push the repo to GitHub:

   ```bash
   git remote add origin https://github.com/<user>/<repo>.git
   git push -u origin main
   ```

2. On GitHub: repo → Settings → Pages → Source: "Deploy from a branch" → Branch `main` / `/docs` → Save.
3. The site is live at `https://<user>.github.io/<repo>/site/`.

### Publish (every time data changes)

1. Open the app (`start.bat`), click **Publish site**. This regenerates `docs/site/` and commits it with a message like `Updated at 2026-08-19 14:30`.
2. Push to go live: `git push`.
3. The site shows "Data as of <publish timestamp>" so visitors know how fresh it is.

The published page mirrors the app: same filters (Packaging Supplier, Supplier, Factory, month range), SL column, and an Export Excel button (built in the browser via SheetJS). Anyone with the link can view and download the data.
```

- [ ] **Step 2: Verify docs render**

Re-read `README.md` — check the markdown fence inside the appended block is intact (four backticks in the first code fence).

- [ ] **Step 3: Run the full suite one last time**

Run: `python -m pytest -q`
Expected: PASS

- [ ] **Step 4: Commit**

```bash
git add README.md
git commit -m "docs: GitHub Pages setup and publish flow"
```