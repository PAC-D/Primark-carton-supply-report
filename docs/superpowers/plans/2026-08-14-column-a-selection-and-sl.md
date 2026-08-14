# Column A Row Selection + SL Column Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the loader's row-selection heuristics with a single "numeric column A" rule and add a fresh 1..N SL index column to the web table and PDF.

**Architecture:** The loader (`loader.py`) becomes the only place row selection happens: a data-sheet row is kept iff its column A cell is numeric. `filtering.py` prepends an `SL` column numbered 1..N on the final table; `pdf.py` treats SL like the other label columns. `app.py` needs no change.

**Tech Stack:** Python 3.14, openpyxl, pandas, reportlab, pytest (all already installed; no new dependencies).

## Global Constraints

- Row selection rule (verbatim from spec): keep a row iff column A holds a numeric value (Python `int`/`float`, excluding `bool`); drop otherwise. No name-based checks remain.
- Known approved consequence: 11 "Total Fashion PJT" subtotal rows have numeric column A and ARE now included (e.g., PARKSCENE figures double). Do not "fix" this.
- SL is a fresh output index 1..N in the final sorted/filtered table — never taken from the workbook.
- Column order is `SL | Packaging Supplier | Supplier | Factory | <months> | Total`.
- SL is blank in the PDF totals row; it does not participate in sums.
- Keep every other behavior: Re-sheet replacement, summary-sheet skip, row-2 month headers, Epyllion label scan, zero-total row hiding, multi-select filters, 24-month cap, latest-24 default.
- All tests must pass after each commit: run `python -m pytest -q` from the project root.
- Commit style: `feat:`, `fix:`, or `docs:` prefix; one commit per task.

---

### Task 1: Numeric column A row selection

**Files:**
- Modify: `loader.py:43-49` (`_parse_sheet` data-row loop)
- Test: `tests/test_loader.py`

**Interfaces:**
- Consumes: nothing new.
- Produces: unchanged `load_records(path) -> (list[Record], list[str])`. `Record` dataclass is NOT changed (SL is not stored).

- [ ] **Step 1: Update and add the failing tests**

In `tests/test_loader.py`, replace `test_total_row_excluded` (the fixture gives the "Total" row a numeric column A, so under the new rule it is now INCLUDED — this is the approved consequence):

```python
def test_total_fashion_pjt_row_included(tmp_path):
    # Approved consequence: subtotal rows like "Total Fashion PJT" carry a
    # numeric column A and are no longer dropped by name checks.
    path = make_workbook(tmp_path, {
        "M&U - 2026": [
            ("A", "Factory One", {(2026, 1): 100}),
            ("A", "Total", {(2026, 1): 9000}),
        ],
    })
    records, _ = load_records(path)
    assert len(records) == 2
```

Add these tests (custom sheets so we control column A):

```python
def test_blank_column_a_row_dropped(tmp_path):
    import datetime
    import openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("M&U - 2026")
    ws.append(["Sl", "Suppliers", "Factory ", None, None])
    ws.cell(row=2, column=4, value=datetime.datetime(2026, 1, 1))
    ws.cell(row=3, column=1, value=1)
    ws.cell(row=3, column=2, value="A")
    ws.cell(row=3, column=3, value="Factory One")
    ws.cell(row=3, column=4, value=100)
    # blank column A -> dropped
    ws.cell(row=4, column=2, value="B")
    ws.cell(row=4, column=3, value="Factory Two")
    ws.cell(row=4, column=4, value=200)
    path = tmp_path / "test.xlsx"
    wb.save(path)
    records, _ = load_records(path)
    assert len(records) == 1
    assert records[0].factory == "Factory One"


def test_text_column_a_row_dropped(tmp_path):
    path = make_workbook(tmp_path, {
        "M&U - 2026": [
            ("A", "Factory One", {(2026, 1): 100}),
        ],
    })
    import openpyxl
    wb = openpyxl.load_workbook(path)
    ws = wb["M&U - 2026"]
    ws.cell(row=3, column=1, value="Sl")
    wb.save(path)
    records, _ = load_records(path)
    assert len(records) == 0
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_loader.py -q`
Expected: FAIL — `test_total_fashion_pjt_row_included` gets 1 record (old code drops "Total"), `test_blank_column_a_row_dropped` gets 2, `test_text_column_a_row_dropped` gets 1.

- [ ] **Step 3: Implement the numeric column A rule**

In `loader.py` `_parse_sheet`, replace the loop body's selection logic (lines ~43-49):

```python
    records = []
    for r in range(3, ws.max_row + 1):
        sl = ws.cell(row=r, column=1).value
        if not isinstance(sl, (int, float)) or isinstance(sl, bool):
            continue
        supplier = str(ws.cell(row=r, column=supplier_col).value or "").strip()
        factory = str(ws.cell(row=r, column=factory_col).value or "").strip()
        row_values = [c.value for c in next(ws.iter_rows(min_row=r, max_row=r))]
        extent = max((i + 1 for i, v in enumerate(row_values) if v is not None), default=0)
        for c, year, month in month_cols:
            if c > extent:
                continue
            records.append(Record(
                packaging_supplier="",  # filled by caller
                supplier=supplier,
                factory=factory,
                year=year,
                month=month,
                cartons=_num(row_values[c - 1]),
            ))
    return records
```

Note the two removed checks: the blank-name skip (`if not supplier and not factory: continue`) and the "total" name skip.

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_loader.py -q`
Expected: PASS (all loader tests, including the other existing tests whose fixtures already use numeric column A).

- [ ] **Step 5: Commit**

```bash
git add tests/test_loader.py loader.py
git commit -m "feat: select data rows by numeric column A"
```

---

### Task 2: SL column in the table

**Files:**
- Modify: `filtering.py:42,45-46,71-78`
- Test: `tests/test_filtering.py`

**Interfaces:**
- Consumes: `load_records(path) -> (list[Record], list[str])` (unchanged from Task 1).
- Produces: `build_table(records, packaging_supplier=None, supplier=None, factory=None, frm=None, to=None)` with first column `"SL"` numbered 1..N on the final table; `packaging_suppliers`, `suppliers_for`, `factories_for` unchanged.

- [ ] **Step 1: Update the failing tests**

In `tests/test_filtering.py`:

```python
def test_build_table_columns_and_total():
    df = build_table(recs(), frm=(2026, 1), to=(2026, 2))
    assert df.columns.tolist() == ["SL", "Packaging Supplier", "Supplier", "Factory", "Jan-26", "Feb-26", "Total"]
    assert len(df) == 3
    assert df["SL"].tolist() == [1, 2, 3]
    row = df[df["Factory"] == "Aspire"].iloc[0]
    assert row["Jan-26"] == 10
    assert row["Feb-26"] == 5
    assert row["Total"] == 15
```

```python
def test_build_table_default_duration():
    df = build_table(recs())
    assert df.columns.tolist()[4] == "Mar-24"
    assert df.columns.tolist()[-2] == "Feb-26"
    assert df.columns.tolist()[-1] == "Total"
    assert len(df.columns) == 29
```

```python
def test_build_table_empty_records():
    df = build_table([])
    assert df.empty
    assert df.columns.tolist() == ["SL", "Packaging Supplier", "Supplier", "Factory", "Total"]
```

```python
def test_build_table_month_beyond_data_is_zero():
    df = build_table(recs(), frm=(2026, 3), to=(2026, 3))
    assert df.empty
    assert df.columns.tolist() == ["SL", "Packaging Supplier", "Supplier", "Factory", "Mar-26", "Total"]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_filtering.py -q`
Expected: FAIL — columns lack "SL", SL list is empty, column counts are 28 not 29.

- [ ] **Step 3: Implement SL numbering**

In `filtering.py`:

```python
BASE_COLUMNS = ["SL", "Packaging Supplier", "Supplier", "Factory"]
```

And in `build_table`, number rows after the zero-total filter (replacing the `rows.append(...)` block):

```python
    rows = []
    for key in sorted(agg):
        cartons = agg[key]
        total = sum(cartons.get(m, 0.0) for m in months)
        if total == 0:
            continue
        rows.append(list(key) + [cartons.get(m, 0.0) for m in months] + [total])
    numbered = [[i] + row for i, row in enumerate(rows, start=1)]
    return pd.DataFrame(numbered, columns=BASE_COLUMNS + [month_label(m) for m in months] + ["Total"])
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_filtering.py -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add tests/test_filtering.py filtering.py
git commit -m "feat: add SL index column to table"
```

---

### Task 3: SL column in the PDF

**Files:**
- Modify: `pdf.py:23-45` (`render_pdf` data/width/total-row logic)
- Test: `tests/test_pdf.py`

**Interfaces:**
- Consumes: `build_table(...)` output — a DataFrame whose first column is `"SL"` (from Task 2).
- Produces: unchanged `render_pdf(df, title, out)`; PDF column order mirrors the DataFrame; SL blank in totals row; SL right-aligned; SL column width auto-sized like the label columns.

- [ ] **Step 1: Update the failing tests**

In `tests/test_pdf.py`, update `make_df` to include SL:

```python
def make_df(rows=200, months=26):
    periods = [(2024 + (8 + i - 1) // 12, (8 + i - 1) % 12 + 1) for i in range(months)]
    cols = ["SL", "Packaging Supplier", "Supplier", "Factory"] + [month_label(p) for p in periods] + ["Total"]
    data = []
    for i in range(rows):
        vals = [i + 1] * months
        data.append([i + 1, "M&U", f"Supplier {i % 10}", f"Factory {i}", *vals, sum(vals)])
    return pd.DataFrame(data, columns=cols)
```

Add a test:

```python
def test_pdf_has_sl_column(tmp_path):
    path = tmp_path / "out.pdf"
    render_pdf(make_df(rows=5, months=3), "M&U — Jan-26 to Mar-26", path)
    text = "".join(page.extract_text() or "" for page in PdfReader(str(path)).pages)
    assert "SL" in text
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_pdf.py -q`
Expected: FAIL — `test_pdf_has_sl_column` finds no "SL" (not rendered); `test_pdf_empty_table` fails because `pdf.py` treats the first three columns as labels but the fixture now has 4 label columns.

- [ ] **Step 3: Implement SL in the PDF**

In `pdf.py`:

```python
LABEL_COLUMNS = ("SL", "Packaging Supplier", "Supplier", "Factory")
```

In `render_pdf`, replace the label checks and width computation:

```python
    widths = (
        [_label_width([row[i] for row in data]) for i in range(4)]
        + [MONTH_COL_WIDTH] * (len(columns) - 4)
    )
```

and in the totals-row loop:

```python
        for c in columns:
            if c in LABEL_COLUMNS:
                total_row.append("")
            else:
                total_row.append(_fmt(sums[c]))
```

(The old `"Total" if c == "Packaging Supplier" else ""` special case becomes a plain blank for all four label columns — the "Total" marker in the totals row disappears. If the PDF total row should still carry the word "Total", the simplest matching behavior is to put it in the "Packaging Supplier" cell, i.e. keep:

```python
            if c in LABEL_COLUMNS:
                total_row.append("Total" if c == "Packaging Supplier" else "")
```

Use this version — it preserves the existing "Total" marker the tests assert on.)

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_pdf.py -q`
Expected: PASS (all pdf tests, including `test_pdf_has_total_row_and_title` which asserts "Total" in text).

- [ ] **Step 5: Commit**

```bash
git add tests/test_pdf.py pdf.py
git commit -m "feat: add SL column to PDF export"
```

---

### Task 4: README, full suite, real-data verification

**Files:**
- Modify: `README.md`
- Test: whole suite + real workbook

**Interfaces:**
- Consumes: all of Tasks 1-3.
- Produces: nothing new.

- [ ] **Step 1: Update README**

In `README.md`, in the "Use" section, change the filter bullet to mention SL:

```markdown
- The table starts with an SL (serial) column, numbered 1, 2, 3... in the current view.
```

In the "Notes" section, replace the notes with:

```markdown
- A row is only read if its column A holds a number; rows with blank column A are skipped. (Subtotal rows such as "Total Fashion PJT" do have numbers and therefore appear in the report.)
```

- [ ] **Step 2: Run the full suite**

Run: `python -m pytest -q`
Expected: PASS (all tests, including the previously existing ones).

- [ ] **Step 3: Verify against the real workbook**

Run this check from the project root (PowerShell):

```powershell
$env:PYTHONPATH = "C:\Users\Shoaib\OneDrive - PacD\Projects\Primark\Primark Carton Supply Report"
python -c "from loader import load_records; from filtering import build_table, default_duration; from pdf import render_pdf; import io; from pypdf import PdfReader; rs, w = load_records('Sales Record V1.xlsx'); print('records:', len(rs), 'warnings:', w); f, t = default_duration(rs); df = build_table(rs, frm=f, to=t); print('shape:', df.shape, 'cols:', df.columns.tolist()[:4]); print('SL:', df['SL'].tolist()[:5], '...'); b = io.BytesIO(); render_pdf(df, 'SL check', b); p = PdfReader(b); print('pages:', len(p.pages), 'SL in pdf:', 'SL' in ''.join(x.extract_text() or '' for x in p.pages))"
```

Expected: `records:` greater than 16,756 (the 11 subtotal rows add ~12 records each), warnings still `['M&U - 2020', 'Uniglory - 2020']`, first four columns `['SL', 'Packaging Supplier', 'Supplier', 'Factory']`, SL starts at 1, PDF renders with SL header.

- [ ] **Step 4: Commit**

```bash
git add README.md
git commit -m "docs: document column A row selection and SL column"
```

---

## Self-Review Notes

- Spec coverage: row selection (Task 1), SL numbering (Task 2), PDF SL (Task 3), docs + verification (Task 4). App needs no change — `st.dataframe` shows the SL column automatically and `render_pdf` receives the same df.
- No placeholders: every step has concrete code/commands.
- Type consistency: `Record` unchanged; `load_records` signature unchanged; `build_table` returns df with `"SL"` first; `render_pdf` consumes that df.
