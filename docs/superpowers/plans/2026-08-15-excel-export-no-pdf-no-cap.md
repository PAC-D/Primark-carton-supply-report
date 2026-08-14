# Excel Export, No PDF, No Month Cap Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Remove the PDF export and the 24-month cap; add a formatted Excel export of the filtered table.

**Architecture:** `filtering.py` drops the cap (`default_duration` returns the full data range; `clamp_duration`/`MAX_MONTHS` deleted). A new `excel_export.py` renders the table into a styled .xlsx (title row, header row, totals row, freeze panes) using openpyxl. `pdf.py`, its tests, and its dependencies are deleted; `app.py` swaps the PDF button for an Excel button.

**Tech Stack:** Python 3.14, pandas, openpyxl, pytest (no new dependencies).

## Global Constraints

- `default_duration(records)` returns `(min period, max period)` over the records' `(year, month)` periods; returns `None` for empty input. `clamp_duration` and `MAX_MONTHS` are deleted.
- Excel export: single sheet named `Report`; row 1 = merged bold title; row 2 = headers (bold, fill `#D9E2F3`); `freeze_panes = "A3"`; data rows use whole numbers (`int(round(float(v)))`) for month/Total columns, label cells as-is; final row = totals (bold, fill `#E2EFDA`) with `"Total"` only in the Packaging Supplier cell and `""` in the other label cells; column widths capped; thin grid borders on the used range.
- Label columns are the first 4: `SL | Packaging Supplier | Supplier | Factory` — import `BASE_COLUMNS` from `filtering` (do not re-declare).
- PDF is fully removed: `pdf.py`, `tests/test_pdf.py`, and the `reportlab`/`pypdf` lines in `requirements.txt` are deleted; `app.py` no longer imports or calls `render_pdf`.
- App: From/To pickers allow any range within the data (`min_value`/`max_value` = earliest/latest data month); no cap caption. Export button label `Export Excel`, filename `carton-report.xlsx`, mime `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.
- All tests pass after each commit: run `python -m pytest -q` from the project root.
- Commit style: `feat:`/`fix:`/`docs:` prefix; one commit per task.
- Shell is Windows PowerShell (no bash). Work from `C:\Users\Shoaib\OneDrive - PacD\Projects\Primark\Primark Carton Supply Report`.

---

### Task 1: Remove the 24-month cap

**Files:**
- Modify: `filtering.py:3-4,24-35` (MAX_MONTHS, default_duration, clamp_duration)
- Test: `tests/test_filtering.py`

**Interfaces:**
- Consumes: nothing new.
- Produces: `default_duration(records) -> (frm, to) | None` — full range. `clamp_duration` no longer exists; `key_of`, `period_of`, `range_months`, `month_label` unchanged. `build_table` and the option-list functions unchanged.

- [ ] **Step 1: Update the failing tests**

In `tests/test_filtering.py`, replace `test_default_duration_latest_24` and `test_default_duration_less_than_24` and delete `test_clamp_duration_max_24`:

```python
def test_default_duration_full_range():
    records = [rec(2021, 1)] + [rec(2026, 7)] + [rec(2026, 7, 0)]
    frm, to = default_duration(records)
    assert frm == (2021, 1)
    assert to == (2026, 7)


def test_default_duration_empty():
    assert default_duration([]) is None


def test_default_duration_single_month():
    records = [rec(2026, 3), rec(2026, 3, 0)]
    frm, to = default_duration(records)
    assert (frm, to) == ((2026, 3), (2026, 3))
```

Update `test_build_table_default_duration`:

```python
def test_build_table_default_duration():
    df = build_table(recs())
    assert df.columns.tolist()[4] == "Jan-26"
    assert df.columns.tolist()[-2] == "Feb-26"
    assert df.columns.tolist()[-1] == "Total"
    assert len(df.columns) == 7
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_filtering.py -q`
Expected: FAIL — `test_default_duration_full_range` gets frm (2024, 8); `test_build_table_default_duration` gets "Mar-24"/29 columns; `test_clamp_duration_max_24` fails to import `clamp_duration`.

- [ ] **Step 3: Implement**

In `filtering.py`:

```python
def default_duration(records):
    if not records:
        return None
    periods = [(r.year, r.month) for r in records]
    return min(periods), max(periods)
```

Delete `MAX_MONTHS = 24` and the entire `clamp_duration` function.

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_filtering.py -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add tests/test_filtering.py filtering.py
git commit -m "feat: remove 24-month cap, default to full data range"
```

---

### Task 2: Excel export module

**Files:**
- Create: `excel_export.py`
- Test: `tests/test_excel_export.py`

**Interfaces:**
- Consumes: `build_table(...)` output (DataFrame with `SL` first, label columns from `filtering.BASE_COLUMNS`); `filtering.BASE_COLUMNS` imported for the label count.
- Produces: `render_excel(df, title, out)` where `out` is a path, `os.PathLike`, or file-like; writes a styled .xlsx; returns None.

- [ ] **Step 1: Write the failing tests**

Create `tests/test_excel_export.py`:

```python
import io

import openpyxl
import pandas as pd

from excel_export import render_excel
from filtering import month_label


def make_df(rows=5, months=3):
    periods = [(2024 + (8 + i - 1) // 12, (8 + i - 1) % 12 + 1) for i in range(months)]
    cols = ["SL", "Packaging Supplier", "Supplier", "Factory"] + [month_label(p) for p in periods] + ["Total"]
    data = []
    for i in range(rows):
        vals = [i + 1] * months
        data.append([i + 1, "M&U", f"Supplier {i % 10}", f"Factory {i}", *vals, sum(vals)])
    return pd.DataFrame(data, columns=cols)


def test_excel_export_creates_workbook(tmp_path):
    path = tmp_path / "out.xlsx"
    render_excel(make_df(), "M&U — Jan-26 to Mar-26", path)
    assert path.exists()
    assert path.stat().st_size > 1000


def test_excel_export_structure(tmp_path):
    path = tmp_path / "out.xlsx"
    df = make_df(rows=4, months=3)
    render_excel(df, "M&U — Jan-26 to Mar-26", path)
    ws = openpyxl.load_workbook(path)["Report"]
    assert ws["A1"].value == "M&U — Jan-26 to Mar-26"
    assert str(ws.merged_cells.ranges[0]).startswith("A1:")
    headers = [ws.cell(row=2, column=c).value for c in range(1, 9)]
    assert headers == df.columns.tolist()
    assert ws.freeze_panes == "A3"
    total_row = [ws.cell(row=7, column=c).value for c in range(1, 9)]
    assert total_row == ["", "Total", "", "", 10, 10, 10, 30]


def test_excel_export_whole_numbers(tmp_path):
    path = tmp_path / "out.xlsx"
    df = make_df(rows=3, months=2)
    df["Jan-26"] = [1.0, 2.5, 3.9]
    df["Total"] = [1.0, 2.5, 3.9]
    render_excel(df, "T", path)
    ws = openpyxl.load_workbook(path)["Report"]
    jan = [ws.cell(row=r, column=5).value for r in range(3, 6)]
    assert jan == [1, 3, 4]
    assert all(isinstance(v, int) for v in jan)


def test_excel_export_accepts_file_object():
    buf = io.BytesIO()
    render_excel(make_df(rows=3, months=2), "T", buf)
    assert len(buf.getvalue()) > 1000
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_excel_export.py -q`
Expected: FAIL — `ModuleNotFoundError: excel_export`.

- [ ] **Step 3: Implement `excel_export.py`**

```python
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from filtering import BASE_COLUMNS

HEADER_BG = "D9E2F3"
TOTAL_BG = "E2EFDA"
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _whole(value):
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return value


def _column_width(values):
    longest = max((len(str(v)) for v in values), default=0)
    return min(longest + 2, 40)


def render_excel(df, title, out):
    if isinstance(out, os.PathLike):
        out = os.fspath(out)
    n = len(BASE_COLUMNS)
    columns = df.columns.tolist()
    data = [columns]
    if len(df):
        for row in df.values.tolist():
            data.append(list(row[:n]) + [_whole(v) for v in row[n:]])
        sums = df.sum(numeric_only=True)
        total_row = []
        for c in columns:
            if c in BASE_COLUMNS:
                total_row.append("Total" if c == "Packaging Supplier" else "")
            else:
                total_row.append(_whole(sums[c]))
        data.append(total_row)

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(columns))
    title_cell = ws.cell(row=1, column=1, value=title)
    title_cell.font = Font(bold=True, size=14)

    for c, value in enumerate(data[0], start=1):
        cell = ws.cell(row=2, column=c, value=value)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor=HEADER_BG)
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center")

    for r, row in enumerate(data[1:], start=3):
        for c, value in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell.border = BORDER
            if c > n:
                cell.alignment = Alignment(horizontal="right")

    last = ws.max_row
    for c in range(1, len(columns) + 1):
        cell = ws.cell(row=last, column=c)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor=TOTAL_BG)

    for c in range(1, len(columns) + 1):
        col = [ws.cell(row=r, column=c).value for r in range(2, last + 1)] + [columns[c - 1]]
        ws.column_dimensions[get_column_letter(c)].width = _column_width(col)

    ws.freeze_panes = "A3"
    wb.save(out)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_excel_export.py -q`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add tests/test_excel_export.py excel_export.py
git commit -m "feat: formatted Excel export of the filtered table"
```

---

### Task 3: Remove the PDF

**Files:**
- Delete: `pdf.py`, `tests/test_pdf.py`
- Modify: `app.py:1-18,114-125` (remove import, PDF button block), `requirements.txt:4-5`

**Interfaces:**
- Consumes: nothing new (the app still imports `render_pdf` — it must stop).
- Produces: no PDF code remains anywhere; `requirements.txt` has no `reportlab`/`pypdf`.

- [ ] **Step 1: Write the failing test**

There is no app test; the failing "test" is the app import. Run:

Run: `python -c "import app"`
Expected: FAIL — ImportError (either `No module named 'pdf'` after the pdf deletion, or `cannot import name 'clamp_duration'`, which Task 1 deleted from filtering.py).

- [ ] **Step 2: Verify it fails**

Run: `python -m pytest -q`
Expected: FAIL — collection error on `tests/test_pdf.py`; `import app` raises ModuleNotFoundError.

- [ ] **Step 3: Implement**

1. Delete the files:

```powershell
Remove-Item pdf.py, tests/test_pdf.py
```

2. In `app.py`, remove the line `from pdf import render_pdf`, remove `clamp_duration` from the `filtering` import list (Task 1 deleted it), and remove the entire export block:

```python
    try:
        buf = io.BytesIO()
        render_pdf(df, title, buf)
        st.download_button(
            "Export PDF",
            data=buf.getvalue(),
            file_name="carton-report.pdf",
            mime="application/pdf",
        )
    except Exception as exc:
        st.error(f"PDF generation failed: {exc}")
```

3. In `requirements.txt`, delete the lines `reportlab>=4.2` and `pypdf>=5`.

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest -q`
Expected: PASS (29 tests).

- [ ] **Step 5: Commit**

```powershell
git add -A
git commit -m "fix: remove PDF export"
```

---

### Task 4: Excel export button in the app

**Files:**
- Modify: `app.py:1-18,55-125`
- Modify: `README.md`

**Interfaces:**
- Consumes: `render_excel(df, title, out)` from Task 2; `default_duration` full-range from Task 1.
- Produces: app with `Export Excel` button, no cap caption, pickers bounded by the data.

- [ ] **Step 1: Update app.py**

1. Imports: add `from excel_export import render_excel`; confirm `clamp_duration` is not in the `filtering` import list (removed in Task 3); keep `key_of` (used by the caption).

2. Picker bounds — replace the hardcoded `min_value=datetime.date(2020, 1, 1)` on both pickers with the earliest data month:

```python
    with f4:
        from_date = st.date_input(
            "From", value=from_default, min_value=from_default, max_value=to_default
        )
    with f5:
        to_date = st.date_input(
            "To", value=to_default, min_value=from_default, max_value=to_default
        )
```

3. Remove the clamp block entirely:

```python
    clamped_frm, clamped_to = clamp_duration(frm, to)
    if (clamped_frm, clamped_to) != (frm, to):
        st.caption(f"Range capped at 24 months — showing {month_label(clamped_frm)} to {month_label(clamped_to)}.")
    frm, to = clamped_frm, clamped_to
```

(so `frm, to` stay as picked, after the `if frm > to: swap`.)

4. Replace the export block with:

```python
    title = (", ".join(pkg_sel) if pkg_sel else "All suppliers") + f" \u2014 {month_label(frm)} to {month_label(to)}"
    try:
        buf = io.BytesIO()
        render_excel(df, title, buf)
        st.download_button(
            "Export Excel",
            data=buf.getvalue(),
            file_name="carton-report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as exc:
        st.error(f"Excel export failed: {exc}")
```

- [ ] **Step 2: Verify the app imports and runs**

Run: `python -c "import app"` — must print nothing and exit 0.

- [ ] **Step 3: Update README**

In `README.md`:
- In "Use", change the export bullet to: `- Click "Export Excel" to download the filtered table as a formatted .xlsx (header row, totals row, whole numbers).`
- In "Use", change the filter bullet to remove "max 24 months": `- Filter by Packaging Supplier, Supplier, Factory (multi-select; empty = All), and any From/To month range (default: full range of the data).`

- [ ] **Step 4: Run the full suite + real-data check**

Run: `python -m pytest -q` — PASS.

Run (PowerShell, from the project root):

```powershell
$env:PYTHONPATH = "C:\Users\Shoaib\OneDrive - PacD\Projects\Primark\Primark Carton Supply Report"
python -c "from loader import load_records; from filtering import build_table, default_duration; from excel_export import render_excel; import io, openpyxl; rs, w = load_records('Sales Record V1.xlsx'); f, t = default_duration(rs); print('range:', f, 'to', t, 'warnings:', w); df = build_table(rs, frm=f, to=t); print('shape:', df.shape); b = io.BytesIO(); render_excel(df, 'Full range check', b); wb = openpyxl.load_workbook(io.BytesIO(b.getvalue())); ws = wb['Report']; print('sheets:', wb.sheetnames, 'freeze:', ws.freeze_panes, 'last row:', ws.max_row); print('total row:', [ws.cell(row=ws.max_row, column=c).value for c in range(1, 6)])"
```

Expected: range = earliest..latest data period (e.g. (2020, 1) to (2026, 12)), warnings unchanged, shape columns = `len(BASE_COLUMNS) + months + 1`, sheet `Report`, freeze `A3`, totals row present with "Total" label.

- [ ] **Step 5: Commit**

```powershell
git add app.py README.md
git commit -m "feat: export table to Excel from the app"
```

---

## Self-Review Notes

- Spec coverage: cap removal (Task 1), Excel module + tests (Task 2), PDF removal incl. deps (Task 3), app wiring + README + real-data check (Task 4). No app tests exist — the app is verified by `import app` plus the real-data run.
- No placeholders: every step has concrete code/commands; test code is verbatim.
- Type consistency: `render_excel(df, title, out)` mirrors the removed `render_pdf` signature; `BASE_COLUMNS` from filtering is the single source of the 4 label columns; `key_of` import stays for the caption; `clamp_duration` is referenced nowhere after Task 3 removes it from the app's import list.