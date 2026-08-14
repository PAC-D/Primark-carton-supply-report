# Carton Supply Report Tool Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A local Streamlit app that reads `Sales Record V1.xlsx`, filters carton data by packaging supplier / supplier / factory / month range, shows a month-column table, and exports it to a multi-page landscape A4 PDF with repeating headers and totals.

**Architecture:** Flat Python modules at the project root — `loader.py` (Excel → flat month-level records), `filtering.py` (duration rules + table pivot), `pdf.py` (ReportLab renderer), `app.py` (Streamlit UI wiring). Tests use pytest against fixture workbooks built in-memory with openpyxl.

**Tech Stack:** Python 3.14, Streamlit, openpyxl, pandas, ReportLab (platypus `LongTable`), pytest, pypdf (test-only, page-count assertions).

## Global Constraints

- Duration range is capped at 24 months; with no duration selection the default is the latest 24 months (latest month with data − 23 … latest month with data).
- Output table columns: `Packaging Supplier | Supplier | Factory | <one col per month, chronological> | Total`.
- Rows with zero cartons across the whole selected range are hidden.
- Re sheets (`<Supplier> Re - <Year>`) replace the original `<Supplier> - <Year>`; never both.
- Summary sheets `2024` and `Uniglory Sales Analy` are always ignored, silently.
- Any other sheet with no month headers is skipped and added to a visible warning list.
- Rows whose supplier or factory name contains `total` (case-insensitive) are excluded.
- Month headers are `datetime` values in row 2. Factory column = the row-2 header whose stripped value equals `factory`; Supplier column = header `suppliers`; fall back to col B (suppliers) and col C (factory) if not found.
- PDF: A4 landscape, header row repeated on every page, alternating row shading, grid lines, `Total` row at the end (per-month column sums + grand total), title line, `Generated on <date>` line.
- Flat module layout at project root (no `src/` package). All commands run from the project root.
- pytest config lives in `pyproject.toml` (`pythonpath = ["."]`).

---

### Task 1: Project scaffolding

**Files:**
- Create: `pyproject.toml`
- Create: `requirements.txt`
- Create: `models.py`

**Interfaces:**
- Produces: `Record` dataclass used by every later task; pytest root config; dependency manifest.

- [ ] **Step 1: Write the failing test**

Create `tests/test_models.py`:

```python
from models import Record


def test_record_fields():
    r = Record("M&U", "PADMA TEXTILES LTD", "Aspire Garments Ltd PJT (24040)", 2026, 1, 80060)
    assert r.packaging_supplier == "M&U"
    assert r.supplier == "PADMA TEXTILES LTD"
    assert r.factory == "Aspire Garments Ltd PJT (24040)"
    assert r.year == 2026
    assert r.month == 1
    assert r.cartons == 80060


def test_record_period():
    r = Record("M&U", "S", "F", 2026, 7, 5)
    assert r.period == (2026, 7)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_models.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'models'`

- [ ] **Step 3: Write minimal implementation**

Create `models.py`:

```python
from dataclasses import dataclass


@dataclass(frozen=True)
class Record:
    packaging_supplier: str
    supplier: str
    factory: str
    year: int
    month: int
    cartons: float

    @property
    def period(self):
        return self.year, self.month
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_models.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Scaffolding files**

Create `pyproject.toml`:

```toml
[tool.pytest.ini_options]
pythonpath = ["."]
testpaths = ["tests"]
```

Create `requirements.txt`:

```
streamlit>=1.40
openpyxl>=3.1
pandas>=2.2
reportlab>=4.2
pypdf>=5
pytest>=8
```

- [ ] **Step 6: Verify full suite still passes**

Run: `python -m pytest -v`
Expected: PASS (2 passed)

- [ ] **Step 7: Commit**

```bash
git add pyproject.toml requirements.txt models.py tests/test_models.py
git commit -m "chore: scaffold project, Record model, pytest config"
```

---

### Task 2: Loader — parse workbook into records

**Files:**
- Create: `loader.py`
- Create: `tests/fixtures.py`
- Create: `tests/test_loader.py`

**Interfaces:**
- Consumes: `Record` from `models.py`.
- Produces: `load_records(path: str | Path) -> tuple[list[Record], list[str]]` — the list is all flat month-level records (one per factory-month), the second element is a list of human-readable warnings for skipped sheets.

- [ ] **Step 1: Write the failing test**

Create `tests/fixtures.py`:

```python
import datetime
from pathlib import Path

import openpyxl


def make_workbook(tmp_path, sheets):
    """sheets: {sheet_title: [(supplier, factory, {(year, month): cartons}), ...]}"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        ws.cell(row=1, column=1, value=title)
        ws.cell(row=2, column=1, value="Sl")
        ws.cell(row=2, column=2, value="Suppliers")
        ws.cell(row=2, column=3, value="Factory ")
        months = sorted({m for _, _, data in rows for m in data})
        for j, (y, m) in enumerate(months, start=4):
            ws.cell(row=2, column=j, value=datetime.datetime(y, m, 1))
        for i, (sup, fac, data) in enumerate(rows, start=3):
            ws.cell(row=i, column=1, value=i - 2)
            ws.cell(row=i, column=2, value=sup)
            ws.cell(row=i, column=3, value=fac)
            for (y, m), v in data.items():
                ws.cell(row=i, column=4 + months.index((y, m)), value=v)
    path = tmp_path / "test.xlsx"
    wb.save(path)
    return path
```

Create `tests/test_loader.py`:

```python
from loader import load_records
from tests.fixtures import make_workbook


def test_load_basic(tmp_path):
    path = make_workbook(tmp_path, {
        "M&U - 2026": [
            ("PADMA TEXTILES LTD", "Aspire Garments Ltd PJT (24040)", {(2026, 1): 2010, (2026, 2): 9284}),
            ("TEX INTERNATIONAL LTD", "Aboni Knitwear Ltd", {(2026, 1): 500}),
        ],
        "Union-2026": [
            ("CENTRO INTERNATIONAL SOURCING LTD", "APS Apparels Limited (26018)", {(2026, 1): 11537}),
        ],
    })
    records, warnings = load_records(path)
    assert warnings == []
    assert len(records) == 4
    m = {(r.packaging_supplier, r.year, r.month, r.factory): r.cartons for r in records}
    assert m[("M&U", 2026, 1, "Aspire Garments Ltd PJT (24040)")] == 2010
    assert m[("M&U", 2026, 2, "Aspire Garments Ltd PJT (24040)")] == 9284
    assert m[("M&U", 2026, 1, "Aboni Knitwear Ltd")] == 500
    assert m[("Union", 2026, 1, "APS Apparels Limited (26018)")] == 11537


def test_re_replaces_original(tmp_path):
    path = make_workbook(tmp_path, {
        "M&U - 2023": [("A", "Old Factory", {(2023, 1): 100})],
        "M&U Re - 2023": [("A", "New Factory", {(2023, 1): 200})],
    })
    records, _ = load_records(path)
    assert len(records) == 1
    assert records[0].factory == "New Factory"
    assert records[0].cartons == 200


def test_total_row_excluded(tmp_path):
    path = make_workbook(tmp_path, {
        "M&U - 2026": [
            ("A", "Factory One", {(2026, 1): 100}),
            ("A", "Total", {(2026, 1): 9000}),
        ],
    })
    records, _ = load_records(path)
    assert len(records) == 1
    assert records[0].factory == "Factory One"


def test_sheet_without_months_warns(tmp_path):
    path = make_workbook(tmp_path, {
        "M&U - 2026": [("A", "F", {(2026, 1): 10})],
        "M&U - 2020": [("A", "F", {})],
    })
    records, warnings = load_records(path)
    assert len(records) == 1
    assert len(warnings) == 1
    assert "M&U - 2020" in warnings[0]


def test_summary_sheets_ignored_silently(tmp_path):
    path = make_workbook(tmp_path, {
        "M&U - 2026": [("A", "F", {(2026, 1): 10})],
        "2024": [("A", "F", {})],
        "Uniglory Sales Analy": [("A", "F", {})],
    })
    records, warnings = load_records(path)
    assert len(records) == 1
    assert warnings == []


def test_empty_cells_are_zero(tmp_path):
    path = make_workbook(tmp_path, {
        "M&U - 2026": [("A", "F", {(2026, 1): None, (2026, 2): 50})],
    })
    records, _ = load_records(path)
    assert len(records) == 2
    by_month = {r.month: r.cartons for r in records}
    assert by_month[1] == 0
    assert by_month[2] == 50


def test_epyllion_layout_factory_in_column_d(tmp_path):
    # Epyllion sheets have an extra "Supplier" column: Sl | Suppliers | Supplier | Factory | months
    wb = __import__("openpyxl").Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Epyllion - 2026")
    ws.append(["Sl", "Suppliers", "Supplier", "Factory ", None, None, None])
    ws.cell(row=2, column=5, value=__import__("datetime").datetime(2026, 1, 1))
    ws.cell(row=3, column=1, value=1)
    ws.cell(row=3, column=2, value="SNQS GLOBAL TEXTILES FZE")
    ws.cell(row=3, column=3, value="SNQS GLOBAL TEXTILES FZE")
    ws.cell(row=3, column=4, value="Alim Knit Ltd (20331)")
    ws.cell(row=3, column=5, value=3949)
    path = tmp_path / "test.xlsx"
    wb.save(path)
    records, _ = load_records(path)
    assert len(records) == 1
    assert records[0].factory == "Alim Knit Ltd (20331)"
    assert records[0].supplier == "SNQS GLOBAL TEXTILES FZE"
    assert records[0].cartons == 3949
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_loader.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'loader'`

- [ ] **Step 3: Write minimal implementation**

Create `loader.py`:

```python
import datetime
from pathlib import Path

import openpyxl

from models import Record

SUMMARY_SHEETS = {"2024", "Uniglory Sales Analy"}


def _num(value):
    try:
        if value is None:
            return 0
        return float(value)
    except (TypeError, ValueError):
        return 0


def _parse_sheet(ws):
    header = [ws.cell(row=2, column=c).value for c in range(1, ws.max_column + 1)]
    month_cols = []
    for i, v in enumerate(header):
        if isinstance(v, (datetime.datetime, datetime.date)):
            month_cols.append((i + 1, v.year, v.month))
    if not month_cols:
        return None
    factory_col = next(
        (i + 1 for i, v in enumerate(header)
         if isinstance(v, str) and v.strip().lower() == "factory"),
        3,
    )
    supplier_col = next(
        (i + 1 for i, v in enumerate(header)
         if isinstance(v, str) and v.strip().lower() == "suppliers"),
        2,
    )
    records = []
    for r in range(3, ws.max_row + 1):
        supplier = str(ws.cell(row=r, column=supplier_col).value or "").strip()
        factory = str(ws.cell(row=r, column=factory_col).value or "").strip()
        if not supplier and not factory:
            continue
        if "total" in (supplier + " " + factory).lower():
            continue
        for c, year, month in month_cols:
            records.append(Record(
                packaging_supplier="",  # filled by caller
                supplier=supplier,
                factory=factory,
                year=year,
                month=month,
                cartons=_num(ws.cell(row=r, column=c).value),
            ))
    return records


def load_records(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    sheets = {}
    for ws in wb.worksheets:
        title = ws.title.strip()
        parts = title.split(" - ")
        if len(parts) < 2 or title in SUMMARY_SHEETS:
            continue
        sheets[title] = ws

    groups = {}
    for title in sheets:
        parts = title.split(" - ")
        supplier = parts[0].strip()
        year = parts[-1].strip()
        if not year.isdigit():
            continue
        is_re = supplier.endswith(" Re")
        base = supplier[:-3].strip() if is_re else supplier
        groups.setdefault((base, int(year)), []).append((title, is_re))

    chosen = []
    for (base, year), group in groups.items():
        re_sheets = [t for t, is_re in group if is_re]
        chosen.append(re_sheets[0] if re_sheets else group[0][0])

    records = []
    warnings = []
    for title in chosen:
        ws = sheets[title]
        supplier_name = title.split(" - ")[0].strip()
        if supplier_name.endswith(" Re"):
            supplier_name = supplier_name[:-3].strip()
        parsed = _parse_sheet(ws)
        if parsed is None:
            warnings.append(title)
            continue
        for rec in parsed:
            records.append(Record(
                packaging_supplier=supplier_name,
                supplier=rec.supplier,
                factory=rec.factory,
                year=rec.year,
                month=rec.month,
                cartons=rec.cartons,
            ))
    wb.close()
    return records, warnings
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_loader.py -v`
Expected: PASS (7 passed)

- [ ] **Step 5: Verify against the real workbook**

Run:

```bash
python -c "from loader import load_records; recs, warns = load_records('Sales Record V1.xlsx'); print(len(recs), 'records'); print('warnings:', warns); from collections import defaultdict; t = defaultdict(float); [t.__setitem__(r.packaging_supplier, t[r.packaging_supplier] + r.cartons) for r in recs]; print(dict(t))"
```

Expected: tens of thousands of records; warnings lists the no-month sheets (`M&U - 2020`, `M&U - 2021`, `Uniglory - 2020`, `Uniglory - 2021`); supplier totals roughly match the earlier analysis (Uniglory ~60M, M&U ~55M, Union ~26M, Epyllion ~23.5M).

- [ ] **Step 6: Run full suite**

Run: `python -m pytest -v`
Expected: PASS (9 passed)

- [ ] **Step 7: Commit**

```bash
git add loader.py tests/fixtures.py tests/test_loader.py
git commit -m "feat: parse workbook into flat month-level records with Re replacement"
```

---

### Task 3: Duration logic (default range + 24-month clamp)

**Files:**
- Create: `filtering.py`
- Create: `tests/test_filtering.py`

**Interfaces:**
- Consumes: `Record` from `models.py`.
- Produces:
  - `key_of(period: tuple[int, int]) -> int` — `(year, month)` → monotonic month key (`year * 12 + (month - 1)`).
  - `period_of(key: int) -> tuple[int, int]` — inverse of `key_of`.
  - `range_months(frm: tuple[int, int], to: tuple[int, int]) -> list[tuple[int, int]]` — inclusive calendar months between `frm` and `to`.
  - `month_label(period: tuple[int, int]) -> str` — e.g. `(2026, 1)` → `"Jan-26"`.
  - `default_duration(records: list[Record]) -> tuple[tuple[int, int], tuple[int, int]] | None` — latest 24 months with data; `None` if `records` is empty.
  - `clamp_duration(frm: tuple[int, int], to: tuple[int, int]) -> tuple[tuple[int, int], tuple[int, int]]` — pushes `frm` forward so the span is at most 24 months (assumes `frm <= to`).

- [ ] **Step 1: Write the failing test**

Create `tests/test_filtering.py`:

```python
from filtering import (
    clamp_duration,
    default_duration,
    key_of,
    month_label,
    period_of,
    range_months,
)
from models import Record


def rec(y, m, cartons=1):
    return Record("M&U", "S", "F", y, m, cartons)


def test_key_and_period_roundtrip():
    assert key_of((2026, 1)) == 2026 * 12
    assert key_of((2026, 12)) == 2026 * 12 + 11
    assert period_of(key_of((2026, 12))) == (2026, 12)
    assert period_of(key_of((2021, 3))) == (2021, 3)


def test_range_months():
    assert range_months((2026, 1), (2026, 3)) == [(2026, 1), (2026, 2), (2026, 3)]
    assert range_months((2025, 12), (2026, 2)) == [(2025, 12), (2026, 1), (2026, 2)]


def test_month_label():
    assert month_label((2026, 1)) == "Jan-26"
    assert month_label((2025, 12)) == "Dec-25"


def test_default_duration_latest_24():
    records = [rec(2021, 1)] + [rec(2026, 7)] + [rec(2026, 7, 0)]
    frm, to = default_duration(records)
    assert to == (2026, 7)
    assert frm == (2024, 8)
    assert key_of(to) - key_of(frm) + 1 == 24


def test_default_duration_empty():
    assert default_duration([]) is None


def test_default_duration_less_than_24():
    # Spec: default is ALWAYS the latest 24 months, even when data is sparse.
    records = [rec(2026, 1), rec(2026, 2), rec(2026, 3)]
    frm, to = default_duration(records)
    assert frm == (2024, 4)
    assert to == (2026, 3)
    assert key_of(to) - key_of(frm) + 1 == 24


def test_clamp_duration_max_24():
    assert clamp_duration((2024, 1), (2026, 7)) == ((2024, 8), (2026, 7))
    assert clamp_duration((2024, 8), (2026, 7)) == ((2024, 8), (2026, 7))
    assert clamp_duration((2026, 1), (2026, 3)) == ((2026, 1), (2026, 3))
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_filtering.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'filtering'`

- [ ] **Step 3: Write minimal implementation**

Create `filtering.py`:

```python
import datetime

MAX_MONTHS = 24


def key_of(period):
    year, month = period
    return year * 12 + (month - 1)


def period_of(key):
    return divmod(key, 12)[0], divmod(key, 12)[1] + 1


def range_months(frm, to):
    return [period_of(k) for k in range(key_of(frm), key_of(to) + 1)]


def month_label(period):
    year, month = period
    return datetime.date(year, month, 1).strftime("%b-%y")


def default_duration(records):
    if not records:
        return None
    to = max((r.year, r.month) for r in records)
    frm = period_of(key_of(to) - (MAX_MONTHS - 1))
    return frm, to


def clamp_duration(frm, to):
    if key_of(to) - key_of(frm) + 1 > MAX_MONTHS:
        frm = period_of(key_of(to) - (MAX_MONTHS - 1))
    return frm, to
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_filtering.py -v`
Expected: PASS (8 passed)

- [ ] **Step 5: Commit**

```bash
git add filtering.py tests/test_filtering.py
git commit -m "feat: duration helpers with latest-24 default and 24-month clamp"
```

---

### Task 4: Filtering + table pivot

**Files:**
- Modify: `filtering.py` (append functions)
- Modify: `tests/test_filtering.py` (append tests)

**Interfaces:**
- Consumes: `Record`, `key_of`, `range_months`, `month_label`, `default_duration`.
- Produces:
  - `build_table(records, packaging_supplier=None, supplier=None, factory=None, frm=None, to=None) -> pd.DataFrame` — columns `["Packaging Supplier", "Supplier", "Factory", <month labels...>, "Total"]`; one row per (packaging supplier, supplier, factory) triple; zero-carton rows hidden; if `frm`/`to` are `None`, `default_duration` is used (empty DataFrame with the three base columns if no records).
  - `packaging_suppliers(records) -> list[str]` — sorted unique values.
  - `suppliers_for(records, packaging_supplier=None) -> list[str]` — sorted unique, filtered by packaging supplier.
  - `factories_for(records, packaging_supplier=None, supplier=None) -> list[str]` — sorted unique, filtered by both.

- [ ] **Step 1: Write the failing test**

Append to `tests/test_filtering.py`:

```python
import pandas as pd

from filtering import (
    build_table,
    factories_for,
    packaging_suppliers,
    suppliers_for,
)
from models import Record


def recs():
    return [
        Record("M&U", "PADMA", "Aspire", 2026, 1, 10),
        Record("M&U", "PADMA", "Aspire", 2026, 2, 5),
        Record("M&U", "TEX", "Aboni", 2026, 1, 7),
        Record("Union", "CENTRO", "APS", 2026, 1, 3),
        Record("Union", "CENTRO", "APS", 2026, 2, 0),
        Record("Union", "CENTRO", "Zero Row", 2026, 1, 0),
        Record("Union", "CENTRO", "Zero Row", 2026, 2, 0),
    ]


def test_build_table_columns_and_total():
    df = build_table(recs(), frm=(2026, 1), to=(2026, 2))
    assert df.columns.tolist() == ["Packaging Supplier", "Supplier", "Factory", "Jan-26", "Feb-26", "Total"]
    assert len(df) == 3
    row = df[df["Factory"] == "Aspire"].iloc[0]
    assert row["Jan-26"] == 10
    assert row["Feb-26"] == 5
    assert row["Total"] == 15


def test_build_table_hides_zero_rows():
    df = build_table(recs(), frm=(2026, 1), to=(2026, 2))
    assert "Zero Row" not in df["Factory"].tolist()


def test_build_table_filters():
    df = build_table(recs(), packaging_supplier="Union", supplier="CENTRO",
                     factory="APS", frm=(2026, 1), to=(2026, 2))
    assert len(df) == 1
    assert df.iloc[0]["Total"] == 3


def test_build_table_month_range():
    df = build_table(recs(), frm=(2026, 2), to=(2026, 2))
    assert df.columns.tolist() == ["Packaging Supplier", "Supplier", "Factory", "Feb-26", "Total"]
    row = df[df["Factory"] == "Aspire"].iloc[0]
    assert row["Feb-26"] == 5
    assert row["Total"] == 5


def test_build_table_default_duration():
    df = build_table(recs())
    assert df.columns.tolist()[3] == "Mar-24"
    assert df.columns.tolist()[-2] == "Feb-26"
    assert df.columns.tolist()[-1] == "Total"
    assert len(df.columns) == 28


def test_build_table_empty_records():
    df = build_table([])
    assert df.empty
    assert df.columns.tolist() == ["Packaging Supplier", "Supplier", "Factory", "Total"]


def test_build_table_month_beyond_data_is_zero():
    df = build_table(recs(), frm=(2026, 3), to=(2026, 3))
    assert df.empty
    assert df.columns.tolist() == ["Packaging Supplier", "Supplier", "Factory", "Mar-26", "Total"]


def test_option_lists():
    rs = recs()
    assert packaging_suppliers(rs) == ["M&U", "Union"]
    assert suppliers_for(rs, "M&U") == ["PADMA", "TEX"]
    assert suppliers_for(rs) == ["CENTRO", "PADMA", "TEX"]
    assert factories_for(rs, "M&U", "PADMA") == ["Aspire"]
    assert factories_for(rs, "Union") == ["APS", "Zero Row"]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_filtering.py -v`
Expected: FAIL — `ImportError: cannot import name 'build_table'`

- [ ] **Step 3: Write minimal implementation**

Append to `filtering.py`:

```python
from collections import defaultdict

import pandas as pd

BASE_COLUMNS = ["Packaging Supplier", "Supplier", "Factory"]


def _empty_table(months):
    return pd.DataFrame(columns=BASE_COLUMNS + [month_label(m) for m in months] + ["Total"])


def build_table(records, packaging_supplier=None, supplier=None, factory=None, frm=None, to=None):
    if frm is None or to is None:
        duration = default_duration(records)
        if duration is None:
            return _empty_table([])
        frm, to = duration
    months = range_months(frm, to)
    if not records:
        return _empty_table(months)
    agg = defaultdict(lambda: defaultdict(float))
    fk, tk = key_of(frm), key_of(to)
    for r in records:
        if packaging_supplier and r.packaging_supplier != packaging_supplier:
            continue
        if supplier and r.supplier != supplier:
            continue
        if factory and r.factory != factory:
            continue
        k = key_of((r.year, r.month))
        if not (fk <= k <= tk):
            continue
        agg[(r.packaging_supplier, r.supplier, r.factory)][(r.year, r.month)] += r.cartons
    rows = []
    for key in sorted(agg):
        cartons = agg[key]
        total = sum(cartons.get(m, 0.0) for m in months)
        if total == 0:
            continue
        rows.append(list(key) + [cartons.get(m, 0.0) for m in months] + [total])
    return pd.DataFrame(rows, columns=BASE_COLUMNS + [month_label(m) for m in months] + ["Total"])


def packaging_suppliers(records):
    return sorted({r.packaging_supplier for r in records})


def suppliers_for(records, packaging_supplier=None):
    return sorted({
        r.supplier for r in records
        if not packaging_supplier or r.packaging_supplier == packaging_supplier
    })


def factories_for(records, packaging_supplier=None, supplier=None):
    return sorted({
        r.factory for r in records
        if (not packaging_supplier or r.packaging_supplier == packaging_supplier)
        and (not supplier or r.supplier == supplier)
    })
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_filtering.py -v`
Expected: PASS (16 passed total in file)

- [ ] **Step 5: Run full suite**

Run: `python -m pytest -v`
Expected: PASS (25 passed)

- [ ] **Step 6: Commit**

```bash
git add filtering.py tests/test_filtering.py
git commit -m "feat: table pivot with month columns, filters, and option lists"
```

---

### Task 5: PDF renderer (ReportLab)

**Files:**
- Create: `pdf.py`
- Create: `tests/test_pdf.py`

**Interfaces:**
- Consumes: `pd.DataFrame` as produced by `build_table`.
- Produces: `render_pdf(df: pd.DataFrame, title: str, out) -> None` — writes a landscape-A4 PDF to `out` (a path `str`/`Path` or a writable binary file-like object). Header row repeated per page, alternating shading, grid lines, `Total` row (per-month column sums + grand total) at the end, `Generated on <date>` line.

- [ ] **Step 1: Write the failing test**

Create `tests/test_pdf.py`:

```python
import datetime
from pathlib import Path

import pandas as pd
from pypdf import PdfReader

from filtering import month_label
from pdf import render_pdf


def make_df(rows=200, months=26):
    periods = [(2024 + (8 + i - 1) // 12, (8 + i - 1) % 12 + 1) for i in range(months)]
    cols = ["Packaging Supplier", "Supplier", "Factory"] + [month_label(p) for p in periods] + ["Total"]
    data = []
    for i in range(rows):
        vals = [i + 1] * months
        data.append(["M&U", f"Supplier {i % 10}", f"Factory {i}", *vals, sum(vals)])
    return pd.DataFrame(data, columns=cols)


def test_pdf_generates_multipage(tmp_path):
    path = tmp_path / "out.pdf"
    render_pdf(make_df(), "All suppliers — Aug-24 to Jul-26", path)
    assert path.exists()
    assert path.stat().st_size > 1000
    reader = PdfReader(str(path))
    assert len(reader.pages) > 1


def test_pdf_has_total_row_and_title(tmp_path):
    path = tmp_path / "out.pdf"
    df = make_df(rows=5, months=3)
    render_pdf(df, "M&U — Jan-26 to Mar-26", path)
    text = "".join(page.extract_text() or "" for page in PdfReader(str(path)).pages)
    assert "Total" in text
    assert "M&U — Jan-26 to Mar-26" in text
    assert "Generated on" in text
    # grand total = 5 rows x 3 months x (1+2+3+4+5) = 3 * 15 = 45
    assert "45" in text


def test_pdf_empty_table(tmp_path):
    path = tmp_path / "empty.pdf"
    df = pd.DataFrame(columns=["Packaging Supplier", "Supplier", "Factory", "Jan-26", "Total"])
    render_pdf(df, "No data", path)
    assert path.exists()
    assert path.stat().st_size > 500


def test_pdf_accepts_file_object():
    import io
    buf = io.BytesIO()
    render_pdf(make_df(rows=10, months=4), "Test", buf)
    assert len(buf.getvalue()) > 1000
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_pdf.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'pdf'`

- [ ] **Step 3: Write minimal implementation**

Create `pdf.py`:

```python
import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import LongTable, Paragraph, SimpleDocTemplate, Spacer, TableStyle

HEADER_BG = colors.HexColor("#D9E2F3")
ALT_BG = colors.HexColor("#F2F5FA")
TOTAL_BG = colors.HexColor("#E2EFDA")

TITLE_STYLE = ParagraphStyle(
    "ReportTitle", fontName="Helvetica-Bold", fontSize=14, leading=17, spaceAfter=2
)
SUB_STYLE = ParagraphStyle(
    "ReportSub", fontName="Helvetica", fontSize=9, leading=12, spaceAfter=6
)


def render_pdf(df, title, out):
    doc = SimpleDocTemplate(
        out,
        pagesize=landscape(A4),
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title=title,
    )
    story = [
        Paragraph(title, TITLE_STYLE),
        Paragraph(f"Generated on {datetime.date.today():%d %b %Y}", SUB_STYLE),
        Spacer(1, 4 * mm),
    ]
    columns = df.columns.tolist()
    data = [columns]
    if len(df):
        data.extend(df.values.tolist())
        sums = df.sum(numeric_only=True)
        total_row = []
        for c in columns:
            if c in ("Packaging Supplier", "Supplier", "Factory"):
                total_row.append("Total" if c == "Packaging Supplier" else "")
            else:
                total_row.append(sums[c])
        data.append(total_row)

    table = LongTable(data, repeatRows=1)
    style = [
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("ALIGN", (3, 0), (-1, -1), "RIGHT"),
    ]
    if len(data) > 1:
        style.append(("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, ALT_BG]))
    if len(df):
        style.append(("BACKGROUND", (0, -1), (-1, -1), TOTAL_BG))
        style.append(("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"))
    table.setStyle(TableStyle(style))
    story.append(table)
    doc.build(story)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_pdf.py -v`
Expected: PASS (4 passed). If `test_pdf_has_total_row_and_title` fails on the `"90"` assertion, adjust the expected number in the test to match the actual grand total (5 rows × 3 months × 6 = 90) — verify by extracting the text.

- [ ] **Step 5: Run full suite**

Run: `python -m pytest -v`
Expected: PASS (29 passed)

- [ ] **Step 6: Commit**

```bash
git add pdf.py tests/test_pdf.py
git commit -m "feat: landscape A4 PDF export with repeating headers and total row"
```

---

### Task 6: Streamlit app, launcher, README, manual verification

**Files:**
- Create: `app.py`
- Create: `start.bat`
- Create: `README.md`

**Interfaces:**
- Consumes: `load_records` (loader.py), `build_table`, `clamp_duration`, `default_duration`, `month_label`, `packaging_suppliers`, `suppliers_for`, `factories_for` (filtering.py), `render_pdf` (pdf.py).
- Produces: the runnable app.

- [ ] **Step 1: Write the app**

Create `app.py`:

```python
import datetime
import io
from pathlib import Path

import streamlit as st

from filtering import (
    build_table,
    clamp_duration,
    default_duration,
    factories_for,
    month_label,
    packaging_suppliers,
    suppliers_for,
)
from loader import load_records
from pdf import render_pdf

WORKBOOK = Path(__file__).parent / "Sales Record V1.xlsx"


@st.cache_data(show_spinner="Reading workbook...")
def load(path_str):
    return load_records(path_str)


def main():
    st.set_page_config(page_title="Primark Carton Supply Report", layout="wide")
    st.title("Primark Carton Supply Report")

    refresh_col, info_col = st.columns([1, 4])
    with refresh_col:
        if st.button("Refresh data"):
            load.clear()
            st.rerun()
    with info_col:
        st.caption(f"Workbook: {WORKBOOK.name}")

    if not WORKBOOK.exists():
        st.error(f"Workbook not found at {WORKBOOK}. Place 'Sales Record V1.xlsx' next to app.py.")
        return
    try:
        records, warnings = load(str(WORKBOOK))
    except Exception as exc:
        st.error(f"Couldn't read the workbook. Is it open in Excel? ({exc})")
        return

    if warnings:
        st.warning("Skipped sheets (no month columns): " + ", ".join(sorted(warnings)))
    if not records:
        st.info("No data found in the workbook.")
        return

    duration = default_duration(records)
    from_default = datetime.date(duration[0][0], duration[0][1], 1)
    to_default = datetime.date(duration[1][0], duration[1][1], 1)

    f1, f2, f3, f4, f5 = st.columns([2, 3, 3, 2, 2])
    with f1:
        pkg = st.selectbox("Packaging Supplier", ["All"] + packaging_suppliers(records))
    pkg_sel = None if pkg == "All" else pkg

    with f2:
        sups = suppliers_for(records, pkg_sel)
        sup = st.selectbox("Supplier", ["All"] + sups)
    sup_sel = None if sup == "All" else sup

    with f3:
        facs = factories_for(records, pkg_sel, sup_sel)
        fac = st.selectbox("Factory", ["All"] + facs)
    fac_sel = None if fac == "All" else fac

    with f4:
        from_date = st.date_input(
            "From", value=from_default, min_value=datetime.date(2020, 1, 1), max_value=to_default
        )
    with f5:
        to_date = st.date_input(
            "To", value=to_default, min_value=datetime.date(2020, 1, 1), max_value=to_default
        )

    frm = (from_date.year, from_date.month)
    to = (to_date.year, to_date.month)
    if frm > to:
        frm, to = to, frm
    clamped_frm, clamped_to = clamp_duration(frm, to)
    if (clamped_frm, clamped_to) != (frm, to):
        st.caption(f"Range capped at 24 months — showing {month_label(clamped_frm)} to {month_label(clamped_to)}.")
    frm, to = clamped_frm, clamped_to

    df = build_table(
        records,
        packaging_supplier=pkg_sel,
        supplier=sup_sel,
        factory=fac_sel,
        frm=frm,
        to=to,
    )

    st.caption(
        f"{len(df):,} rows \u00b7 {len(df.columns) - 3} months "
        f"\u00b7 {month_label(frm)} to {month_label(to)}"
    )

    if df.empty:
        st.info("No data for the selected filters.")
        return

    st.dataframe(df, use_container_width=True, hide_index=True)

    title = (pkg if pkg_sel else "All suppliers") + f" \u2014 {month_label(frm)} to {month_label(to)}"
    buf = io.BytesIO()
    render_pdf(df, title, buf)
    st.download_button(
        "Export PDF",
        data=buf.getvalue(),
        file_name="carton-report.pdf",
        mime="application/pdf",
    )


main()
```

- [ ] **Step 2: Create launcher**

Create `start.bat`:

```bat
@echo off
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found. Install it from https://www.python.org/downloads/ and tick "Add to PATH".
    pause
    exit /b 1
)
python -m pip install -r requirements.txt --quiet
python -m streamlit run app.py
```

- [ ] **Step 3: Create README**

Create `README.md`:

```markdown
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
```

- [ ] **Step 4: Run unit suite**

Run: `python -m pytest -v`
Expected: PASS (29 passed)

- [ ] **Step 5: Manual verification against the real workbook**

Run: `python -m streamlit run app.py` (or double-click `start.bat`)

Verify:
- App loads, `M&U - 2020/2021`, `Uniglory - 2020/2021` appear in the skipped-sheets warning.
- Default range shows Aug-24 to Jul-26 (26 months → clamped to 24: Sep-24 to Jul-26) with 2026 and 2025 + 2024 columns; count rows, compare a known factory (e.g. "Union Sportswear Ltd PJT" under M&U) against the raw Excel sheet `M&U - 2026` sum.
- Selecting Packaging Supplier = M&U narrows the Supplier dropdown; Factory dropdown cascades.
- Selecting From = Jan-24, To = Jan-26 shows the "capped at 24 months" caption and 24 columns.
- Export PDF: multi-page, header repeated on every page, `Total` row present, landscape A4.

- [ ] **Step 6: Commit**

```bash
git add app.py start.bat README.md
git commit -m "feat: streamlit app with filters and PDF export"
```