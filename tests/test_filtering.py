from filtering import (
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


import pandas as pd

from filtering import (
    build_table,
    factories_for,
    packaging_suppliers,
    suppliers_for,
)


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
    assert df.columns.tolist() == ["SL", "Packaging Supplier", "Supplier", "Factory", "Jan-26", "Feb-26", "Total"]
    assert len(df) == 3
    assert df["SL"].tolist() == [1, 2, 3]
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


def test_build_table_multiple_filters():
    df = build_table(recs(), packaging_supplier=["M&U", "Union"],
                     frm=(2026, 1), to=(2026, 2))
    assert len(df) == 3
    df = build_table(recs(), factory=["Aspire", "APS"], frm=(2026, 1), to=(2026, 2))
    assert len(df) == 2
    df = build_table(recs(), packaging_supplier=[], frm=(2026, 1), to=(2026, 2))
    assert len(df) == 3


def test_build_table_month_range():
    df = build_table(recs(), frm=(2026, 2), to=(2026, 2))
    assert df.columns.tolist() == ["SL", "Packaging Supplier", "Supplier", "Factory", "Feb-26", "Total"]
    row = df[df["Factory"] == "Aspire"].iloc[0]
    assert row["Feb-26"] == 5
    assert row["Total"] == 5


def test_build_table_default_duration():
    df = build_table(recs())
    assert df.columns.tolist()[4] == "Jan-26"
    assert df.columns.tolist()[-2] == "Feb-26"
    assert df.columns.tolist()[-1] == "Total"
    assert len(df.columns) == 7


def test_build_table_empty_records():
    df = build_table([])
    assert df.empty
    assert df.columns.tolist() == ["SL", "Packaging Supplier", "Supplier", "Factory", "Total"]


def test_build_table_month_beyond_data_is_zero():
    df = build_table(recs(), frm=(2026, 3), to=(2026, 3))
    assert df.empty
    assert df.columns.tolist() == ["SL", "Packaging Supplier", "Supplier", "Factory", "Mar-26", "Total"]


def test_option_lists():
    rs = recs()
    assert packaging_suppliers(rs) == ["M&U", "Union"]
    assert suppliers_for(rs, "M&U") == ["PADMA", "TEX"]
    assert suppliers_for(rs) == ["CENTRO", "PADMA", "TEX"]
    assert factories_for(rs, "M&U", "PADMA") == ["Aspire"]
    assert factories_for(rs, "Union") == ["APS", "Zero Row"]


def test_option_lists_multiple():
    rs = recs()
    assert suppliers_for(rs, ["M&U", "Union"]) == ["CENTRO", "PADMA", "TEX"]
    assert suppliers_for(rs, []) == ["CENTRO", "PADMA", "TEX"]
    assert factories_for(rs, ["M&U", "Union"], ["CENTRO"]) == ["APS", "Zero Row"]