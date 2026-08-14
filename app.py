import datetime
import io
from pathlib import Path

import streamlit as st

from filtering import (
    build_table,
    default_duration,
    factories_for,
    key_of,
    month_label,
    packaging_suppliers,
    suppliers_for,
)
from loader import load_records

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
        pkg = st.multiselect(
            "Packaging Supplier", packaging_suppliers(records), placeholder="All"
        )
    pkg_sel = pkg or None

    with f2:
        sups = suppliers_for(records, pkg_sel)
        sup = st.multiselect("Supplier", sups, placeholder="All")
    sup_sel = sup or None

    with f3:
        facs = factories_for(records, pkg_sel, sup_sel)
        fac = st.multiselect("Factory", facs, placeholder="All")
    fac_sel = fac or None

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

    df = build_table(
        records,
        packaging_supplier=pkg_sel,
        supplier=sup_sel,
        factory=fac_sel,
        frm=frm,
        to=to,
    )

    st.caption(
        f"{len(df):,} rows \u00b7 {key_of(to) - key_of(frm) + 1} months "
        f"\u00b7 {month_label(frm)} to {month_label(to)}"
    )

    if df.empty:
        st.info("No data for the selected filters.")
        return

    st.dataframe(df, use_container_width=True, hide_index=True)

    title = (", ".join(pkg_sel) if pkg_sel else "All suppliers") + f" \u2014 {month_label(frm)} to {month_label(to)}"


main()
