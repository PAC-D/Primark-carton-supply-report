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


from collections import defaultdict

import pandas as pd

BASE_COLUMNS = ["Packaging Supplier", "Supplier", "Factory"]


def _as_list(value):
    if value is None:
        return None
    if isinstance(value, str):
        return [value]
    return list(value)


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
    pkg_sel = _as_list(packaging_supplier)
    sup_sel = _as_list(supplier)
    fac_sel = _as_list(factory)
    agg = defaultdict(lambda: defaultdict(float))
    fk, tk = key_of(frm), key_of(to)
    for r in records:
        if pkg_sel and r.packaging_supplier not in pkg_sel:
            continue
        if sup_sel and r.supplier not in sup_sel:
            continue
        if fac_sel and r.factory not in fac_sel:
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
    pkg_sel = _as_list(packaging_supplier)
    return sorted({
        r.supplier for r in records
        if not pkg_sel or r.packaging_supplier in pkg_sel
    })


def factories_for(records, packaging_supplier=None, supplier=None):
    pkg_sel = _as_list(packaging_supplier)
    sup_sel = _as_list(supplier)
    return sorted({
        r.factory for r in records
        if (not pkg_sel or r.packaging_supplier in pkg_sel)
        and (not sup_sel or r.supplier in sup_sel)
    })