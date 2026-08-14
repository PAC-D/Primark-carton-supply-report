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