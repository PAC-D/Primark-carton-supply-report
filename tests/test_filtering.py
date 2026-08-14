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