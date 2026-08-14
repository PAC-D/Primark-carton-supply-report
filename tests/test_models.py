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
