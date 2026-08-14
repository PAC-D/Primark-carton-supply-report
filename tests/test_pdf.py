import io
from pathlib import Path

import pandas as pd
from pypdf import PdfReader

from filtering import month_label
from pdf import render_pdf


def make_df(rows=200, months=26):
    periods = [(2024 + (8 + i - 1) // 12, (8 + i - 1) % 12 + 1) for i in range(months)]
    cols = ["SL", "Packaging Supplier", "Supplier", "Factory"] + [month_label(p) for p in periods] + ["Total"]
    data = []
    for i in range(rows):
        vals = [i + 1] * months
        data.append([i + 1, "M&U", f"Supplier {i % 10}", f"Factory {i}", *vals, sum(vals)])
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
    # pypdf cannot round-trip the em-dash from the WinAnsi content stream, so
    # assert the title's extractable components rather than the exact string.
    assert "M&U" in text
    assert "M&U;" not in text
    assert "Jan-26 to Mar-26" in text
    assert "Generated on" in text
    # grand total = 5 rows x 3 months x (1+2+3+4+5) = 3 * 15 = 45
    assert "45" in text


def test_pdf_has_sl_column(tmp_path):
    path = tmp_path / "out.pdf"
    render_pdf(make_df(rows=5, months=3), "M&U — Jan-26 to Mar-26", path)
    text = "".join(page.extract_text() or "" for page in PdfReader(str(path)).pages)
    assert "SL" in text
    assert sum(1 for line in text.splitlines() if line == "15") == 4


def test_pdf_keeps_numeric_label_columns_unmangled(tmp_path):
    path = tmp_path / "out.pdf"
    df = make_df(rows=3, months=3)
    df.loc[0, "Factory"] = "12.5"
    render_pdf(df, "M&U — Jan-26 to Mar-26", path)
    text = "".join(page.extract_text() or "" for page in PdfReader(str(path)).pages)
    assert "12.5" in text
    assert not any(line == "12" for line in text.splitlines())


def test_pdf_empty_table(tmp_path):
    path = tmp_path / "empty.pdf"
    df = pd.DataFrame(columns=["Packaging Supplier", "Supplier", "Factory", "Jan-26", "Total"])
    render_pdf(df, "No data", path)
    assert path.exists()
    assert path.stat().st_size > 500


def test_pdf_accepts_file_object():
    buf = io.BytesIO()
    render_pdf(make_df(rows=10, months=4), "Test", buf)
    assert len(buf.getvalue()) > 1000


def test_pdf_page_width_grows_with_columns(tmp_path):
    p3 = tmp_path / "a.pdf"
    p12 = tmp_path / "b.pdf"
    render_pdf(make_df(rows=5, months=3), "T", p3)
    render_pdf(make_df(rows=5, months=12), "T", p12)
    w3 = PdfReader(str(p3)).pages[0].mediabox.width
    w12 = PdfReader(str(p12)).pages[0].mediabox.width
    assert w12 > w3


def test_pdf_no_decimals(tmp_path):
    path = tmp_path / "out.pdf"
    render_pdf(make_df(rows=5, months=3), "T", path)
    text = "".join(page.extract_text() or "" for page in PdfReader(str(path)).pages)
    assert ".0" not in text
    assert "45" in text