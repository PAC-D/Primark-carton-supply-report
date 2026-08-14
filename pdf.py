import datetime
import os
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import LongTable, Paragraph, SimpleDocTemplate, Spacer, TableStyle

HEADER_BG = colors.HexColor("#D9E2F3")
ALT_BG = colors.HexColor("#F2F5FA")
TOTAL_BG = colors.HexColor("#E2EFDA")

FONT_SIZE = 8
MONTH_COL_WIDTH = 16 * mm

TITLE_STYLE = ParagraphStyle(
    "ReportTitle", fontName="Helvetica-Bold", fontSize=14, leading=17, spaceAfter=2
)
SUB_STYLE = ParagraphStyle(
    "ReportSub", fontName="Helvetica", fontSize=9, leading=12, spaceAfter=6
)


def _fmt(value):
    try:
        return f"{int(round(float(value)))}"
    except (TypeError, ValueError):
        return str(value)


def _label_width(values):
    longest = max((str(v) for v in values), key=len, default="")
    return min(max(len(longest) * 5.5 + 10, 20 * mm), 70 * mm)


def render_pdf(df, title, out):
    if isinstance(out, os.PathLike):
        out = os.fspath(out)
    columns = df.columns.tolist()
    data = [columns]
    if len(df):
        for row in df.values.tolist():
            data.append(list(row[:3]) + [_fmt(v) for v in row[3:]])
        sums = df.sum(numeric_only=True)
        total_row = []
        for c in columns:
            if c in ("Packaging Supplier", "Supplier", "Factory"):
                total_row.append("Total" if c == "Packaging Supplier" else "")
            else:
                total_row.append(_fmt(sums[c]))
        data.append(total_row)

    widths = (
        [_label_width([row[i] for row in data]) for i in range(3)]
        + [MONTH_COL_WIDTH] * (len(columns) - 3)
    )
    doc = SimpleDocTemplate(
        out,
        pagesize=(sum(widths) + 20 * mm, A4[1]),
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title=title,
    )
    story = [
        Paragraph(escape(title), TITLE_STYLE),
        Paragraph(f"Generated on {datetime.date.today():%d %b %Y}", SUB_STYLE),
        Spacer(1, 4 * mm),
    ]

    table = LongTable(data, repeatRows=1, colWidths=widths)
    style = [
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), HEADER_BG),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), FONT_SIZE),
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