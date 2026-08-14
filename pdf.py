import datetime
import os
from xml.sax.saxutils import escape

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
    if isinstance(out, os.PathLike):
        out = os.fspath(out)
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
        Paragraph(escape(title), TITLE_STYLE),
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