"""
TNC Filtration Service Estimator — reference exports.

Two outputs, both built live from the workbook's References sheet:
  * references_csv(refs)  -> bytes   (Citation, Full reference, Link)
  * references_pdf(refs)  -> bytes   (formatted bibliography, AquaPIP styling)

Kept Streamlit-independent so it can be tested on its own.
"""

import io
import os
import csv
import datetime

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, HRFlowable,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Same palette as AquaPIP's report.py
TEAL = colors.HexColor("#0B4F5C")
INK = colors.HexColor("#1A2B2F")
MUTED = colors.HexColor("#5B6B70")

BODY_FONT, BOLD_FONT = "Helvetica", "Helvetica-Bold"
for base, bold, paths in [
    ("DejaVu", "DejaVu-Bold",
     ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
      "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]),
]:
    try:
        if all(os.path.exists(p) for p in paths):
            pdfmetrics.registerFont(TTFont(base, paths[0]))
            pdfmetrics.registerFont(TTFont(bold, paths[1]))
            BODY_FONT, BOLD_FONT = base, bold
    except Exception:
        pass


def _esc(s):
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def references_csv(refs):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Citation", "Full reference", "Link"])
    for r in refs:
        w.writerow([r.cite, r.full, r.link])
    return buf.getvalue().encode("utf-8-sig")


def references_pdf(refs):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=LETTER,
        leftMargin=0.75 * inch, rightMargin=0.75 * inch,
        topMargin=0.6 * inch, bottomMargin=0.6 * inch,
        title="Filtration Service Estimator — References",
    )
    ss = getSampleStyleSheet()
    title = ParagraphStyle("t", parent=ss["Title"], fontName=BOLD_FONT,
                           textColor=TEAL, fontSize=20, leading=23,
                           spaceAfter=2, alignment=TA_LEFT)
    sub = ParagraphStyle("st", fontName=BODY_FONT, textColor=MUTED,
                         fontSize=9.5, leading=13, spaceAfter=10)
    item = ParagraphStyle("it", fontName=BODY_FONT, textColor=INK,
                          fontSize=9.5, leading=13, spaceAfter=9,
                          leftIndent=14, firstLineIndent=-14)
    link = ParagraphStyle("lk", fontName=BODY_FONT, textColor=TEAL,
                          fontSize=8.5, leading=11, spaceAfter=2, leftIndent=14)
    foot = ParagraphStyle("ft", fontName=BODY_FONT, textColor=MUTED,
                          fontSize=7.5, leading=10)

    story = [Paragraph("Filtration Service Estimator — References", title)]
    date = datetime.date.today().strftime("%d %b %Y")
    story.append(Paragraph(
        f"Clearance-rate and length-weight sources &nbsp;·&nbsp; "
        f"generated {date} &nbsp;·&nbsp; The Nature Conservancy", sub))
    story.append(HRFlowable(width="100%", thickness=1, color=TEAL, spaceAfter=10))

    for r in sorted(refs, key=lambda x: x.cite.lower()):
        full = _esc(r.full) or _esc(r.cite)
        story.append(Paragraph(full, item))
        if r.link:
            story.append(Paragraph(
                f'<a href="{r.link}" color="#0B4F5C"><u>{_esc(r.link)}</u></a>', link))

    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=0.4,
                            color=colors.HexColor("#D7DEE0"), spaceAfter=6))
    story.append(Paragraph(
        "Source: TNC Filtration Service Estimator workbook "
        "(data/Clearance_rate_estimation_tool_TNC.xlsx). This reference list "
        "regenerates automatically when the workbook is updated.", foot))

    doc.build(story)
    return buf.getvalue()


if __name__ == "__main__":
    import filtration_logic as fl
    here = os.path.dirname(os.path.abspath(__file__))
    b = fl.load_workbook_bundle(os.path.join(here, "data", fl.DEFAULT_WORKBOOK))
    with open(os.path.join(here, "_refs_test.pdf"), "wb") as f:
        f.write(references_pdf(b["references"]))
    with open(os.path.join(here, "_refs_test.csv"), "wb") as f:
        f.write(references_csv(b["references"]))
    print(f"Wrote _refs_test.pdf and _refs_test.csv for {len(b['references'])} refs")
