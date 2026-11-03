"""
TNC Filtration Service Estimator — reference exports.

Two outputs, both built live from the workbook's References sheet:
  * references_csv(refs)  -> bytes   (Citation, Full reference, Link, Free copy)
  * references_pdf(refs)  -> bytes   (formatted reference list, AquaPIP styling)

Each reference's main link is its DOI where one exists; "Free copy" is an
optional second link (a free version of a paper whose DOI lands on a paywall),
read from an optional "Free link" column in the References sheet.

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
from reportlab.platypus import SimpleDocTemplate, Paragraph, HRFlowable
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Same palette as AquaPIP's report.py
TEAL = colors.HexColor("#0B4F5C")
INK = colors.HexColor("#1A2B2F")
MUTED = colors.HexColor("#5B6B70")

BODY_FONT, BOLD_FONT = "Helvetica", "Helvetica-Bold"
_DEJAVU = ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
           "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]
try:
    if all(os.path.exists(p) for p in _DEJAVU):
        pdfmetrics.registerFont(TTFont("DejaVu", _DEJAVU[0]))
        pdfmetrics.registerFont(TTFont("DejaVu-Bold", _DEJAVU[1]))
        BODY_FONT, BOLD_FONT = "DejaVu", "DejaVu-Bold"
except Exception:
    pass

GAP = 10      # space between one reference and the next (pt)
TIGHT = 1.5   # space between a reference and its own link(s) (pt)


def _esc(s):
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _free(r):
    return getattr(r, "free_link", "") or ""


def references_csv(refs):
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["Citation", "Full reference", "Link", "Free copy"])
    for r in refs:
        w.writerow([r.cite, r.full, r.link, _free(r)])
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
                          fontSize=9.5, leading=13, spaceAfter=GAP,
                          leftIndent=14, firstLineIndent=-14)
    link = ParagraphStyle("lk", fontName=BODY_FONT, textColor=TEAL,
                          fontSize=8.5, leading=11, spaceAfter=GAP, leftIndent=14)

    story = [Paragraph("Filtration Service Estimator — References", title)]
    date = datetime.date.today().strftime("%d %b %Y")
    story.append(Paragraph(
        "Sources for the filtration equations and length-to-weight conversions. "
        f"The Nature Conservancy, {date}.", sub))
    story.append(HRFlowable(width="100%", thickness=1, color=TEAL, spaceAfter=10))

    for r in sorted(refs, key=lambda x: x.cite.lower()):
        links = [("", r.link)] if r.link else []
        if _free(r):
            links.append(("Free copy: ", _free(r)))
        # The reference sits tight against its own link(s); the gap comes AFTER
        # the last link, so each link reads as part of the reference above it.
        st_item = ParagraphStyle("itx", parent=item,
                                 spaceAfter=TIGHT if links else GAP)
        story.append(Paragraph(_esc(r.full) or _esc(r.cite), st_item))
        for i, (prefix, url) in enumerate(links):
            last = i == len(links) - 1
            st_link = ParagraphStyle("lkx", parent=link,
                                     spaceAfter=GAP if last else TIGHT)
            story.append(Paragraph(
                f'{prefix}<a href="{_esc(url)}" color="#0B4F5C"><u>{_esc(url)}</u></a>',
                st_link))

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
