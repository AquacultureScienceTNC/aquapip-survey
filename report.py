"""
AquaPIP — MEL Monitoring Survey
PDF report generator.

build_pdf(submission, matches, code2label) -> bytes

Mirrors the "Your MEL Protocol Recommendations" slide:
  * Page 1  = summary (survey responses + one recommended protocol per indicator)
  * Then     = one detail block per unique recommended protocol, including the
               method-summary field (AK) and the V2 download placeholders.

Cost/effort are drawn as vector circles so they render regardless of the fonts
available in the deployment environment.
"""

import io
import os
import re
import datetime

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable,
    KeepTogether, PageBreak, Flowable,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import mel_logic as ml

# --------------------------------------------------------------------------- #
# Fonts — register DejaVu for wide glyph coverage if present; else fall back.
# --------------------------------------------------------------------------- #
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

TEAL = colors.HexColor("#0B4F5C")
INK = colors.HexColor("#1A2B2F")
MUTED = colors.HexColor("#5B6B70")
HAIR = colors.HexColor("#D7DEE0")


def first_url(text):
    m = re.search(r"https?://\S+", text or "")
    return m.group(0).rstrip(" ;)") if m else ""


# --------------------------------------------------------------------------- #
# Vector cost/effort dots
# --------------------------------------------------------------------------- #
class Dots(Flowable):
    def __init__(self, filled, total=3, r=3.1, gap=9, color=TEAL):
        super().__init__()
        self.filled = max(0, min(total, filled))
        self.total = total
        self.r = r
        self.gap = gap
        self.color = color
        self.width = gap * total
        self.height = r * 2 + 1

    def draw(self):
        c = self.canv
        y = self.height / 2
        for i in range(self.total):
            x = self.r + i * self.gap
            c.setStrokeColor(self.color)
            c.setLineWidth(0.8)
            if i < self.filled:
                c.setFillColor(self.color)
                c.circle(x, y, self.r, stroke=1, fill=1)
            else:
                c.setFillColor(colors.white)
                c.circle(x, y, self.r, stroke=1, fill=1)


def _tier_swatch(hex_color, size=8):
    """A small filled square (tier colour) as a flowable."""
    class _Sw(Flowable):
        def __init__(self):
            super().__init__()
            self.width = size
            self.height = size

        def draw(self):
            self.canv.setFillColor(colors.HexColor(hex_color))
            self.canv.setStrokeColor(HAIR)
            self.canv.rect(0, 0, size, size, stroke=1, fill=1)
    return _Sw()


# --------------------------------------------------------------------------- #
# Styles
# --------------------------------------------------------------------------- #
def _styles():
    ss = getSampleStyleSheet()
    s = {}
    s["title"] = ParagraphStyle("t", parent=ss["Title"], fontName=BOLD_FONT,
                                textColor=TEAL, fontSize=22, leading=25,
                                spaceAfter=2, alignment=TA_LEFT)
    s["subtitle"] = ParagraphStyle("st", fontName=BODY_FONT, textColor=MUTED,
                                   fontSize=9.5, leading=13, spaceAfter=10)
    s["h2"] = ParagraphStyle("h2", fontName=BOLD_FONT, textColor=TEAL,
                             fontSize=11, leading=14, spaceBefore=10,
                             spaceAfter=5, tracking=0)
    s["label"] = ParagraphStyle("lb", fontName=BOLD_FONT, textColor=MUTED,
                                fontSize=7.5, leading=10)
    s["val"] = ParagraphStyle("vl", fontName=BODY_FONT, textColor=INK,
                              fontSize=9, leading=12)
    s["body"] = ParagraphStyle("bd", fontName=BODY_FONT, textColor=INK,
                               fontSize=9, leading=12.5, spaceAfter=3)
    s["small"] = ParagraphStyle("sm", fontName=BODY_FONT, textColor=MUTED,
                                fontSize=7.7, leading=10)
    s["proto"] = ParagraphStyle("pr", fontName=BOLD_FONT, textColor=INK,
                                fontSize=10, leading=12.5)
    s["code"] = ParagraphStyle("cd", fontName=BOLD_FONT, textColor=TEAL,
                               fontSize=8.5, leading=11)
    s["footer"] = ParagraphStyle("ft", fontName=BODY_FONT, textColor=MUTED,
                                 fontSize=7, leading=9)
    return s


def _tier_pill(rec, S):
    """A small coloured tier badge as a one-cell table."""
    txt = Paragraph(
        f"<font color='#1A2B2F'><b>T{rec['tier']} · {rec['tier_label']}</b></font>"
        if rec["tier"] else "<b>Unrated</b>", S["small"])
    t = Table([[txt]], colWidths=[1.15 * inch])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(rec["tier_color"])),
        ("BOX", (0, 0), (-1, -1), 0.5, HAIR),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]))
    return t


# --------------------------------------------------------------------------- #
# Main builder
# --------------------------------------------------------------------------- #
def build_pdf(submission, matches, code2label):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=LETTER,
        leftMargin=0.7 * inch, rightMargin=0.7 * inch,
        topMargin=0.6 * inch, bottomMargin=0.6 * inch,
        title="MEL Protocol Recommendations",
    )
    S = _styles()
    story = []

    farm = submission.get("farm_name") or "Your farm"
    date = datetime.date.today().strftime("%d %b %Y")
    story.append(Paragraph("Your MEL Protocol Recommendations", S["title"]))
    story.append(Paragraph(
        f"{farm} &nbsp;·&nbsp; generated {date} &nbsp;·&nbsp; "
        f"from your Regenerative Aquaculture MEL survey", S["subtitle"]))
    story.append(HRFlowable(width="100%", thickness=1, color=TEAL,
                            spaceBefore=0, spaceAfter=8))

    # ---- Survey responses ------------------------------------------------- #
    story.append(Paragraph("YOUR SURVEY RESPONSES", S["h2"]))
    prof = submission.get("profile", {})
    prof_pairs = []
    if submission.get("aqua_type"):
        prof_pairs.append(("Aquaculture type", submission["aqua_type"]))
    for fkey, label, _vh, _rk in ml.PROFILE_FIELDS:
        if fkey == "aqua_type":
            continue
        v = prof.get(fkey)
        if v:
            prof_pairs.append((label, v))
    # 2-column grid of label/value
    grid = []
    for i in range(0, len(prof_pairs), 2):
        left = prof_pairs[i]
        right = prof_pairs[i + 1] if i + 1 < len(prof_pairs) else ("", "")
        grid.append([
            Paragraph(left[0].upper(), S["label"]),
            Paragraph(str(left[1]), S["val"]),
            Paragraph(right[0].upper(), S["label"]),
            Paragraph(str(right[1]), S["val"]),
        ])
    if grid:
        t = Table(grid, colWidths=[1.15 * inch, 2.0 * inch, 1.15 * inch, 2.0 * inch])
        t.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ]))
        story.append(t)

    goals = submission.get("goals") or ml.goals_from_codes(submission.get("codes", []))
    story.append(Spacer(1, 4))
    story.append(Paragraph("MEL GOALS", S["label"]))
    story.append(Paragraph(" · ".join(goals) if goals else "—", S["val"]))
    if submission.get("farmer_goals"):
        story.append(Spacer(1, 3))
        story.append(Paragraph("INTERESTED IN MEASURING", S["label"]))
        story.append(Paragraph(" · ".join(submission["farmer_goals"]), S["val"]))

    # ---- Recommended protocols summary ------------------------------------ #
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=0.5, color=HAIR, spaceAfter=6))
    n_sel = len(submission.get("codes", []))
    story.append(Paragraph(
        f"RECOMMENDED FARMER-FRIENDLY PROTOCOLS &nbsp; "
        f"<font size=8 color='#5B6B70'>· one best practitioner-runnable "
        f"protocol per indicator ({n_sel} selected)</font>", S["h2"]))

    header = [Paragraph(x, S["label"]) for x in
              ["INDICATOR", "TIER", "PROTOCOL", "COST", "EFFORT"]]
    data = [header]
    row_styles = []
    r_i = 1
    for code, cands in matches.items():
        label = code2label.get(code, code)
        if not cands:
            data.append([
                Paragraph(label, S["code"]),
                Paragraph("—", S["small"]),
                Paragraph("<i>No protocol in the current database for this "
                          "indicator yet.</i>", S["small"]),
                Paragraph("—", S["small"]), Paragraph("—", S["small"]),
            ])
            row_styles.append(("TEXTCOLOR", (0, r_i), (-1, r_i), MUTED))
            r_i += 1
            continue
        rec = cands[0]["row"]
        badge = " ★" if cands[0]["badge"] else ""
        data.append([
            Paragraph(label + badge, S["code"]),
            _tier_pill(rec, S),
            Paragraph(_protocol_title(rec), S["val"]),
            Dots(rec["cost_ord"], color=TEAL),
            Dots(rec["effort_ord"], color=TEAL),
        ])
        r_i += 1

    tbl = Table(data, colWidths=[1.25 * inch, 1.25 * inch, 2.7 * inch,
                                 0.5 * inch, 0.5 * inch], repeatRows=1)
    base_style = [
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, 0), (-1, 0), 0.6, TEAL),
        ("LINEBELOW", (0, 1), (-1, -2), 0.3, HAIR),
    ]
    tbl.setStyle(TableStyle(base_style + row_styles))
    story.append(tbl)

    # legend
    story.append(Spacer(1, 5))
    legend = ("Colour = who can run it:  "
              "T4 Practitioner · T3 Partnership · T2 Researcher · T1 Reference.   "
              "Dots ●○○→●●● = cost / effort (low→high).   ★ = strong fit to your farm.")
    story.append(Paragraph(legend, S["small"]))

    # ---- Detail pages ----------------------------------------------------- #
    story.append(PageBreak())
    story.append(Paragraph("Protocol details", S["title"]))
    story.append(Paragraph("Full method, equipment and sampling plan for each "
                           "recommended protocol.", S["subtitle"]))
    story.append(HRFlowable(width="100%", thickness=1, color=TEAL, spaceAfter=8))

    # unique protocols across the top picks, remembering which indicators each covers
    seen = {}
    order = []
    for code, cands in matches.items():
        if not cands:
            continue
        rec = cands[0]["row"]
        rid = rec["row"]
        if rid not in seen:
            seen[rid] = {"rec": rec, "codes": [code], "n_alt": len(cands)}
            order.append(rid)
        else:
            seen[rid]["codes"].append(code)

    for rid in order:
        blk = seen[rid]
        story.append(_detail_block(blk["rec"], blk["codes"], blk["n_alt"],
                                   code2label, S))
        story.append(Spacer(1, 6))
        story.append(HRFlowable(width="100%", thickness=0.4, color=HAIR,
                                spaceAfter=8))

    story.append(Spacer(1, 4))
    story.append(Paragraph(
        "Source: TNC Regenerative Aquaculture MEL literature database "
        "(data/mel_database.xlsx) · Usability Tier System, Skill v1.5. "
        "Recommendations regenerate automatically when the database is updated.",
        S["footer"]))

    doc.build(story)
    return buf.getvalue()


def _protocol_title(rec):
    """Prefer a concise protocol title; fall back to paper title."""
    return rec["title"] or "(untitled)"


def _detail_block(rec, codes, n_alt, code2label, S):
    parts = []
    code_str = " · ".join(codes)
    parts.append(Paragraph(code_str, S["code"]))
    parts.append(Paragraph(_protocol_title(rec), S["proto"]))
    meta = " · ".join(x for x in [rec["authors"], rec["year"], rec["publication"]] if x)
    if meta:
        parts.append(Paragraph(meta, S["small"]))
    parts.append(Spacer(1, 3))

    # tier + cost + effort line
    tline = Table([[
        _tier_pill(rec, S),
        Paragraph(f"<b>Skill:</b> {rec['skill']}", S["small"]),
        Paragraph("<b>Cost:</b> " + (rec["cost_label"] or "—"), S["small"]),
        Dots(rec["cost_ord"], color=TEAL),
        Paragraph("<b>Effort:</b> " + (rec["effort_label"] or "—"), S["small"]),
        Dots(rec["effort_ord"], color=TEAL),
    ]], colWidths=[1.2 * inch, 1.0 * inch, 1.0 * inch, 0.5 * inch,
                   1.0 * inch, 0.5 * inch])
    tline.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                               ("LEFTPADDING", (0, 0), (-1, -1), 0),
                               ("TOPPADDING", (0, 0), (-1, -1), 1),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
    parts.append(tline)

    def field(lbl, val, style=None):
        if not val:
            return
        parts.append(Paragraph(lbl.upper(), S["label"]))
        parts.append(Paragraph(str(val), style or S["body"]))

    field("What it measures", rec["measures"])

    # Method summary = AK. If empty, say pending (V2 will fill it).
    if rec["method_summary"]:
        field("Method summary", rec["method_summary"])
    else:
        parts.append(Paragraph("METHOD SUMMARY", S["label"]))
        parts.append(Paragraph("<i>Method summary pending — will appear once the "
                               "database includes the Protocol Method Summary "
                               "field.</i>", S["small"]))

    field("Equipment", rec["equipment"])
    field("Statistical approach", rec["stats"])
    if rec["exec_practitioner"]:
        field("Executable by practitioner?", rec["exec_practitioner"])
    field("Notes / caveats", rec["notes"])

    url = first_url(rec["url"])
    if url:
        parts.append(Paragraph("SOURCE", S["label"]))
        parts.append(Paragraph(
            f'<a href="{url}" color="#0B4F5C"><u>{url}</u></a>', S["small"]))

    # ---- V2 download placeholders ---------------------------------------- #
    parts.append(Spacer(1, 4))
    parts.append(Paragraph("DOWNLOADS", S["label"]))
    dl_rows = []
    for lbl, val in [("Normalized protocol", rec["url_template"]),
                     ("Field data sheet", rec["url_datasheet"]),
                     ("Stats workbook (Excel)", rec["url_workbook"])]:
        u = first_url(val)
        if u:
            cell = Paragraph(f'{lbl}: <a href="{u}" color="#0B4F5C"><u>download</u></a>',
                             S["small"])
        else:
            cell = Paragraph(f'{lbl}: <font color="#9AA7AB">coming in V2 '
                             f'(not yet available)</font>', S["small"])
        dl_rows.append([cell])
    dlt = Table(dl_rows, colWidths=[5.0 * inch])
    dlt.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0),
                             ("TOPPADDING", (0, 0), (-1, -1), 1),
                             ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
    parts.append(dlt)

    if n_alt > 1:
        parts.append(Spacer(1, 2))
        parts.append(Paragraph(
            f"<i>{n_alt - 1} additional protocol(s) in the database also match "
            f"the indicator(s) above.</i>", S["small"]))

    return KeepTogether(parts)


# --------------------------------------------------------------------------- #
# Self-test
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    rows, vocab, headers = ml.load_database(os.path.join(here, "data", "mel_database.xlsx"))
    groups = ml.build_indicator_options(vocab.get("MEL Indicator", []), rows)
    code2label, label2code, code2goal = ml.flatten_options(groups)
    profile = {"aqua_type": "Seaweed", "species": "Saccharina (sugar kelp)",
               "algae_type": "Brown", "farm_structure": "Longline (surface)",
               "farm_scale": "Medium (1–20 ha)", "coculture": "Monoculture",
               "climate_zone": "Temperate Atlantic",
               "site_depth": "Shallow nearshore (2–10 m)",
               "wave_energy": "Semi-exposed",
               "water_clarity": "Clear (visual methods feasible)"}
    codes = ["H&B 1.1.1", "WQ 1.2.1", "WQ 1.1.1", "CC 2.1.1", "CC 1.1.1"]
    matches = ml.match_protocols(rows, codes, profile)
    submission = {"farm_name": "Tidewater Kelp Co. · Casco Bay, Maine",
                  "profile": profile, "codes": codes,
                  "farmer_goals": ["Blue Carbon — ecosystem-positive farming"],
                  "goals": ml.goals_from_codes(codes)}
    pdf = build_pdf(submission, matches, code2label)
    out = os.path.join(here, "_sample_report.pdf")
    with open(out, "wb") as f:
        f.write(pdf)
    print(f"Wrote {out} ({len(pdf):,} bytes)")
