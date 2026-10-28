"""
AquaPIP - indicator pictures.

The pictures live in assets/indicators/<name>.svg (one per sector x indicator; the name is
the 'Picture' column of the Readiness sheet). Edit or replace a file and the app and the
PDFs pick it up on the next run.

svg_text(name)            -> SVG markup for the page (or "" if the file is missing)
drawing(name, width_pt)   -> reportlab Drawing for PDFs (or None)

The PDF converter handles the subset the pictures use (rect, circle, ellipse, path with
M/L/H/V/C/Q/T/A/Z, text, g, inline style or presentation attributes), so no extra
dependency is needed. Curves are flattened to short line segments.
"""

import math
import os
import re
from xml.etree import ElementTree as ET

ART_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "indicators")


def svg_path(name):
    return os.path.join(ART_DIR, name + ".svg") if name else ""


def svg_text(name):
    p = svg_path(name)
    if not p or not os.path.exists(p):
        return ""
    with open(p, encoding="utf-8") as f:
        return f.read()


# --------------------------------------------------------------------------- #
# SVG path -> polylines
# --------------------------------------------------------------------------- #
_TOK = re.compile(r"[MmLlHhVvQqTtCcAaZz]|-?\d*\.?\d+(?:[eE]-?\d+)?")


def _arc(x1, y1, rx, ry, phi, fa, fs, x2, y2, n=16):
    if rx == 0 or ry == 0:
        return [(x2, y2)]
    phi = math.radians(phi)
    cp, sp = math.cos(phi), math.sin(phi)
    dx, dy = (x1 - x2) / 2, (y1 - y2) / 2
    x1p, y1p = cp * dx + sp * dy, -sp * dx + cp * dy
    lam = x1p ** 2 / rx ** 2 + y1p ** 2 / ry ** 2
    if lam > 1:
        rx, ry = rx * math.sqrt(lam), ry * math.sqrt(lam)
    num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p
    den = rx * rx * y1p * y1p + ry * ry * x1p * x1p
    co = math.sqrt(max(0.0, num / den)) if den else 0.0
    if fa == fs:
        co = -co
    cxp, cyp = co * rx * y1p / ry, -co * ry * x1p / rx
    cx = cp * cxp - sp * cyp + (x1 + x2) / 2
    cy = sp * cxp + cp * cyp + (y1 + y2) / 2

    def ang(ux, uy, vx, vy):
        return math.atan2(ux * vy - uy * vx, ux * vx + uy * vy)
    t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
    dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry)
    if not fs and dt > 0:
        dt -= 2 * math.pi
    elif fs and dt < 0:
        dt += 2 * math.pi
    return [(cx + rx * math.cos(t1 + dt * i / n) * cp - ry * math.sin(t1 + dt * i / n) * sp,
             cy + rx * math.cos(t1 + dt * i / n) * sp + ry * math.sin(t1 + dt * i / n) * cp) for i in range(1, n + 1)]


def path_polylines(d):
    """-> [(points, closed)] in SVG coordinates."""
    t = _TOK.findall(d or "")
    i, subs, cur, start = 0, [], None, (0.0, 0.0)
    x = y = 0.0
    cmd, lastc = None, None

    def num():
        nonlocal i
        v = float(t[i]); i += 1
        return v
    while i < len(t):
        if t[i].isalpha():
            cmd = t[i]; i += 1
            if cmd in "Zz":
                if cur:
                    subs.append((cur, True))
                cur, (x, y), lastc = None, start, None
                continue
        if cmd is None:
            break
        rel, c = cmd.islower(), cmd.upper()
        if c == "M":
            if cur:
                subs.append((cur, False))
            nx, ny = num(), num()
            x, y = (x + nx, y + ny) if rel else (nx, ny)
            cur, start, lastc = [(x, y)], (x, y), None
            cmd = "l" if rel else "L"
        elif c in "LHV":
            if c == "L":
                nx, ny = num(), num(); x, y = (x + nx, y + ny) if rel else (nx, ny)
            elif c == "H":
                nx = num(); x = x + nx if rel else nx
            else:
                ny = num(); y = y + ny if rel else ny
            cur.append((x, y)); lastc = None
        elif c in "QT":
            if c == "Q":
                a, b = num(), num()
                qx, qy = (x + a, y + b) if rel else (a, b)
            else:
                qx, qy = (2 * x - lastc[0], 2 * y - lastc[1]) if lastc else (x, y)
            nx, ny = num(), num()
            ex, ey = (x + nx, y + ny) if rel else (nx, ny)
            for k in range(1, 11):
                s = k / 10.0
                cur.append(((1 - s) ** 2 * x + 2 * (1 - s) * s * qx + s * s * ex,
                            (1 - s) ** 2 * y + 2 * (1 - s) * s * qy + s * s * ey))
            lastc, x, y = (qx, qy), ex, ey
        elif c == "C":
            v = [num() for _ in range(6)]
            if rel:
                v = [v[0] + x, v[1] + y, v[2] + x, v[3] + y, v[4] + x, v[5] + y]
            for k in range(1, 13):
                s = k / 12.0
                cur.append(((1 - s) ** 3 * x + 3 * (1 - s) ** 2 * s * v[0] + 3 * (1 - s) * s * s * v[2] + s ** 3 * v[4],
                            (1 - s) ** 3 * y + 3 * (1 - s) ** 2 * s * v[1] + 3 * (1 - s) * s * s * v[3] + s ** 3 * v[5]))
            x, y, lastc = v[4], v[5], None
        elif c == "A":
            rx, ry, phi, fa, fs, nx, ny = [num() for _ in range(7)]
            ex, ey = (x + nx, y + ny) if rel else (nx, ny)
            cur += _arc(x, y, rx, ry, phi, int(fa), int(fs), ex, ey)
            x, y, lastc = ex, ey, None
        else:
            i += 1
    if cur:
        subs.append((cur, False))
    return subs


# --------------------------------------------------------------------------- #
# SVG -> reportlab Drawing
# --------------------------------------------------------------------------- #
def _attrs(el, inherited):
    a = dict(inherited)
    for k, v in el.attrib.items():
        if k != "style":
            a[k] = v
    for decl in (el.get("style") or "").split(";"):
        if ":" in decl:
            k, v = decl.split(":", 1)
            a[k.strip()] = v.strip()
    return a


def drawing(name, width_pt):
    """Picture `name` as a reportlab Drawing `width_pt` wide, or None."""
    txt = svg_text(name)
    if not txt:
        return None
    try:
        from reportlab.graphics.shapes import Drawing, Group, Path, Rect, Ellipse, String
        from reportlab.lib import colors
        root = ET.fromstring(txt)
    except Exception:
        return None
    vb = [float(v) for v in (root.get("viewBox") or "0 0 320 150").split()]
    vw, vh = vb[2], vb[3]
    s = width_pt / vw
    H = vh * s
    d = Drawing(width_pt, H)
    shapes = Group()
    texts = []

    def col(v):
        if not v or v == "none":
            return None
        try:
            return colors.HexColor(v) if v.startswith("#") else colors.toColor(v)
        except Exception:
            return None

    def paint(shp, a):
        f = col(a.get("fill", "#000000"))
        shp.fillColor = f
        if f is not None and a.get("fill-opacity"):
            shp.fillOpacity = float(a["fill-opacity"])
        sc = col(a.get("stroke", "none"))
        shp.strokeColor = sc
        if sc is not None:
            sw = float(str(a.get("stroke-width", "1")).replace("px", ""))
            shp.strokeWidth = sw * s
            if a.get("stroke-linecap") == "round":
                shp.strokeLineCap = 1
            if a.get("stroke-linejoin") == "round":
                shp.strokeLineJoin = 1
            if a.get("stroke-dasharray"):
                da = [float(v) * s for v in a["stroke-dasharray"].replace(",", " ").split()]
                shp.strokeDashArray = [max(0.01, da[0]), da[1]] if len(da) >= 2 else None
        return shp

    X = lambda v: float(v) * s
    Y = lambda v: H - float(v) * s

    def walk(node, inherited):
        for el in node:
            tag = el.tag.split("}")[-1]
            a = _attrs(el, inherited)
            if tag == "g":
                walk(el, {k: v for k, v in a.items() if k not in ("x", "y", "d", "cx", "cy", "r", "rx", "ry", "width", "height")})
            elif tag == "rect":
                x, y, w, h = (float(a.get(k, 0)) for k in ("x", "y", "width", "height"))
                rx = float(a.get("rx", 0))
                shapes.add(paint(Rect(X(x), Y(y + h), w * s, h * s, rx=rx * s, ry=rx * s), a))
            elif tag in ("circle", "ellipse"):
                rx = float(a.get("r", a.get("rx", 0))); ry = float(a.get("r", a.get("ry", 0)))
                shapes.add(paint(Ellipse(X(a["cx"]), Y(a["cy"]), rx * s, ry * s), a))
            elif tag == "path":
                p = Path()
                for pts, closed in path_polylines(a.get("d")):
                    if len(pts) < 2:
                        continue
                    p.moveTo(X(pts[0][0]), Y(pts[0][1]))
                    for px, py in pts[1:]:
                        p.lineTo(X(px), Y(py))
                    if closed:
                        p.closePath()
                if a.get("fill", "#000000") == "none":
                    a = dict(a); a["fill"] = "none"
                shapes.add(paint(p, a))
            elif tag == "text":
                size = float(str(a.get("font-size", "12")).replace("px", "")) * s
                anchor = {"middle": "middle", "end": "end"}.get(a.get("text-anchor"), "start")
                bold = str(a.get("font-weight", "")) in ("700", "bold")
                texts.append(String(X(a.get("x", 0)), Y(a.get("y", 0)), el.text or "",
                                    fontName="Helvetica-Bold" if bold else "Helvetica", fontSize=size,
                                    fillColor=col(a.get("fill", "#000000")), textAnchor=anchor))
    walk(root, {})
    # use the report's DejaVu fonts when registered (they carry the subscript in "O\u2082");
    # otherwise fall back to Helvetica and plain digits
    try:
        from reportlab.pdfbase import pdfmetrics
        pdfmetrics.getFont("DejaVu-Bold")
        bold_font, body_font, plain = "DejaVu-Bold", "DejaVu", False
    except Exception:
        bold_font, body_font, plain = "Helvetica-Bold", "Helvetica", True
    for t in texts:
        t.fontName = bold_font if t.fontName.endswith("Bold") else body_font
        if plain:
            t.text = t.text.replace("\u2082", "2")
    d.add(shapes)
    for t in texts:
        d.add(t)
    return d
