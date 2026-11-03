"""
TNC Filtration Service Estimator  (Streamlit app)

Companion to AquaPIP. A farmer picks a bivalve species, enters shell height (one
size or several size classes with how many animals in each), dry weight, and water
temperature, and gets the estimated volume of water their stock can clear — tied to
MEL Objective 1.3 (farmed biomass improves light penetration / reduces hypoxia).

EVERYTHING is read live from ONE workbook so the tool updates the moment the
science does:  data/Clearance_rate_estimation_tool_TNC.xlsx
Add a species, change an equation, add a reference or a length-weight conversion,
edit the How-to / caveats text — reboot the app and it is reflected here, with no
code change (see filtration_logic.py header for the mechanism).
"""

import os
import glob
import math
import streamlit as st
import pandas as pd

import filtration_logic as fl
import filtration_report as fr
import filtration_storage as fs

# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
st.set_page_config(page_title="TNC Filtration Service Estimator", page_icon="💧",
                   layout="wide", initial_sidebar_state="collapsed")

HERE = os.path.dirname(os.path.abspath(__file__))
WB_PATH = os.path.join(HERE, "data", fl.DEFAULT_WORKBOOK)
# Protocol PDFs belong in data/protocols/. Uploading a folder through GitHub's
# web page can nest it one level deeper (data/data/protocols/), so look there too.
PROTOCOL_DIRS = [os.path.join(HERE, "data", "protocols"),
                 os.path.join(HERE, "data", "data", "protocols")]
PROTOCOL_DIR = PROTOCOL_DIRS[0]

# Set this to your deployed AquaPIP URL to show a cross-link button (leave blank
# to hide it). e.g. "https://aquapip.streamlit.app"
AQUAPIP_URL = ""

# Set this to a contact / data-sharing form URL (e.g. the same Microsoft Form you
# use for AquaPIP) to show a "Get in touch" button at the bottom. Blank = hidden.
CONTACT_URL = "https://forms.cloud.microsoft/Pages/ResponsePage.aspx?id=wW2-eY7Xu0uyK9mUwKQXp6WdY3q1w15BvDRlgn30McZURVlVWEtVV1I3MlZZRUoxQlFRVDZUTzBBTS4u"

# --------------------------------------------------------------------------- #
# Other filtration tools, linked by species  (edit here)
#   "redirect": used while this estimator can't calculate the species. Steps 2-4
#               are shown greyed out and the farmer is pointed to the other tool.
#               A redirect species that isn't in the workbook is added to the
#               species list. As soon as the workbook has an equation for the
#               species, its card switches to "compare" automatically.
#   "compare":  the estimate runs (after "Continue here"), with a link to compare.
# Keys are scientific names in lower case; they are matched against the workbook.
# --------------------------------------------------------------------------- #
FARM_URL = "http://www.farmscale.org/"
FARM_WHY = ("It models your farm using water data from your site, such as "
            "chlorophyll.")
OTHER_TOOLS = {
    "crassostrea virginica": {
        "name": "Crassostrea virginica", "common": "eastern oyster",
        "mode": "redirect", "tool": "Rutgers Oyster Eco-Serve Calculator",
        "url": "https://itsappserver.sebs.rutgers.edu/Oysterecoservecalc/",
        "why": ("It is built from feeding measurements taken on working oyster "
                "farms from Massachusetts to Virginia."),
    },
    "mytilus galloprovincialis": {
        "name": "Mytilus galloprovincialis", "common": "Mediterranean mussel",
        "mode": "redirect", "tool": "FARM model", "url": FARM_URL,
        "why": FARM_WHY,
    },
    "crassostrea gigas": {
        "name": "Crassostrea gigas", "common": "Pacific oyster",
        "mode": "compare", "tool": "FARM model", "url": FARM_URL, "why": FARM_WHY,
    },
    "mytilus edulis": {
        "name": "Mytilus edulis", "common": "blue mussel",
        "mode": "compare", "tool": "FARM model", "url": FARM_URL, "why": FARM_WHY,
    },
    "ruditapes philippinarum": {
        "name": "Ruditapes philippinarum", "common": "Manila clam",
        "mode": "compare", "tool": "FARM model", "url": FARM_URL, "why": FARM_WHY,
    },
}

TEAL = "#0B4F5C"

# Same visual language as AquaPIP (app.py), plus a caveat card + metric styles.
st.markdown(f"""
<style>
  .block-container {{ padding-top: 2rem; padding-bottom: 3rem; max-width: 1150px; }}
  #MainMenu, footer {{ visibility: hidden; }}
  h1, h2, h3 {{ font-family: Georgia,'Times New Roman',serif; color:{TEAL}; }}
  .hero h1 {{ margin-bottom:.1rem; font-size:1.85rem; }}
  .hero p  {{ color:#5B6B70; margin-top:0; font-size:.95rem; }}
  .step {{ font-family:Georgia,serif; color:{TEAL}; font-weight:700; font-size:1.12rem; margin:1.2rem 0 .2rem; }}
  .step.off {{ color:#AEB9BD; }}
  .card {{ background:#fff; border:1px solid #E1E7E9; border-left:5px solid {TEAL};
           border-radius:0 10px 10px 0; padding:.8rem 1rem; margin:.5rem 0; }}
  .k {{ font-size:.66rem; font-weight:700; letter-spacing:.03em; text-transform:uppercase; color:#5B6B70; margin-top:.45rem; }}
  .v {{ font-size:.9rem; color:#1A2B2F; }}
  .tag {{ background:#EEF3F4; color:#33474C; font-size:.74rem; padding:.12rem .5rem; border-radius:999px; margin-left:.3rem; }}
  .learnbox {{ background:#F3F6F7; border:1px solid #E1E7E9; border-radius:10px; padding:.6rem .8rem; }}
  .formula {{ background:#F3F6F7; border:1px solid #E1E7E9; border-radius:8px;
              padding:.55rem .8rem; font-family:'DejaVu Sans Mono',monospace;
              font-size:.86rem; color:#1A2B2F; }}
  .caveat {{ background:#FFF7E6; border:1px solid #F2D591; border-left:5px solid #E0A82E;
             border-radius:0 10px 10px 0; padding:.75rem 1rem; margin:.6rem 0; }}
  .caveat .hd {{ color:#8A5A00; font-weight:700; font-family:Georgia,serif; margin-bottom:.25rem; }}
  .caveat p {{ color:#5B4a1f; font-size:.9rem; margin:.15rem 0; }}
  .caveat .sub {{ color:#8A5A00; font-weight:700; font-family:Georgia,serif;
                 margin:.6rem 0 .2rem; padding-top:.55rem; border-top:1px solid #F2D591; }}
  .metricbig {{ background:#EAF3F1; border:1px solid #CFE4DE; border-radius:12px;
                padding:1rem 1.15rem; text-align:center; }}
  .metricbig .num {{ font-family:Georgia,serif; color:{TEAL}; font-size:2.0rem; font-weight:700; line-height:1.1; }}
  .metricbig .lab {{ color:#5B6B70; font-size:.8rem; text-transform:uppercase; letter-spacing:.03em; }}
  .clslabel {{ padding-top:.55rem; }}
</style>
""", unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# Data (cached on the workbook's modified-time -> reloads when the file changes)
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner=False)
def get_bundle(_mtime):
    return fl.load_workbook_bundle(WB_PATH)


if not os.path.exists(WB_PATH):
    st.error(f"Missing `data/{fl.DEFAULT_WORKBOOK}`. Add the workbook to the "
             "repo's data/ folder.")
    st.stop()

bundle = get_bundle(os.path.getmtime(WB_PATH))
SPECIES = bundle["species"]
CONVS = bundle["conversions"]
REFS = bundle["references"]


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #
def step(txt, off=False):
    cls = "step off" if off else "step"
    st.markdown(f"<div class='{cls}'>{txt}</div>", unsafe_allow_html=True)


def _flip(key):
    """Button callback: toggle a True/False setting kept between reruns."""
    st.session_state[key] = not st.session_state.get(key, False)


def _continue_here(name):
    """Button callback: the farmer chose to carry on in this estimator."""
    st.session_state["continue_for"] = name


def other_tool_card(t, species=None, continued=False, mode=None):
    """'Farming X?' card + a link to the other tool (see OTHER_TOOLS near the
    top). Compare cards also get a 'Continue here' button; the steps below stay
    greyed out until it is clicked. `mode` is the card's effective mode."""
    redirect = (mode or t["mode"]) == "redirect"
    if redirect:
        body = (t.get("note") or f"This estimator doesn't include {t['common']} "
                f"yet.") + f" Try the {t['tool']} instead. {t['why']}"
    else:
        body = f"The {t['tool']} may give a more site-specific estimate. {t['why']}"
    st.markdown(
        "<div class='card'>"
        f"<b style='color:{TEAL}'>Farming {t['common']}?</b><br>"
        f"<span class='v'>{body}</span></div>",
        unsafe_allow_html=True)
    if redirect:
        st.link_button(f"Go to the {t['tool']} ↗", t["url"], type="primary")
        return
    b1, b2, _ = st.columns([1.5, 1, 1.5])
    with b1:
        st.link_button(f"Compare with the {t['tool']} ↗", t["url"],
                       use_container_width=True)
    with b2:
        if continued:
            st.button("✓ Continuing here", disabled=True, key="cont_done",
                      use_container_width=True)
        else:
            st.button("Continue here", key="cont_btn", on_click=_continue_here,
                      args=(species or t["name"],), use_container_width=True)


def greyed_next_steps(note="The steps below are greyed out for this species."):
    """Steps 2-4, shown disabled: the farmer sees what this estimator asks for,
    but is pointed to the other tool linked above."""
    st.caption(note)
    step("2 · Shell height & how many", off=True)
    st.radio("Size entry", ["One shell height", "Several size classes"],
             horizontal=True, label_visibility="collapsed", disabled=True,
             key="off_size")
    g1, g2 = st.columns(2)
    with g1:
        st.number_input("Mean shell height (mm)", value=50.0, disabled=True,
                        key="off_shell")
    with g2:
        st.number_input("Number of bivalves", value=1000, disabled=True,
                        key="off_count")
    step("3 · Dry weight", off=True)
    st.radio("Dry-weight source", ["Use average dry weight",
                                   "Use dry weight by size class",
                                   "Estimate dry weight from my shell height"],
             label_visibility="collapsed", disabled=True, key="off_dw")
    step("4 · Water temperature", off=True)
    st.number_input("Water temperature (°C)", value=18.0, disabled=True,
                    key="off_temp")


def sig(x, n=3):
    """Format a number to n significant figures with thousands separators."""
    if x is None:
        return "—"
    try:
        x = float(x)
    except (TypeError, ValueError):
        return "—"
    if x == 0:
        return "0"
    from math import log10, floor
    d = n - 1 - floor(log10(abs(x)))
    d = max(0, d)
    return f"{round(x, d):,.{d}f}"


def fmt_int(x):
    try:
        return f"{int(round(float(x))):,}"
    except (TypeError, ValueError):
        return "—"


def _num(x):
    """Coerce an editor/number value to float or None (handles NaN)."""
    if x is None:
        return None
    try:
        f = float(x)
    except (TypeError, ValueError):
        return None
    return None if math.isnan(f) else f


def conversions_for(species_name):
    key = fl.norm_species(species_name).lower()
    return [c for c in CONVS if c.species.lower() == key and c.valid]


def protocol_files():
    """Real files in the protocol folders (placeholders ignored). Checks every
    folder in PROTOCOL_DIRS; if a file name appears twice, the first copy wins."""
    out, seen = [], set()
    for d in PROTOCOL_DIRS:
        if not os.path.isdir(d):
            continue
        for p in sorted(glob.glob(os.path.join(d, "*"))):
            base = os.path.basename(p)
            if base.lower() in ("readme.md", ".gitkeep", "readme.txt") or base in seen:
                continue
            if os.path.isfile(p):
                seen.add(base)
                out.append(p)
    return out


# --------------------------------------------------------------------------- #
# Dry-weight protocol + data sheet: one click downloads both (as one zip)
# --------------------------------------------------------------------------- #
DRY_WEIGHT_PROTOCOL_PDF = "Bivalve_dry_weight_protocol.pdf"
DRY_WEIGHT_DATASHEET_PDF = "Bivalve_dry_weight_data_sheet.pdf"
DRY_WEIGHT_ZIP_NAME = "Bivalve_dry_weight_protocol_and_data_sheet.zip"


def _find_in_data(name):
    """Last resort: an exact file name anywhere under data/."""
    hits = sorted(glob.glob(os.path.join(HERE, "data", "**", name), recursive=True))
    return hits[0] if hits else None


def dry_weight_files():
    """(protocol_path, datasheet_path) from data/protocols/; either may be None.
    The exact file names above win. Otherwise a file whose name mentions 'sheet'
    is the data sheet, and one mentioning protocol/dry/weight/dtw is the protocol."""
    files = protocol_files()
    by_name = {os.path.basename(p): p for p in files}
    proto = by_name.get(DRY_WEIGHT_PROTOCOL_PDF) or _find_in_data(DRY_WEIGHT_PROTOCOL_PDF)
    sheet = by_name.get(DRY_WEIGHT_DATASHEET_PDF) or _find_in_data(DRY_WEIGHT_DATASHEET_PDF)
    for p in files:
        b = os.path.basename(p).lower()
        if "sheet" in b:
            sheet = sheet or p
        elif any(k in b for k in ("protocol", "dry", "weight", "dtw")):
            proto = proto or p
    return proto, sheet


def dry_weight_protocol():
    """Kept for backwards compatibility: the protocol file path, or None."""
    return dry_weight_files()[0]


@st.cache_data(show_spinner=False)
def _zip_files(paths, mtimes):
    """One in-memory zip holding the given files. `mtimes` is in the signature
    only so the cached zip refreshes when a PDF in the repo is replaced."""
    import io
    import zipfile
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            z.write(p, arcname=os.path.basename(p))
    return buf.getvalue()
  
def dry_weight_protocol_button(key):
    """One click downloads the protocol AND its data sheet. Streamlit sends one
    file per click, so both PDFs travel together in a single zip. If only one
    of the two files is present it downloads on its own; if neither is, a
    greyed-out 'coming soon' button shows. `key` keeps the widget unique across
    the places this is shown."""
    proto, sheet = dry_weight_files()
    files = [p for p in (proto, sheet) if p]
    if not files:
        st.button("📄 Dry-weight protocol (coming soon)", disabled=True, key=key,
                  help="Not yet available — drop the protocol PDF into "
                       "data/protocols/ in the repo and it will activate here.")
        return
    if len(files) == 1:
        p = files[0]
        label = ("📄 Download dry-weight protocol" if p == proto
                 else "📄 Download dry-weight data sheet")
        with open(p, "rb") as fh:
            st.download_button(label, data=fh.read(),
                               file_name=os.path.basename(p),
                               mime="application/pdf", key=key)
        return
    data = _zip_files(tuple(files), tuple(os.path.getmtime(p) for p in files))
    st.download_button("📄 Download dry-weight protocol + data sheet", data=data,
                       file_name=DRY_WEIGHT_ZIP_NAME, mime="application/zip",
                       key=key,
                       help="Two PDFs in one download: the step-by-step protocol "
                            "and a printable data sheet (one sheet per size class).")




# --------------------------------------------------------------------------- #
# Header
# --------------------------------------------------------------------------- #
st.markdown(
    "<div class='hero'><h1>Filtration Service Estimator</h1>"
    "<p>Estimate the maximum volume of water your bivalves can filter — a measure "
    "of how your farm helps clear the water, improving light for seagrass and algae "
    "and easing low-oxygen stress in the bay.</p></div>",
    unsafe_allow_html=True)

hcols = st.columns([3, 1])
with hcols[1]:
    if AQUAPIP_URL:
        st.link_button("↩ Back to AquaPIP", AQUAPIP_URL, use_container_width=True)
    if CONTACT_URL:
        st.link_button("💬 Questions or feedback?", CONTACT_URL,
                       use_container_width=True)
st.divider()

# Welcome / what this tool is (fixed framing — not workbook-driven, so it stays
# correct and is never lost when the data workbook is updated).
st.markdown(
    "<div class='learnbox'>"
    "<div class='v' style='margin:.15rem 0'><b>Welcome to the TNC Filtration "
    "Services estimator tool.</b></div>"
    "<div class='v' style='margin:.35rem 0'>This tool is designed to allow growers "
    "to approximate the potential maximum volume of water cleared by the bivalves "
    "being grown on site. It brings together existing published data and "
    "calculations of size-specific filtration rates to make general estimations "
    "about the filtration capacity of stock held (biomass) on a farm at any point "
    "in time. It does not value the effect of water filtration on the broader "
    "ecosystem, rather it guides a general understanding of the process of "
    "filtration that contributes to a range of ecosystem services such as "
    "improving/reducing suspended matter and regulating phytoplankton.</div>"
    "</div>",
    unsafe_allow_html=True)

# How to use (web tool steps)
with st.expander("How to use this tool"):
    st.markdown(
        "1. **Choose your species.**\n"
        "2. **Enter shell height** — one size, or several size classes with how "
        "many animals are in each.\n"
        "3. **Enter dry weight** — use your own measured weights, estimate them "
        "from shell height where a conversion exists, or apply one value to all. "
        "(Not needed for species whose filtration is calculated from shell height "
        "directly.)\n"
        "4. **Enter your water temperature.**\n"
        "5. Read **what one animal filters**, then **what your farm filters**, "
        "each with its caveats.")


# --------------------------------------------------------------------------- #
# 1 — Species
# --------------------------------------------------------------------------- #
step("1 · Choose your species")
sp_names = [s.name for s in SPECIES]
# Redirect-only species (e.g. eastern oyster) join the list next to their genus.
_known = {fl.norm_species(n).lower() for n in sp_names}
for _key, _t in OTHER_TOOLS.items():
    if _t["mode"] == "redirect" and _key not in _known:
        _genus = _t["name"].split()[0].lower() + " "
        _same = [i for i, n in enumerate(sp_names) if n.lower().startswith(_genus)]
        sp_names.insert(_same[-1] + 1 if _same else len(sp_names), _t["name"])
sel_name = st.selectbox("Species", sp_names, index=None,
                        placeholder="Select the bivalve you farm...",
                        label_visibility="collapsed")
if not sel_name:
    st.info("Pick your species to start the estimate.")
    st.stop()

# Species handled by another tool. A "redirect" species is only redirected while
# the workbook has no equation for it; once it does, it becomes a "compare".
_key = fl.norm_species(sel_name).lower()
other = OTHER_TOOLS.get(_key)
_sp = next((s for s in SPECIES if s.name == sel_name), None)
_can_calc = bool(_sp and _sp.calculable)
other_mode = None
if other:
    other_mode = ("compare" if other["mode"] == "compare" or _can_calc
                  else "redirect")
if other_mode == "redirect":
    other_tool_card(other, mode="redirect")
    greyed_next_steps()
    st.stop()

sp = _sp

# Not-yet-calculable species: grey out, with a plain-language status.
if not sp.calculable:
    st.markdown(
        f"<div class='card' style='border-left-color:#B9C2C5'>"
        f"<b style='color:{TEAL}'>{sp.name}</b><br>"
        f"<span style='color:#5B6B70'>{sp.farmer_status} "
        f"It will appear here as the evidence grows.</span></div>",
        unsafe_allow_html=True)
    st.stop()

# Formula + per-species reference/notes button (side by side)
fcol1, fcol2 = st.columns([3, 1])
with fcol1:
    _eq_on = st.session_state.get("show_eq", False)
    st.button("Hide equation" if _eq_on else "Show equation", key="eq_btn",
              on_click=_flip, args=("show_eq",))
    if _eq_on:
        # Written out from the formula the tool actually runs, so it can never
        # drift from the calculation. Excel syntax, ready to copy.
        _eq, _legend = sp.display_equation()
        st.code(_eq, language=None)
        st.caption(_legend + " To run it yourself, for example in Excel, swap "
                   "each letter for your own value or a cell reference.")
with fcol2:
    with st.popover("📄 Reference & notes", use_container_width=True):
        st.markdown(f"**{sp.name}**")
        matched = fl.references_for(sp.reference, REFS)
        if matched:
            st.markdown("**Reference(s)**")
            for r in matched:
                st.markdown(f"- {r.full or r.cite}"
                            + (f"  \n  [{r.link}]({r.link})" if r.link else "")
                            + (f"  \n  Free copy: [{r.free_link}]({r.free_link})"
                               if r.free_link else ""))
        elif sp.reference:
            st.markdown(f"**Reference:** {sp.reference}")
        if sp.notes:
            st.markdown("**Notes**")
            st.caption(sp.notes)
        if not matched and not sp.reference and not sp.notes:
            st.caption("No reference or notes recorded for this species yet.")

# Another tool covers this species too: offer it, and keep the estimate below
# greyed out until the farmer chooses "Continue here".
if other_mode == "compare":
    _here = st.session_state.get("continue_for") == sp.name
    other_tool_card(other, species=sp.name, continued=_here, mode="compare")
    if not _here:
        greyed_next_steps("The steps below are greyed out until you choose "
                          "Continue here.")
        st.stop()


# Dry-weight source options — defined here so Step 2 can react to the choice the
# farmer makes in Step 3 (Streamlit persists the radio via its key across reruns).
OPT_ONE = "Use average dry weight — download the protocol below for the correct way to obtain it"
OPT_MEAS = "Use dry weight by size class — download the protocol below for the correct way to obtain it"
OPT_EST = "Estimate dry weight from my shell height"

# Why dry weight isn't asked for, for species whose equation uses shell height.
LENGTH_NOTES = {
    "mytilus edulis": (
        "Not needed for blue mussel. Evidence suggests its filtration is better "
        "predicted by shell height than by meat weight, which rises and falls with "
        "the mussel's condition through the year. You can skip straight to "
        "temperature."),
}
_dw_src_prev = st.session_state.get("dwsrc", OPT_ONE)
# Shell height plays no part in a dry-weight species' filtration equation when a
# single dry weight is applied to every animal — so it is greyed out in that case.
shell_unused = (not sp.uses_length) and (_dw_src_prev == OPT_ONE)


# --------------------------------------------------------------------------- #
# 2 — Shell height & how many  (defines the size classes)
# --------------------------------------------------------------------------- #
step("2 · Shell height & how many")
st.caption("Enter the mean shell height of your bivalves and how many you have. "
           "Use one size, or split your stock into several size classes.")

size_mode = st.radio("Size entry", ["One shell height", "Several size classes"],
                     horizontal=True, label_visibility="collapsed")

if shell_unused:
    st.info("Shell height is greyed out because you've chosen **one dry weight for "
            "all** in the next step. For this species, filtration is then calculated "
            "from dry weight and temperature only, so shell height isn't used. To "
            "enter sizes, choose *estimate dry weight from shell height* or "
            "*measured dry weights by size class* below.")

size_classes = []                       # [{shell_mm, count}, ...]  — source of truth
if size_mode == "One shell height":
    c1, c2 = st.columns(2)
    with c1:
        _sh = st.number_input("Mean shell height (mm)", min_value=0.0,
                              max_value=400.0, value=50.0, step=1.0, format="%.1f",
                              disabled=shell_unused)
    with c2:
        _n = st.number_input("Number of bivalves", min_value=0, value=1000,
                            step=100, format="%d")
    size_classes.append({"shell_mm": None if shell_unused else _num(_sh),
                         "count": _num(_n)})
else:
    seed = pd.DataFrame([{"Shell height (mm)": 50.0, "Number of bivalves": 1000},
                         {"Shell height (mm)": 30.0, "Number of bivalves": 1000}])
    ed = st.data_editor(
        seed, num_rows="dynamic", use_container_width=True, hide_index=True,
        key="sizeclasses",
        column_config={
            "Shell height (mm)": st.column_config.NumberColumn(
                min_value=0.0, max_value=400.0, step=1.0, format="%.1f",
                disabled=shell_unused),
            "Number of bivalves": st.column_config.NumberColumn(
                min_value=0, step=100, format="%d"),
        })
    for _, r in ed.iterrows():
        sm, ct = _num(r.get("Shell height (mm)")), _num(r.get("Number of bivalves"))
        if sm is None and ct is None:
            continue
        size_classes.append({"shell_mm": None if shell_unused else sm, "count": ct})

if not size_classes:
    st.info("Add at least one size class to continue.")
    st.stop()


# --------------------------------------------------------------------------- #
# 3 — Dry weight  (size classes carry over from step 2)
# --------------------------------------------------------------------------- #
step("3 · Dry weight")

dtw_list = [None] * len(size_classes)   # dry tissue weight (g) per size class
dw_src = None                           # set when a dry-weight source is chosen

if sp.uses_length:
    st.info(LENGTH_NOTES.get(
        fl.norm_species(sp.name).lower(),
        "Not needed for this species. Its filtration equation uses shell height "
        "directly, so you can skip straight to temperature."))
else:
    convs = conversions_for(sp.name)
    src = st.radio("Dry-weight source", [OPT_ONE, OPT_MEAS, OPT_EST],
                   label_visibility="collapsed", key="dwsrc")
    dw_src = src
    st.caption("This species' equation uses dry tissue weight (g).")

    if src == OPT_ONE:
        st.caption("A single average dry tissue weight is applied to every animal. "
                   "The default is a placeholder — enter your own value if you have "
                   "it. Use the protocol below for the correct way to measure it.")
        dry_weight_protocol_button("proto_one")
        dtw_all = st.number_input("Average dry tissue weight (g)", min_value=0.0,
                                  value=1.0, step=0.1, format="%.3f")
        dtw_list = [dtw_all] * len(size_classes)

    elif src == OPT_MEAS:
        st.caption("Enter a measured dry tissue weight for each size class carried "
                   "over from above. Use the protocol below for the correct way to "
                   "measure it.")
        dry_weight_protocol_button("proto_meas")
        for i, sc in enumerate(size_classes):
            cA, cB = st.columns([2, 1])
            with cA:
                st.markdown(
                    f"<div class='v clslabel'><b>Class {i+1}</b> · "
                    f"{sig(sc['shell_mm'],3)} mm · {fmt_int(sc['count'])} individuals"
                    f"</div>", unsafe_allow_html=True)
            with cB:
                dtw_list[i] = st.number_input(
                    "Dry weight (g)", min_value=0.0, value=1.0, step=0.1,
                    format="%.3f", key=f"dtwmeas_{i}", label_visibility="collapsed")

        # ---- Invitation to contribute these measurements -------------------- #
        st.markdown(
            "<div class='card' style='margin-top:.6rem'>"
            "<b style='color:#0B4F5C'>Help improve the calculator</b><br>"
            "<span class='v'>We're always working to improve our length-to-dry-weight "
            "estimates so we can give farmers the most accurate filtration rates. If "
            "you'd like to contribute to this work, you can save your entries to be "
            "used in this calculator. Your name and email are optional, but help us "
            "credit you in any future publications or works. Thank you!</span></div>",
            unsafe_allow_html=True)

        if st.session_state.get("contrib_saved"):
            st.success("Thank you — your measurements have been saved and will help "
                       "improve the tool.")
        else:
            with st.popover("✍️  Contribute my measurements",
                            use_container_width=False):
                st.markdown("**Contribute your dry-weight measurements**")
                st.caption("We're always working to improve our length-to-dry-weight "
                           "estimates so we can give farmers the most accurate "
                           "filtration rates. Tick the box below to save your entries "
                           "to be used in this calculator. Your name and email are "
                           "optional, but help us credit you in any future "
                           "publications or works. Thank you!")
                consent = st.checkbox(
                    "Yes — save my species, size classes, dry weights, temperature "
                    "and numbers, and use them to improve the calculator.")
                cn1, cn2 = st.columns(2)
                with cn1:
                    contrib_name = st.text_input("Name (optional)", key="contrib_name")
                with cn2:
                    contrib_email = st.text_input("Email (optional)", key="contrib_email")
                if not fs.is_configured():
                    st.caption("⚠ Data collection isn't switched on for this "
                               "deployment yet — ask the tool administrator to add "
                               "the contributions sheet in the app settings.")
                if st.button("Save my data", type="primary", disabled=not consent,
                             key="contrib_submit"):
                    payload = {
                        "species": sp.name,
                        "temperature_c": st.session_state.get("watertemp"),
                        "name": (contrib_name or "").strip(),
                        "email": (contrib_email or "").strip(),
                        "classes": [{"shell_mm": sc["shell_mm"],
                                     "dtw_g": dtw_list[i],
                                     "count": sc["count"]}
                                    for i, sc in enumerate(size_classes)],
                    }
                    ok, status = fs.save_contribution(payload)
                    if ok:
                        st.session_state["contrib_saved"] = True
                        st.rerun()
                    else:
                        st.warning("Sorry, we couldn't save that right now. Please "
                                   "try again later.")

    else:  # OPT_EST — estimate from shell height using a length-weight conversion
        if convs:
            if len(convs) == 1:
                chosen = convs[0]
                st.caption(f"Using conversion **{chosen.label()}** — "
                           f"`{chosen.eq_text}`")
            else:
                labels = [c.label() for c in convs]
                pick = st.selectbox("Length-weight conversion to use", labels)
                chosen = convs[labels.index(pick)]
                st.caption(f"`{chosen.eq_text}`")
            preview = []
            for i, sc in enumerate(size_classes):
                d = None
                if sc["shell_mm"] is not None:
                    try:
                        d = chosen.dtw_from_length(sc["shell_mm"])
                    except fl.ExprError:
                        d = None
                dtw_list[i] = d
                preview.append({
                    "Class": i + 1, "Shell height (mm)": sc["shell_mm"],
                    "Estimated dry weight (g)": round(d, 3) if d is not None else None,
                    "Number of bivalves": fmt_int(sc["count"])})
            st.dataframe(pd.DataFrame(preview), use_container_width=True,
                         hide_index=True)
        else:
            st.warning("No length-weight conversion is available yet for this "
                       "species, so shell height can't be turned into dry weight. "
                       "Use *measured dry weights* or *one dry weight for all* for "
                       "now — conversions will be added as data become available.")
            st.stop()


# --------------------------------------------------------------------------- #
# 4 — Water temperature
# --------------------------------------------------------------------------- #
step("4 · Water temperature")
st.caption("The water temperature at your site. Filtration changes with "
           "temperature, so use a value for the period you care about.")
temp = st.number_input("Water temperature (°C)", min_value=0.0, max_value=40.0,
                       value=18.0, step=0.5, format="%.1f", key="watertemp")


# --------------------------------------------------------------------------- #
# Results
# --------------------------------------------------------------------------- #
rows_for_calc = [{"shell_mm": sc["shell_mm"], "count": sc["count"],
                  "dtw_g": dtw_list[i]} for i, sc in enumerate(size_classes)]
results, totals, notes = fl.compute_rows(sp, temp, rows_for_calc)

# ---- What one animal filters ----------------------------------------------- #
# Every equation gives litres per animal per hour; the farm figures below are
# simply that rate x the number of animals.
step("What one animal filters")
st.caption("Filtration rate (FR) is the volume of water one animal clears of "
           "particles each hour. In the scientific literature it is often called "
           "the clearance rate.")

show_shell = sp.uses_length or dw_src in (OPT_MEAS, OPT_EST)
ind_rows = []
for i, e in enumerate(results):
    row = {"Size class": i + 1}
    if show_shell:
        row["Shell height (mm)"] = sig(e["shell_mm"], 3) if e["shell_mm"] is not None else "—"
    if not sp.uses_length:
        row["Dry weight (g)"] = sig(e["dtw_g"], 3) if e["dtw_g"] is not None else "—"
    row["Filtration rate (L) per individual per hour"] = (
        sig(e["cr_lph"], 3) if e["cr_lph"] is not None else "—")
    if e["error"]:
        row["Note"] = e["error"]
    ind_rows.append(row)
if ind_rows:
    st.dataframe(pd.DataFrame(ind_rows), use_container_width=True, hide_index=True)

# The species' own validity range, where the workbook notes give one.
_rng = sp.valid_temp_range
if _rng and not (_rng[0] <= temp <= _rng[1]):
    st.warning(f"This species' equation is only valid between {_rng[0]:g} and "
               f"{_rng[1]:g} °C. Your temperature ({temp:.1f} °C) is outside that "
               "range, so treat these numbers with extra caution.")
for n in notes:
    st.caption("⚠ " + n)

st.markdown(
    "<div class='caveat'><div class='hd'>⚠ Caveats for the rate per animal</div>"
    "<p>These rates are a best-case maximum, worked out from water temperature "
    "and animal size alone. On a real farm, a bivalve often filters less.</p>"
    "<p><b>Food levels.</b> When food (phytoplankton and other particles) is "
    "plentiful, bivalves are satiated more quickly and tend to filter less "
    "water.</p>"
    "<p><b>Water conditions.</b> Particle size, salinity, the amount of suspended "
    "material, pollutants and the time of day can all lower filtration.</p>"
    "<p><b>Limited evidence.</b> The equations often draw on, and are "
    "extrapolated from, one or a small number of published studies. This limits "
    "confidence in the numbers, but it is also a valuable area for "
    "species-specific and farm-based research. If you have data that could "
    "inform this tool, please get in touch using the contact form at the bottom "
    "of the page.</p></div>",
    unsafe_allow_html=True)


# ---- What your farm filters ------------------------------------------------ #
step("What your farm filters")
m1, m2 = st.columns(2)
with m1:
    st.markdown(f"<div class='metricbig'><div class='num'>{sig(totals['filt_m3d'],3)}</div>"
                f"<div class='lab'>cubic metres per day (m³/day)</div></div>",
                unsafe_allow_html=True)
with m2:
    st.markdown(f"<div class='metricbig'><div class='num'>{sig(totals['filt_lph'],3)}</div>"
                f"<div class='lab'>litres per hour (L/h)</div></div>",
                unsafe_allow_html=True)
st.caption(f"Each animal's filtration rate × **{fmt_int(totals['count'])}** "
           f"animals, at **{temp:.1f} °C**.")

farm_rows = []
for i, e in enumerate(results):
    farm_rows.append({
        "Size class": i + 1,
        "Number of bivalves": fmt_int(e["count"]),
        "Filtration (L/h)": sig(e["filt_lph"], 3) if e["filt_lph"] is not None else "—",
        "Filtration m³/day": sig(e["filt_m3d"], 3) if e["filt_m3d"] is not None else "—",
    })
if len(farm_rows) > 1:
    st.dataframe(pd.DataFrame(farm_rows), use_container_width=True, hide_index=True)

st.markdown(
    "<div class='caveat'><div class='hd'>⚠ Caveats for the farm total</div>"
    "<p>The farm total assumes every animal filters at its maximum rate all the "
    "time, in a perfectly mixed water column. It does not account for animals "
    "re-filtering water that others have already filtered, or for how water "
    "flow, tides and residence time move water through the farm.</p>"
    "<p>It shows the volume of water your stock could clear, not the effect of "
    "that filtering on the wider ecosystem.</p>"
    "<div class='sub'>Farm filtration is not equivalent to wild and restored "
    "bivalves</div>"
    "<p>It is important to recognise that ecosystem services supported by an "
    "aquaculture farm are not directly equivalent to those provided by wild and "
    "restored bivalve communities (Munroe et al. 2026, Lefcheck et al. 2021, "
    "Lunstrum et al. 2018). Protecting and enhancing bivalve populations in the "
    "wild remains the critical pathway to support filtration in marine "
    "ecosystems.</p></div>",
    unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# References (download all)
# --------------------------------------------------------------------------- #
step("References")
st.caption("Every source behind the filtration equations and length-to-weight "
           "conversions.")
rc1, rc2 = st.columns(2)
with rc1:
    try:
        st.download_button("⬇ Download all references (PDF)",
                           data=fr.references_pdf(REFS),
                           file_name="Filtration_tool_references.pdf",
                           mime="application/pdf", type="primary",
                           use_container_width=True)
    except Exception as e:
        st.info(f"PDF export temporarily unavailable ({type(e).__name__}).")
with rc2:
    st.download_button("⬇ Download all references (CSV)",
                       data=fr.references_csv(REFS),
                       file_name="Filtration_tool_references.csv",
                       mime="text/csv", use_container_width=True)

with st.expander(f"View all {len(REFS)} references"):
    for r in sorted(REFS, key=lambda x: x.cite.lower()):
        line = f"**{r.cite}** — {r.full}" if r.full else f"**{r.cite}**"
        if r.link:
            line += f"  \n[{r.link}]({r.link})"
        if r.free_link:
            line += f"  \nFree copy: [{r.free_link}]({r.free_link})"
        st.markdown(line)


# --------------------------------------------------------------------------- #
# Contact / share data
# --------------------------------------------------------------------------- #
step("Get in touch")
st.caption("Have data that could improve this tool, or a question? We'd love to "
           "hear from you — species-specific and farm-based measurements are a "
           "valuable area for research.")
if CONTACT_URL:
    st.link_button("✉  Contact us / share your data", CONTACT_URL)
else:
    st.info("A contact form will be linked here. (Set CONTACT_URL near the top of "
            "filtration_app.py — e.g. the same Microsoft Form used for AquaPIP.)")

st.divider()
st.caption("A companion to AquaPIP, from The Nature Conservancy.")
