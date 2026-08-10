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

# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
st.set_page_config(page_title="TNC Filtration Service Estimator", page_icon="💧",
                   layout="wide", initial_sidebar_state="collapsed")

HERE = os.path.dirname(os.path.abspath(__file__))
WB_PATH = os.path.join(HERE, "data", fl.DEFAULT_WORKBOOK)
PROTOCOL_DIR = os.path.join(HERE, "data", "protocols")

# Set this to your deployed AquaPIP URL to show a cross-link button (leave blank
# to hide it). e.g. "https://aquapip.streamlit.app"
AQUAPIP_URL = ""

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
HOWTO = bundle["howto"]


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #
def step(txt):
    st.markdown(f"<div class='step'>{txt}</div>", unsafe_allow_html=True)


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
    """Any real files dropped into data/protocols/ (ignore placeholders)."""
    if not os.path.isdir(PROTOCOL_DIR):
        return []
    out = []
    for p in sorted(glob.glob(os.path.join(PROTOCOL_DIR, "*"))):
        base = os.path.basename(p)
        if base.lower() in ("readme.md", ".gitkeep", "readme.txt"):
            continue
        if os.path.isfile(p):
            out.append(p)
    return out


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
st.divider()

# Intro (live from the workbook's "How to use" sheet)
if HOWTO["intro"]:
    st.markdown("<div class='learnbox'>" +
                "".join(f"<div class='v' style='margin:.15rem 0'>{t}</div>"
                        for t in HOWTO["intro"]) + "</div>",
                unsafe_allow_html=True)

# How-to steps (live) in an expander
if HOWTO["steps"]:
    with st.expander("How to use this tool"):
        for i, s in enumerate(HOWTO["steps"], 1):
            st.markdown(f"{i}. {s}")

# Caveats — always visible, prominent (live from the workbook)
if HOWTO["caveats"]:
    st.markdown(
        "<div class='caveat'><div class='hd'>⚠ Important caveats — please read</div>" +
        "".join(f"<p>{c}</p>" for c in HOWTO["caveats"]) + "</div>",
        unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# 1 — Species
# --------------------------------------------------------------------------- #
step("1 · Choose your species")
sp_names = [s.name for s in SPECIES]
sel_name = st.selectbox("Species", sp_names, index=None,
                        placeholder="Select the bivalve you farm...",
                        label_visibility="collapsed")
if not sel_name:
    st.info("Pick your species to start the estimate.")
    st.stop()

sp = next(s for s in SPECIES if s.name == sel_name)

# Not-yet-calculable species: grey out, show the sheet's own message.
if not sp.calculable:
    st.markdown(
        f"<div class='card' style='border-left-color:#B9C2C5'>"
        f"<b style='color:{TEAL}'>{sp.name}</b><br>"
        f"<span style='color:#5B6B70'>{sp.status_message} "
        f"This species will become calculable here as the evidence base grows.</span></div>",
        unsafe_allow_html=True)
    st.stop()

# Formula + per-species reference/notes button (side by side)
fcol1, fcol2 = st.columns([3, 1])
with fcol1:
    if sp.in_full:
        st.markdown(f"<div class='formula'>{sp.in_full}</div>", unsafe_allow_html=True)
with fcol2:
    with st.popover("📄 Reference & notes", use_container_width=True):
        st.markdown(f"**{sp.name}**")
        matched = fl.references_for(sp.reference, REFS)
        if matched:
            st.markdown("**Reference(s)**")
            for r in matched:
                st.markdown(f"- {r.full or r.cite}"
                            + (f"  \n  [{r.link}]({r.link})" if r.link else ""))
        elif sp.reference:
            st.markdown(f"**Reference:** {sp.reference}")
        if sp.notes:
            st.markdown("**Notes**")
            st.caption(sp.notes)
        if not matched and not sp.reference and not sp.notes:
            st.caption("No reference or notes recorded for this species yet.")


# --------------------------------------------------------------------------- #
# 2 — Shell height & how many  (defines the size classes)
# --------------------------------------------------------------------------- #
step("2 · Shell height & how many")
st.caption("Enter the mean shell height of your bivalves and how many you have. "
           "Use one size, or split your stock into several size classes.")

size_mode = st.radio("Size entry", ["One shell height", "Several size classes"],
                     horizontal=True, label_visibility="collapsed")

size_classes = []                       # [{shell_mm, count}, ...]  — source of truth
if size_mode == "One shell height":
    c1, c2 = st.columns(2)
    with c1:
        _sh = st.number_input("Mean shell height (mm)", min_value=0.0,
                              max_value=400.0, value=50.0, step=1.0, format="%.1f")
    with c2:
        _n = st.number_input("Number of bivalves", min_value=0, value=1000,
                            step=100, format="%d")
    size_classes.append({"shell_mm": _num(_sh), "count": _num(_n)})
else:
    seed = pd.DataFrame([{"Shell height (mm)": 50.0, "Number of bivalves": 1000},
                         {"Shell height (mm)": 30.0, "Number of bivalves": 1000}])
    ed = st.data_editor(
        seed, num_rows="dynamic", use_container_width=True, hide_index=True,
        key="sizeclasses",
        column_config={
            "Shell height (mm)": st.column_config.NumberColumn(
                min_value=0.0, max_value=400.0, step=1.0, format="%.1f"),
            "Number of bivalves": st.column_config.NumberColumn(
                min_value=0, step=100, format="%d"),
        })
    for _, r in ed.iterrows():
        sm, ct = _num(r.get("Shell height (mm)")), _num(r.get("Number of bivalves"))
        if sm is None and ct is None:
            continue
        size_classes.append({"shell_mm": sm, "count": ct})

if not size_classes:
    st.info("Add at least one size class to continue.")
    st.stop()


# --------------------------------------------------------------------------- #
# 3 — Dry weight  (size classes carry over from step 2)
# --------------------------------------------------------------------------- #
step("3 · Dry weight")

dtw_list = [None] * len(size_classes)   # dry tissue weight (g) per size class

if sp.uses_length:
    st.info("Not needed for this species — its filtration is calculated from shell "
            "height directly, so you can skip straight to temperature.")
else:
    convs = conversions_for(sp.name)
    OPT_ONE = "Use one dry weight for all"
    OPT_MEAS = "I have my own dry weights by size class (more accurate)"
    OPT_EST = "Estimate dry weight from my shell height"
    src = st.radio("Dry-weight source", [OPT_ONE, OPT_MEAS, OPT_EST],
                   label_visibility="collapsed")
    st.caption("This species' equation uses dry tissue weight (g).")

    if src == OPT_ONE:
        st.caption("A single dry tissue weight is applied to every animal. The "
                   "default is a placeholder — enter your own value if you have it.")
        dtw_all = st.number_input("Dry tissue weight (g)", min_value=0.0,
                                  value=1.0, step=0.1, format="%.3f")
        dtw_list = [dtw_all] * len(size_classes)

    elif src == OPT_MEAS:
        st.caption("Your size classes carry over from above — just add the "
                   "**measured** dry tissue weight for each.")
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
                       value=18.0, step=0.5, format="%.1f")


# --------------------------------------------------------------------------- #
# Results
# --------------------------------------------------------------------------- #
rows_for_calc = [{"shell_mm": sc["shell_mm"], "count": sc["count"],
                  "dtw_g": dtw_list[i]} for i, sc in enumerate(size_classes)]
results, totals, notes = fl.compute_rows(sp, temp, rows_for_calc)

step("Estimated water cleared")
m1, m2 = st.columns(2)
with m1:
    st.markdown(f"<div class='metricbig'><div class='num'>{sig(totals['filt_m3d'],3)}</div>"
                f"<div class='lab'>cubic metres per day (m³/day)</div></div>",
                unsafe_allow_html=True)
with m2:
    st.markdown(f"<div class='metricbig'><div class='num'>{sig(totals['filt_lph'],3)}</div>"
                f"<div class='lab'>litres per hour (L/h)</div></div>",
                unsafe_allow_html=True)
st.caption(f"Total across **{fmt_int(totals['count'])}** animals at "
           f"**{temp:.1f} °C**. This is a best-case maximum (see caveats above).")

# Per-size-class breakdown: shell height, dry weight, number, then the three stats.
show = []
for e in results:
    row = {"Shell height (mm)": sig(e["shell_mm"], 3) if e["shell_mm"] is not None else "—"}
    if not sp.uses_length:
        row["Dry weight (g)"] = sig(e["dtw_g"], 3) if e["dtw_g"] is not None else "—"
    row["Number of bivalves"] = fmt_int(e["count"])
    row["Filtration rate (L) per individual per hour"] = (
        sig(e["cr_lph"], 3) if e["cr_lph"] is not None else "—")
    row["Filtration (L/h)"] = sig(e["filt_lph"], 3) if e["filt_lph"] is not None else "—"
    row["Filtration m³/day"] = sig(e["filt_m3d"], 3) if e["filt_m3d"] is not None else "—"
    if e["error"]:
        row["Note"] = e["error"]
    show.append(row)

if show:
    st.dataframe(pd.DataFrame(show), use_container_width=True, hide_index=True)

for n in notes:
    st.caption("⚠ " + n)


# --------------------------------------------------------------------------- #
# Protocols · estimating dry tissue weight (greyed until a file is added)
# --------------------------------------------------------------------------- #
step("Protocols · estimating dry tissue weight")
st.caption("Standardised field/lab protocols for measuring dry tissue weight will "
           "appear here for download once added.")
pfiles = protocol_files()
if pfiles:
    pcols = st.columns(min(3, len(pfiles)))
    for i, p in enumerate(pfiles):
        with pcols[i % len(pcols)]:
            with open(p, "rb") as fh:
                st.download_button(f"⬇ {os.path.basename(p)}", data=fh.read(),
                                   file_name=os.path.basename(p),
                                   use_container_width=True, key=f"proto_{i}")
else:
    st.button("Dry-weight protocol (coming soon)", disabled=True,
              use_container_width=False,
              help="Not yet available — drop a protocol file into data/protocols/ "
                   "in the repo and it will appear here automatically.")


# --------------------------------------------------------------------------- #
# References (download all)
# --------------------------------------------------------------------------- #
step("References")
st.caption("Every clearance-rate and length-weight source used by this tool.")
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
        st.markdown(line)

st.divider()
st.caption("Source: TNC Filtration Service Estimator workbook "
           "(data/Clearance_rate_estimation_tool_TNC.xlsx). Species, equations, "
           "conversions and references update automatically when the workbook is "
           "updated. Companion to AquaPIP · The Nature Conservancy.")
