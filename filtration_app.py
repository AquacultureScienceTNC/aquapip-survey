"""
TNC Filtration Service Estimator  (Streamlit app)

Companion to AquaPIP. A farmer picks a bivalve species, enters water temperature
and one or more size classes (shell height and/or dry weight, plus how many
animals), and gets the estimated maximum volume of water their stock can clear —
tied to MEL Objective 1.3 (farmed biomass improves light penetration / reduces
hypoxia).

EVERYTHING is read live from ONE workbook so the tool updates the moment the
science does:  data/Clearance_rate_estimation_tool_TNC.xlsx
Add a species, change an equation, add a reference or a length-weight conversion,
edit the How-to / caveats text — reboot the app and it is reflected here, with no
code change (see filtration_logic.py header for the mechanism).
"""

import os
import glob
import streamlit as st

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
# 2 — Water temperature (site-level)
# --------------------------------------------------------------------------- #
step("2 · Water temperature")
st.caption("The temperature of the water at your site. Filtration rate changes "
           "with temperature, so use a value that reflects the period you care about.")
temp = st.number_input("Water temperature (°C)", min_value=0.0, max_value=40.0,
                       value=18.0, step=0.5, format="%.1f")


# --------------------------------------------------------------------------- #
# 3 — Size classes + dry-weight source
# --------------------------------------------------------------------------- #
step("3 · Your bivalves")

sp_slug = fl.norm_species(sp.name).lower().replace(" ", "_")
convs = conversions_for(sp.name)
rows_for_calc = []          # list of {shell_mm, count, dtw_g}
derived_dtw = []            # parallel list of derived DTW (estimate mode) or None

if sp.uses_length:
    # Length-based equation (e.g. Mytilus edulis): shell height drives CR directly.
    st.caption("This species' equation uses **shell height** directly, so no dry "
               "weight is needed. Add a row per size class — the totals are summed.")
    import pandas as pd
    seed = pd.DataFrame([{"Shell height (mm)": 50.0, "Number of bivalves": 1000}])
    ed = st.data_editor(
        seed, num_rows="dynamic", use_container_width=True, hide_index=True,
        key=f"sc_{sp_slug}_len",
        column_config={
            "Shell height (mm)": st.column_config.NumberColumn(
                min_value=0.0, max_value=400.0, step=1.0, format="%.1f"),
            "Number of bivalves": st.column_config.NumberColumn(
                min_value=0, step=100, format="%d"),
        })
    for _, r in ed.iterrows():
        rows_for_calc.append({"shell_mm": r.get("Shell height (mm)"),
                              "count": r.get("Number of bivalves"),
                              "dtw_g": None})
        derived_dtw.append(None)

else:
    # Dry-weight-based equation: choose how dry weight is supplied.
    opt_one = "Use one dry weight for all"
    opt_est = "Estimate dry weight from my shell height"
    opt_meas = "I have my own dry weights by size class (more accurate)"
    options = [opt_one, opt_meas, opt_est]

    st.caption("This species' equation uses **dry tissue weight (g)**. Choose how "
               "to provide it:")
    src = st.radio("Dry-weight source", options, label_visibility="collapsed")

    if src == opt_est and not convs:
        st.warning("No length-weight conversion is available yet for this species, "
                   "so shell height can't be turned into dry weight. Use *measured "
                   "dry weights* or *one dry weight for all* for now — conversions "
                   "will be added as the data become available.")

    import pandas as pd

    if src == opt_one:
        st.caption("A single dry tissue weight is applied to every animal. The "
                   "default is a placeholder — enter your own value if you have it.")
        oc1, oc2 = st.columns(2)
        with oc1:
            dtw_one = st.number_input("Dry tissue weight (g)", min_value=0.0,
                                      value=1.0, step=0.1, format="%.3f")
        with oc2:
            n_one = st.number_input("Number of bivalves", min_value=0, value=1000,
                                    step=100, format="%d")
        rows_for_calc.append({"shell_mm": None, "count": n_one, "dtw_g": dtw_one})
        derived_dtw.append(None)

    elif src == opt_meas:
        st.caption("Enter one row per size class with its **measured** dry tissue "
                   "weight and how many animals are in that class. Rows are summed.")
        seed = pd.DataFrame([{"Dry weight (g)": 1.0, "Number of bivalves": 1000}])
        ed = st.data_editor(
            seed, num_rows="dynamic", use_container_width=True, hide_index=True,
            key=f"sc_{sp_slug}_meas",
            column_config={
                "Dry weight (g)": st.column_config.NumberColumn(
                    min_value=0.0, step=0.1, format="%.3f"),
                "Number of bivalves": st.column_config.NumberColumn(
                    min_value=0, step=100, format="%d"),
            })
        for _, r in ed.iterrows():
            rows_for_calc.append({"shell_mm": None,
                                  "count": r.get("Number of bivalves"),
                                  "dtw_g": r.get("Dry weight (g)")})
            derived_dtw.append(None)

    else:  # estimate from shell height
        chosen_conv = None
        if convs:
            if len(convs) == 1:
                chosen_conv = convs[0]
                st.caption(f"Using conversion: **{chosen_conv.label()}** — "
                           f"`{chosen_conv.eq_text}`")
            else:
                labels = [c.label() for c in convs]
                pick = st.selectbox("Length-weight conversion to use", labels)
                chosen_conv = convs[labels.index(pick)]
                st.caption(f"`{chosen_conv.eq_text}`")
        st.caption("Enter one row per size class by **shell height**. Dry weight is "
                   "estimated from the conversion and shown in the results below.")
        seed = pd.DataFrame([{"Shell height (mm)": 50.0, "Number of bivalves": 1000}])
        ed = st.data_editor(
            seed, num_rows="dynamic", use_container_width=True, hide_index=True,
            key=f"sc_{sp_slug}_est",
            column_config={
                "Shell height (mm)": st.column_config.NumberColumn(
                    min_value=0.0, max_value=400.0, step=1.0, format="%.1f"),
                "Number of bivalves": st.column_config.NumberColumn(
                    min_value=0, step=100, format="%d"),
            })
        for _, r in ed.iterrows():
            L = r.get("Shell height (mm)")
            dtw = None
            if chosen_conv and L is not None:
                try:
                    dtw = chosen_conv.dtw_from_length(L)
                except fl.ExprError:
                    dtw = None
            rows_for_calc.append({"shell_mm": L,
                                  "count": r.get("Number of bivalves"),
                                  "dtw_g": dtw})
            derived_dtw.append(dtw)


# --------------------------------------------------------------------------- #
# Results
# --------------------------------------------------------------------------- #
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

# Per-row breakdown table
import pandas as pd
show = []
for i, e in enumerate(results):
    row = {}
    if sp.uses_length:
        row["Shell height (mm)"] = e["shell_mm"]
    else:
        if e.get("dtw_g") is not None:
            row["Dry weight (g)"] = round(e["dtw_g"], 3)
        if derived_dtw[i] is not None and not sp.uses_length:
            row["Shell height (mm)"] = e["shell_mm"]
    row["Number of bivalves"] = fmt_int(e["count"])
    row["Per animal (L/hr)"] = sig(e["cr_lph"], 3) if e["cr_lph"] is not None else "—"
    row["Filtration (L/h)"] = sig(e["filt_lph"], 3) if e["filt_lph"] is not None else "—"
    row["Filtration (m³/day)"] = sig(e["filt_m3d"], 3) if e["filt_m3d"] is not None else "—"
    if e["error"]:
        row["Note"] = e["error"]
    show.append(row)

if show:
    st.dataframe(pd.DataFrame(show), use_container_width=True, hide_index=True)

for n in notes:
    st.caption("⚠ " + n)


# --------------------------------------------------------------------------- #
# 4 — Dry-weight protocol downloads (greyed until a file is added)
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
# 5 — Length-weight conversions library
# --------------------------------------------------------------------------- #
step("Length-weight conversions")
st.caption("Shell height → dry tissue weight conversions in the library. More "
           "species will be added; each is documented with its location, season "
           "and source.")
if CONVS:
    import pandas as pd

    def _conv_link(c):
        hits = fl.references_for(c.reference, REFS)
        return hits[0].link if hits and hits[0].link else ""

    dfc = pd.DataFrame([{
        "Species": c.species, "Location": c.location, "Season": c.season,
        "Conversion (mm → g dry weight)": c.eq_text,
        "Reference": c.reference, "Source": _conv_link(c),
    } for c in CONVS])
    st.dataframe(dfc, use_container_width=True, hide_index=True,
                 column_config={"Source": st.column_config.LinkColumn(
                     "Source", display_text="open")})
else:
    st.info("No length-weight conversions in the library yet.")


# --------------------------------------------------------------------------- #
# 6 — References (download all)
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
