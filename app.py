"""
AquaPIP — MEL Monitoring Survey  (Streamlit app)

Flow:  survey  ->  learn_indicators (indicator education)  ->
       learn_classes (how to read the results)  ->  results

Reads two workbooks from data/ (both by header name, so they update live):
  - mel_database.xlsx            : the literature / protocol database
  - MEL_Indicator_Reference.xlsx : sector-filtered indicators + teaching fields + species
"""

import os
import streamlit as st

import mel_logic as ml
import report
import storage

# --------------------------------------------------------------------------- #
# Config + constants
# --------------------------------------------------------------------------- #
st.set_page_config(page_title="AquaPIP - MEL Monitoring Survey", page_icon="🌊",
                   layout="wide", initial_sidebar_state="collapsed")

HERE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(HERE, "data", "mel_database.xlsx")
REF_PATH = os.path.join(HERE, "data", "MEL_Indicator_Reference.xlsx")
COVER_IMG = os.path.join(HERE, "assets", "mel_framework_cover.png")
FRAMEWORK_URL = ("https://www.aquaculturescience.org/content/dam/tnc/nature/en/"
                 "documents/MEL_Framework_TNC_Final_MedRes.pdf")
TEAL = "#0B4F5C"
AREA_COLOR = {"Habitat & Biodiversity": "#2F8F5B",
              "Water Quality": "#2E7D9A",
              "Climate Change": "#0B4F5C"}
GROUP_KEYS = {"Habitat & Biodiversity": "ind_hb",
              "Water Quality": "ind_wq",
              "Climate Change": "ind_cc"}
CONSENT_TEXT = ("You can store my responses. Your responses will only be used for "
                "internal tracking purposes, and not for any other uses unless "
                "explicitly authorized in writing by you.")

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
  .chip {{ color:#fff; font-size:.72rem; font-weight:700; padding:.12rem .5rem; border-radius:6px; font-family:monospace; }}
  .k {{ font-size:.66rem; font-weight:700; letter-spacing:.03em; text-transform:uppercase; color:#5B6B70; margin-top:.45rem; }}
  .v {{ font-size:.9rem; color:#1A2B2F; }}
  .proto {{ background:#fff; border:1px solid #E1E7E9; border-left:5px solid {TEAL}; border-radius:0 10px 10px 0;
            padding:.7rem .95rem; margin:.35rem 0 .55rem; }}
  .pill {{ color:#1A2B2F; font-weight:700; font-size:.74rem; padding:.12rem .5rem; border-radius:999px; border:1px solid rgba(0,0,0,.06); }}
  .tag {{ background:#EEF3F4; color:#33474C; font-size:.74rem; padding:.12rem .5rem; border-radius:999px; margin-left:.3rem; }}
  .fit {{ background:{TEAL}; color:#fff; font-size:.7rem; font-weight:700; padding:.1rem .45rem; border-radius:999px; }}
  .sq {{ display:inline-block; width:.7rem; height:.7rem; border-radius:2px; margin-right:.4rem; vertical-align:middle; border:1px solid rgba(0,0,0,.08); }}
  .learnbox {{ background:#F3F6F7; border:1px solid #E1E7E9; border-radius:10px; padding:.6rem .8rem; }}
</style>
""", unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# Data
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner=False)
def get_data(_m_db, _m_ref):
    rows, vocab, headers = ml.load_database(DB_PATH)
    ref = ml.load_reference(REF_PATH)
    prof_opts = ml.profile_options(vocab, rows)
    return rows, vocab, prof_opts, ref


for p, name in [(DB_PATH, "data/mel_database.xlsx"),
                (REF_PATH, "data/MEL_Indicator_Reference.xlsx")]:
    if not os.path.exists(p):
        st.error(f"Missing `{name}`. Make sure it's in the repo's data/ folder.")
        st.stop()

rows, vocab, prof_opts, ref = get_data(os.path.getmtime(DB_PATH), os.path.getmtime(REF_PATH))
AQUA_OPTS = vocab.get("Aquaculture Type") or [
    "Seaweed", "Mollusk", "Echinoderm", "Finfish", "Multi-trophic (IMTA)", "Not specific"]

if "view" not in st.session_state:
    st.session_state.view = "survey"


# --------------------------------------------------------------------------- #
# Small render helpers
# --------------------------------------------------------------------------- #
def step(txt):
    st.markdown(f"<div class='step'>{txt}</div>", unsafe_allow_html=True)


def dots(o):
    return ml.dots(o)


def indicator_card(ind):
    """Full teaching card for one reference indicator (interim page 1)."""
    color = AREA_COLOR.get(ind["area"], TEAL)
    fields = [("Goal", ind["goal"]), ("Objective", ind["objective"]),
              ("What is measured", ind["measured"]), ("Complexity", ind["complexity"]),
              ("Metric", ind["metric"]), ("Suggested method", ind["method"]),
              ("Proxy / additional method", ind["proxy"]),
              ("Frequency / timing", ind["frequency"]),
              ("Location of sampling", ind["location"])]
    body = "".join(f"<div class='k'>{k}</div><div class='v'>{v}</div>"
                   for k, v in fields if v)
    return (f"<div class='card' style='border-left-color:{color}'>"
            f"<div style='display:flex;align-items:center;gap:.5rem;margin-bottom:.2rem'>"
            f"<span class='chip' style='background:{color}'>{ind['code']}</span>"
            f"<span style='font-weight:700;color:#1A2B2F;font-size:1rem'>{ind['label']}</span></div>"
            f"{body}</div>")


# --------------------------------------------------------------------------- #
# VIEW 1 - SURVEY
# --------------------------------------------------------------------------- #
def render_survey():
    st.markdown("<div class='hero'><h1>Regenerative Aquaculture - MEL Monitoring Survey</h1>"
                "<p>Tell us about your farm and what you'd like to measure. We'll match you to "
                "monitoring protocols from the TNC MEL evidence base - with the most farmer-friendly "
                "options first.</p></div>", unsafe_allow_html=True)
    st.divider()

    # --- Step 1: aquaculture type (drives everything) ---
    step("1 - What do you farm?")
    aqua = st.selectbox("Aquaculture type", AQUA_OPTS, index=None,
                        placeholder="Select your farm type...", key="aqua_type",
                        label_visibility="collapsed")
    if not aqua:
        st.info("Pick your farm type to see the indicators and species that apply to it.")
        return

    ind_groups = ml.sector_indicator_groups(ref, aqua)
    code2label, label2code, code2goal = ml.flatten_options(ind_groups)
    available_codes = set(code2label)
    species_opts = ml.species_options_for(ref, aqua)

    # reactive reset when the farm type changes
    if st.session_state.get("_prev_aqua") != aqua:
        st.session_state["_prev_aqua"] = aqua
        st.session_state["prof_species"] = "- any -"
        st.session_state["_prev_fg"] = None
        for k in GROUP_KEYS.values():
            st.session_state[k] = []

    # --- Step 2: goals ---
    step("2 - What are you interested in measuring?")
    st.caption("Optional. Each choice pre-selects a set of MEL indicators you can fine-tune in step 4.")
    fg = st.multiselect("Goals", list(ml.FARMER_GOALS.keys()), key="farmer_goals",
                        label_visibility="collapsed", placeholder="Choose one or more...")
    for g in fg:
        st.caption(f"**{g}** - {ml.FARMER_GOALS[g]['blurb']}")

    # seed indicators from goals ∩ sector when goals (or farm type) changed
    if st.session_state.get("_prev_fg") != fg:
        st.session_state["_prev_fg"] = list(fg)
        derived = ml.codes_for_farmer_goals(fg, available_codes)
        for area, key in GROUP_KEYS.items():
            st.session_state[key] = [code2label[c] for c in derived if code2goal.get(c) == area]

    # --- Step 3: about your farm ---
    step("3 - About your farm")
    st.caption("Optional - leave any field on \u201c- any -\u201d. These never hide protocols; they "
               "move the best-fitting ones up and add a \u201cfits your farm\u201d marker.")
    farm_name = st.text_input("Farm / site name (optional)", key="farm_name_input",
                              placeholder="e.g. Tidewater Kelp Co., Casco Bay")
    profile = {}
    c1, c2 = st.columns(2)
    with c1:
        sp = st.selectbox("Cultivated species", ["- any -"] + species_opts, key="prof_species")
        if sp and sp != "- any -":
            profile["species"] = sp
    with c2:
        if ml.algae_applicable(aqua):
            al = st.selectbox("Algae type", ["- any -"] + (prof_opts.get("algae_type") or []),
                              key="prof_algae")
            if al and al != "- any -":
                profile["algae_type"] = al
        else:
            st.selectbox("Algae type", ["Not applicable (non-seaweed system)"],
                         disabled=True, key="prof_algae_na")

    rest = [("farm_structure", "Farm structure"), ("farm_scale", "Farm scale"),
            ("coculture", "Co-culture context"), ("climate_zone", "Climate zone"),
            ("site_depth", "Site depth"), ("wave_energy", "Wave energy"),
            ("water_clarity", "Water clarity")]
    rc = st.columns(2)
    for i, (fkey, label) in enumerate(rest):
        with rc[i % 2]:
            v = st.selectbox(label, ["- any -"] + (prof_opts.get(fkey) or []), key=f"prof_{fkey}")
            if v and v != "- any -":
                profile[fkey] = v

    # --- Step 4: indicators (sector-filtered) ---
    step("4 - Refine your MEL indicators")
    st.caption(f"Showing indicators that apply to **{aqua}**. Pre-filled from your goals - "
               f"add or remove any.")
    gc = st.columns(3)
    for i, area in enumerate(["Habitat & Biodiversity", "Water Quality", "Climate Change"]):
        labels = [it["label"] for it in ind_groups[area]]
        with gc[i]:
            st.markdown(f"**{area}**")
            if labels:
                st.multiselect(area, labels, key=GROUP_KEYS[area],
                               label_visibility="collapsed", placeholder="none selected")
            else:
                st.caption("_Not applicable for this farm type._")
                st.session_state[GROUP_KEYS[area]] = []

    codes = []
    for area in ["Habitat & Biodiversity", "Water Quality", "Climate Change"]:
        for it in ind_groups[area]:
            if it["label"] in st.session_state.get(GROUP_KEYS[area], []) and it["code"] not in codes:
                codes.append(it["code"])
    st.caption(f"**{len(codes)}** indicator(s) selected.")

    # --- Step 5: consent + continue ---
    step("5 - Save & continue")
    consent = st.checkbox(CONSENT_TEXT, value=False, key="consent")
    if st.button("Continue to your indicators  ->", type="primary"):
        if not codes:
            st.warning("Please select at least one indicator (or pick a goal in step 2 to pre-fill them).")
        else:
            sub = {"farm_name": (farm_name or "").strip(), "aqua_type": aqua,
                   "farmer_goals": list(fg), "codes": codes, "profile": profile,
                   "goals": ml.goals_from_codes(codes), "code2label": code2label,
                   "consent": bool(consent)}
            st.session_state["submission"] = sub
            st.session_state["_save_status"] = (
                storage.save_response(sub) if consent else (None, "no_consent"))
            st.session_state["view"] = "learn_indicators"
            st.rerun()


# --------------------------------------------------------------------------- #
# VIEW 2 - LEARN INDICATORS (interim page 1)
# --------------------------------------------------------------------------- #
def render_learn_indicators():
    sub = st.session_state.get("submission")
    if not sub:
        st.session_state.view = "survey"; st.rerun(); return
    aqua = sub["aqua_type"]

    nav = st.columns(2)
    with nav[0]:
        if st.button("<-  Want to add other indicators? Return to survey", use_container_width=True):
            st.session_state.view = "survey"; st.rerun()
    with nav[1]:
        if st.button("Ready to continue?  ->", type="primary", use_container_width=True):
            st.session_state.view = "learn_classes"; st.rerun()

    st.markdown(f"<div class='hero'><h1>Your MEL indicators</h1>"
                f"<p>A quick primer on what each indicator you chose is really tracking, and how it's "
                f"measured - for a <b>{aqua}</b> farm.</p></div>", unsafe_allow_html=True)

    # framework "learn more" box with thumbnail
    b1, b2 = st.columns([1, 5])
    with b1:
        if os.path.exists(COVER_IMG):
            st.image(COVER_IMG, use_container_width=True)
    with b2:
        st.markdown("<div class='learnbox'><b>Interested in learning more about these indicators?</b><br>"
                    "Read them in full detail in the TNC MEL Framework.</div>", unsafe_allow_html=True)
        st.link_button("Open the MEL Framework (PDF)  \u2197", FRAMEWORK_URL)
    st.divider()

    sector_inds = ml.indicators_for_sector(ref, aqua)
    selected = [i for i in sector_inds if i["code"] in sub["codes"]]
    others = [i for i in sector_inds if i["code"] not in sub["codes"]]

    st.subheader("The indicators you selected")
    if not selected:
        st.info("You haven't selected any indicators yet - go back and choose some, or explore the "
                "full set below.")
    for ind in selected:
        st.markdown(indicator_card(ind), unsafe_allow_html=True)

    if others:
        st.markdown("#### Curious about the other indicators? Explore these below.")
        for ind in others:
            with st.expander(f"{ind['code']}  \u00b7  {ind['label']}"):
                st.markdown(indicator_card(ind), unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# VIEW 3 - LEARN CLASSES (interim page 2)
# --------------------------------------------------------------------------- #
def render_learn_classes():
    if not st.session_state.get("submission"):
        st.session_state.view = "survey"; st.rerun(); return

    nav = st.columns(2)
    with nav[0]:
        if st.button("<-  Back to indicators", use_container_width=True):
            st.session_state.view = "learn_indicators"; st.rerun()
    with nav[1]:
        if st.button("See my recommendations  ->", type="primary", use_container_width=True):
            st.session_state.view = "results"; st.rerun()

    st.markdown("<div class='hero'><h1>Understanding your protocol recommendations</h1>"
                "<p>A quick guide to the labels you'll see on the next page.</p></div>",
                unsafe_allow_html=True)

    st.markdown(
        "On the next page you'll see a set of monitoring methods - we call them **protocols** - chosen "
        "to match the MEL indicators you picked. Each comes from the published record: peer-reviewed "
        "scientific studies, plus *grey* and *white* literature such as government manuals, industry "
        "guides, and technical reports. We've read each source and sorted them so the most "
        "farmer-friendly options rise to the top. Here's what the labels mean.")

    st.markdown("##### Who can run it - the tier (T4-T1)")
    for n, meta in ml.TIER_META.items():
        st.markdown(f"<span class='sq' style='background:{meta['color']}'></span>"
                    f"<b>T{n} - {meta['label']}</b>", unsafe_allow_html=True)
    st.markdown(
        "- **T4 - Practitioner** (green): you can run it essentially on your own, with gear and skills "
        "common on a working farm - visual transects, baited camera drops, basket traps.\n"
        "- **T3 - Partnership** (pale green): you do the fieldwork, but one step - usually the lab "
        "analysis - is best handed to a partner such as a local university or a commercial lab.\n"
        "- **T2 - Researcher** (amber): assumes a specialist team throughout, or specialized "
        "equipment and analysis (satellite mapping, eDNA, carbonate chemistry). It's here so you can "
        "see what's possible and know who to call.\n"
        "- **T1 - Reference** (coral): not a step-by-step method, but useful background - a review or "
        "synthesis to help you understand the topic.")
    st.caption("A higher tier isn't \u201cbetter\u201d - it just takes more specialized hands. Many farms "
               "start with T4 and T3 methods and bring in partners over time.")

    st.markdown("##### Skill")
    st.markdown("Beside the tier you'll see **Basic, Intermediate, or Advanced** - the plain-language "
                "read of the tier (how much specialized training the method assumes). Basic tracks with "
                "T4, Intermediate with T3, Advanced with T2. If you're already comfortable with a "
                "technique, trust that experience; the skill level is a guide, not a gate.")

    st.markdown("##### Cost")
    st.markdown("Shown as dots (low \u25cf to high \u25cf\u25cf\u25cf), covering what it takes to get the "
                "data - equipment plus any lab or analysis fees. Where a method's cost spans a range, we "
                "show the upper end so you're never caught short on budget. A planning guide, not a quote.")

    st.markdown("##### Effort")
    st.markdown("Also dots, but about your **time** - how long the sampling or deployment runs. A "
                "one-day survey is low; a season of repeated visits, medium; year-round or continuous "
                "monitoring, high. This is field time, not analysis time.")

    st.markdown("##### \u2605 Strong fit to your farm")
    st.markdown("Some protocols carry a star, meaning the method lines up well with the farm details "
                "you gave - your species, system, scale, and setting. It's a nudge toward options "
                "proven in conditions like yours; it never hides the others, and an un-starred protocol "
                "can still be a great choice.")

    st.markdown("##### On each protocol you'll also find")
    st.markdown("- **Set-up** - what you'll need on hand to begin (equipment and arrangements like "
                "booking lab time).\n"
                "- **Summary** - a plain-language description of how the method is carried out on a farm.\n"
                "- **Protocol details** - the fuller picture: analysis, caveats, and a link to the "
                "original source.\n"
                "- **Downloads** *(coming soon)* - where available, a **normalized protocol** (the method "
                "rewritten as clean, farm-ready steps), a **field data sheet** (a printable form for your "
                "measurements), and a **stats workbook** (a spreadsheet set up to crunch the numbers).")
    st.caption("Take what's useful, leave what isn't, and start where you're comfortable.")


# --------------------------------------------------------------------------- #
# VIEW 4 - RESULTS
# --------------------------------------------------------------------------- #
def _tier_legend():
    with st.popover("What do T4-T1 mean?"):
        for n, meta in ml.TIER_META.items():
            st.markdown(f"<span class='sq' style='background:{meta['color']}'></span>"
                        f"<b>T{n} - {meta['label']}</b> \u00b7 {meta['skill']}",
                        unsafe_allow_html=True)
        st.caption("T4 you can run yourself \u00b7 T3 fieldwork by you + a lab step by a partner \u00b7 "
                   "T2 specialist team throughout \u00b7 T1 background reference (not run).")


def _cse_legend():
    with st.popover("Cost, skill & effort"):
        st.markdown("**Cost** \u25cf\u25cb\u25cb -> \u25cf\u25cf\u25cf - equipment + lab/analysis fees "
                    "(ranges shown at the upper end).")
        st.markdown("**Effort** \u25cf\u25cb\u25cb -> \u25cf\u25cf\u25cf - your field time; low = a day, "
                    "high = year-round / continuous.")
        st.markdown("**Skill** - Basic (T4) \u00b7 Intermediate (T3) \u00b7 Advanced (T2): how much "
                    "specialized training the method assumes.")
        st.markdown("**\u2605 fits your farm** - the method matches the farm details you gave.")


def render_results():
    sub = st.session_state.get("submission")
    if not sub:
        st.session_state.view = "survey"; st.rerun(); return
    code2label = sub.get("code2label", {})

    top = st.columns([3, 1])
    with top[0]:
        farm = sub.get("farm_name") or "Your farm"
        st.markdown(f"<div class='hero'><h1>Your MEL Protocol Recommendations</h1>"
                    f"<p>{farm} \u00b7 {sub['aqua_type']} \u00b7 one best practitioner-runnable protocol "
                    f"per selected indicator.</p></div>", unsafe_allow_html=True)
    with top[1]:
        if st.button("<-  Edit answers", use_container_width=True):
            st.session_state.view = "survey"; st.rerun()

    ok, _ = st.session_state.get("_save_status", (None, ""))
    if ok:
        st.caption("\u2713 Your responses were saved for internal tracking.")

    lg = st.columns([1, 1, 3])
    with lg[0]:
        _tier_legend()
    with lg[1]:
        _cse_legend()
    st.divider()

    matches = ml.match_protocols(rows, sub["codes"], sub["profile"], farmer_aqua=sub["aqua_type"])
    left, right = st.columns([2, 1])

    with left:
        st.subheader("Recommended protocols")
        for code, cands in matches.items():
            label = code2label.get(code, code)
            if not cands:
                st.markdown(f"<div class='proto' style='border-left-color:#CCC'>"
                            f"<b style='color:{TEAL}'>{label}</b><br>"
                            f"<span style='color:#5B6B70;font-size:.86rem'>No protocol in the current "
                            f"database for this indicator and farm type yet.</span></div>",
                            unsafe_allow_html=True)
                continue
            best = cands[0]; r = best["row"]
            fit = "<span class='fit'>\u2605 fits your farm</span>" if best["badge"] else ""
            st.markdown(
                f"<div class='proto'><div style='color:{TEAL};font-weight:700;font-size:.85rem'>"
                f"{label} {fit}</div>"
                f"<div style='font-weight:700;color:#1A2B2F;margin:.15rem 0 .35rem'>{r['title']}</div>"
                f"<span class='pill' style='background:{r['tier_color']}'>T{r['tier']} \u00b7 {r['tier_label']}</span>"
                f"<span class='tag'>Skill: {r['skill']}</span>"
                f"<span class='tag'>Cost {dots(r['cost_ord'])}</span>"
                f"<span class='tag'>Effort {dots(r['effort_ord'])}</span></div>",
                unsafe_allow_html=True)
            if len(cands) > 1:
                with st.expander(f"{len(cands) - 1} more option(s) for {code}"):
                    for alt in cands[1:]:
                        ar = alt["row"]; star = " \u2605" if alt["badge"] else ""
                        st.markdown(f"<span class='sq' style='background:{ar['tier_color']}'></span>"
                                    f"<b>T{ar['tier']} {ar['tier_label']}</b>{star} - {ar['title']} "
                                    f"<span style='color:#5B6B70'>\u00b7 cost {dots(ar['cost_ord'])} "
                                    f"\u00b7 effort {dots(ar['effort_ord'])}</span>", unsafe_allow_html=True)

    with right:
        st.subheader("Your survey")
        html = ["<div class='learnbox'>"]
        html.append("<div class='k'>Farm type</div><div class='v'>" + sub["aqua_type"] + "</div>")
        if sub.get("farmer_goals"):
            html.append("<div class='k'>Interested in measuring</div><div class='v'>"
                        + "<br>".join(sub["farmer_goals"]) + "</div>")
        html.append("<div class='k'>MEL goals</div><div class='v'>"
                    + (" \u00b7 ".join(sub["goals"]) or "-") + "</div>")
        for fkey, label, _vh, _rk in ml.PROFILE_FIELDS:
            if sub["profile"].get(fkey):
                html.append(f"<div class='k'>{label}</div><div class='v'>{sub['profile'][fkey]}</div>")
        html.append("<div class='k'>Indicators selected</div>")
        for code in sub["codes"]:
            cc = matches.get(code, [])
            color = cc[0]["row"]["tier_color"] if cc else "#CCC"
            html.append(f"<div class='v'><span class='sq' style='background:{color}'></span>"
                        f"{code2label.get(code, code)}</div>")
        html.append("</div>")
        st.markdown("".join(html), unsafe_allow_html=True)

    st.divider()
    try:
        pdf = report.build_pdf(sub, matches, code2label)
        fname = (sub.get("farm_name") or "MEL").replace(" ", "_")[:40]
        st.download_button("Download PDF report", data=pdf,
                           file_name=f"{fname}_MEL_recommendations.pdf",
                           mime="application/pdf", type="primary")
    except Exception as e:
        st.info(f"PDF export is temporarily unavailable ({type(e).__name__}).")

    st.divider()
    st.subheader("Full protocol details")
    shown = set()
    for code, cands in matches.items():
        if not cands:
            continue
        r = cands[0]["row"]
        if r["row"] in shown:
            continue
        shown.add(r["row"])
        covers = [code2label.get(c, c) for c, cc in matches.items()
                  if cc and cc[0]["row"]["row"] == r["row"]]
        with st.expander(f"{r['title']}  \u00b7  T{r['tier']} {r['tier_label']}"):
            _detail(r, covers)


def _detail(r, covers):
    st.markdown(f"**Covers:** {' \u00b7 '.join(covers)}")
    meta = " \u00b7 ".join(x for x in [r["authors"], r["year"], r["publication"]] if x)
    if meta:
        st.caption(meta)
    st.markdown(f"<span class='pill' style='background:{r['tier_color']}'>T{r['tier']} \u00b7 {r['tier_label']}</span> "
                f"<span class='tag'>Skill: {r['skill']}</span> "
                f"<span class='tag'>Cost {dots(r['cost_ord'])} ({r['cost_label']})</span> "
                f"<span class='tag'>Effort {dots(r['effort_ord'])} ({r['effort_label']})</span>",
                unsafe_allow_html=True)
    st.write("")

    def f(lbl, val):
        if val:
            st.markdown(f"**{lbl}**"); st.write(val)

    f("Set-up (equipment)", r["equipment"])
    if r["method_summary"]:
        f("Summary", r["method_summary"])
    else:
        st.markdown("**Summary**")
        st.caption("Method summary pending - appears once the database includes it for this protocol.")
    f("Statistical approach", r["stats"])
    f("Executable by practitioner?", r["exec_practitioner"])
    f("Notes / caveats", r["notes"])
    url = report.first_url(r["url"])
    if url:
        st.markdown(f"**Source:** [{url}]({url})")

    st.markdown("**Downloads**")
    d = st.columns(3)
    for col, lbl, val in [(d[0], "Normalized protocol", r["url_template"]),
                          (d[1], "Field data sheet", r["url_datasheet"]),
                          (d[2], "Stats workbook", r["url_workbook"])]:
        u = report.first_url(val)
        with col:
            if u:
                st.link_button(f"Download {lbl}", u, use_container_width=True)
            else:
                st.button(f"{lbl} (V2)", disabled=True, use_container_width=True,
                          key=f"dl_{r['row']}_{lbl}", help="Coming in V2 - not yet available.")


# --------------------------------------------------------------------------- #
# Router
# --------------------------------------------------------------------------- #
v = st.session_state.view
if v == "survey":
    render_survey()
elif v == "learn_indicators":
    render_learn_indicators()
elif v == "learn_classes":
    render_learn_classes()
else:
    render_results()
