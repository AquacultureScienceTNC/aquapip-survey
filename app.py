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
import monitoring

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
    mon = monitoring.load(REF_PATH)   # "what are you already monitoring?" screener
    return rows, vocab, prof_opts, ref, mon


for p, name in [(DB_PATH, "data/mel_database.xlsx"),
                (REF_PATH, "data/MEL_Indicator_Reference.xlsx")]:
    if not os.path.exists(p):
        st.error(f"Missing `{name}`. Make sure it's in the repo's data/ folder.")
        st.stop()

rows, vocab, prof_opts, ref, MON = get_data(os.path.getmtime(DB_PATH), os.path.getmtime(REF_PATH))
AQUA_OPTS = vocab.get("Aquaculture Type") or [
    "Seaweed", "Mollusk", "Echinoderm", "Finfish", "Multi-trophic (IMTA)", "Not specific"]
CODE_NAMES = ml.code_name_map(ref)   # code -> farmer-facing name (for scrubbing free text)

if "view" not in st.session_state:
    st.session_state.view = "survey"


# --------------------------------------------------------------------------- #
# Small render helpers
# --------------------------------------------------------------------------- #
def step(txt):
    st.markdown(f"<div class='step'>{txt}</div>", unsafe_allow_html=True)


def dots(o):
    return ml.dots(o)


# widget keys whose values must survive leaving and returning to the survey view
PERSIST_KEYS = ["aqua_type", "farmer_goals", "farm_name_input",
                "prof_species", "prof_algae", "prof_farm_structure", "prof_farm_scale",
                "prof_coculture", "prof_climate_zone", "prof_site_depth",
                "prof_wave_energy", "prof_water_clarity",
                "ind_hb", "ind_wq", "ind_cc",
                "consent", "resp_gps", "resp_name", "resp_contact"]


def _persist_restore():
    """Streamlit clears a widget's stored value when that widget isn't rendered
    (which happens the moment you leave the survey view). Restore each answer
    from its shadow copy before the widgets are built, so going back keeps them."""
    for k in PERSIST_KEYS:
        sk = "_keep_" + k
        if sk in st.session_state and k not in st.session_state:
            st.session_state[k] = st.session_state[sk]


def _persist_save():
    """Mirror the current answers into shadow keys Streamlit won't garbage-collect."""
    for k in PERSIST_KEYS:
        if k in st.session_state:
            st.session_state["_keep_" + k] = st.session_state[k]


def _reset_survey():
    """Clear all survey answers for a fresh start (Start-over button)."""
    keys = (PERSIST_KEYS + ["_keep_" + k for k in PERSIST_KEYS]
            + ["_prev_aqua", "_prev_fg", "prof_algae_na", "p1",
               "submission", "primary_override", "_save_status", "monitoring_sel"])
    for k in keys:
        st.session_state.pop(k, None)
    for k in list(st.session_state):          # dynamic monitoring keys + shadows
        if k.startswith("mon_") or k.startswith("_keep_mon_"):
            st.session_state.pop(k, None)
    st.session_state["view"] = "survey"


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
    _persist_restore()
    st.markdown("<div class='hero'><h1>Regenerative Aquaculture - MEL Monitoring Survey</h1>"
                "<p>Tell us about your farm and what you'd like to measure. We'll match you to "
                "monitoring protocols from the TNC MEL evidence base - with the most farmer-friendly "
                "options first.</p></div>", unsafe_allow_html=True)
    st.divider()

    # --- Step 1: aquaculture type (drives everything) ---
    step("1 - What do you farm?")
    _idx = {} if "aqua_type" in st.session_state else {"index": None}
    aqua = st.selectbox("Aquaculture type", AQUA_OPTS,
                        placeholder="Select your farm type...", key="aqua_type",
                        label_visibility="collapsed", **_idx)
    if not aqua:
        st.info("Pick your farm type to see the indicators and species that apply to it.")
        return

    ind_groups = ml.sector_indicator_groups(ref, aqua)
    code2label, label2code, code2goal = ml.flatten_options(ind_groups)
    available_codes = set(code2label)
    species_opts = ml.species_options_for(ref, aqua)
    structure_opts = ml.structure_options_for(ref, aqua)

    # reactive reset when the farm type changes
    if st.session_state.get("_prev_aqua") != aqua:
        st.session_state["_prev_aqua"] = aqua
        st.session_state["prof_species"] = "- any -"
        st.session_state["prof_farm_structure"] = "- any -"
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
        if ml.algae_applicable(aqua) and not profile.get("species"):
            al = st.selectbox("Algae type", ["- any -"] + (prof_opts.get("algae_type") or []),
                              key="prof_algae")
            if al and al != "- any -":
                profile["algae_type"] = al
        else:
            reason = ("Not needed - you selected a species" if profile.get("species")
                      else "Not applicable (non-seaweed system)")
            st.selectbox("Algae type", [reason], disabled=True, key="prof_algae_na")

    c3, c4 = st.columns(2)
    with c3:
        stv = st.selectbox("Farm structure", ["- any -"] + structure_opts, key="prof_farm_structure")
        if stv and stv != "- any -":
            profile["farm_structure"] = stv
    with c4:
        fs = st.selectbox("Farm scale", ["- any -"] + (prof_opts.get("farm_scale") or []),
                          key="prof_farm_scale")
        if fs and fs != "- any -":
            profile["farm_scale"] = fs

    rest = [("coculture", "Co-culture context"), ("climate_zone", "Climate zone"),
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

    bcol = st.columns([2, 1])
    with bcol[0]:
        go = st.button("Next: what you already monitor  ->", type="primary", use_container_width=True)
    with bcol[1]:
        if st.button("Start over (clear answers)", use_container_width=True):
            _reset_survey(); st.rerun()

    _persist_save()   # capture every answer before navigating away from the survey

    if go:
        if not codes:
            st.warning("Please select at least one indicator (or pick a goal in step 2 to pre-fill them).")
        else:
            st.session_state["p1"] = {
                "farm_name": (farm_name or "").strip(), "aqua_type": aqua,
                "farmer_goals": list(fg), "codes": codes, "profile": profile,
                "goals": ml.goals_from_codes(codes), "code2label": code2label}
            st.session_state["view"] = "monitoring"
            st.rerun()


# --------------------------------------------------------------------------- #
# VIEW 1b - WHAT YOU ALREADY MONITOR  (+ optional responses)
# --------------------------------------------------------------------------- #
def render_monitoring():
    _persist_restore()
    p1 = st.session_state.get("p1")
    if not p1:                       # reached without completing page 1
        st.session_state["view"] = "survey"; st.rerun(); return

    st.markdown(
        "<div class='hero'><h1>What are you already monitoring?</h1>"
        "<p>Most farms already collect some data - a temperature logger here, an occasional "
        "water sample there - but rarely in the systematic, repeatable way that turns "
        "observations into evidence of a co-benefit.</p></div>", unsafe_allow_html=True)
    st.divider()

    st.markdown(
        "Tell us what you already track. We'll move protocols that **build on methods you "
        "already run** toward the top - shortening the path from what you do today to a "
        "defensible **Habitat, Water Quality, or Climate** outcome. This only reorders your "
        "recommendations; it never hides a protocol, and skipping it is completely fine.")
    st.caption("Open a category and choose anything you already do - even occasionally, or not every season.")

    mon_selected = monitoring.render_screener(MON)

    st.divider()

    # --- optional responses (moved here: the last thing, after all other inputs) ---
    step("Add your responses (optional)")
    st.markdown("The Global Aquaculture Team is working to gain a better understanding of what "
                "farmers are looking to monitor on their farms. If you click this option, your "
                "responses will be added to a database for us to review.")
    consent = st.checkbox("Yes - add my responses to the database", key="consent")
    gps = cname = cinfo = ""
    if consent:
        cc1, cc2, cc3 = st.columns(3)
        with cc1:
            gps = st.text_input("GPS coordinates", key="resp_gps",
                                 placeholder="e.g. 44.1234, -68.5678")
        with cc2:
            cname = st.text_input("Your name", key="resp_name")
        with cc3:
            cinfo = st.text_input("Contact information", key="resp_contact",
                                  placeholder="email or phone")

    bcol = st.columns([1, 2, 1])
    with bcol[0]:
        back = st.button("<-  Back", use_container_width=True)
    with bcol[1]:
        go = st.button("Next  ->  MEL Framework Overview", type="primary", use_container_width=True)
    with bcol[2]:
        if st.button("Start over", use_container_width=True):
            _reset_survey(); st.rerun()

    _persist_save()

    if back:
        st.session_state["view"] = "survey"; st.rerun()
    if go:
        sub = dict(p1)
        sub.update({"gps": gps.strip(), "contact_name": cname.strip(),
                    "contact_info": cinfo.strip(), "consent": bool(consent),
                    "mon_sel": mon_selected})
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
    selected_codes = set(sub["codes"])

    nav = st.columns(2)
    with nav[0]:
        if st.button("<-  Want to add other indicators? Return to survey", use_container_width=True):
            st.session_state.view = "survey"; st.rerun()
    with nav[1]:
        if st.button("Ready to continue?  ->", type="primary", use_container_width=True):
            st.session_state.view = "learn_classes"; st.rerun()

    st.markdown(f"<div class='hero'><h1>Your MEL indicators</h1>"
                f"<p>The full MEL indicator set for a <b>{aqua}</b> farm, laid out as in the framework. "
                f"The indicators you selected are highlighted; the rest are greyed so you can see them "
                f"in context.</p></div>", unsafe_allow_html=True)

    fields = [("Goal", "goal"), ("Objective", "objective"), ("Indicator", "measured"),
              ("Complexity", "complexity"), ("Metric", "metric"), ("Suggested method", "method"),
              ("Proxy / additional method", "proxy"), ("Frequency / timing", "frequency"),
              ("Location of sampling", "location")]

    def matrix(area, inds):
        color = AREA_COLOR.get(area, TEAL)

        def cstyle(sel):
            return (f"background:#fff;color:#1A2B2F;border-top:3px solid {color};"
                    if sel else "background:#F5F6F7;color:#9AA5A8;border-top:3px solid #E4E8E9;")

        head = (f"<td style='min-width:118px;background:{color};color:#fff;font-weight:700;"
                f"padding:.45rem .55rem;vertical-align:bottom'>{area}</td>")
        for ind in inds:
            sel = ind["code"] in selected_codes
            name = ind.get("name") or ind["label"]
            if sel:
                chip = (f"<span style='background:{color};color:#fff;font-weight:700;font-size:.84rem;"
                        f"padding:.2rem .5rem;border-radius:6px;display:inline-block'>{name}</span>")
            else:
                chip = f"<span style='color:#9AA5A8;font-weight:600;font-size:.84rem'>{name}</span>"
            head += (f"<td style='{cstyle(sel)}min-width:158px;padding:.5rem .55rem;vertical-align:bottom'>"
                     f"{chip}</td>")
        rows_html = f"<tr>{head}</tr>"
        for flabel, fkey in fields:
            cells = (f"<td style='background:#EEF3F4;color:#33474C;font-weight:700;font-size:.66rem;"
                     f"text-transform:uppercase;letter-spacing:.02em;padding:.35rem .55rem;"
                     f"vertical-align:top;white-space:nowrap'>{flabel}</td>")
            for ind in inds:
                sel = ind["code"] in selected_codes
                val = ind.get(fkey) or "\u2014"
                cells += (f"<td style='{cstyle(sel)}padding:.35rem .55rem;vertical-align:top;"
                          f"font-size:.78rem;line-height:1.3'>{val}</td>")
            rows_html += f"<tr>{cells}</tr>"
        return (f"<div style='overflow-x:auto;border:1px solid #E1E7E9;border-radius:10px;"
                f"margin:.4rem 0 1.1rem'><table style='border-collapse:collapse;width:100%'>"
                f"{rows_html}</table></div>")

    sector_inds = ml.indicators_for_sector(ref, aqua)
    if not selected_codes:
        st.info("You haven't selected any indicators yet - nothing is highlighted below. Go back to "
                "add some, or read the full set here for context.")
    for area in ["Habitat & Biodiversity", "Water Quality", "Climate Change"]:
        inds = [i for i in sector_inds if i["area"] == area]
        if inds:
            st.markdown(matrix(area, inds), unsafe_allow_html=True)

    # framework "learn more" box - at the BOTTOM
    st.divider()
    b1, b2 = st.columns([1, 5])
    with b1:
        if os.path.exists(COVER_IMG):
            st.image(COVER_IMG, use_container_width=True)
    with b2:
        st.markdown("<div class='learnbox'><b>Interested in learning more about these indicators?</b><br>"
                    "Read them in full detail in the TNC MEL Framework.</div>", unsafe_allow_html=True)
        st.link_button("Open the MEL Framework (PDF)  \u2197", FRAMEWORK_URL)


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

    st.markdown("<div class='hero'><h1>Understanding our protocol recommendations</h1>"
                "<p>A quick guide to the labels you'll see on the next page.</p></div>",
                unsafe_allow_html=True)

    st.markdown(
        "On the next page you'll see a set of **protocols** - chosen "
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

    sel = sub.get("mon_sel", [])
    mon_scorer = (lambda row: monitoring.match(
        row["method_category"], row["method_summary"], sel)) if sel else None
    matches = ml.match_protocols(rows, sub["codes"], sub["profile"],
                                 farmer_aqua=sub["aqua_type"], mon_scorer=mon_scorer)

    # user overrides: move the chosen protocol to the front for that indicator,
    # so the card, the details section, and the PDF all follow the choice.
    overrides = st.session_state.setdefault("primary_override", {})
    for code, cands in matches.items():
        ov = overrides.get(code)
        if ov is not None and cands:
            idx = next((i for i, c in enumerate(cands) if c["row"]["row"] == ov), None)
            if idx:
                cands.insert(0, cands.pop(idx))

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
            mon_cov = best.get("mon_covered") or []
            mon_chip = ("<span class='fit' style='background:#3E7CB1;margin-left:.3rem'>"
                        "\u2713 builds on your monitoring: " + ", ".join(mon_cov)
                        + "</span>") if mon_cov else ""
            st.markdown(
                f"<div class='proto'><div style='color:{TEAL};font-weight:700;font-size:.85rem'>"
                f"{label} {fit}{mon_chip}</div>"
                f"<div style='font-weight:700;color:#1A2B2F;margin:.15rem 0 .35rem'>{r['title']}</div>"
                f"<span class='pill' style='background:{r['tier_color']}'>T{r['tier']} \u00b7 {r['tier_label']}</span>"
                f"<span class='tag'>Skill: {r['skill']}</span>"
                f"<span class='tag'>Cost {dots(r['cost_ord'])}</span>"
                f"<span class='tag'>Effort {dots(r['effort_ord'])}</span></div>",
                unsafe_allow_html=True)
            if len(cands) > 1:
                with st.expander(f"{len(cands) - 1} more option(s) for {label}"):
                    st.caption("Click any option to view its full details and, if it suits your farm "
                               "better, choose it as your primary \u2014 that updates the PDF report and "
                               "the Full protocol details section below.")
                    for alt in cands[1:]:
                        ar = alt["row"]; star = " \u2605" if alt["badge"] else ""
                        with st.popover(f"T{ar['tier']} \u00b7 {ar['title'][:70]}",
                                        use_container_width=True):
                            _detail(ar, [label], keyns=f"alt{code}", stacked=True, show_title=True)
                            st.divider()
                            st.markdown("**Would you like to choose this as your primary protocol for "
                                        f"{label} instead?** Selecting this will update the PDF report "
                                        "and the Full protocol details section below.")
                            if st.button("Yes, make this my primary", type="primary",
                                         key=f"mkp_{code}_{ar['row']}"):
                                st.session_state["primary_override"][code] = ar["row"]
                                st.rerun()

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

        # second box: what they already monitor, grouped by category
        st.subheader("Your ongoing monitoring")
        mon_sel = sub.get("mon_sel") or []
        if mon_sel:
            grouped = {}
            for sm in mon_sel:
                grouped.setdefault(sm.get("cat", "Other"), []).append(sm["sub"])
            h2 = ["<div class='learnbox'>"]
            for catname, subs in grouped.items():
                h2.append(f"<div class='k'>{catname}</div><div class='v'>"
                          + ", ".join(subs) + "</div>")
            h2.append("</div>")
            st.markdown("".join(h2), unsafe_allow_html=True)
            st.caption("Protocols that build on these are marked "
                       "\u201c\u2713 builds on your monitoring\u201d and moved up within their tier.")
        else:
            st.markdown("<div class='learnbox'><div class='v'>No ongoing monitoring "
                        "added \u2014 recommendations are ordered by tier and farm fit.</div></div>",
                        unsafe_allow_html=True)

    st.divider()
    try:
        pdf = report.build_pdf(sub, matches, code2label, name_map=CODE_NAMES)
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
            _detail(r, covers, keyns="main")


def _detail(r, covers, keyns="", stacked=False, show_title=False):
    if show_title and r.get("title"):
        st.markdown(f"##### {r['title']}")
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
        val = ml.strip_indicator_codes(val, CODE_NAMES)
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
    dls = [("Normalized protocol", r["url_template"]),
           ("Field data sheet", r["url_datasheet"]),
           ("Stats workbook", r["url_workbook"])]
    cols = [None, None, None] if stacked else st.columns(3)
    for i, (lbl, val) in enumerate(dls):
        u = report.first_url(val)
        ctx = st.container() if stacked else cols[i]
        with ctx:
            if u:
                st.link_button(f"Download {lbl}", u, use_container_width=True)
            else:
                st.button(f"{lbl} (V2)", disabled=True, use_container_width=True,
                          key=f"dl_{keyns}_{r['row']}_{lbl}", help="Coming in V2 - not yet available.")


# --------------------------------------------------------------------------- #
# Router
# --------------------------------------------------------------------------- #
v = st.session_state.view
if v == "survey":
    render_survey()
elif v == "monitoring":
    render_monitoring()
elif v == "learn_indicators":
    render_learn_indicators()
elif v == "learn_classes":
    render_learn_classes()
else:
    render_results()
