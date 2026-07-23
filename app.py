"""
AquaPIP — MEL Monitoring Survey  (Streamlit app)

Run locally:   streamlit run app.py
Deploy:        push this folder to GitHub, then deploy on share.streamlit.io
               (see README.md for click-by-click steps).

Flow
----
  view == "survey"   the two-part survey (goals -> farm -> indicators -> consent)
  view == "results"  the recommendations page + PDF download

The app reads data/mel_database.xlsx live and by column header name, so updating
that file (same filename) and pushing will refresh both the survey options and
the recommendations automatically.
"""

import os
import streamlit as st

import mel_logic as ml
import report
import storage

# --------------------------------------------------------------------------- #
# Page config + styling
# --------------------------------------------------------------------------- #
st.set_page_config(
    page_title="AquaPIP — MEL Monitoring Survey",
    page_icon="🌊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

TEAL = "#0B4F5C"
DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "data", "mel_database.xlsx")

st.markdown(f"""
<style>
  .block-container {{ padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1150px; }}
  #MainMenu, footer {{ visibility: hidden; }}
  h1, h2, h3 {{ font-family: Georgia, 'Times New Roman', serif; color: {TEAL}; }}
  .app-hero h1 {{ margin-bottom: .1rem; font-size: 1.9rem; }}
  .app-hero p  {{ color: #5B6B70; margin-top: 0; font-size: .95rem; }}
  .step-label {{ font-family: Georgia, serif; color: {TEAL}; font-weight: 700;
                 font-size: 1.15rem; margin: 1.2rem 0 .2rem; }}
  .step-help  {{ color: #5B6B70; font-size: .85rem; margin: 0 0 .5rem; }}
  .proto-card {{ background:#fff; border:1px solid #E1E7E9; border-left:5px solid {TEAL};
                 border-radius:10px; padding:.75rem .95rem; margin:.35rem 0 .55rem;
                 box-shadow:0 1px 2px rgba(11,79,92,.05); }}
  .proto-ind   {{ color:{TEAL}; font-weight:700; font-size:.82rem; letter-spacing:.02em; }}
  .proto-title {{ color:#1A2B2F; font-weight:700; font-size:.98rem; margin:.15rem 0 .4rem; }}
  .proto-meta  {{ display:flex; flex-wrap:wrap; gap:.4rem; align-items:center; }}
  .tier-pill   {{ color:#1A2B2F; font-weight:700; font-size:.74rem;
                  padding:.12rem .5rem; border-radius:999px; border:1px solid rgba(0,0,0,.06); }}
  .chip        {{ background:#EEF3F4; color:#33474C; font-size:.74rem;
                  padding:.12rem .5rem; border-radius:999px; }}
  .chip .dot   {{ letter-spacing:1px; }}
  .badge-fit   {{ background:{TEAL}; color:#fff; font-size:.7rem; font-weight:700;
                  padding:.1rem .45rem; border-radius:999px; }}
  .resp-box    {{ background:#F3F6F7; border:1px solid #E1E7E9; border-radius:10px;
                  padding:.8rem .95rem; }}
  .resp-box .k {{ color:#5B6B70; font-size:.68rem; font-weight:700; letter-spacing:.03em;
                  text-transform:uppercase; margin-top:.5rem; }}
  .resp-box .v {{ color:#1A2B2F; font-size:.86rem; }}
  .ind-line    {{ font-size:.86rem; color:#1A2B2F; margin:.12rem 0; }}
  .sq          {{ display:inline-block; width:.7rem; height:.7rem; border-radius:2px;
                  margin-right:.4rem; vertical-align:middle; border:1px solid rgba(0,0,0,.08); }}
  .legend      {{ color:#5B6B70; font-size:.8rem; }}
</style>
""", unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
# Data (cached; busts when the workbook's mtime changes)
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner=False)
def get_data(_mtime):
    rows, vocab, headers = ml.load_database(DB_PATH)
    ind_groups = ml.build_indicator_options(vocab.get("MEL Indicator", []), rows)
    prof_opts = ml.profile_options(vocab, rows)
    return rows, vocab, ind_groups, prof_opts


try:
    mtime = os.path.getmtime(DB_PATH)
except OSError:
    st.error(f"Database not found at `{DB_PATH}`. "
             f"Make sure data/mel_database.xlsx is present.")
    st.stop()

rows, vocab, ind_groups, prof_opts = get_data(mtime)
code2label, label2code, code2goal = ml.flatten_options(ind_groups)
available_codes = set(code2label)

GROUP_KEYS = {
    "Habitat & Biodiversity": "ind_hb",
    "Water Quality": "ind_wq",
    "Climate Change": "ind_cc",
}

if "view" not in st.session_state:
    st.session_state.view = "survey"


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _dots_html(ordinal):
    return f"<span class='dot'>{ml.dots(ordinal)}</span>"


def selected_codes_ordered():
    """Union of the three indicator multiselects, mapped labels -> codes,
    in group order (H&B, WQ, CC), preserving within-group order."""
    out = []
    for goal in ["Habitat & Biodiversity", "Water Quality", "Climate Change"]:
        for it in ind_groups[goal]:
            if it["label"] in st.session_state.get(GROUP_KEYS[goal], []):
                if it["code"] not in out:
                    out.append(it["code"])
    return out


# --------------------------------------------------------------------------- #
# SURVEY VIEW
# --------------------------------------------------------------------------- #
def render_survey():
    st.markdown(
        "<div class='app-hero'><h1>Regenerative Aquaculture — MEL Monitoring Survey</h1>"
        "<p>Tell us about your farm and what you want to measure. "
        "We'll match you to monitoring protocols from the TNC MEL evidence base — "
        "ranked so the most farmer-friendly options come first.</p></div>",
        unsafe_allow_html=True)
    st.divider()

    # ---- Step 1: what do you want to measure? (farmer-facing goals) ------- #
    st.markdown("<div class='step-label'>1 · What are you interested in measuring?</div>",
                unsafe_allow_html=True)
    st.markdown("<div class='step-help'>Pick one or more. Each choice pre-selects a "
                "set of MEL indicators that you can fine-tune in step 3.</div>",
                unsafe_allow_html=True)
    farmer_goals = st.multiselect(
        "Goals", options=list(ml.FARMER_GOALS.keys()),
        key="farmer_goals", label_visibility="collapsed",
        placeholder="Choose what you'd like to measure…")
    for g in farmer_goals:
        st.caption(f"**{g}** — {ml.FARMER_GOALS[g]['blurb']}")

    # ---- Auto-sync goals -> indicator selections (BEFORE those widgets) --- #
    if st.session_state.get("_prev_fg") != farmer_goals:
        derived = ml.codes_for_farmer_goals(farmer_goals, available_codes)
        for goal, key in GROUP_KEYS.items():
            st.session_state[key] = [
                code2label[c] for c in derived if code2goal.get(c) == goal]
        st.session_state["_prev_fg"] = list(farmer_goals)

    # ---- Step 2: about your farm ----------------------------------------- #
    st.markdown("<div class='step-label'>2 · About your farm</div>",
                unsafe_allow_html=True)
    st.markdown("<div class='step-help'>Optional — leave any field on “— any —”. "
                "These never hide protocols; they just move the best-fitting ones "
                "up and add a “fits your farm” marker.</div>",
                unsafe_allow_html=True)

    farm_name = st.text_input("Farm / site name (optional)",
                              key="farm_name_input",
                              placeholder="e.g. Tidewater Kelp Co., Casco Bay")

    profile = {}
    cols = st.columns(2)
    for i, (fkey, label, _vh, _rk) in enumerate(ml.PROFILE_FIELDS):
        opts = ["— any —"] + prof_opts.get(fkey, [])
        with cols[i % 2]:
            choice = st.selectbox(label, options=opts, index=0, key=f"prof_{fkey}")
        if choice and choice != "— any —":
            profile[fkey] = choice

    # ---- Step 3: refine indicators --------------------------------------- #
    st.markdown("<div class='step-label'>3 · Refine your MEL indicators</div>",
                unsafe_allow_html=True)
    st.markdown("<div class='step-help'>Pre-filled from step 1. Add or remove any "
                "indicator. Grouped by MEL goal.</div>", unsafe_allow_html=True)

    gcols = st.columns(3)
    group_display = {
        "Habitat & Biodiversity": "Habitat & Biodiversity",
        "Water Quality": "Water Quality",
        "Climate Change": "Climate Change",
    }
    for i, goal in enumerate(["Habitat & Biodiversity", "Water Quality", "Climate Change"]):
        labels = [it["label"] for it in ind_groups[goal]]
        with gcols[i]:
            st.markdown(f"**{group_display[goal]}**")
            st.multiselect(goal, options=labels, key=GROUP_KEYS[goal],
                           label_visibility="collapsed",
                           placeholder="none selected")

    codes = selected_codes_ordered()
    st.caption(f"**{len(codes)}** indicator(s) selected across "
               f"{len(ml.goals_from_codes(codes))} MEL goal(s).")

    # ---- Step 4: consent + submit ---------------------------------------- #
    st.markdown("<div class='step-label'>4 · Save & generate</div>",
                unsafe_allow_html=True)
    consent = st.checkbox(
        "You can store my responses. Your responses will only be used for "
        "internal tracking purposes, and not for any other uses unless "
        "explicitly authorized in writing by you.",
        value=False, key="consent")

    c1, c2 = st.columns([1, 3])
    with c1:
        go = st.button("Generate my recommendations  →", type="primary",
                       use_container_width=True)
    if go:
        if not codes:
            st.warning("Please select at least one indicator in step 3 "
                       "(or pick a goal in step 1 to pre-fill them).")
        else:
            submission = {
                "farm_name": (farm_name or "").strip(),
                "farmer_goals": list(farmer_goals),
                "codes": codes,
                "profile": profile,
                "goals": ml.goals_from_codes(codes),
                "consent": bool(consent),
            }
            st.session_state["submission"] = submission
            if consent:
                st.session_state["_save_status"] = storage.save_response(submission)
            else:
                st.session_state["_save_status"] = (None, "no_consent")
            st.session_state["view"] = "results"
            st.rerun()


# --------------------------------------------------------------------------- #
# RESULTS VIEW
# --------------------------------------------------------------------------- #
def render_results():
    sub = st.session_state.get("submission")
    if not sub:
        st.session_state["view"] = "survey"
        st.rerun()
        return

    top = st.columns([3, 1])
    with top[0]:
        farm = sub.get("farm_name") or "Your farm"
        st.markdown(f"<div class='app-hero'><h1>Your MEL Protocol Recommendations</h1>"
                    f"<p>{farm} · one best practitioner-runnable protocol per "
                    f"selected indicator.</p></div>", unsafe_allow_html=True)
    with top[1]:
        if st.button("←  Edit answers", use_container_width=True):
            st.session_state["view"] = "survey"
            st.rerun()

    ok, status = st.session_state.get("_save_status", (None, ""))
    if ok:
        st.caption("✓ Your responses were saved for internal tracking.")

    # legend
    parts = []
    for n, meta in ml.TIER_META.items():
        parts.append(f"<span class='sq' style='background:{meta['color']}'></span>"
                     f"T{n} {meta['label']}")
    st.markdown("<div class='legend'>" + " &nbsp; ".join(parts) +
                " &nbsp;|&nbsp; ●○○→●●● cost / effort (low→high) "
                "&nbsp;|&nbsp; ★ strong fit to your farm</div>",
                unsafe_allow_html=True)
    st.divider()

    matches = ml.match_protocols(rows, sub["codes"], sub["profile"])

    left, right = st.columns([2, 1])

    # ---- left: recommended protocols ------------------------------------- #
    with left:
        st.subheader("Recommended protocols")
        for code, cands in matches.items():
            label = code2label.get(code, code)
            if not cands:
                st.markdown(f"<div class='proto-card' style='border-left-color:#CCC'>"
                            f"<div class='proto-ind'>{label}</div>"
                            f"<div class='v' style='color:#5B6B70;font-size:.86rem'>"
                            f"No protocol in the current database for this indicator "
                            f"yet.</div></div>", unsafe_allow_html=True)
                continue
            best = cands[0]
            r = best["row"]
            badge = "<span class='badge-fit'>★ fits your farm</span>" if best["badge"] else ""
            st.markdown(
                f"<div class='proto-card'>"
                f"<div class='proto-ind'>{label} {badge}</div>"
                f"<div class='proto-title'>{r['title']}</div>"
                f"<div class='proto-meta'>"
                f"<span class='tier-pill' style='background:{r['tier_color']}'>"
                f"T{r['tier']} · {r['tier_label']}</span>"
                f"<span class='chip'>Skill: {r['skill']}</span>"
                f"<span class='chip'>Cost {_dots_html(r['cost_ord'])}</span>"
                f"<span class='chip'>Effort {_dots_html(r['effort_ord'])}</span>"
                f"</div></div>", unsafe_allow_html=True)
            if len(cands) > 1:
                with st.expander(f"{len(cands) - 1} more option(s) for {code}"):
                    for alt in cands[1:]:
                        ar = alt["row"]
                        star = " ★" if alt["badge"] else ""
                        st.markdown(
                            f"<div class='ind-line'>"
                            f"<span class='sq' style='background:{ar['tier_color']}'></span>"
                            f"<b>T{ar['tier']} {ar['tier_label']}</b>{star} — {ar['title']} "
                            f"<span style='color:#5B6B70'>· cost {ml.dots(ar['cost_ord'])} "
                            f"· effort {ml.dots(ar['effort_ord'])}</span></div>",
                            unsafe_allow_html=True)

    # ---- right: your responses ------------------------------------------- #
    with right:
        st.subheader("Your survey")
        html = ["<div class='resp-box'>"]
        if sub.get("farmer_goals"):
            html.append("<div class='k'>Interested in measuring</div>")
            html.append("<div class='v'>" + "<br>".join(sub["farmer_goals"]) + "</div>")
        html.append("<div class='k'>MEL goals</div>")
        html.append("<div class='v'>" + (" · ".join(sub["goals"]) or "—") + "</div>")
        prof = sub.get("profile", {})
        for fkey, label, _vh, _rk in ml.PROFILE_FIELDS:
            if prof.get(fkey):
                html.append(f"<div class='k'>{label}</div>")
                html.append(f"<div class='v'>{prof[fkey]}</div>")
        html.append("<div class='k'>Indicators selected</div>")
        for code in sub["codes"]:
            cands = matches.get(code, [])
            color = cands[0]["row"]["tier_color"] if cands else "#CCC"
            html.append(f"<div class='ind-line'><span class='sq' "
                        f"style='background:{color}'></span>{code2label.get(code, code)}</div>")
        html.append("</div>")
        st.markdown("".join(html), unsafe_allow_html=True)

    # ---- PDF download ---------------------------------------------------- #
    st.divider()
    try:
        pdf_bytes = report.build_pdf(sub, matches, code2label)
        fname = (sub.get("farm_name") or "MEL").replace(" ", "_")[:40]
        st.download_button("⬇  Download PDF report", data=pdf_bytes,
                           file_name=f"{fname}_MEL_recommendations.pdf",
                           mime="application/pdf", type="primary")
    except Exception as e:
        st.info(f"PDF export is temporarily unavailable ({type(e).__name__}). "
                f"Your recommendations are shown above.")

    # ---- full details ---------------------------------------------------- #
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
        with st.expander(f"{r['title']}  ·  T{r['tier']} {r['tier_label']}"):
            _render_detail(r, covers)


def _render_detail(r, covers):
    meta = " · ".join(x for x in [r["authors"], r["year"], r["publication"]] if x)
    st.markdown(f"**Covers:** {' · '.join(covers)}")
    if meta:
        st.caption(meta)
    st.markdown(
        f"<span class='tier-pill' style='background:{r['tier_color']}'>"
        f"T{r['tier']} · {r['tier_label']}</span> "
        f"<span class='chip'>Skill: {r['skill']}</span> "
        f"<span class='chip'>Cost {ml.dots(r['cost_ord'])} ({r['cost_label']})</span> "
        f"<span class='chip'>Effort {ml.dots(r['effort_ord'])} ({r['effort_label']})</span>",
        unsafe_allow_html=True)
    st.write("")

    def field(lbl, val):
        if val:
            st.markdown(f"**{lbl}**")
            st.write(val)

    field("What it measures", r["measures"])
    if r["method_summary"]:
        field("Method summary", r["method_summary"])
    else:
        st.markdown("**Method summary**")
        st.caption("Method summary pending — appears once the database includes "
                   "the Protocol Method Summary field.")
    field("Equipment", r["equipment"])
    field("Statistical approach", r["stats"])
    field("Executable by practitioner?", r["exec_practitioner"])
    field("Notes / caveats", r["notes"])

    url = report.first_url(r["url"])
    if url:
        st.markdown(f"**Source:** [{url}]({url})")

    # V2 downloads
    st.markdown("**Downloads**")
    d1, d2, d3 = st.columns(3)
    for col, lbl, val in [(d1, "Normalized protocol", r["url_template"]),
                          (d2, "Field data sheet", r["url_datasheet"]),
                          (d3, "Stats workbook", r["url_workbook"])]:
        u = report.first_url(val)
        with col:
            if u:
                st.link_button(f"⬇ {lbl}", u, use_container_width=True)
            else:
                st.button(f"{lbl} (V2)", disabled=True, use_container_width=True,
                          key=f"dl_{r['row']}_{lbl}",
                          help="Coming in V2 — not yet available.")


# --------------------------------------------------------------------------- #
# Router
# --------------------------------------------------------------------------- #
if st.session_state.view == "survey":
    render_survey()
else:
    render_results()
