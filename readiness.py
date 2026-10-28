"""
AquaPIP - "How close are you?" page (Streamlit).

Sits between "What are you already monitoring?" and the MEL Framework overview. For every
indicator the farmer picked it shows which of their current monitoring counts toward it, a
closeness score they can sharpen by answering a few questions, what's left to reach the
Framework minimum, and indicators their monitoring could also count toward. A two-page
readout downloads as a PDF.

Scoring lives in readiness_logic.py; recipes (parts, questions, wording) live in the
'Readiness' sheet of MEL_Indicator_Reference.xlsx; pictures in assets/indicators/.

render_page(sub, base_codes, recipes, ref, rows, stretch) -> "back" | "next" | None
  Updates `sub` in place: codes, code2label, goals, readout, readout_summary,
  readout_monitoring (read by report.build_pdf / build_readout_pdf).
"""

import html

import streamlit as st

import mel_logic as ml
import readiness_logic as RL
import indicator_art
import report

TEAL = "#0B4F5C"
AREA_CHIP = {"Habitat & Biodiversity": "#277A4D", "Water Quality": "#2A7390", "Climate Change": "#0B4F5C"}

K_ANS, K_SEL, K_REM, K_ADD, K_FOCUS, K_MOVE = ("rd_answers", "rd_selected", "rd_removed",
                                              "rd_added", "rd_focus", "rd_last_move")

CSS = """
<style>
.rd-chip{font-family:ui-monospace,Menlo,Consolas,monospace;font-size:.74rem;font-weight:700;color:#fff;
  padding:2px 8px;border-radius:6px;white-space:nowrap;margin-right:.45rem}
.rd-name{font-weight:700;color:#1A2B2F;font-size:1.02rem}
.rd-score{font-family:Georgia,'Times New Roman',serif;font-weight:700;color:#0B4F5C;font-size:1.12rem;float:right}
.rd-tag{color:#4A5B60;font-size:.92rem;margin-top:.1rem}
.rd-muted{color:#5B6B70;font-size:.88rem}
.rd-bar{position:relative;background:#E3EAEC;border-radius:999px;overflow:hidden;margin:.5rem 0 .4rem;height:10px}
.rd-bar.big{height:14px}
.rd-fill{position:absolute;left:0;top:0;bottom:0;background:#0B4F5C}
.rd-range{position:absolute;top:0;bottom:0;background:repeating-linear-gradient(135deg,rgba(11,79,92,.55) 0 2px,rgba(11,79,92,.14) 2px 6px)}
.rd-dim{opacity:.45}
.rd-linked{background:#E4F1F4;border-radius:8px;padding:.15rem .35rem}
.rd-band{font-weight:700;color:#1A2B2F;margin:.9rem 0 .15rem}
.rd-band span{font-weight:400;color:#5B6B70;font-size:.88rem;margin-left:.5rem}
.rd-pic svg{display:block;width:100%;height:auto;border:1px solid #DDE5E7;border-radius:10px;background:#fff}
.rd-big{font-family:Georgia,'Times New Roman',serif;font-size:2.4rem;line-height:1;font-weight:700;color:#0B4F5C}
.rd-does{list-style:none;margin:.2rem 0 .4rem;padding:0}
.rd-does li{margin:.35rem 0;display:flex;gap:.55rem;align-items:flex-start}
.rd-ico{flex:none;width:1.25rem;height:1.25rem;border-radius:50%;display:inline-flex;align-items:center;
  justify-content:center;font-size:.8rem;font-weight:700;margin-top:.1rem}
.rd-ico.core{background:#0B4F5C;color:#fff}
.rd-ico.ctx{border:2px solid #5B6B70;color:#5B6B70}
.rd-ico.miss{border:2px dashed #5B6B70;color:#5B6B70}
.rd-note{padding:.6rem .8rem;border-radius:10px;background:#E4F1F4;color:#16343B;margin:.3rem 0}
.rd-done{padding:.6rem .8rem;border-radius:10px;background:#E6F2EB;color:#1D5536;font-weight:600;margin:.3rem 0}
.rd-first{padding:.7rem .9rem;border-radius:10px;border:1px solid #DDE5E7;background:#fff;margin:.3rem 0}
.rd-sq{display:inline-block;width:11px;height:11px;border-radius:2px;border:1px solid rgba(0,0,0,.15);
  margin:0 .35rem 0 .7rem;vertical-align:middle}
.rd-fw{display:grid;grid-template-columns:auto 1fr;gap:.25rem .8rem;font-size:.9rem}
.rd-fw b{color:#5B6B70;font-weight:400}
.rd-sug{border:1px dashed #8FA7AE;border-radius:12px;padding:.8rem .9rem;background:#fff;margin-bottom:.4rem}
</style>
"""


# --------------------------------------------------------------------------- #
# Small helpers
# --------------------------------------------------------------------------- #
def _e(t):
    return html.escape(t or "", quote=False)


def _slug(code):
    return code.replace("&", "").replace(" ", "_").replace(".", "_")


def _ord(n):
    return "%d%s" % (n, "tsnrhtdd"[(n // 10 % 10 != 1) * (n % 10 < 4) * n % 10::4])


def _cols(spec, **kw):
    try:
        return st.columns(spec, vertical_alignment="center", **kw)
    except TypeError:
        return st.columns(spec, **kw)


def _chip(code):
    color = AREA_CHIP.get(ml.goal_of_code(code), TEAL)
    return "<span class='rd-chip' style='background:%s'>%s</span>" % (color, _e(code))


def _bar(sc, big=False):
    lo, hi = max(0.0, sc["lo"]), max(0.0, sc["hi"])
    return ("<div class='rd-bar%s'><span class='rd-fill' style='width:%.1f%%'></span>"
            "<span class='rd-range' style='left:%.1f%%;width:%.1f%%'></span></div>"
            % (" big" if big else "", lo, lo, max(0.0, hi - lo)))


def reset():
    """Forget this page's answers (called from the app's Start over)."""
    for k in list(st.session_state):
        if k.startswith("rd_"):
            st.session_state.pop(k, None)


def _proto_counts(rows, codes, aqua):
    out = {}
    if not codes:
        return out
    try:
        m = ml.match_protocols(rows, list(codes), {}, farmer_aqua=aqua)
    except Exception:
        return out
    for code, cands in m.items():
        cnt = {}
        for c in cands:
            r = c["row"]
            if r.get("is_protocol"):
                cnt[r["tier"]] = cnt.get(r["tier"], 0) + 1
        out[code] = cnt
    return out


# --------------------------------------------------------------------------- #
# Page
# --------------------------------------------------------------------------- #
def render_page(sub, base_codes, recipes, ref, rows, stretch=None):
    stretch = stretch or {}
    st.markdown(CSS, unsafe_allow_html=True)
    aqua = sub.get("aqua_type") or ""
    sectors = ml.sectors_for_aqua(aqua)
    mon_labels = [s["sub"] for s in (sub.get("mon_sel") or []) if s.get("sub")]
    have = set(mon_labels)

    answers = st.session_state.setdefault(K_ANS, {})
    removed = list(st.session_state.setdefault(K_REM, []))
    added = list(st.session_state.setdefault(K_ADD, []))
    active = RL.active_codes(list(base_codes or []), set(removed), added)

    scored = [(c, RL.recipe_for(recipes, sectors, c)) for c in active]
    unscored = [c for c, r in scored if not r]
    scored = [(c, r, RL.score(r, have, answers.get(c))) for c, r in scored if r]
    ranked = [e for _i, e in RL.rank(scored)]
    order = [c for c, _r, _s in ranked]
    sel = st.session_state.get(K_SEL)
    if sel not in order:
        sel = order[0] if order else None

    sugs = RL.suggestions(recipes, sectors, active, set(removed), have, answers)
    fw = {}
    for ind in ref.get("indicators", []):
        if ind["sector"] in sectors:
            fw.setdefault(ind["code"], ind)

    # ---- keep the submission in step (results, learn pages and PDFs read it) ----
    c2l = dict(sub.get("code2label") or {})
    for c in added:
        r = RL.recipe_for(recipes, sectors, c)
        if r:
            c2l.setdefault(c, r["name"])
    n_build = sum(1 for _c, _r, s in ranked if s["M"] > 0)
    n_act = len(ranked)
    summary = ("%d of your %d indicators build on what you already do%s"
               % (n_build, n_act, (", and %d %s new for you." % (n_act - n_build, "is" if n_act - n_build == 1 else "are"))
                  if n_act - n_build else ".")) if n_act else "You\u2019ve removed every indicator."
    sub["codes"] = active
    sub["code2label"] = c2l
    sub["goals"] = ml.goals_from_codes(active)
    sub["readout"] = RL.readout(recipes, sectors, active, have, answers)
    sub["readout_summary"] = summary
    sub["readout_monitoring"] = mon_labels

    # ---- hero ----
    st.markdown("<div class='hero'><h1>How close are you?</h1><p>We compared what you already monitor with "
                "what the MEL Framework asks for on each indicator you picked. The closer you are, the higher "
                "it sits. Answer the questions on each one to sharpen its score, remove any you\u2019d rather "
                "not take on yet, and download your readout when you\u2019re done.</p></div>",
                unsafe_allow_html=True)
    if not recipes:
        st.info("The readiness guide isn\u2019t set up yet (no \u2018Readiness\u2019 sheet in "
                "MEL_Indicator_Reference.xlsx). You can carry on to the next page.")
    st.markdown("**%s**" % _e(summary))

    # ---- what you already monitor ----
    links = {m: {"core": [], "ctx": [], "sug": []} for m in mon_labels}
    for c, r, _s in ranked:
        for m in RL.core_methods(r, have):
            links[m]["core"].append(r["name"])
        for m in RL.context_methods(r, have):
            links[m]["ctx"].append(r["name"])
    for c, r, _s in sugs:
        for m in RL.core_methods(r, have):
            links[m]["sug"].append(r["name"])

    def badge(m):
        L = links.get(m, {})
        if L.get("core"):
            return "counts for %d" % len(L["core"])
        if L.get("ctx"):
            return "context for %d" % len(L["ctx"])
        if L.get("sug"):
            return "could count for %d" % len(L["sug"])
        return "not linked"

    focus = None
    st.markdown("<div class='step'>What you already monitor</div>", unsafe_allow_html=True)
    if mon_labels:
        lab = "Tap one to see which of your indicators it counts toward."
        if hasattr(st, "pills"):
            focus = st.pills(lab, mon_labels, selection_mode="single", key=K_FOCUS,
                             format_func=lambda m: "%s \u00b7 %s" % (m, badge(m)))
        else:
            pick = st.selectbox(lab, ["\u2014"] + mon_labels, key=K_FOCUS,
                                format_func=lambda m: m if m == "\u2014" else "%s \u00b7 %s" % (m, badge(m)))
            focus = None if pick == "\u2014" else pick
        if focus:
            L = links.get(focus, {})
            join = lambda xs: xs[0] if len(xs) == 1 else ", ".join(xs[:-1]) + " and " + xs[-1]
            note = ""
            if L.get("core"):
                note = "%s counts toward %s." % (focus, join(L["core"]))
            if L.get("ctx"):
                note += (" It also adds" if note else "%s adds" % focus) + " useful context for %s, but doesn\u2019t measure %s." % (
                    join(L["ctx"]), "them" if len(L["ctx"]) > 1 else "it")
            if L.get("sug"):
                note += (" It could also count toward" if note else "%s could count toward" % focus) + " %s, suggested below." % join(L["sug"])
            if not note:
                note = "%s isn\u2019t linked to any indicator you picked." % focus
            st.markdown("<div class='rd-note'>%s</div>" % _e(note), unsafe_allow_html=True)
    else:
        st.caption("You didn\u2019t tick any monitoring on the previous page, so every indicator starts at 0%. "
                   "Go back to add what you already do.")

    # ---- dial-in progress + readout download ----
    done, total = RL.dial(recipes, sectors, active, have, answers)
    complete = done >= total
    pc = st.columns([3, 1.4])
    with pc[0]:
        st.progress((done / total) if total else 1.0,
                    text=("All %d questions answered. Your readout is ready." % total if total and complete else
                          "%d of %d questions answered. Open questions print as ranges." % (done, total) if total else
                          "No follow-up questions yet."))
    with pc[1]:
        try:
            pdf = report.build_readout_pdf(sub)
            fname = (sub.get("farm_name") or "AquaPIP").replace(" ", "_")[:40]
            st.download_button("Download my readout (PDF)", data=pdf, file_name="%s_readout.pdf" % fname,
                               mime="application/pdf", type="primary" if complete else "secondary",
                               key="rd_download", **stretch)
        except Exception as e:   # never block the page on the PDF
            st.caption("Readout PDF unavailable (%s)." % type(e).__name__)

    # ---- list + detail ----
    move = st.session_state.get(K_MOVE)
    linked = set()
    if focus:
        L = links.get(focus, {})
        linked = set(L.get("core", [])) | set(L.get("ctx", []))
    left, right = st.columns([1.2, 1], gap="large")
    with left:
        st.markdown("<div class='step'>Your indicators, closest first</div>", unsafe_allow_html=True)
        if unscored:
            st.caption("Not scored yet (no readiness guide for %s)." % ", ".join(unscored))
        for key, name, rng, _low in RL.BANDS:
            items = [(i, c, r, s) for i, (c, r, s) in enumerate(ranked) if RL.band(s) == key]
            if not items:
                continue
            st.markdown("<div class='rd-band'>%s<span>%s</span></div>" % (name, _e(rng)), unsafe_allow_html=True)
            for i, c, r, s in items:
                core = RL.core_methods(r, have)
                ctx = RL.context_methods(r, have)
                uses = ("Builds on: " + ", ".join(core)) if core else (("Related: " + ", ".join(ctx)) if ctx else
                                                                         "Nothing you do feeds this yet")
                bits = []
                if 0 < s["M"] < 1:
                    bits.append("%d of %d parts" % (s["covered"], s["n_parts"]))
                if s["M"] > 0 and r["questions"]:
                    bits.append("%d question%s to go" % (s["open_n"], "" if s["open_n"] == 1 else "s") if s["open_n"]
                                else ("all answered" if bits else "All answered"))
                if move and move[0] == c:
                    bits.insert(0, "<b style='color:#0B4F5C'>%s</b>" % (
                        "Just added" if move[1] is None else ("Up from " if move[2] < move[1] else "Down from ") + _ord(move[1] + 1)))
                cls = "" if not focus else (" rd-linked" if r["name"] in linked else " rd-dim")
                with st.container(border=True):
                    a, b = _cols([5, 1.25])
                    with a:
                        st.markdown(
                            "<div class='%s'><span class='rd-score'>%s</span>%s<span class='rd-name'>%s</span>"
                            "<div class='rd-tag'>%s</div>%s<div class='rd-muted'>%s%s</div></div>"
                            % (cls.strip(), RL.score_text(s), _chip(c), _e(r["name"]), _e(r["tagline"]), _bar(s),
                               _e(uses), (" \u00b7 " + " \u00b7 ".join(bits)) if bits else ""),
                            unsafe_allow_html=True)
                    with b:
                        if st.button("Showing" if c == sel else "Open", key="rd_open_" + _slug(c),
                                     type="primary" if c == sel else "secondary", **stretch):
                            st.session_state[K_SEL] = c
                            st.session_state.pop(K_MOVE, None)
                            st.rerun()

    with right:
        if sel:
            _detail(sel, recipes, sectors, have, answers, fw, rows, aqua, ranked, added, removed, stretch)
        else:
            st.info("You\u2019ve removed every indicator. Add one back below, or go back and pick new ones.")

    # ---- suggestions ----
    if sugs:
        st.markdown("<div class='step'>Your monitoring could also count toward</div>", unsafe_allow_html=True)
        st.caption("These indicators aren\u2019t on your list yet, but something you already monitor feeds them. "
                   "Add any you\u2019d like to work toward; you can remove them again later.")
        cols = st.columns(3)
        for n, (c, r, s) in enumerate(sugs):
            core = RL.core_methods(r, have)
            with cols[n % 3]:
                st.markdown(
                    "<div class='rd-sug'>%s<span class='rd-name'>%s</span><div class='rd-tag'>%s</div>%s"
                    "<div><b style='font-family:Georgia,serif;color:#0B4F5C'>%s</b></div>"
                    "<div class='rd-muted'>Builds on: %s%s</div></div>"
                    % (_chip(c), _e(r["name"]), _e(r["tagline"]), _bar(s), RL.score_text(s), _e(", ".join(core)),
                       (" (%d of %d parts)" % (s["covered"], s["n_parts"])) if s["M"] < 1 else ""),
                    unsafe_allow_html=True)
                if st.button("Add to my indicators", key="rd_add_" + _slug(c), **stretch):
                    st.session_state[K_ADD] = added + [c]
                    st.session_state[K_SEL] = c
                    st.session_state[K_MOVE] = (c, None, None)
                    st.rerun()

    # ---- removed for now ----
    gone = [c for c in removed if c in (base_codes or [])]
    if gone:
        st.markdown("<div class='step'>Removed for now</div>", unsafe_allow_html=True)
        cols = st.columns(3)
        for n, c in enumerate(gone):
            r = RL.recipe_for(recipes, sectors, c)
            with cols[n % 3]:
                st.markdown("%s<span class='rd-name'>%s</span>" % (_chip(c), _e(r["name"] if r else c)),
                            unsafe_allow_html=True)
                if st.button("Add back", key="rd_back_" + _slug(c), **stretch):
                    st.session_state[K_REM] = [x for x in removed if x != c]
                    st.session_state[K_SEL] = c
                    st.rerun()

    with st.expander("How is this score worked out?"):
        st.markdown("Doing the measurement itself is worth **60 points**. The other **40** come from how your "
                    "sampling compares with the MEL Framework\u2019s minimum for that indicator: where you sample, "
                    "how often, and how you record it.")
        st.markdown("If you do only part of what an indicator needs, such as one of its three surveys, the score is "
                    "scaled to that part. Until you answer an indicator\u2019s questions, its bar shows the range it "
                    "could fall in.")
        st.markdown("Measurements that only add context, like water clarity alongside chlorophyll, are shown but "
                    "don\u2019t add points.")

    st.divider()
    nav = st.columns([1, 2])
    with nav[0]:
        if st.button("<-  Back", key="rd_nav_back", **stretch):
            return "back"
    with nav[1]:
        if st.button("Next  ->  MEL Framework Overview", type="primary", key="rd_nav_next", **stretch):
            return "next"
    return None


def _detail(code, recipes, sectors, have, answers, fw, rows, aqua, ranked, added, removed, stretch):
    r = RL.recipe_for(recipes, sectors, code)
    a = answers.get(code, {})
    s = RL.score(r, have, a)
    with st.container(border=True):
        st.markdown("%s<div style='font-family:Georgia,serif;font-size:1.4rem;font-weight:700;color:#0B4F5C;"
                    "margin:.3rem 0 .4rem'>%s</div>" % (_chip(code), _e(r["name"])), unsafe_allow_html=True)
        svg = indicator_art.svg_text(r["picture"])
        if svg:
            st.markdown("<div class='rd-pic'>%s</div>" % svg, unsafe_allow_html=True)
        if r["caption"]:
            st.caption(r["caption"])
        if r["summary"]:
            st.markdown(_e(r["summary"]))

        band = RL.BAND_NAME[RL.band(s)]
        note = ("Nothing you already do measures this yet." if s["M"] <= 0 else
                ("Answer %d more question%s to pin this down." % (s["open_n"], "" if s["open_n"] == 1 else "s"))
                if s["open_n"] else "Based on your answers.")
        if 0 < s["M"] < 1:
            note += " You do %d of the %d parts this indicator needs." % (s["covered"], s["n_parts"])
        st.markdown("<span class='rd-big'>%s</span> &nbsp; <b>%s</b>%s<div class='rd-muted'>%s</div>"
                    % (RL.score_text(s), band, _bar(s, big=True), _e(note)), unsafe_allow_html=True)

        # what you already do
        items = []
        core_set = set()
        for p in r["parts"]:
            hit = [m for m in p["methods"] if m in have]
            core_set.update(hit)
            if hit:
                items.append(("core", "\u2713", ", ".join(hit),
                              ("Covers: " + p["label"][0].lower() + p["label"][1:]) if len(r["parts"]) > 1 else
                              "This is the measurement itself"))
            else:
                items.append(("miss", "+", p["label"], "Not yet" + ("; " + p["hint"] if p["hint"] else "")))
        for m in r["context"]:
            if m in have and m not in core_set:
                items.append(("ctx", "\u2022", m, "Useful context; doesn\u2019t add points"))
        st.markdown("**What you already do**")
        st.markdown("<ul class='rd-does'>%s</ul>" % "".join(
            "<li><span class='rd-ico %s'>%s</span><span><b>%s</b><br><span class='rd-muted'>%s</span></span></li>"
            % (k, ic, _e(t), _e(sub_)) for k, ic, t, sub_ in items), unsafe_allow_html=True)

        # dial it in
        if s["M"] > 0 and r["questions"]:
            st.markdown("**Dial it in**")
            st.caption("Your answers move this score and the order of your list. Tap an answer again to clear it.")
            pts = RL.question_points(r)
            changed = False
            new_a = dict(a)
            for q in r["questions"]:
                opts = [lab for lab, _c in q["options"]]
                cur = a.get(q["key"])
                cur = cur if isinstance(cur, int) and 0 <= cur < len(opts) else None
                wkey = "rd_q_%s_%s" % (_slug(code), q["key"])
                label = "%s (%d pts)" % (q["text"], pts.get(q["key"], 0))
                if hasattr(st, "pills"):
                    val = st.pills(label, opts, selection_mode="single",
                                   default=opts[cur] if cur is not None else None, key=wkey)
                else:
                    val = st.radio(label, opts, index=cur, horizontal=True, key=wkey)
                new = opts.index(val) if val in opts else None
                if new != cur:
                    changed = True
                    if new is None:
                        new_a.pop(q["key"], None)
                    else:
                        new_a[q["key"]] = new
            if changed:
                before = [c for c, _r, _s in ranked]
                answers = dict(answers); answers[code] = new_a
                st.session_state[K_ANS] = answers
                after = []
                for _i, (c, _r, _s) in RL.rank([(c, rr, RL.score(rr, have, answers.get(c))) for c, rr, _s in ranked]):
                    after.append(c)
                f, t = before.index(code), after.index(code)
                st.session_state[K_MOVE] = (code, f, t) if f != t else None
                st.rerun()

        tds = RL.todos(r, s, have, a)
        if tds:
            st.markdown("**To reach 100%**")
            st.markdown("\n".join("- %s %s" % (_e(t), ("**%s**" % g) if g else "") for t, g in tds))
        elif s["M"] > 0 and s["open_n"]:
            st.caption("Answer the questions above to see what, if anything, is left.")
        elif s["M"] >= 1 and not s["open_n"]:
            st.markdown("<div class='rd-done'>You meet the Framework\u2019s minimum for this indicator.</div>",
                        unsafe_allow_html=True)
        if r["note"] and s["M"] >= 1:
            st.markdown("<div class='rd-note'>%s</div>" % _e(r["note"]), unsafe_allow_html=True)
        if s["M"] <= 0 and r["first"]:
            st.markdown("<div class='rd-first'><b style='color:#0B4F5C'>A first step</b><br>%s</div>" % _e(r["first"]),
                        unsafe_allow_html=True)

        # what the Framework asks for (verbatim from MEL_Indicators)
        f = fw.get(code)
        if f:
            st.markdown("**What the Framework asks for**")
            diff = (f.get("complexity") or "").split(" - ")[0].strip()
            rows_ = [("Method", f.get("method")), ("How often", f.get("frequency")),
                     ("Where", f.get("location")), ("Difficulty", diff)]
            st.markdown("<div class='rd-fw'>%s</div>" % "".join(
                "<b>%s</b><span>%s</span>" % (k, _e(v)) for k, v in rows_ if v), unsafe_allow_html=True)

        cnt = _proto_counts(rows, [code], aqua).get(code, {})
        n = sum(cnt.values())
        if n:
            parts = "".join("<span class='rd-sq' style='background:%s'></span>%d %s (T%d)"
                            % (meta["color"], cnt[t], meta["label"], t)
                            for t, meta in ml.VISIBLE_TIER_META.items() if cnt.get(t))
            st.markdown("<div style='font-size:.9rem'><b>AquaPIP has %d protocol%s for this:</b>%s</div>"
                        % (n, "" if n == 1 else "s", parts), unsafe_allow_html=True)
        else:
            st.caption("No AquaPIP protocols for this yet.")

        if s["M"] <= 0:
            st.caption("Not ready to take this on? Removing it keeps your list focused on what you\u2019re closer to. "
                       "You can add it back.")
        if st.button("Remove this indicator", key="rd_remove_" + _slug(code)):
            if code in added:
                st.session_state[K_ADD] = [c for c in added if c != code]
            else:
                st.session_state[K_REM] = removed + [code]
            st.session_state.pop(K_SEL, None)
            st.session_state.pop(K_MOVE, None)
            st.rerun()
