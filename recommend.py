#!/usr/bin/env python3
"""Recommendation engine for the ODL PM dashboard.

Auto-derives actionable capacity / performance recommendations from the built
data, merges in manual recommendations (recommendations_manual.json), and
overlays the team's tracked status from statuses.json. IDs are STABLE across
rebuilds so a status set today survives tomorrow's data refresh.

Status lifecycle:  Not Started -> In Progress -> Achieved   (also: Dismissed)

Each recommendation = {
  id, category, severity (high|medium|low), scope_type, scope, title, detail,
  metric, suggested_action, source (auto|manual), gid?, months?
  + tracked: status, owner, target_month, notes, evidence_url, updated_at,
             updated_by, first_seen
}
"""
import re

STATUS_VALUES = ["Not Started", "In Progress", "Achieved", "Dismissed"]
SEV_RANK = {"high": 0, "medium": 1, "low": 2}
_MON = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def mlabel(m):
    y, mo = m.split("-")
    return f"{_MON[int(mo)]} {y}"


def span_str(months):
    months = sorted(set(months))
    if not months:
        return ""
    spans, start, prev = [], months[0], months[0]
    for m in months[1:]:
        if _next(prev) == m:
            prev = m
        else:
            spans.append((start, prev))
            start = prev = m
    spans.append((start, prev))
    return ", ".join(mlabel(s) if s == e else f"{mlabel(s)}–{mlabel(e)}" for s, e in spans)


def _next(m):
    y, mo = map(int, m.split("-"))
    return f"{y+1:04d}-01" if mo == 12 else f"{y:04d}-{mo+1:02d}"


def _months_ago(m, n):
    y, mo = map(int, m.split("-"))
    idx = y * 12 + (mo - 1) - n
    return f"{idx // 12:04d}-{idx % 12 + 1:02d}"


def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", (s or "").lower()).strip("-")[:48]


def _future(months, now, n=None):
    f = [m for m in months if m >= now]
    return f[:n] if n else f


def build_auto(data, now):
    cap = data["capacity"]
    months = data["meta"]["months"]
    full = data["meta"]["full_monthly_points"]
    recs = []

    # ---- R1/R2 team over / tight capacity (forward 9 months) ----
    for scope in ("Design", "Media", "PM"):
        pcts = cap["pct"].get(scope, {})
        win = _future(months, now, 9)
        over = [m for m in win if pcts.get(m) and pcts[m] > 1.0]
        tight = [m for m in win if pcts.get(m) and 0.9 <= pcts[m] <= 1.0]
        if over:
            worst = max(over, key=lambda m: pcts[m])
            recs.append(dict(
                id=f"cap-over-{scope.lower()}", category="capacity-team",
                severity="high", scope_type="team", scope=scope,
                title=f"{scope} team over capacity",
                detail=f"{scope} is scheduled above 100% in {span_str(over)} "
                       f"(peak {round(pcts[worst]*100)}% in {mlabel(worst)}).",
                metric=f"peak {round(pcts[worst]*100)}%",
                suggested_action="Rebalance assignments, pull in another team member, "
                                 "or defer/de-scope a project in these months."))
        elif tight:
            worst = max(tight, key=lambda m: pcts[m])
            recs.append(dict(
                id=f"cap-tight-{scope.lower()}", category="capacity-team",
                severity="medium", scope_type="team", scope=scope,
                title=f"{scope} team running tight",
                detail=f"{scope} is 90–100% allocated in {span_str(tight)} "
                       f"(peak {round(pcts[worst]*100)}% in {mlabel(worst)}). Little slack for new work.",
                metric=f"peak {round(pcts[worst]*100)}%",
                suggested_action="Avoid adding scope here; hold buffer for overruns."))

    # ---- R3/R4/R6 per person ----
    for nm, pp in cap["person"].items():
        if not pp["has_capacity"]:
            fut = [m for m in _future(months, now) if pp["scheduled"].get(m, 0) > 0.05]
            if fut:
                recs.append(dict(
                    id=f"data-nocap-{slug(nm)}", category="data-quality",
                    severity="medium", scope_type="person", scope=nm,
                    title=f"{nm} has scheduled work but no capacity row",
                    detail=f"{nm} is assigned {span_str(fut)} but has no entry on the "
                           f"Estimated Capacity tab, so their load isn't counted in any team total.",
                    metric=f"{round(sum(pp['scheduled'][m] for m in fut),1)} pts",
                    suggested_action="Add them to Estimated Capacity (Points), or convert "
                                     "the assignment to an unstaffed placeholder."))
            continue
        rem = pp["remaining"]
        over = [m for m in _future(months, now, 6) if rem.get(m) is not None and rem[m] < -0.5]
        if over:
            worst = min(over, key=lambda m: rem[m])
            recs.append(dict(
                id=f"person-over-{slug(nm)}", category="capacity-person",
                severity="high", scope_type="person", scope=nm,
                title=f"{nm} over capacity",
                detail=f"{nm} ({pp['team']}) is over-allocated in {span_str(over)} "
                       f"(over by {abs(rem[worst])} pts / {round(abs(rem[worst])*32)}h in {mlabel(worst)}).",
                metric=f"over by {abs(rem[worst])} pts",
                suggested_action="Redistribute a deliverable to a teammate or shift timing."))
        # exclude the current (partly-elapsed) month; require >=2 future months free
        near = [m for m in _future(months, now, 4) if m > now]
        under = [m for m in near if pp["estimated"].get(m, 0) > 0 and rem.get(m, 0) > 1.5]
        if len(under) >= 2 and not over:   # don't flag someone both over and under
            recs.append(dict(
                id=f"person-under-{slug(nm)}", category="capacity-person",
                severity="low", scope_type="person", scope=nm,
                title=f"{nm} under-utilized soon",
                detail=f"{nm} ({pp['team']}) has ≥1.5 pts (≈48h+) free in {span_str(under)}.",
                metric=f"{round(max(rem[m] for m in under),1)} pts free",
                suggested_action="Assign upcoming work or pull a project forward."))

    # ---- R5 unstaffed committed work ----
    for team in ("Design", "Media", "PM"):
        un = cap["unassigned"].get(team, {})
        fut = [m for m in _future(months, now, 9) if un.get(m, 0) > 0.05]
        if fut:
            recs.append(dict(
                id=f"unstaffed-{team.lower()}", category="capacity-unstaffed",
                severity="medium", scope_type="team", scope=team,
                title=f"Unstaffed {team} work committed",
                detail=f"{round(sum(un[m] for m in fut),1)} pts of {team} work in "
                       f"{span_str(fut)} is committed but not assigned to a named person.",
                metric=f"{round(sum(un[m] for m in fut),1)} pts",
                suggested_action="Assign an owner or confirm the work is real before it slips."))

    # ---- Asana hygiene & performance (matched + asana-only) ----
    report_cutoff = _months_ago(now, 9)   # only nag about recent completions
    seen, arecs = set(), []
    for p in data["projects"]:
        a = p.get("asana")
        if a and a.get("gid") and a["gid"] not in seen:
            seen.add(a["gid"]); arecs.append((p["name"], a))
    for a in data["asana_only_projects"]:
        if a.get("gid") and a["gid"] not in seen:
            seen.add(a["gid"]); arecs.append((a["name"], a))

    for name, a in arecs:
        gid = a["gid"]
        # (Impact-Tracker-outdated items were intentionally dropped — low-signal
        # data-hygiene noise that cluttered the tracker.)
        lc = (a.get("last_completed") or "")[:7]
        if (a.get("status") == "Complete" and a.get("post_status") in (None, "No report produced")
                and lc and lc >= report_cutoff):
            recs.append(dict(
                id=f"report-{gid}", category="post-project", severity="low",
                scope_type="project", scope=name, gid=gid,
                title=f"No post-project report: {name}",
                detail=f"Project is Complete (last activity {mlabel(lc)}) but has no post-project "
                       f"evaluation/report on record.",
                metric="no report", suggested_action="Send the reflection survey and produce the evaluation report."))
        if a.get("status") == "On Hold":
            recs.append(dict(
                id=f"hold-{gid}", category="on-hold", severity="medium",
                scope_type="project", scope=name, gid=gid,
                title=f"On hold: {name}",
                detail="Project is On Hold in Asana.",
                metric="on hold", suggested_action="Confirm a restart date or formally close it to free capacity."))

    # ---- R10 estimation calibration (grouped) ----
    mism = [(p["name"], p["est_size"], p["actual_size"]) for p in data["projects"]
            if p.get("est_size") and p.get("actual_size") and p["est_size"] != p["actual_size"]]
    if mism:
        ex = "; ".join(f"{n} (est {e}→act {a})" for n, e, a in mism[:6])
        recs.append(dict(
            id="calib-size", category="calibration", severity="low", scope_type="global",
            scope="Estimation", title=f"{len(mism)} projects landed off their estimated size",
            detail=f"The estimated and actual size estimate differ for {len(mism)} projects. e.g. {ex}.",
            metric=f"{len(mism)} projects",
            suggested_action="Review these in a retro to tighten future sizing estimates."))

    # ---- R11 workbook drift (grouped) ----
    if data.get("drift"):
        worst = max(data["drift"], key=lambda d: abs(d["delta"]))
        recs.append(dict(
            id="hygiene-drift", category="data-hygiene", severity="medium", scope_type="global",
            scope="Workbook", title="Team Capacity tab out of sync with Projects tab",
            detail=f"{len(data['drift'])} month/scope cells differ by >2 pts between the workbook's "
                   f"hand-maintained Team Capacity tab and the consistent recompute from the Projects "
                   f"tab (worst: {worst['scope']} {mlabel(worst['month'])}, "
                   f"tab={worst['workbook']} vs computed={worst['recomputed']}).",
            metric=f"{len(data['drift'])} cells",
            suggested_action="Re-point the Team Capacity tab formulas at the Projects tab, "
                             "or retire that tab in favor of this dashboard."))

    for r in recs:
        r.setdefault("source", "auto")
        r.setdefault("gid", None)
    return recs


def build_faculty_feedback(data, now):
    """Recommendations grounded in actual faculty feedback — one per project that
    has a reflection report we extracted a summary from. The detail is what
    faculty said; the live report is linked (doc_url) so the team can act on it
    (capture a testimonial, log any improvement). This is what lets the
    Recommendations tab draw on real perspectives data, not just capacity/Asana
    hygiene signals."""
    # rating-style lines ("ODL Team Rating", "Learning Designer - 5 (Extremely
    # Good Support)") are scores, not action items — keep them out of the action.
    RATING = re.compile(r"(?i)(^odl team|\s[-–]\s*\d|\brating\b|extremely good|good support)")
    by_proj = {}
    for r in data.get("reflections", []):
        if r.get("type") != "report":
            continue
        proj, summ = r.get("project"), (r.get("summary") or [])
        if not proj or not summ:
            continue
        by_proj.setdefault(proj, []).append(r)

    def digit_ratio(rep):  # lower = more prose-like (vs a raw Qualtrics number dump)
        s = " ".join((rep.get("summary") or [])[:2])
        return sum(c.isdigit() for c in s) / max(1, len(s))

    # the reports carry their OWN "Recommendations / Lessons Learned / What could
    # have gone better" sections — distill those into the action rather than dumping
    # the whole summary, so the tracker shows real, project-grounded recommendations.
    REC_HDR = re.compile(r"(?i)\b(recommendations?|lessons?\s+learned|what could have gone "
                         r"better|project improvements?|do differently|to improve)\b")
    # where the extractor concatenated the NEXT report section — cut there so the
    # "Project Details / Timeframe / Description" boilerplate doesn't bleed in.
    BOUNDARY = re.compile(r"(?i)\b(project details|project timeframe|project description|"
                          r"outcomes? summary|survey results|key findings|date of retrospective|"
                          r"participants)\b")

    def trim_section(t):
        t = re.sub(r"\s+", " ", str(t)).strip(" ;:-–/")
        m = BOUNDARY.search(t)
        if m and m.start() >= 12:
            t = t[:m.start()].rstrip(" ;:-–/")
        return t

    def cap(t, n):                      # word-boundary truncate + ellipsis (no mid-word cuts)
        t = trim_section(t)
        if len(t) <= n:
            return t
        cut = t[:n]
        sp = cut.rfind(" ")
        return (cut[:sp] if sp > 40 else cut).rstrip(" ;:,.-–/") + "…"

    def distill(summ):
        out = []
        for s in summ:
            s = str(s).strip()
            if not REC_HDR.search(s) or RATING.search(s):
                continue
            body = s
            for _ in range(3):  # peel compound headers, e.g. "What could have gone better / Lessons Learned;"
                new = re.sub(r"(?i)^\s*[/:;\-–]*\s*(recommendations?|lessons?\s+learned|what could "
                             r"have gone better|project improvements?|do differently|to improve)"
                             r"\s*[/:;\-–]*\s*", "", body, count=1)
                if new == body:
                    break
                body = new
            body = trim_section(body)
            if len(body.split()) >= 4:
                out.append(body)
        return out

    recs = []
    for proj, reports in by_proj.items():
        r = min(reports, key=digit_ratio)   # one rec per project, prefer the cleanest summary
        summ = r.get("summary") or []
        distilled = distill(summ)
        if distilled:
            detail = cap("  •  ".join(distilled[:2]), 520)
            action = cap(distilled[0], 300)
            title = f"Apply lessons from {proj}"
        else:
            # prefer summary/takeaway content that doesn't START with report boilerplate
            clean = [s for s in summ if not BOUNDARY.match(str(s).strip())] or summ
            detail = cap(" ".join(clean[:2]), 500)
            take = [t for t in (r.get("takeaways") or [])    # skip empty list-lead-in preambles ("… identified:") + boilerplate
                    if len(t.split()) >= 6 and not RATING.search(t)
                    and not str(t).rstrip().endswith(":") and not BOUNDARY.match(str(t).strip())]
            action = (cap(take[0], 300) if take else
                      "Review the reflection report; capture a testimonial and log any improvement actions.")
            title = f"Faculty feedback — {proj}"
        recs.append(dict(
            id=f"faculty-fb-{slug(proj)}", category="faculty-feedback", severity="low",
            scope_type="project", scope=proj, title=title,
            detail=detail, metric=(r.get("completion_rate") or "reflection report"),
            suggested_action=action, source="auto", gid=None,
            doc_url=(r.get("doc_url") or r.get("drive_url") or r.get("rel"))))
    recs.sort(key=lambda x: x["scope"])
    return recs


def merge(data, manual, statuses, now, today):
    auto = build_auto(data, now) + build_faculty_feedback(data, now)
    clean = []
    for m in (manual or []):
        if not isinstance(m, dict) or not (m.get("title") or m.get("id")):
            continue   # skip malformed hand-edited entries instead of crashing
        m.setdefault("source", "manual")
        m.setdefault("severity", "medium")
        m.setdefault("category", "manual")
        m.setdefault("scope_type", "global")
        if not m.get("id"):
            m["id"] = "manual-" + slug(m.get("title", "untitled"))
        m.setdefault("title", m.get("id"))   # id-only hand-edits must not break the sort
        clean.append(m)
    recs = auto + clean
    # overlay tracked status; seed new ones
    for r in recs:
        st = statuses.setdefault(r["id"], {})
        st.setdefault("status", "Not Started")
        st.setdefault("first_seen", today)
        r["status"] = st.get("status", "Not Started")
        r["owner"] = st.get("owner", "")
        r["target_month"] = st.get("target_month", "")
        r["notes"] = st.get("notes", "")
        r["evidence_url"] = st.get("evidence_url", "")
        r["updated_at"] = st.get("updated_at", "")
        r["updated_by"] = st.get("updated_by", "")
        r["first_seen"] = st.get("first_seen", today)
    # active auto ids (to detect resolved/stale on the UI side)
    active_ids = {r["id"] for r in recs}
    for rid, st in statuses.items():
        st["_active"] = rid in active_ids
    recs.sort(key=lambda r: (SEV_RANK.get(r.get("severity"), 9),
                             {"auto": 0, "manual": 1}.get(r.get("source"), 2), r.get("title", "")))
    return recs, statuses
