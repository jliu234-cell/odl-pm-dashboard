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


def _rnorm(s):
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def _future(months, now, n=None):
    f = [m for m in months if m >= now]
    return f[:n] if n else f


def build_auto(data, now):
    cap = data["capacity"]
    months = data["meta"]["months"]
    full = data["meta"]["full_monthly_points"]
    recs = []

    # ---- R1/R2 team over / tight capacity (forward 9 months) ----
    # Framed by TEAM, not by naming individuals (director feedback: lead with the
    # per-team situation and a facilitation ask, never an "underutilized people" list).
    flagged = set()   # teams already called stretched/tight, so we don't also call them "has room"
    for scope in ("Design", "Media", "PM"):
        pcts = cap["pct"].get(scope, {})
        win = _future(months, now, 9)
        over = [m for m in win if pcts.get(m) and pcts[m] > 1.0]
        tight = [m for m in win if pcts.get(m) and 0.9 <= pcts[m] <= 1.0]
        if over:
            flagged.add(scope)
            worst = max(over, key=lambda m: pcts[m])
            recs.append(dict(
                id=f"cap-over-{scope.lower()}", category="capacity-team",
                severity="high", scope_type="team", scope=scope,
                title=f"{scope} is booked beyond capacity",
                detail=f"The {scope} team's allocated hours run over 100% of capacity in "
                       f"{span_str(over)} (peak {round(pcts[worst]*100)}% in {mlabel(worst)}).",
                metric=f"peak {round(pcts[worst]*100)}%",
                suggested_action="Help the team rebalance — shift a deliverable, bring in "
                                 "another person, or move a timeline in these months."))
        elif tight:
            flagged.add(scope)
            worst = max(tight, key=lambda m: pcts[m])
            recs.append(dict(
                id=f"cap-tight-{scope.lower()}", category="capacity-team",
                severity="medium", scope_type="team", scope=scope,
                title=f"{scope} is running close to full",
                detail=f"The {scope} team is 90–100% allocated in {span_str(tight)} "
                       f"(peak {round(pcts[worst]*100)}% in {mlabel(worst)}), so there's little slack for new work.",
                metric=f"peak {round(pcts[worst]*100)}%",
                suggested_action="Hold a buffer here for overruns before committing new scope."))

    # ---- team headroom (facilitation, by team — replaces per-person "under-utilized") ----
    for scope in ("Design", "Media", "PM"):
        if scope in flagged:
            continue
        pcts = cap["pct"].get(scope, {})
        win = _future(months, now, 4)
        light = [m for m in win if pcts.get(m) is not None and pcts[m] < 0.7]
        if len(light) >= 2:
            lo = min(light, key=lambda m: pcts[m])
            recs.append(dict(
                id=f"cap-room-{scope.lower()}", category="capacity-team",
                severity="low", scope_type="team", scope=scope,
                title=f"{scope} has room to take on more",
                detail=f"The {scope} team is below a full load in {span_str(light)} "
                       f"(down to {round(pcts[lo]*100)}% in {mlabel(lo)}) — a good window to pull "
                       f"upcoming work forward or start something new.",
                metric=f"as low as {round(pcts[lo]*100)}%",
                suggested_action=f"Route upcoming {scope.lower()} work here, or schedule a project "
                                 f"the team could pick up."))

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
        # someone carrying MORE than a full load — a support signal, framed gently
        over = [m for m in _future(months, now, 6) if rem.get(m) is not None and rem[m] < -0.25]
        if over:
            worst = min(over, key=lambda m: rem[m])
            recs.append(dict(
                id=f"person-over-{slug(nm)}", category="capacity-person",
                severity="medium", scope_type="person", scope=nm,
                title=f"{nm} is carrying more than a full load",
                detail=f"{nm} ({pp['team']})'s allocated hours run over capacity in {span_str(over)} "
                       f"(about {round(abs(rem[worst])*32)}h over in {mlabel(worst)}) — worth a check-in "
                       f"so the workload stays sustainable.",
                metric=f"~{round(abs(rem[worst])*32)}h over",
                suggested_action="Check in and, if it's too much, rebalance a deliverable or shift a timeline."))
        # (Per-person under-utilization is intentionally not flagged — headroom is
        #  surfaced by TEAM above, per director feedback: no "underutilized people" list.)

    # ---- R5 unstaffed committed work (the sheet's "<team> Unassigned" rows) ----
    for team in ("Design", "Media", "PM"):
        un = cap["unassigned"].get(team, {})
        fut = [m for m in _future(months, now, 9) if un.get(m, 0) > 0.05]
        if fut:
            hrs = round(sum(un[m] for m in fut) * 32)
            recs.append(dict(
                id=f"unstaffed-{team.lower()}", category="capacity-unstaffed",
                severity="medium", scope_type="team", scope=team,
                title=f"{team} work is committed but has no owner yet",
                detail=f"About {hrs}h of {team} work in {span_str(fut)} is allocated to an "
                       f"“Unassigned” slot rather than a named person.",
                metric=f"~{hrs}h unassigned",
                suggested_action="Give it an owner, or confirm the work is real before it slips."))

    # ---- Asana hygiene & performance (matched + asana-only) ----
    # A project already HAS a reflection report if it's in the Project Reflection
    # Reports 2025 Drive folder (reflection_key_considerations.json). Match on a
    # normalized token-subset so "MS ACMS" == "MS-ACMS Faculty Videos", "Responsible
    # & Ethical AI" == "Responsible and Ethical AI", etc. — so we never falsely nag
    # a project whose report actually exists.
    report_tok = [set(_rnorm(r.get("project")).split())
                  for r in (data.get("reflection_considerations") or {}).get("reports", [])
                  if r.get("project")]

    def _has_report(name):
        nt = set(_rnorm(name).split())
        if not nt:
            return False
        for rt in report_tok:
            if nt == rt or (len(nt & rt) >= 2 and (nt <= rt or rt <= nt)):
                return True
        return False

    # We can't retroactively get a reflection for a long-closed project, so instead
    # of nagging about missing reports we surface RECENTLY-wrapped projects with no
    # reflection yet — "it's about time to collect faculty feedback while it's fresh."
    reflect_cutoff = _months_ago(now, 4)
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
        lc = (a.get("last_completed") or "")[:7]
        if (a.get("status") == "Complete" and a.get("post_status") in (None, "No report produced")
                and lc and lc >= reflect_cutoff and not _has_report(name)):
            recs.append(dict(
                id=f"report-{gid}", category="post-project", severity="low",
                scope_type="project", scope=name, gid=gid,
                title=f"Time to collect faculty feedback: {name}",
                detail=f"{name} wrapped up recently ({mlabel(lc)}) and doesn't have a reflection on "
                       f"record yet — a good moment to gather the faculty partner's feedback while it's fresh.",
                metric="due for reflection",
                suggested_action="Send the reflection survey to the faculty partner and schedule the retro."))
        if a.get("status") == "On Hold":
            recs.append(dict(
                id=f"hold-{gid}", category="on-hold", severity="medium",
                scope_type="project", scope=name, gid=gid,
                title=f"On hold: {name}",
                detail="Project is On Hold in Asana.",
                metric="on hold", suggested_action="Confirm a restart date or formally close it to free capacity."))

    # ---- R10 estimate accuracy: planned hours (Capacity sheet) vs actual ----
    ea = data.get("est_actual") or {}
    def _h(x):
        return round(x) if isinstance(x, (int, float)) else "—"
    over = [r for r in ea.get("projects", []) if r.get("over")]
    if over:
        ex = "; ".join(f"{r['name']} ({_h(r.get('actual_hours'))}h vs {_h(r.get('planned_hours'))}h planned)"
                       for r in over[:5])
        recs.append(dict(
            id="calib-over-budget", category="calibration", severity="medium", scope_type="global",
            scope="Estimation", title=f"{len(over)} active project(s) over their planned hours",
            detail=f"Logged time has passed ~120% of the planned (Capacity-sheet) hours for "
                   f"{len(over)} active project(s): {ex}. That crowds out other work or hides overtime.",
            metric=f"{len(over)} over plan",
            suggested_action="Check scope/timeline on these and re-estimate the remaining work."))
    n_active = ea.get("n_active", 0)
    n_pa = ea.get("n_with_planned", 0)
    if n_active and n_pa < max(3, n_active // 2):
        recs.append(dict(
            id="calib-hours-coverage", category="calibration", severity="medium", scope_type="global",
            scope="Estimation", title="Most active projects have no planned hours to track against",
            detail=f"Only {n_pa} of {n_active} active projects have planned hours on the Capacity "
                   f"Allocations sheet, so planned-vs-actual can't be measured for the rest. Listing "
                   f"each person's projects on the sheet turns the timesheet into real calibration.",
            metric=f"{n_pa}/{n_active} with a plan",
            suggested_action="Make sure active projects are listed against people on the Capacity "
                             "Allocations sheet so planned hours flow through."))

    # ---- R11 status-update-derived problems -> recommendation items (with evidence
    #      + Asana links) so the Recommendations tab draws on the weekly updates too ----
    ws = (data.get("meta", {}).get("sources", {}) or {}).get("asana_ws", "https://app.asana.com")
    for f in (data.get("status_problems") or {}).get("flagged", [])[:12]:
        gid = f.get("gid")
        reasons = f.get("reasons") or []
        recs.append(dict(
            id=f"risk-{gid or slug(f['name'])}", category="on-hold" if any("On hold" in x for x in reasons) else "post-project",
            severity="high" if f.get("level", 0) >= 3 else "medium",
            scope_type="project", scope=f["name"], gid=gid,
            title=f"Project at risk: {f['name']}",
            detail=" · ".join(reasons)[:500] + (f" (as of {f['date']})" if f.get("date") else ""),
            metric="flagged at risk",
            suggested_action="Read the weekly update and follow up on the blocker or timeline.",
            doc_url=(ws + "/project/" + gid) if gid else None))

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
    # NB: the old per-report faculty-feedback recs (build_faculty_feedback) are
    # retired. They duplicated the Recommendations tab's "Key considerations for
    # future projects" section (which is agent-verified from the reports folder and
    # correctly separates reports that have no retrospective, e.g. Parent
    # Empowerment) and mis-fired on those empty reports. Reflection content now
    # lives only in that section + the Director Brief.
    auto = build_auto(data, now)
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
