#!/usr/bin/env python3
"""ODL PM Capacity & Performance Dashboard — build/ETL step.

Reads the human-maintained capacity workbook (ODL Project and Capacity
Planning.xlsx) and the nightly Asana snapshot (../odl_estimator/data_all/) and
emits a single ``data.json``.  ``render.py`` then bakes that into a
self-contained ``index.html``.

Capacity model (mirrors the workbook's Explanations tab):
  1 point = 32 hours = ~1 productive week.  4 points/person/month = full load.

Canonical figures are recomputed *consistently* so they always reconcile:
  estimated[scope][m] = Σ person estimated capacity (the workbook's
      'Estimated Capacity (Points)' tab — its per-person values sum EXACTLY to
      the workbook's Total, verified at build time).
  scheduled[scope][m] = Σ point allocations on the 'Projects' tab, by role-team.
  remaining = estimated − scheduled ;  pct = scheduled / estimated.
The workbook's own 'Team Capacity' tab is kept as a *reference* overlay; where
it diverges from the consistent recompute we surface a drift flag (the tab is
maintained by hand and has drifted — a real data-hygiene signal, not hidden).

No invented numbers: every figure traces to a sheet cell or snapshot field.

Run:  python3 build.py            # writes data.json (+ validation report)
      python3 build.py --report   # validation report only, no write
"""
import os, sys, csv, json, re, datetime, argparse
import recommend

HERE = os.path.dirname(os.path.abspath(__file__))
MANUAL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "recommendations_manual.json")
STATUS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "statuses.json")
ND_DEPT_FILE = os.path.join(HERE, "nd_departments.json")  # ND-website dept enrichment
DRIVE_LINKS_FILE = os.path.join(HERE, "reflection_drive_links.json")  # Drive URLs for Reflection/ PDFs
BRIEF_INPUTS_FILE = os.path.join(HERE, "brief_inputs.json")  # manual weekly inputs for the Director Brief
ROOT = os.path.dirname(HERE)
XLSX = os.path.join(ROOT, "ODL Project and Capacity Planning.xlsx")
ASANA_DIR = os.path.join(ROOT, "odl_estimator", "data_all")
REFLECTION_DIR = os.path.join(ROOT, "Reflection")   # downloaded reflection PDFs
REFLECTION_REL = "../Reflection/"                    # link path relative to index.html
POINT_HOURS = 32
FULL_MONTHLY_POINTS = 4

# person -> team, seeded from the workbook's per-person sheet names
# (x_LD_*, x_MP_*, x_PM_*, x_Intern_*) which encode each person's team.
ROSTER_SEED = {
    "Design": ["Yi", "Kuangchen", "Bri", "Janet (Temp)", "Janet"],
    "Media":  ["Matthew", "Tim", "Adam", "Adam - Freelance", "Kevin", "Colin",
               "Derrick", "KC", "Naomi"],
    "PM":     ["Michael", "Annie", "Jordan", "Lawrence", "Janyl", "Michael T", "Sonia"],
    "Intern": ["Nina", "Maddie", "Minyoung"],
}
TEAM_ORDER = ["Design", "Media", "PM", "Intern", "Other"]
SCOPES = ["Total", "Design", "Media", "PM", "Intern", "Other"]

# "person" cells that are really unstaffed placeholders, not a named individual
GENERIC_PERSONS = {
    "design", "media", "pm", "graphics", "design (generic)", "media (generic)",
    "pm (generic)", "design team", "media team", "pm team", "graphics team",
    "student", "student worker", "design sw", "media sw", "tbd", "unassigned",
    "intern", "freelance", "media intern", "design intern", "xr intern", "ld intern",
}


# explicit, reviewed name aliases (the same person spelled different ways across
# tabs). Only collapse pairs we've confirmed — never blanket suffix-stripping.
PERSON_ALIASES = {
    "nina - ld intern": "Nina",
    "maddie - media intern": "Maddie",
    "minyoung - xr intern": "Minyoung",
}


def canon_person(n):
    """Collapse confirmed spelling variants so one individual isn't split across
    keys: trailing punctuation ('Michael T.' -> 'Michael T') + the reviewed
    PERSON_ALIASES. Conservative — no heuristic role/note-suffix stripping."""
    if not n:
        return n
    s = n.strip().rstrip(".").strip()
    return PERSON_ALIASES.get(s.lower(), s)


def role_to_team(role):
    r = (role or "").strip().lower()
    if not r:
        return "Other"
    if "develop" in r:                       # developer / development -> Media (dev)
        return "Media"
    if re.search(r"\bpm\b", r) or "project manager" in r:   # word boundary: not "develoPMent"
        return "PM"
    if "design" in r:
        return "Design"
    if "media" in r or "graphic" in r:
        return "Media"
    if "intern" in r:
        return "Intern"
    return "Other"


def is_generic_person(name):
    """True for unstaffed placeholders (e.g. 'Media', 'Graphics Team',
    'XR Developer'), NOT for named contributors (e.g. 'Adam - Freelance')."""
    s = (name or "").strip().lower()
    if not s:
        return True
    if s in GENERIC_PERSONS:
        return True
    if s.endswith(" team") or "(generic)" in s:
        return True
    return s in ("developer", "xr developer", "lead media", "lead design",
                 "lead designer", "designer", "graphics")


def ym(dt):
    if isinstance(dt, (datetime.datetime, datetime.date)):
        return f"{dt.year:04d}-{dt.month:02d}"
    return None


def norm(s):
    # drop apostrophes BEFORE splitting so "Qur'an"->"quran" (not "qur an"),
    # "Roberto's"->"robertos" — keeps such tokens matchable.
    s = re.sub(r"['’‘`]", "", (s or "").lower())
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def as_num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip().replace("%", "")
        try:
            return float(s)
        except ValueError:
            return None
    return None


def load_wb():
    import openpyxl
    wb = openpyxl.load_workbook(XLSX, data_only=True, read_only=True)
    return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)]
            for ws in wb.worksheets}


# --------------------------------------------------------------------------- #
#  Projects sheet  ->  projects[], roles, allocations, phases
# --------------------------------------------------------------------------- #
PROJ_MONTH_START_COL = 11  # 0-based; col12 (=index 11) is the first month, 2024-09


def parse_projects(rows):
    hdr = rows[1]
    month_cols = {i: ym(c) for i, c in enumerate(hdr)
                  if i >= PROJ_MONTH_START_COL and ym(c)}
    months = [month_cols[i] for i in sorted(month_cols)]

    projects, cur = [], None
    SECTION_HINTS = ("projects in progress", "completed", "on hold", "backlog",
                     "potential", "archived", "not started", "future")

    def flush():
        nonlocal cur
        if cur and cur["roles"]:
            projects.append(cur)
        cur = None

    for ri in range(2, len(rows)):
        r = rows[ri]
        if not r:
            continue
        c1 = r[0]
        label = (str(c1).strip() if c1 is not None else "")
        indented = isinstance(c1, str) and c1[:1] in (" ", "\t")
        proj_col = (str(r[1]).strip() if len(r) > 1 and r[1] is not None else "")
        has_attr = any([(len(r) > 2 and r[2]), (len(r) > 4 and r[4]),
                        (len(r) > 8 and r[8]), (len(r) > 9 and r[9])])

        if indented:  # role row
            if cur is None:
                cur = new_project(proj_col or label)
            person = r[10] if len(r) > 10 else None
            person = canon_person(str(person).strip()) if person not in (None, "") else None
            alloc = {month_cols[i]: as_num(r[i]) for i in month_cols
                     if i < len(r) and as_num(r[i]) not in (None, 0.0)}
            cur["roles"].append({
                "role": label, "team": role_to_team(label),
                "person": person, "generic": is_generic_person(person) or person is None,
                "alloc": alloc,
            })
            continue

        if not label:
            continue
        low = label.lower()
        if not has_attr and not proj_col and any(h in low for h in SECTION_HINTS):
            flush()
            continue
        # project header row
        flush()
        cur = new_project(label)
        cur["est_size"] = clean_size(r[2] if len(r) > 2 else None)
        cur["actual_size"] = clean_size(r[3] if len(r) > 3 else None)
        cur["approx_points"] = as_num(r[4]) if len(r) > 4 else None
        cur["type"] = (str(r[6]).strip() if len(r) > 6 and r[6] else None)
        cur["start"] = ym(r[8]) if len(r) > 8 else None
        cur["end"] = ym(r[9]) if len(r) > 9 else None
        cur["phases"] = [{"month": month_cols[i], "label": str(r[i]).strip()}
                         for i in sorted(month_cols)
                         if i < len(r) and isinstance(r[i], str) and r[i].strip()]
    flush()

    for p in projects:
        tot = {}
        for role in p["roles"]:
            for m, v in role["alloc"].items():
                tot[m] = round(tot.get(m, 0.0) + v, 4)
        p["points_by_month"] = tot
        p["total_points"] = round(sum(tot.values()), 3)
        ms = sorted(tot)
        p["first_month"] = ms[0] if ms else p.get("start")
        p["last_month"] = ms[-1] if ms else p.get("end")
        p["staffed_points"] = round(sum(v for role in p["roles"] if not role["generic"]
                                        for v in role["alloc"].values()), 3)
        p["unstaffed_points"] = round(p["total_points"] - p["staffed_points"], 3)
    return projects, months


def clean_size(v):
    if v is None:
        return None
    s = str(v).strip().upper()
    return s if s in ("XS", "S", "M", "L", "XL", "XXL", "UNIQUE") else (str(v).strip() or None)


def new_project(name):
    return {"name": name, "est_size": None, "actual_size": None,
            "approx_points": None, "type": None, "start": None, "end": None,
            "phases": [], "roles": []}


# --------------------------------------------------------------------------- #
#  Estimated Capacity (Points)  ->  per-person monthly estimated capacity
# --------------------------------------------------------------------------- #
def parse_person_estimated(rows):
    month_map, out = None, {}
    for r in rows:
        if not r:
            continue
        dts = [(j, ym(c)) for j, c in enumerate(r) if ym(c)]
        if len(dts) >= 6 and month_map is None:
            month_map = dict(dts)
            continue
        if month_map and isinstance(r[0], str):
            nm = canon_person(r[0].strip())
            if not nm or nm.lower() in ("total", "estimated capacity"):
                continue
            vals = {m: as_num(r[j]) for j, m in month_map.items()
                    if j < len(r) and as_num(r[j]) is not None}
            if vals:
                out[nm] = vals
    return out


# --------------------------------------------------------------------------- #
#  Team Capacity sheet  ->  workbook reference overlay
# --------------------------------------------------------------------------- #
TC_BLOCKS = [("team capacity", "Total"), ("design team capacity", "Design"),
             ("media team estimated capacity", "Media"),
             ("project manager team capacity", "PM"),
             ("intern team capacity", "Intern")]
TC_METRICS = {"estimated capacity": "estimated", "scheduled capacity": "scheduled",
              "difference": "difference", "difference (hours)": "difference_hours",
              "% allocated": "pct"}


def parse_team_capacity(rows):
    n, starts = len(rows), []
    for ri in range(n):
        c1 = rows[ri][0] if rows[ri] else None
        if isinstance(c1, str):
            low = c1.strip().lower()
            for hint, scope in TC_BLOCKS:
                if low.startswith(hint):
                    starts.append((ri, scope))
                    break
    starts.append((n, None))
    out = {}
    for idx in range(len(starts) - 1):
        ri, scope = starts[idx]
        end = starts[idx + 1][0]
        month_map, block = None, {k: {} for k in TC_METRICS.values()}
        for rj in range(ri, end):
            r = rows[rj]
            if not r:
                continue
            dts = [(j, ym(c)) for j, c in enumerate(r) if ym(c)]
            if len(dts) >= 6 and month_map is None:
                month_map = dict(dts)
                continue
            if isinstance(r[0], str) and month_map:
                key = TC_METRICS.get(r[0].strip().lower())
                if key:
                    for j, m in month_map.items():
                        v = as_num(r[j]) if j < len(r) else None
                        if v is not None:
                            block[key][m] = v
        out[scope] = block
    return out


# --------------------------------------------------------------------------- #
#  Tshirt Project Types  ->  standard staffing profiles per size
# --------------------------------------------------------------------------- #
def parse_tshirt(rows):
    sizes, cur = {}, None
    SIZE_NAMES = {"large": "L", "medium": "M", "small": "S", "xs": "XS",
                  "extra large": "XL", "xl": "XL"}
    for r in rows:
        if not r or not isinstance(r[0], str) or not r[0].strip():
            continue
        head = r[0].strip().lower()
        if head in SIZE_NAMES:
            cur = SIZE_NAMES[head]
            sizes[cur] = {"roles": {}, "n_months": 0}
            continue
        if cur:
            vals = [as_num(x) for x in r[1:]]
            vals = [(v if v is not None else 0.0) for v in vals]
            while vals and vals[-1] == 0.0:
                vals.pop()
            if vals:
                sizes[cur]["roles"][r[0].strip()] = vals
                sizes[cur]["n_months"] = max(sizes[cur]["n_months"], len(vals))
    for s, d in sizes.items():
        d["total_points"] = round(sum(sum(v) for v in d["roles"].values()), 3)
        d["peak_monthly"] = round(max((sum(v[i] for v in d["roles"].values() if i < len(v))
                                       for i in range(d["n_months"])), default=0), 3)
    return sizes


# --------------------------------------------------------------------------- #
#  Asana snapshot
# --------------------------------------------------------------------------- #
def _stats(vals):
    vals = sorted(v for v in vals if v is not None)
    if not vals:
        return None
    n = len(vals)
    med = vals[n // 2] if n % 2 else (vals[n // 2 - 1] + vals[n // 2]) / 2
    return {"n": n, "mean": round(sum(vals) / n, 2), "median": round(med, 2),
            "min": round(min(vals), 2), "max": round(max(vals), 2)}


def _median(vals):
    vals = sorted(v for v in vals if v is not None)
    if not vals:
        return 0.0
    n = len(vals)
    return vals[n // 2] if n % 2 else (vals[n // 2 - 1] + vals[n // 2]) / 2


IMPACT_WANT = {"Faculty Satisfaction Index": "fsi", "Net Promoter Score": "nps",
               "Student Reach / Year": "reach", "Status": "status",
               "Academic Department": "dept", "Project Type": "type",
               "Total Assets": "assets", "Faculty Collaborators": "faculty",
               "ODL Departments Involved": "odl_depts", "Summary": "summary",
               "Link to Reflection Report": "reflink",
               # dates + per-year effort drive the Faculty "by year" view
               "Start Date": "start", "End Date": "end",
               "Total Hours": "hours_total", "Total Hours (2024)": "h2024",
               "Total Hours (2025)": "h2025", "Total Hours (2026)": "h2026"}
IMPACT_NUMERIC = {"fsi", "nps", "reach", "assets",
                  "hours_total", "h2024", "h2025", "h2026"}
IMPACT_DATE = {"start", "end"}


def parse_impact_tracker():
    """All tasks on the Asana 'Impact Tracker' board, keyed by gid (each task ==
    one ODL project's impact record), with the fields the dashboard surfaces."""
    path = os.path.join(ASANA_DIR, "task_custom_fields.csv")
    if not os.path.exists(path):
        return {}
    tasks = {}
    for r in csv.DictReader(open(path, encoding="utf-8")):
        if (r.get("project_name") or "").strip() != "Impact Tracker":
            continue
        field = IMPACT_WANT.get(r.get("custom_field_name"))
        if not field:
            continue
        key = r["task_gid"]
        rec = tasks.setdefault(key, {"name": (r.get("task_name") or "").strip()})
        num = as_num(r.get("number_value"))
        raw = (r.get("display_value") or r.get("text_value") or "").strip()
        # date custom fields carry a clean ISO date in date_value (display_value
        # is an ISO timestamp); prefer the date_value.
        date = (r.get("date_value") or "").strip()
        if field in IMPACT_NUMERIC:
            val = num
        elif field in IMPACT_DATE:
            val = date or raw[:10]
        else:
            val = raw
        if val not in (None, ""):
            rec[field] = val
    return tasks


def parse_faculty_ratings(impact):
    """Per-project faculty satisfaction from the Impact Tracker (tasks with an
    FSI). Fuller perspectives data lives in Qualtrics (access pending)."""
    if not impact:
        return None
    ratings = []
    for d in impact.values():
        if d.get("fsi") is None:
            continue  # only projects with a faculty rating
        ratings.append({"project": d.get("name", ""),
                        "fsi": d.get("fsi"), "nps": d.get("nps"),
                        "reach": d.get("reach"), "status": d.get("status"),
                        "dept": d.get("dept"), "type": d.get("type"),
                        "assets": d.get("assets"), "faculty": d.get("faculty"),
                        "reflink": d.get("reflink")})
    ratings.sort(key=lambda x: (-(x["fsi"] or 0), x["project"]))
    # distribution buckets for FSI (5-point scale)
    buckets = {"5.0": 0, "4.5–4.9": 0, "4.0–4.4": 0, "<4.0": 0}
    for r in ratings:
        f = r["fsi"]
        buckets["5.0" if f >= 5 else "4.5–4.9" if f >= 4.5 else "4.0–4.4" if f >= 4.0 else "<4.0"] += 1
    return {"source": "Asana Impact Tracker", "scale": 5,
            "fsi": _stats([r["fsi"] for r in ratings]),
            "nps": _stats([r["nps"] for r in ratings if isinstance(r["nps"], (int, float))]),
            "distribution": buckets, "ratings": ratings}


def build_departments(impact, nd):
    """Group Impact Tracker projects by academic department, so the Faculty tab
    can show which ND departments ODL has worked / is working with and on what.
    Departments come from Asana's 'Academic Department'; blanks are filled from
    nd_departments.json (inferred from nd.edu, tagged 'nd.edu' vs 'asana'). Each
    department also carries a one-line context blurb when nd_departments.json
    provides one."""
    proj_dept = (nd or {}).get("project_dept") or {}
    ctx = (nd or {}).get("department_context") or {}
    if not isinstance(proj_dept, dict):
        proj_dept = {}
    if not isinstance(ctx, dict):
        ctx = {}
    groups = {}
    for d in impact.values():
        name = d.get("name", "")
        if not name:
            continue
        dept = (d.get("dept") or "").strip()
        source, conf, url = ("asana" if dept else None), None, None
        if not dept:
            pd = proj_dept.get(name)
            if isinstance(pd, dict):
                dept = (pd.get("dept") or "").strip()
                if dept:
                    source, conf, url = "nd.edu", pd.get("confidence"), pd.get("source")
        groups.setdefault(dept or "—", []).append({
            "project": name, "faculty": d.get("faculty"), "type": d.get("type"),
            "fsi": d.get("fsi"), "status": d.get("status"),
            "reflink": d.get("reflink"), "summary": d.get("summary"),
            "dept_source": source, "dept_confidence": conf, "dept_url": url})
    depts = []
    for dname, projs in groups.items():
        if dname == "—":
            continue
        fsis = [p["fsi"] for p in projs if isinstance(p["fsi"], (int, float))]
        c = ctx.get(dname) or {}
        if not isinstance(c, dict):
            c = {}
        depts.append({"dept": dname, "n": len(projs),
                      "avg_fsi": round(sum(fsis) / len(fsis), 2) if fsis else None,
                      "official_name": c.get("official_name", ""),
                      "college": c.get("college", ""),
                      "context": c.get("blurb", ""),
                      "context_source": c.get("source", ""),
                      "projects": sorted(projs, key=lambda p: (-(p["fsi"] or 0), p["project"]))})
    depts.sort(key=lambda d: (-d["n"], d["dept"]))
    unknown = sorted(p["project"] for p in groups.get("—", []))
    return {"departments": depts, "unknown_count": len(unknown),
            "unknown_projects": unknown, "nd_loaded": bool(nd)}


def _resolve_dept(d, proj_dept):
    """Department for an Impact Tracker record: Asana's 'Academic Department',
    else inferred from nd.edu (nd_departments.json). Returns (dept, source)."""
    dept = (d.get("dept") or "").strip()
    if dept:
        return dept, "asana"
    pd = proj_dept.get(d.get("name", ""))
    if isinstance(pd, dict):
        dd = (pd.get("dept") or "").strip()
        if dd:
            return dd, "nd.edu"
    return "", None


def build_faculty_years(impact, nd):
    """Group Impact Tracker projects by the calendar year(s) they were active, so
    the Faculty tab can show what ODL worked on each year — with each project's
    department alongside. A project's active years span its Start Date → End Date
    (inclusive); projects with only one date use that year; projects with neither
    fall into an 'unknown' bucket. Per-year effort (Total Hours (YYYY)) is carried
    where Asana has it. Departments resolve the same way as the by-department view
    (Asana, then nd.edu)."""
    if not impact:
        return {"years": [], "unknown": [], "min_year": None, "max_year": None}
    proj_dept = (nd or {}).get("project_dept") or {}
    if not isinstance(proj_dept, dict):
        proj_dept = {}

    def yr(s):
        s = (s or "").strip()
        return int(s[:4]) if len(s) >= 4 and s[:4].isdigit() else None

    by_year, unknown = {}, []
    for d in impact.values():
        name = d.get("name", "")
        if not name:
            continue
        dept, dsrc = _resolve_dept(d, proj_dept)
        sy, ey = yr(d.get("start")), yr(d.get("end"))
        hours_by_year = {y: d.get(f"h{y}") for y in (2024, 2025, 2026)
                         if isinstance(d.get(f"h{y}"), (int, float)) and d.get(f"h{y}")}
        rec = {"project": name, "dept": dept or "—", "dept_source": dsrc,
               "faculty": d.get("faculty"), "type": d.get("type"),
               "fsi": d.get("fsi"), "status": d.get("status"),
               "reflink": d.get("reflink"), "assets": d.get("assets"),
               "hours_total": d.get("hours_total"), "hours_by_year": hours_by_year,
               "start": d.get("start"), "end": d.get("end")}
        if sy is None and ey is None:
            unknown.append(rec)
            continue
        lo = sy if sy is not None else ey
        hi = ey if ey is not None else sy
        if hi < lo:
            lo, hi = hi, lo
        # cap the span so a stray far-future end date can't explode the range
        for y in range(lo, min(hi, lo + 12) + 1):
            by_year.setdefault(y, []).append(rec)

    years = []
    for y in sorted(by_year, reverse=True):
        projs = by_year[y]
        fsis = [p["fsi"] for p in projs if isinstance(p["fsi"], (int, float))]
        dept_breakdown = {}
        for p in projs:
            dept_breakdown[p["dept"]] = dept_breakdown.get(p["dept"], 0) + 1
        hrs = round(sum(p["hours_by_year"].get(y, 0) for p in projs), 1)
        years.append({
            "year": y, "n": len(projs),
            "avg_fsi": round(sum(fsis) / len(fsis), 2) if fsis else None,
            "hours": hrs or None,
            "n_depts": len([k for k in dept_breakdown if k != "—"]),
            "dept_breakdown": dict(sorted(dept_breakdown.items(),
                                          key=lambda kv: (-kv[1], kv[0]))),
            "projects": sorted(projs, key=lambda p: (-(p["fsi"] or 0), p["project"]))})
    ys = [y["year"] for y in years]
    return {"years": years, "unknown": sorted(unknown, key=lambda p: p["project"]),
            "min_year": min(ys) if ys else None, "max_year": max(ys) if ys else None}


# --------------------------------------------------------------------------- #
#  Data-driven staffing profiles  ->  "how many projects can we take" demand
# --------------------------------------------------------------------------- #
PLAN_TEAMS = ["Design", "Media", "PM", "Intern"]
SIZE_ORDER = ["XS", "S", "M", "L", "XL"]


def build_size_profiles(projects, tshirt):
    """For each T-shirt size, the *typical* per-team staffing footprint computed
    from the real projects of that size on the Projects tab. This replaces the
    workbook's frozen 'Tshirt Project Types' template as the unit of demand in the
    planner so the estimate tracks what projects actually cost now (and, unlike
    the template, includes real PM time).

    The per-team headline demand (``teams``) is the **median** across projects of
    that team's total points — robust to a single mis-sized project (a few real
    projects are entered far larger/longer than their size implies). The
    month-by-month ``profile`` (used to stage a project) is the mean shape over a
    **representative duration** (the median project span, so one 18-month outlier
    can't smear an XS over 18 months), scaled so each team's profile sums to that
    team's median demand — keeping the staged month-by-month effect consistent
    with the estimator's headline number. Sizes with too few real projects fall
    back to the workbook template, tagged as such.

    Note: only the four staffable teams (Design/Media/PM/Intern) are modelled —
    the same teams the estimator measures surplus for — so a project's 'Other'
    allocation is not counted; projects whose work is *entirely* 'Other' are
    dropped and counted in ``dropped_other`` for the validation report."""
    by_size = {s: [] for s in SIZE_ORDER}
    dropped_other = {s: 0 for s in SIZE_ORDER}
    for p in projects:
        sz = (p.get("est_size") or "").upper()
        if sz not in by_size or not p.get("roles"):
            continue
        ms = sorted({m for role in p["roles"] for m in role["alloc"]})
        if not ms:
            continue
        idx = {m: i for i, m in enumerate(ms)}
        perteam = {t: [0.0] * len(ms) for t in PLAN_TEAMS}
        for role in p["roles"]:
            t = role["team"]
            if t not in perteam:
                continue
            for m, v in role["alloc"].items():
                perteam[t][idx[m]] += (v or 0.0)
        if sum(sum(a) for a in perteam.values()) <= 0:
            # has allocations, but all on the 'Other' team — not modelled
            if sum(sum(role["alloc"].values()) for role in p["roles"]) > 0:
                dropped_other[sz] += 1
            continue
        by_size[sz].append({"L": len(ms), "perteam": perteam})

    def wb_fallback(sz):
        wb = tshirt.get(sz)
        if not wb:
            return None
        prof = {t: [] for t in PLAN_TEAMS}
        for role, arr in wb["roles"].items():
            t = role_to_team(role)
            if t not in prof:
                continue
            for k, v in enumerate(arr):
                while len(prof[t]) <= k:
                    prof[t].append(0.0)
                prof[t][k] += (v or 0.0)
        ln = max((len(a) for a in prof.values()), default=0)
        for t in prof:
            prof[t] += [0.0] * (ln - len(prof[t]))
        teams = {t: round(sum(a), 3) for t, a in prof.items() if sum(a) > 0}
        peak = round(max((sum(prof[t][k] for t in prof) for k in range(ln)), default=0), 3)
        return {"n": 0, "source": "workbook template",
                "profile": {t: [round(x, 3) for x in prof[t]] for t in PLAN_TEAMS},
                "teams": teams, "months": ln, "peak_monthly": peak,
                "total_pts": round(sum(teams.values()), 3), "median_total": None}

    # POOLED team mix (point-weighted across ALL sized projects). This is the
    # stable Design/Media/PM/Intern split applied to every size. The old model
    # took each size's per-team MEDIAN independently, which on tiny samples
    # (XL n=2 — two atypical dev-heavy projects) produced a non-monotonic ladder
    # (XL had 16 Design / ~0 Media) and made the estimator claim it could absorb
    # MORE XL than XS. A size's *total* effort is reliable and rises with size, so
    # we keep the median grand total per size and split it by the pooled mix — the
    # per-team demand is then guaranteed to rise monotonically with size.
    pooled_sum = {t: 0.0 for t in PLAN_TEAMS}
    for s2 in SIZE_ORDER:
        for r in by_size[s2]:
            for t in PLAN_TEAMS:
                pooled_sum[t] += sum(r["perteam"][t])
    pooled_tot = sum(pooled_sum.values())
    pooled_frac = ({t: pooled_sum[t] / pooled_tot for t in PLAN_TEAMS} if pooled_tot > 1e-9
                   else {t: 0.0 for t in PLAN_TEAMS})

    out, order = {}, []
    prev_total = 0.0
    for sz in SIZE_ORDER:
        recs = by_size[sz]
        if len(recs) < 2:                      # too thin to be representative
            fb = wb_fallback(sz)
            if not fb:
                continue
            # bring the fallback onto the SAME model as real sizes: keep the
            # template's total but split it by the pooled team mix (so it carries
            # PM time and the right shape) and feed it through the monotonic clamp,
            # so a thin sample can never re-invert the ladder.
            size_total = max(fb["total_pts"] or 0.0, prev_total)
            prev_total = size_total
            repL = fb["months"] or 1
            tot_shape = [sum(fb["profile"][t][k] for t in PLAN_TEAMS) for k in range(repL)]
            ssum = sum(tot_shape)
            tot_shape = ([x / ssum for x in tot_shape] if ssum > 1e-9
                         else [1.0 / repL] * repL)        # uniform if template has no shape
            teams = {t: round(size_total * pooled_frac.get(t, 0.0), 3)
                     for t in PLAN_TEAMS if pooled_frac.get(t, 0.0) > 1e-9}
            profile = {t: [round(tot_shape[k] * teams.get(t, 0.0), 3) for k in range(repL)]
                       for t in PLAN_TEAMS}
            out[sz] = {"n": 0, "source": "workbook template (pooled split)",
                       "profile": profile, "teams": teams, "months": repL,
                       "peak_monthly": round(max((sum(profile[t][k] for t in PLAN_TEAMS)
                                                  for k in range(repL)), default=0), 3),
                       "total_pts": round(sum(teams.values()), 3), "median_total": None,
                       "team_mix": {t: round(pooled_frac[t], 3) for t in PLAN_TEAMS if pooled_frac[t] > 1e-9}}
            order.append(sz)
            continue
        n = len(recs)
        # representative duration = median project span (caps an outlier's smear)
        repL = max(1, int(round(_median([r["L"] for r in recs]))))
        # mean monthly shape per team (WHEN the work tends to happen)
        shape = {t: [0.0] * repL for t in PLAN_TEAMS}
        for r in recs:
            for t in PLAN_TEAMS:
                for k in range(repL):
                    shape[t][k] += (r["perteam"][t][k] if k < r["L"] else 0.0)
        for t in PLAN_TEAMS:
            shape[t] = [x / n for x in shape[t]]
        # size total = median grand total (robust, monotonic); kept non-decreasing
        # across the ladder so a noisy small sample can never invert the ordering.
        med_total = _median([sum(sum(r["perteam"][t]) for t in PLAN_TEAMS) for r in recs])
        size_total = max(med_total, prev_total)
        prev_total = size_total
        # split that total by the pooled team mix
        teams = {}
        for t in PLAN_TEAMS:
            d = round(size_total * pooled_frac.get(t, 0.0), 3)
            if d > 1e-9:
                teams[t] = d
        # profile: scale each team's shape to its demand; if this size never logged
        # a team's work (zero shape) but the pooled mix gives it demand, spread it
        # uniformly so staging still reflects the corrected footprint.
        profile = {}
        for t in PLAN_TEAMS:
            dem = teams.get(t, 0.0)
            s = sum(shape[t])
            if dem <= 0:
                profile[t] = [0.0] * repL
            elif s > 1e-9:
                profile[t] = [round(x * dem / s, 3) for x in shape[t]]
            else:
                profile[t] = [round(dem / repL, 3)] * repL
        while repL > 0 and all(profile[t][repL - 1] < 1e-9 for t in PLAN_TEAMS):
            for t in PLAN_TEAMS:
                profile[t].pop()
            repL -= 1
        peak = round(max((sum(profile[t][k] for t in PLAN_TEAMS) for k in range(repL)),
                         default=0), 3)
        out[sz] = {"n": n, "source": "actual projects",
                   "profile": {t: profile[t] for t in PLAN_TEAMS},
                   "teams": teams, "months": repL, "peak_monthly": peak,
                   "total_pts": round(sum(teams.values()), 3),
                   "median_total": round(med_total, 3),
                   "team_mix": {t: round(pooled_frac[t], 3) for t in PLAN_TEAMS if pooled_frac[t] > 1e-9}}
        order.append(sz)
    return {"order": order, "sizes": out, "plan_teams": PLAN_TEAMS,
            "team_mix": {t: round(pooled_frac[t], 3) for t in PLAN_TEAMS if pooled_frac[t] > 1e-9},
            "dropped_other": {s: c for s, c in dropped_other.items() if c}}


def _pdf_text(path):
    try:
        import fitz
        return "\n".join(p.get_text() for p in fitz.open(path))
    except Exception:
        try:
            import pypdf
            return "\n".join(pg.extract_text() or "" for pg in pypdf.PdfReader(path).pages)
        except Exception:
            return ""


_ZW = re.compile(r"[​‌‍⁠­]")
_BULLET = re.compile(r"^\s*[●•◦▪\-\*•]\s*")
_SUM_HDR = re.compile(r"summary|highlights|what (went )?well|feedback", re.I)
_TAKE_HDR = re.compile(r"takeaway|action item|recommendation|next step|improv|lesson", re.I)


def _clean(s):
    return _ZW.sub("", (s or "")).replace("’", "'").replace("“", '"').replace("”", '"').strip()


def _report_sections(txt):
    """Pull Summary-style and Takeaway/Action-item bullets from an ODL report."""
    lines = [_clean(l) for l in txt.split("\n")]
    summary, takeaways, cur, buf = [], [], None, ""

    def flush():
        nonlocal buf
        b = re.sub(r"\s*[○◦▪]\s*", "; ", buf).strip(" ;")
        b = re.sub(r"\s+", " ", b)
        if len(b) > 8 and not b.lower().endswith(".pdf"):
            (summary if cur == "s" else takeaways if cur == "t" else []).append(b)
        buf = ""

    for l in lines:
        if not l:
            continue
        head = l.rstrip(":").strip()
        if len(head) <= 60 and _TAKE_HDR.search(head) and not _BULLET.match(l):
            flush(); cur = "t"; continue
        if len(head) <= 60 and _SUM_HDR.search(head) and not _BULLET.match(l):
            flush(); cur = "s"; continue
        if cur is None:
            continue
        if _BULLET.match(l):
            flush(); buf = _BULLET.sub("", l)
        else:
            buf += " " + l
    flush()
    cap = lambda xs: [x[:400] for x in xs[:6]]
    return cap(summary), cap(takeaways)


def parse_reflections(projects, faculty, impact):
    """Index the downloaded reflection PDFs, extract report summaries, match each
    to a project, and attach the live Google-Doc URL (Asana 'Link to Reflection
    Report') so the dashboard opens the real document — not a local file path
    that won't resolve from Drive/Canvas. Surveys are linked only (raw dumps)."""
    if not os.path.isdir(REFLECTION_DIR):
        return []
    # name -> {tokens, live doc URL}, from workbook projects + faculty + Impact Tracker
    cand = {}

    def add(name, reflink=None):
        if not name:
            return
        info = cand.get(name)
        if info is None:
            cand[name] = {"toks": set(norm(name).split()), "reflink": reflink or None}
        elif reflink and not info["reflink"]:
            info["reflink"] = reflink

    for p in projects:
        toks = set(norm(p["name"]).split())
        fac = (p.get("asana") or {}).get("faculty") or ""
        toks |= set(norm(fac).split())
        cand[p["name"]] = {"toks": toks, "reflink": None}
    for r in (faculty or {}).get("ratings", []):
        add(r["project"], r.get("reflink"))
    for d in (impact or {}).values():
        add(d.get("name"), d.get("reflink"))
    # Stop only structural / doc-type / temporal noise — NOT content tokens like
    # "ai", "video", or numerals ("ii"), which are exactly what discriminate
    # short project names. Doc-type words are dropped so a reflection never
    # matches a generic Asana task like "Post-Project Evaluation Project".
    STOP = {"the", "and", "for", "of", "a", "an", "to", "in", "on", "with",
            "project", "report", "reports", "survey", "surveys", "response",
            "responses", "feedback", "results", "result", "evaluation",
            "reflection", "retrospective", "summary", "qualtrics", "link", "links",
            "summer", "sprint", "sprints", "2021", "2023", "2024", "2025"}

    def match(title):
        t = set(norm(title).split()) - STOP
        if len(t) < 2:
            return None
        # Score = (#shared tokens, owns a live reflink, fraction of the candidate's
        # own tokens covered). The reflink term breaks overlap ties toward the
        # project that actually has a reflection doc (e.g. the faculty redesign,
        # not a same-named intern project); frac favours the tightest match.
        best, bscore = None, (1, 0, 0.0)
        for name, info in cand.items():
            ctoks = info["toks"] - STOP
            inter = ctoks & t
            if len(inter) < 2:
                continue
            score = (len(inter), 1 if info.get("reflink") else 0,
                     len(inter) / max(1, len(ctoks)))
            if score > bscore:
                best, bscore = name, score
        return best

    # Google Drive view URLs for the Reflection/ PDFs, so links open from
    # Drive/Canvas (a relative ../Reflection/ path can't). Match by exact
    # filename, then a normalized fallback (curly apostrophes etc.).
    drive_links = load_json(DRIVE_LINKS_FILE, {})
    if not isinstance(drive_links, dict):
        drive_links = {}
    drive_links = {k: v for k, v in drive_links.items() if isinstance(v, str) and not k.startswith("_")}
    drive_norm = {norm(k): v for k, v in drive_links.items()}

    import urllib.parse
    out = []
    for fn in sorted(os.listdir(REFLECTION_DIR)):
        if not fn.lower().endswith(".pdf"):
            continue
        low = fn.lower()
        base = re.sub(r"\.pdf$", "", fn, flags=re.I).strip().lower()
        # a doc whose name ends in "survey" (e.g. "Calc I Reflection Report
        # Survey") IS a survey form, even though it contains "reflection report".
        is_survey = base.endswith("survey") or (
            ("survey" in low or "response" in low) and "reflection report" not in low
            and "evaluation report" not in low and "retrospective" not in low)
        label = re.sub(r"\.pdf$", "", fn, flags=re.I)
        label = re.sub(r"(?i)\s*-?\s*google docs.*$", "", label)
        label = re.sub(r"(?i)\.docx.*$", "", label).strip(" -_")
        mname = match(label)
        # doc_url = the project's source reflection-report Google Doc (Asana),
        # only for actual reports (a survey is a different artifact);
        # drive_url = the Drive copy of THIS PDF (works for reports + surveys).
        rec = {"file": fn, "rel": REFLECTION_REL + urllib.parse.quote(fn),
               "drive_url": drive_links.get(fn) or drive_norm.get(norm(fn)),
               "doc_url": (cand.get(mname, {}).get("reflink") if (mname and not is_survey) else None),
               "type": "survey" if is_survey else "report",
               "label": label, "project": mname}
        if not is_survey:
            txt = _pdf_text(os.path.join(REFLECTION_DIR, fn))
            summ, take = _report_sections(txt)
            cr = re.search(r"completion rate[:\s]*([0-9]+%[^\n]{0,30})", _clean(txt), re.I)
            rec["summary"] = summ
            rec["takeaways"] = take
            rec["completion_rate"] = (cr.group(1).strip() if cr else None)
        out.append(rec)
    return out


def load_asana():
    path = os.path.join(ASANA_DIR, "projects.csv")
    if not os.path.exists(path):
        return [], None
    rows = list(csv.DictReader(open(path, encoding="utf-8")))
    snap = datetime.datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d")
    return rows, snap


def asana_record(r):
    g = lambda k: (r.get(k) or "").strip() or None
    return {
        "gid": g("project_gid"), "name": g("project_name"),
        "archived": (r.get("archived") == "True"),
        "status": g("cf::ODL Project Status"), "ndl_status": g("cf::NDL Project Status"),
        "current_phase": g("cf::Current Phase") or g("cf::Progress"),
        "project_type": g("cf::Project Type") or g("cf::Type of Project"),
        "est_size": g("cf::Estimated T-Shirt Size"), "actual_size": g("cf::Actual T-Shirt Size"),
        "impact_tracker": g("cf::Impact Tracker Status"), "post_status": g("cf::Post-Project Status"),
        "owner": g("owner"), "point_person": g("cf::Project Point Person"),
        "faculty": g("cf::Faculty Members"),
        "dept": g("cf::Academic Department") or g("cf::Academic Department / Institute"),
        "start_on": g("start_on"), "due_on": g("due_on"),
        "created_at": (g("created_at") or "")[:10] or None,
        "last_completed": g("last_completed"),
    }


def _pseudo_asana(rec):
    """Asset-spreadsheet / placeholder records that shouldn't match real projects."""
    n = (rec["name"] or "").lower()
    return n.endswith("asset spreadsheet") or rec.get("status") == "Asset Spreadsheet"


def join_asana(projects, asana_rows):
    recs = [asana_record(r) for r in asana_rows]
    by_norm = {}
    for rec in recs:
        if rec["name"] and not _pseudo_asana(rec):
            by_norm.setdefault(norm(rec["name"]), []).append(rec)
    used = set()

    def pick(cands):  # prefer not-yet-used, then non-archived
        for r in sorted(cands, key=lambda r: (r["gid"] in used, r["archived"], r["gid"])):
            if r["gid"] not in used:
                return r
        return None

    # 1) exact normalized match, one-to-one
    for p in projects:
        p["asana"] = None
        cands = by_norm.get(norm(p["name"]))
        if cands:
            r = pick(cands)
            if r:
                p["asana"] = r
                used.add(r["gid"])
    # 2) fuzzy fallback: full token-subset, >=2 shared tokens, one-to-one
    for p in projects:
        if p["asana"]:
            continue
        ptok = set(norm(p["name"]).split())
        if len(ptok) < 2:
            continue
        best, blen = None, 0
        for an, cands in by_norm.items():
            atok = set(an.split())
            inter = ptok & atok
            if len(inter) >= 2 and (inter == ptok or inter == atok) and len(inter) > blen:
                r = pick(cands)
                if r:
                    best, blen = r, len(inter)
        if best:
            p["asana"] = best
            used.add(best["gid"])

    matched = {p["asana"]["gid"] for p in projects if p.get("asana")}
    # keep non-archived records (incl. Complete) so post-project/impact rules see them
    asana_only = [r for r in recs if r["gid"] not in matched
                  and not r["archived"] and not _pseudo_asana(r)]
    return asana_only


# --------------------------------------------------------------------------- #
#  People roster
# --------------------------------------------------------------------------- #
def enrich_projects(projects, now):
    for p in projects:
        a = p.get("asana") or {}
        st = a.get("status")
        last, first = p.get("last_month"), p.get("first_month")
        has_future = bool(last and last >= now)
        if st in ("In Progress", "On Hold"):
            active = True
        elif st == "Complete":
            active = has_future  # Asana says done but the workbook still has future work
        else:
            active = has_future
        p["active"] = active
        if st:
            p["status_display"] = st
        elif not last:
            p["status_display"] = "Unscheduled"
        elif last < now:
            p["status_display"] = "Past"
        elif first and first > now:
            p["status_display"] = "Upcoming"
        else:
            p["status_display"] = "Active (plan)"
        cur = None
        for ph in p["phases"]:
            if ph["month"] <= now:
                cur = ph["label"]
        p["current_phase"] = a.get("current_phase") or cur


def build_people(projects, person_est):
    team_of = {nme: t for t, names in ROSTER_SEED.items() for nme in names}
    counts = {}
    for p in projects:
        for role in p["roles"]:
            per = role["person"]
            if per and not role["generic"]:
                counts.setdefault(per, {})
                counts[per][role["team"]] = counts[per].get(role["team"], 0) + 1
    names = (set(team_of) | set(counts) | set(person_est)) - set(
        n for n in (set(counts) | set(person_est)) if is_generic_person(n))
    people = []
    for per in sorted(names):
        if per in team_of:
            team = team_of[per]
        elif re.search(r"\bintern\b", per.lower()):   # 'Nina - LD Intern' etc.
            team = "Intern"
        elif per in counts and counts[per]:
            team = max(counts[per], key=counts[per].get)
        else:
            team = "Other"
        people.append({"name": per, "team": team, "is_intern": team == "Intern"})
    return people


# --------------------------------------------------------------------------- #
#  Canonical, internally-consistent capacity model
# --------------------------------------------------------------------------- #
def build_capacity(projects, person_est, people, months):
    team_of = {p["name"]: p["team"] for p in people}

    estimated = {s: {} for s in SCOPES}
    for nm, d in person_est.items():
        if is_generic_person(nm):
            continue
        t = team_of.get(nm, "Other")
        for m, v in d.items():
            estimated["Total"][m] = round(estimated["Total"].get(m, 0) + v, 4)
            if t in estimated:
                estimated[t][m] = round(estimated[t].get(m, 0) + v, 4)

    scheduled = {s: {} for s in SCOPES}
    person_sched, unassigned = {}, {t: {} for t in TEAM_ORDER}
    for p in projects:
        for role in p["roles"]:
            t = role["team"]
            for m, v in role["alloc"].items():
                scheduled["Total"][m] = round(scheduled["Total"].get(m, 0) + v, 4)
                if t in scheduled:
                    scheduled[t][m] = round(scheduled[t].get(m, 0) + v, 4)
                if role["generic"]:
                    unassigned[t][m] = round(unassigned[t].get(m, 0) + v, 4)
                else:
                    d = person_sched.setdefault(role["person"], {})
                    d[m] = round(d.get(m, 0) + v, 4)

    # remaining is None for a month with no estimated capacity (e.g. pre-2025),
    # so those months never read as false "over capacity". sorted() -> deterministic output.
    remaining = {s: {m: (round(estimated[s].get(m, 0) - scheduled[s].get(m, 0), 4)
                         if estimated[s].get(m) is not None else None)
                     for m in sorted(set(estimated[s]) | set(scheduled[s]))} for s in SCOPES}
    pct = {s: {m: (round(scheduled[s].get(m, 0) / estimated[s][m], 4)
                   if estimated[s].get(m) else None)
               for m in sorted(set(estimated[s]) | set(scheduled[s]))} for s in SCOPES}

    person = {}
    for nm in sorted(set(person_est) | set(person_sched)):
        if is_generic_person(nm):
            continue
        est = {m: v for m, v in person_est.get(nm, {}).items()}
        sch = {m: v for m, v in person_sched.get(nm, {}).items()}
        rem = {m: round(est.get(m, 0) - sch.get(m, 0), 4) for m in sorted(set(est) | set(sch))}
        person[nm] = {"team": team_of.get(nm, "Other"), "has_capacity": bool(est),
                      "estimated": est, "scheduled": sch, "remaining": rem}
    return {"scopes": SCOPES, "estimated": estimated, "scheduled": scheduled,
            "remaining": remaining, "pct": pct, "unassigned": unassigned,
            "person": person}


# --------------------------------------------------------------------------- #
#  Compute the full data payload (shared by the CLI build and the live server)
# --------------------------------------------------------------------------- #
def compute_data(do_recs=True, write_status=True, verbose=False):
    """Parse the workbook + Asana snapshot and return the complete ``data`` dict
    the dashboard renders. Re-reads its sources every call, so when teammates
    update their time (the workbook is a live Google-Drive sheet; the Asana CSVs
    are refreshed by the pull pipeline) the numbers are current — this is what
    ``serve.py`` calls per request to serve real-time capacity.

      do_recs       also derive/merge recommendations (set False for a fast
                    report-only pass).
      write_status  persist newly-seen recommendation IDs back to statuses.json
                    (the canonical build does; the live server must NOT, so it
                    never clobbers the shared file on every page load).
      verbose       print the build-validation report.
    """
    sheets = load_wb()
    projects, months = parse_projects(sheets["Projects"])
    person_est = parse_person_estimated(sheets["Estimated Capacity (Points)"])
    tshirt = parse_tshirt(sheets["Tshirt Project Types"])
    wb_ref = parse_team_capacity(sheets["Team Capacity"])
    asana_rows, asana_snap = load_asana()
    asana_only = join_asana(projects, asana_rows)
    impact = parse_impact_tracker()
    faculty = parse_faculty_ratings(impact)
    nd = load_json(ND_DEPT_FILE, {})
    if not isinstance(nd, dict):
        nd = {}
    departments = build_departments(impact, nd)
    faculty_years = build_faculty_years(impact, nd)
    people = build_people(projects, person_est)
    cap = build_capacity(projects, person_est, people, months)
    now_ym = datetime.date.today().isoformat()[:7]
    enrich_projects(projects, now_ym)
    reflections = parse_reflections(projects, faculty, impact)
    size_profiles = build_size_profiles(projects, tshirt)

    mset = set(months)
    for s in SCOPES:
        mset |= set(cap["estimated"][s]) | set(cap["scheduled"][s])
    all_months = sorted(mset)

    # drift: consistent recompute vs workbook Team Capacity tab
    drift = []
    for scope in SCOPES:
        ref = wb_ref.get(scope, {})
        for m in sorted(set(cap["scheduled"][scope]) | set(ref.get("scheduled", {}))):
            rc, ex = cap["scheduled"][scope].get(m), ref.get("scheduled", {}).get(m)
            if rc is not None and ex is not None and abs(rc - ex) > 2.0:
                drift.append({"month": m, "scope": scope,
                              "recomputed": round(rc, 2), "workbook": round(ex, 2),
                              "delta": round(rc - ex, 2)})

    if verbose:
        _print_validation(projects, people, all_months, asana_only, asana_snap,
                          tshirt, faculty, reflections, departments, faculty_years,
                          size_profiles, wb_ref, cap, drift, now_ym)

    today = datetime.date.today().isoformat()
    data = {
        "meta": {"generated_at": datetime.datetime.now().isoformat(timespec="seconds"),
                 "today": today,
                 "point_hours": POINT_HOURS, "full_monthly_points": FULL_MONTHLY_POINTS,
                 "months": all_months, "asana_snapshot_date": asana_snap,
                 "source_xlsx": os.path.basename(XLSX),
                 "status_values": recommend.STATUS_VALUES},
        "teams": TEAM_ORDER, "people": people,
        "capacity": cap, "workbook_reference": wb_ref, "drift": drift,
        "projects": projects, "asana_only_projects": asana_only, "tshirt": tshirt,
        "size_profiles": size_profiles,
        "faculty": faculty, "reflections": reflections, "departments": departments,
        "faculty_years": faculty_years,
        # manual weekly inputs for the Director Brief (intake queue, All-Hands
        # shout-outs/wins, curated round-up stories, timesheet-compliance %).
        # The brief's auto-derived sections compute from the data above in the UI.
        "brief_inputs": load_json(BRIEF_INPUTS_FILE, {}),
    }

    if do_recs:
        # ---- recommendations: auto-derive + manual, overlay tracked status ----
        manual = load_json(MANUAL_FILE, [])
        if not isinstance(manual, list):
            if verbose:
                print(f"  WARN: {os.path.basename(MANUAL_FILE)} is not a JSON array; ignoring.")
            manual = []
        statuses = load_json(STATUS_FILE, {})
        if not isinstance(statuses, dict):
            if verbose:
                print(f"  WARN: {os.path.basename(STATUS_FILE)} is not a JSON object; ignoring.")
            statuses = {}
        recs, statuses = recommend.merge(data, manual, statuses, now_ym, today)
        data["recommendations"] = recs
        if write_status:
            with open(STATUS_FILE, "w", encoding="utf-8") as f:
                json.dump(statuses, f, indent=1)
        if verbose:
            by_status = {}
            for r in recs:
                by_status[r["status"]] = by_status.get(r["status"], 0) + 1
            print(f"recommendations: {len(recs)} ({sum(1 for r in recs if r['source']=='auto')} auto, "
                  f"{sum(1 for r in recs if r['source']=='manual')} manual)  status={by_status}")
    else:
        data["recommendations"] = []
    return data


def _print_validation(projects, people, all_months, asana_only, asana_snap, tshirt,
                      faculty, reflections, departments, faculty_years, size_profiles,
                      wb_ref, cap, drift, now):
    print("=" * 72)
    print("ODL PM DASHBOARD — BUILD VALIDATION")
    print("=" * 72)
    print(f"projects: {len(projects)} | roles: {sum(len(p['roles']) for p in projects)}"
          f" | people: {len(people)} | months: {all_months[0]}..{all_months[-1]}")
    print(f"asana matched: {sum(1 for p in projects if p.get('asana'))}/{len(projects)}"
          f" | asana-only active: {len(asana_only)} | snapshot: {asana_snap}")
    print(f"t-shirt sizes: {sorted(tshirt)} | Other-team people: "
          f"{[p['name'] for p in people if p['team']=='Other']}")
    sp = size_profiles.get("sizes", {})
    print("size profiles (planner demand, median): " + (", ".join(
        f"{sz}={sp[sz]['source'][:6]}(n={sp[sz]['n']},{sp[sz]['total_pts']}pt/{sp[sz]['months']}mo)"
        for sz in size_profiles.get("order", [])) or "none"))
    drop = size_profiles.get("dropped_other") or {}
    if drop:
        print(f"   (size profiles dropped {sum(drop.values())} 'Other'-only project(s): {drop})")
    if faculty and faculty.get("fsi"):
        print(f"faculty ratings (Asana Impact Tracker): {faculty['fsi']['n']} projects, "
              f"FSI mean {faculty['fsi']['mean']}/5, NPS mean "
              f"{faculty['nps']['mean'] if faculty.get('nps') else '-'}")
    fy = faculty_years.get("years", [])
    if fy:
        print(f"faculty by year: {len(fy)} years ({faculty_years['min_year']}–{faculty_years['max_year']}), "
              f"{len(faculty_years.get('unknown', []))} projects with no dates · "
              + ", ".join(f"{y['year']}:{y['n']}" for y in fy[:8]))
    nref = len(reflections)
    if nref:
        rep = sum(1 for r in reflections if r["type"] == "report")
        mat = sum(1 for r in reflections if r.get("project"))
        live = sum(1 for r in reflections if r.get("doc_url"))
        print(f"reflection PDFs: {nref} ({rep} reports, {nref-rep} surveys), {mat} matched to a project,"
              f" {live} with a live Google-Doc link")
    dd = departments["departments"]
    by_src = sum(1 for d in dd for p in d["projects"] if p["dept_source"] == "nd.edu")
    print(f"departments (Impact Tracker): {len(dd)} depts grouping "
          f"{sum(d['n'] for d in dd)} projects ({by_src} dept assignments from nd.edu, "
          f"{departments['unknown_count']} projects still without a department); "
          f"nd_departments.json {'loaded' if departments['nd_loaded'] else 'NOT present (run nd-department-enrich)'}")
    # estimated-total exact-match check
    bad = [m for m in all_months
           if wb_ref['Total']['estimated'].get(m) is not None
           and abs(cap['estimated']['Total'].get(m, 0) - wb_ref['Total']['estimated'][m]) > 0.05]
    print(f"estimated Total vs workbook: {'EXACT MATCH all months' if not bad else 'MISMATCH '+str(bad)}")
    print(f"\nCAPACITY @ {now} (recomputed, consistent):")
    for s in SCOPES:
        e = cap['estimated'][s].get(now); sc = cap['scheduled'][s].get(now)
        r = cap['remaining'][s].get(now); pc = cap['pct'][s].get(now)
        print(f"   {s:<7} est={e} sched={sc} remaining={r} alloc={round(pc*100) if pc else '-'}%")
    print(f"\nworkbook drift (|recompute−Team Capacity tab|>2pt): {len(drift)} month/scope cells")
    over = [(nm, m, pp['remaining'][m]) for nm, pp in cap['person'].items()
            for m in pp['remaining'] if pp['remaining'][m] is not None and pp['remaining'][m] < -0.5 and m >= now]
    print(f"individuals over capacity (remaining<−0.5, {now}+): {len(over)} person-months")


# --------------------------------------------------------------------------- #
#  Main
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--report", action="store_true")
    args = ap.parse_args()

    data = compute_data(do_recs=not args.report, write_status=not args.report,
                        verbose=True)
    if args.report:
        return

    out = os.path.join(HERE, "data.json")
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=1, default=str)
    print(f"\nwrote {out}  ({os.path.getsize(out)//1024} KB)")

    # bake into the self-contained index.html
    try:
        import render
        path, n = render.render()
        print(f"wrote {path}  ({n // 1024} KB)")
    except FileNotFoundError:
        print("  (template.html not found yet — run render.py after creating it)")


def load_json(path, default):
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"  WARN: could not read {os.path.basename(path)}: {e}")
    return default


if __name__ == "__main__":
    main()
