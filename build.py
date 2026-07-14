#!/usr/bin/env python3
"""ODL PM Capacity & Performance Dashboard — build/ETL step.

Reads two live sources and emits a single ``data.json`` that ``render.py`` bakes
into a self-contained ``index.html``:

  1. the "Capacity Allocations" Google Sheet (exported to
     ``capacity_allocations.csv``) — one row per person per month with Capacity
     Hours, Allocated Hours, and the projects they're allocated to. This is the
     capacity model AND the per-project planned-hours budget (see below).
  2. the nightly Asana snapshot (``../odl_estimator/data_all/``) — projects.csv,
     tasks_raw.csv, time_entries.csv, status_updates.csv, task_custom_fields.csv.

The Excel workbook ("ODL Project and Capacity Planning.xlsx") was DROPPED
entirely (director order, 2026-07). Nothing here reads it anymore.

Capacity model (hours-based; the front-end toggles pts/hrs by ×32):
  1 point = 32 hours ≈ 1 productive week.  4 points/person/month = a full load.
  estimated[scope][m] = Σ named-people capacity hours from the Capacity sheet.
  scheduled[scope][m] = Σ allocated hours from the Capacity sheet.
  remaining = estimated − scheduled ;  pct = scheduled / estimated.

The project list = Asana projects.csv (source of truth) UNIONED with the project
names appearing in the Capacity sheet; overhead / non-project boards are excluded
from project-level panels (see OVERHEAD_PATTERNS) but stay in capacity totals.

No invented numbers: every figure traces to a sheet cell or snapshot field.

Run:  python3 build.py            # writes data.json (+ validation report)
      python3 build.py --report   # validation report only, no write
      python3 build.py --public   # also writes the redacted public variant
"""
import os, sys, csv, json, re, datetime, argparse
import recommend

HERE = os.path.dirname(os.path.abspath(__file__))
MANUAL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "recommendations_manual.json")
STATUS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "statuses.json")
ND_DEPT_FILE = os.path.join(HERE, "nd_departments.json")  # ND-website dept enrichment
DRIVE_LINKS_FILE = os.path.join(HERE, "reflection_drive_links.json")  # Drive URLs for Reflection/ PDFs
BRIEF_INPUTS_FILE = os.path.join(HERE, "brief_inputs.json")  # manual weekly inputs for the Director Brief
REFLECTION_KC_FILE = os.path.join(HERE, "reflection_key_considerations.json")  # "key considerations for future projects" pulled from the reflection reports (Drive)
REFLECTION_THEMES_FILE = os.path.join(HERE, "reflection_themes.json")  # curated synthesis of recurring takeaways across the reports
ROOT = os.path.dirname(HERE)
# Sources default to the local Google-Drive layout, but each can be overridden by
# an env var so the cloud refresh (GitHub Actions) can point them at fetched
# copies — see fetch_drive.py and .github/workflows/refresh.yml.
# Capacity numbers + per-project planned hours come from the live "Capacity
# Allocations" Google Sheet (per-person, per-month capacity vs allocated HOURS),
# exported to CSV here. The tab we want is gid 1025292964.
CAPACITY_CSV = os.environ.get("ODL_CAPACITY_CSV") or os.path.join(HERE, "capacity_allocations.csv")
CAPSHEET_URL = os.environ.get("ODL_CAPACITY_URL") or "https://docs.google.com/spreadsheets/d/1YD9b8vLnglbA5pmFO6HsE-bvMq7wluHv1P1wY4FCujw/edit?gid=1025292964"
ASANA_DIR = os.environ.get("ODL_ASANA_DIR") or os.path.join(ROOT, "odl_estimator", "data_all")
# the estimator's calibrated hours quartiles power the Plan tab (archetype ranges)
CALIBRATION_FILE = os.path.join(ASANA_DIR, "derived", "calibration.json")
REFLECTION_DIR = os.environ.get("ODL_REFLECTION_DIR") or os.path.join(ROOT, "Reflection")   # downloaded reflection PDFs
REFLECTION_REL = "../Reflection/"                    # link path relative to index.html
REFLECTION_FOLDER_URL = "https://drive.google.com/drive/folders/1taQ01ykkfwJEcaJ3JIEAv2A6dwTn_0JW"
# Live source URLs for the "source ↗" links in the UI (override via env if files move).
ASANA_PROJECT_BASE = "https://app.asana.com/0/"
ASANA_HOME = "https://app.asana.com/"
# workspace-scoped Asana URL base ("…/1/<workspace_gid>/…"), for project + task deep links
ASANA_WS = os.environ.get("ODL_ASANA_WS") or "https://app.asana.com/1/228221773618853"
ASANA_IMPACT_BOARD = os.environ.get("ODL_IMPACT_BOARD_GID") or "1211592424221769"  # the Impact Tracker board, for per-task links
# the "NDL Project Tracking & Awareness" board — where intake requests become projects
NDL_BOARD_GID = os.environ.get("ODL_NDL_BOARD_GID") or "1207871566050072"
NDL_BOARD_URL = ASANA_WS + "/project/" + NDL_BOARD_GID + "/list/1207871653620573"
# intake sections on that board we surface (edit here to add/remove a queue)
INTAKE_SECTIONS = ["Received Requests - Triage Needed",
                   "Received Requests - Incentives Requests",
                   "Leadership Review or Approval Needed",
                   "Under Review by Unit"]
POINT_HOURS = 32
FULL_MONTHLY_POINTS = 4

# Names that are NOT real projects: overhead, admin, PD/PTO, and non-project Asana
# boards. Excluded from every project-level panel, but their hours still count in
# the capacity totals (they're real committed time). Easy to edit — add a lowercase
# substring/prefix to hide a new overhead board. Matched case-insensitively as an
# exact name, a prefix, or a contained substring.
OVERHEAD_PATTERNS = [
    "z-",                                 # z-Professional Development, z-Out of Office (prefix)
    "time off", "out of office", "vacation",
    "intern/student worker supervising",
    "impact tracker",
    "ndl project tracking",               # the intake/tracking board itself
    "capacity planning",
    "asana clean up",
    "odl project reflection report",
    "proposals and conferences prep",
    "odl media studio sessions",
    "educause lab",
    # generic non-project Asana boards / scaffolding
    "test -", "test-",                    # "Test - XR Working Group Project"
    "project requests & intake", "standard project template", "template -",
    "asset spreadsheet", "consultations", "michael's pm tasks",
    "digital learning task force", "odl project estimations", "2024 projects",
]
OVERHEAD_EXACT = {"test", "teaching", "consultations", "2024 projects"}


def is_overhead(name):
    """True for overhead / non-project boards (kept in capacity totals, hidden from
    the project list). See OVERHEAD_PATTERNS."""
    low = (name or "").strip().lower()
    if not low:
        return True
    if low in OVERHEAD_EXACT:
        return True
    return any(low == p or low.startswith(p) or p in low for p in OVERHEAD_PATTERNS)

# person -> team, seeded from the workbook's per-person sheet names
# (x_LD_*, x_MP_*, x_PM_*, x_Intern_*) which encode each person's team.
ROSTER_SEED = {
    "Design": ["Yi", "Kuangchen", "Bri", "Janet (Temp)", "Janet"],
    "Media":  ["Matthew", "Tim", "Adam", "Adam - Freelance", "Kevin", "Colin",
               "Derrick", "KC", "Naomi"],
    # Lawrence left ODL — intentionally NOT here, so he never appears person-level in
    # the current-state UI. His historical logged hours still count in project totals.
    "PM":     ["Michael", "Annie", "Jordan", "Janyl", "Michael T", "Sonia"],
    "Intern": ["Nina", "Maddie", "Minyoung"],
}
TEAM_ORDER = ["Design", "Media", "PM", "Intern", "Other"]
SCOPES = ["Total", "Design", "Media", "PM", "Intern", "Other"]

# Full names on the "Capacity Allocations" sheet -> team. This is the current
# team (interns and departed members like Lawrence simply aren't on the sheet).
CAPSHEET_TEAM = {
    "Kuangchen Hsu": "Design", "Yi Lu": "Design", "Brianna Stines": "Design",
    "Colin Gallagher": "Media", "KC Frye": "Media", "Kevin DeCloedt": "Media",
    "Matthew Simmons": "Media", "Derrick Patrick": "Media",
    "Annie Conaghan": "PM", "Michael Lerma": "PM",
}
# People to hide from person-level views (leads whose project hours aren't
# tracked). Team capacity totals are left intact. Per director feedback
# 2026-07-14 — KC: "we don't put [hours] for him."
PERSON_HIDE = {"kc frye", "kc"}
def _hidden_person(nm):
    return (nm or "").strip().lower() in PERSON_HIDE
_MONTHS_FULL = {m.lower(): i for i, m in enumerate(
    ["", "January", "February", "March", "April", "May", "June", "July",
     "August", "September", "October", "November", "December"])}

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
        rec = tasks.setdefault(key, {"name": (r.get("task_name") or "").strip(), "gid": key})
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


def _load_hours_log(path, phase_col="canonical_phase"):
    """Read one hours CSV into the compact, string-interned log the dashboard buckets
    into daily/weekly/biweekly/monthly views, grouped by person, project, or phase.
    Each entry is ``[date, author_idx, project_idx, phase_idx, hours]`` — the only
    columns any downstream consumer needs (actuals are matched by project *name*, and
    compliance by *entry_author*). ``phase_col`` is the phase column name, since the
    Asana Timesheet export calls it ``_phase`` rather than ``canonical_phase``.
    Returns {} when the file is missing or has no usable rows (no "source" key —
    the caller stamps that)."""
    if not os.path.exists(path):
        return {}
    people, projects, phases = {}, {}, {}        # value -> stable index
    def idx(d, v):
        v = (v or "").strip() or "—"
        return d.setdefault(v, len(d))
    entries = []
    for r in csv.DictReader(open(path, encoding="utf-8")):
        date = (r.get("entry_date") or "").strip()[:10]
        hrs = as_num(r.get("hours"))
        if len(date) != 10 or not hrs:       # need a real date + non-zero hours
            continue
        hrs = round(hrs, 2)
        entries.append([date, idx(people, r.get("entry_author")),
                        idx(projects, r.get("project_name")),
                        idx(phases, r.get(phase_col)), hrs])
    if not entries:
        return {}
    entries.sort(key=lambda e: e[0])
    inv = lambda d: [k for k, _ in sorted(d.items(), key=lambda kv: kv[1])]
    return {"people": inv(people), "projects": inv(projects), "phases": inv(phases),
            "entries": entries, "date_min": entries[0][0], "date_max": entries[-1][0],
            "manual_count": 0,
            "total_hours": round(sum(e[4] for e in entries), 1)}


def parse_time_entries():
    """Compact, string-interned hours log for the dashboard, from Asana time-tracking.

    Two possible sources, whichever is FRESHER wins (later max entry_date; ties and a
    missing/empty timesheet fall back to time_entries):
      * ``time_entries.csv`` — per-project-page time entries (rebuilt nightly by the
        estimator's asana_pull). The old workbook "Time Log" backup tab was never
        rolled out and is dropped.
      * the Asana sidebar "Timesheet" report export (``ODL_TIMESHEET_CSV`` env override,
        else ``data_all/timesheet.csv``) — the director considers this more accurate
        than the per-project-page entries, so it is preferred once it is at least as
        fresh. Its phase column is ``_phase``; everything else lines up.

    Stamps ``source`` = "asana_timesheet" | "asana_time_entries"; all other fields are
    identical in shape regardless of source. Returns {} when neither source has entries.

    NB: as of writing the last per-page entry is 2026-06-24 and the Timesheet export is
    staler (2026-06-02), so time_entries still wins — recent windows read low, honestly."""
    # Prefer the workspace-wide timesheet (EVERY person's Asana time-tracking, pulled
    # via the workspace endpoint by regen_timesheet.py in the estimator refresh) — the
    # per-task time_entries.csv only captured whoever logged on tasks we fetched, so it
    # missed most people. Fall back to the per-task file when the workspace one is absent.
    te_path = os.path.join(ASANA_DIR, "timesheet_ws.csv")
    if not os.path.exists(te_path):
        te_path = os.path.join(ASANA_DIR, "time_entries.csv")
    te = _load_hours_log(te_path)
    ts_path = os.environ.get("ODL_TIMESHEET_CSV") or os.path.join(ASANA_DIR, "timesheet.csv")
    ts = _load_hours_log(ts_path, phase_col="_phase")
    # prefer the Timesheet export only when strictly fresher; ties / missing -> time_entries.
    if ts and ts.get("date_max", "") > te.get("date_max", ""):
        ts["source"] = "asana_timesheet"
        return ts
    if not te:
        return {}
    te["source"] = "asana_time_entries"
    return te


def parse_faculty_ratings(impact):
    """Per-project faculty satisfaction from the Impact Tracker (tasks with an
    FSI). Fuller perspectives data lives in Qualtrics (access pending)."""
    if not impact:
        return None
    ratings = []
    for d in impact.values():
        if d.get("fsi") is None:
            continue  # only projects with a faculty rating
        ratings.append({"project": d.get("name", ""), "gid": d.get("gid"),
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
            "project": name, "gid": d.get("gid"), "faculty": d.get("faculty"), "type": d.get("type"),
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
        rec = {"project": name, "gid": d.get("gid"), "dept": dept or "—", "dept_source": dsrc,
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
            # auto-extract the report's own "Key considerations for future projects" /
            # "Areas for improvement" bullets, so a NEW report dropped in the folder
            # flows into the Recommendations tab on the next nightly rebuild. The
            # curated reflection_key_considerations.json overrides these where present.
            rec["key_considerations"] = _kc_bullets(txt)
            rec["completion_rate"] = (cr.group(1).strip() if cr else None)
        out.append(rec)
    return out


# headers that open a "key considerations / areas for improvement" section, and the
# ones that end it (the next section) — used to auto-pull those bullets from a report.
_KC_HDR = re.compile(r"key considerations for future|considerations for future"
                     r"|areas? (for|of) improvement|areas? to improve"
                     r"|recommendations? for future", re.I)
_KC_END = re.compile(r"(?i)^(project details|project timeframe|timeframe|participants"
                     r"|date of|survey results|outcomes? summary|next steps?|memorable"
                     r"|what went well|highlights|summary)\b")


def _kc_bullets(txt):
    """Pull the bullet text under a report's 'Key considerations for future projects'
    / 'Areas for improvement' heading. Returns [] when the report has no such section
    (e.g. a report with no retrospective)."""
    lines = [_clean(l) for l in txt.split("\n")]
    out, collecting, buf = [], False, ""

    def flush():
        nonlocal buf
        b = re.sub(r"\s+", " ", re.sub(r"\s*[○◦▪]\s*", "; ", buf)).strip(" ;")
        if len(b) > 12 and not b.lower().endswith(".pdf"):
            out.append(b[:400])
        buf = ""

    for l in lines:
        if not l:
            continue
        head = l.rstrip(":").strip()
        if len(head) <= 70 and _KC_HDR.search(head) and not _BULLET.match(l):
            flush(); collecting = True; continue
        if collecting and len(head) <= 70 and not _BULLET.match(l) and _KC_END.search(head):
            flush(); collecting = False; continue
        if not collecting:
            continue
        if _BULLET.match(l):
            flush(); buf = _BULLET.sub("", l)
        else:
            buf += " " + l
    flush()
    return out[:8]


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
        # real per-project hours straight from Asana (used for the displayed "project
        # size" — NOT the capacity-sheet even-split). Durations like "336h 55m".
        "actual_h": parse_asana_duration(g("cf::Actual time (total)")),
        "est_h": parse_asana_duration(g("cf::Estimated time (total)")),
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


def _is_real_project_rec(rec):
    """An Asana record that represents an actual ODL project (not a container /
    template / test board). We include it in the project list when it carries any
    real project metadata."""
    if _pseudo_asana(rec) or is_overhead(rec.get("name")):
        return False
    return bool(rec.get("status") or rec.get("ndl_status") or rec.get("point_person")
                or rec.get("faculty") or rec.get("est_size") or rec.get("actual_size")
                or rec.get("due_on") or rec.get("post_status"))


def _capsheet_plan(cap_entries):
    """Even-split each (person, month) row's Allocated Hours across the projects it
    lists, and accumulate per project -> planned hours, per-person monthly hours,
    unassigned monthly hours, and the set of months the project is planned in.

    The even split is an ASSUMPTION (the sheet doesn't say how a person's hours
    divide across their several projects that month) — stated verbatim in the UI's
    'How we estimate' methodology. Overhead names are kept here (they still consume
    capacity) but filtered out of the project LIST later."""
    plan = {}
    for e in cap_entries:
        projs = [p for p in e["projects"] if p]
        if not projs:
            continue
        share = (e["alloc_h"] / len(projs)) if e["alloc_h"] else 0.0
        for pr in projs:
            d = plan.setdefault(pr, {"planned": 0.0, "alloc": {}, "unassigned": {},
                                     "months": set()})
            d["months"].add(e["month"])
            if not share:
                continue
            d["planned"] += share
            if e["generic"]:
                d["unassigned"][e["month"]] = round(d["unassigned"].get(e["month"], 0.0) + share, 4)
            else:
                a = d["alloc"].setdefault(e["person"], {"team": e["team"], "by": {}})
                a["by"][e["month"]] = round(a["by"].get(e["month"], 0.0) + share, 4)
    return plan


def build_projects(asana_rows, cap_entries):
    """The project list = Asana projects.csv (real-project records) UNIONED with the
    project names on the Capacity sheet. Each project carries its Asana record (when
    matched), and — reconstructed from the Capacity sheet's even-split allocation —
    per-person monthly hours (as 'roles'), planned hours, and the plan window.
    Overhead / non-project boards are excluded (see OVERHEAD_PATTERNS). Returns
    (projects, asana_only=[])."""
    recs = [asana_record(r) for r in asana_rows]
    real = [r for r in recs if _is_real_project_rec(r)]
    by_norm = {}
    for rec in real:
        by_norm.setdefault(norm(rec["name"]), []).append(rec)
    plan = _capsheet_plan(cap_entries)

    # candidate display names: capsheet projects (non-overhead) first, then any
    # real Asana project not already represented by a capsheet name.
    def match_asana(name, used):
        cands = by_norm.get(norm(name))
        if cands:
            for r in sorted(cands, key=lambda r: (r["gid"] in used, r["archived"], r["gid"])):
                if r["gid"] not in used:
                    return r
        # fuzzy: exact token-subset, >=2 shared tokens
        ptok = set(norm(name).split())
        if len(ptok) < 2:
            return None
        best, blen = None, 0
        for an, cs in by_norm.items():
            atok = set(an.split())
            inter = ptok & atok
            if len(inter) >= 2 and (inter == ptok or inter == atok) and len(inter) > blen:
                for r in sorted(cs, key=lambda r: (r["gid"] in used, r["archived"], r["gid"])):
                    if r["gid"] not in used:
                        best, blen = r, len(inter)
                        break
        return best

    projects, used = [], set()
    order = [pr for pr in plan if not is_overhead(pr)]
    order.sort()
    for pr in order:
        a = match_asana(pr, used)
        if a:
            used.add(a["gid"])
        projects.append(_new_project(pr, a, plan.get(pr)))
    # Asana real projects not tied to a capsheet name -> plan-less projects
    for rec in real:
        if rec["gid"] in used:
            continue
        used.add(rec["gid"])
        projects.append(_new_project(rec["name"], rec, None))
    return projects, []


def _new_project(name, asana, planrec):
    PH = POINT_HOURS
    a = asana or {}
    roles, points_by_month, staffed = [], {}, 0.0
    plan_months = set()
    planned_hours = 0.0
    if planrec:
        planned_hours = round(planrec["planned"], 2)
        plan_months = set(planrec["months"])
        for person, pa in sorted(planrec["alloc"].items()):
            alloc = {m: round(h / PH, 4) for m, h in pa["by"].items() if h}
            if not alloc:
                continue
            roles.append({"role": pa["team"], "team": pa["team"], "person": person,
                          "generic": False, "alloc": alloc})
            for m, v in alloc.items():
                points_by_month[m] = round(points_by_month.get(m, 0.0) + v, 4)
                staffed += v
        # unassigned committed work -> one generic role carrying the parked hours
        if planrec["unassigned"]:
            alloc = {m: round(h / PH, 4) for m, h in planrec["unassigned"].items() if h}
            if alloc:
                roles.append({"role": "Unassigned", "team": "Other", "person": None,
                              "generic": True, "alloc": alloc})
                for m, v in alloc.items():
                    points_by_month[m] = round(points_by_month.get(m, 0.0) + v, 4)
    total_points = round(planned_hours / PH, 3)
    # ---- displayed "project size" = REAL Asana hours, with its source shown
    # (director feedback 2026-07-14). Potential/unmatched projects with no Asana
    # hours are NOT estimated (no capacity-sheet even-split fabricated as a size).
    actual_h, est_h = a.get("actual_h"), a.get("est_h")
    if a.get("gid") and actual_h:
        size_h, size_src = round(actual_h, 1), "logged in Asana (time-tracking)"
    elif a.get("gid") and est_h:
        size_h, size_src = round(est_h, 1), "estimated in Asana"
    elif a.get("gid"):
        size_h, size_src = None, "no hours entered in Asana yet"
    else:
        size_h, size_src = None, "potential project — not scoped in Asana"
    ms = sorted(m for m in plan_months)
    first = ms[0] if ms else ((a.get("start_on") or "")[:7] or None)
    last = ms[-1] if ms else ((a.get("due_on") or "")[:7] or None)
    return {
        "name": name, "asana": a or None,
        "est_size": a.get("est_size"), "actual_size": a.get("actual_size"),
        "size_h": size_h, "size_src": size_src,
        "actual_h": (round(actual_h, 1) if actual_h else None),
        "est_h": (round(est_h, 1) if est_h else None),
        "type": a.get("project_type"),
        "start": (a.get("start_on") or "")[:7] or first,
        "end": (a.get("due_on") or "")[:7] or last,
        "phases": [],   # month-by-month phases lived only in the dropped workbook
        "roles": roles,
        "points_by_month": points_by_month,
        "total_points": total_points,
        "planned_hours": planned_hours if planrec else None,
        "first_month": first, "last_month": last,
        "plan_months": ms,
        "staffed_points": round(staffed, 3),
        "unstaffed_points": round(total_points - staffed, 3),
        "current_phase": a.get("current_phase"),
    }


# --------------------------------------------------------------------------- #
#  Active / status classification
# --------------------------------------------------------------------------- #
def enrich_projects(projects, updates, now, today):
    """active = the project is planned in the Capacity sheet in the current month or
    any later month, OR (it's not archived in Asana AND EITHER it has an Asana status
    update in the last 30 days OR its ODL Project Status is "In Progress"). The
    In-Progress clause catches active projects whose PMs haven't posted a recent
    status update. Documented in a UI tooltip."""
    try:
        cutoff = (datetime.date.fromisoformat(today) - datetime.timedelta(days=30)).isoformat()
    except Exception:
        cutoff = today
    for p in projects:
        a = p.get("asana") or {}
        gid = a.get("gid")
        has_future_plan = any(m >= now for m in p.get("plan_months", []))
        ups = updates.get(gid) if gid else None
        recent_update = bool(ups and ups[0].get("date", "") >= cutoff)
        st = a.get("status")
        in_progress = bool(st and st.strip().lower() == "in progress")
        p["active"] = bool(has_future_plan or ((not a.get("archived")) and (recent_update or in_progress)))
        last, first = p.get("last_month"), p.get("first_month")
        if st:
            p["status_display"] = st
        elif not p.get("plan_months"):
            p["status_display"] = "Unscheduled"
        elif last and last < now:
            p["status_display"] = "Past"
        elif first and first > now:
            p["status_display"] = "Upcoming"
        else:
            p["status_display"] = "Active (plan)"
        p["current_phase"] = a.get("current_phase")


# --------------------------------------------------------------------------- #
#  Capacity Allocations sheet  ->  the live, hours-based capacity model
# --------------------------------------------------------------------------- #
def month_to_ym(s):
    """'June 2026' -> '2026-06'."""
    parts = (s or "").strip().split()
    if len(parts) == 2 and parts[0].lower() in _MONTHS_FULL and parts[1].isdigit():
        return f"{int(parts[1]):04d}-{_MONTHS_FULL[parts[0].lower()]:02d}"
    return None


def is_capsheet_generic(name):
    """The '<team> Unassigned' placeholder rows — committed work with no named owner."""
    return "unassign" in (name or "").strip().lower()


def capsheet_team(name):
    """Team for a Capacity-Allocations row: the reviewed name->team map, the team
    encoded in an 'Unassigned' placeholder, else a first-name fall back to the
    seed roster (so a newly-added teammate still lands on the right team)."""
    n = (name or "").strip()
    if n in CAPSHEET_TEAM:
        return CAPSHEET_TEAM[n]
    low = n.lower()
    if "unassign" in low:
        if "design" in low:
            return "Design"
        if "media" in low:
            return "Media"
        if re.search(r"\bpm\b", low) or "project manager" in low:
            return "PM"
        if "intern" in low:
            return "Intern"
        return "Other"
    first = (n.split() or [n])[0].lower()
    for team, names in ROSTER_SEED.items():
        if any(x.lower().startswith(first) or first.startswith(x.lower()) for x in names):
            return team
    return "Other"


def parse_capacity_sheet(path):
    """Read the exported 'Capacity Allocations' CSV: one row per person per month
    with Capacity Hours, Allocated Hours, and the projects they're allocated to.
    Returns a flat list of {person, team, generic, month, cap_h, alloc_h,
    projects[]} — or None when the CSV isn't present (build then has no capacity)."""
    if not os.path.exists(path):
        return None
    rows = list(csv.reader(open(path, encoding="utf-8")))
    hi = None
    for i, r in enumerate(rows):
        cells = [str(c or "").strip().lower() for c in r]
        if "person" in cells and "capacity hours" in cells:
            hi = i
            break
    if hi is None:
        return None
    hdr = [str(c or "").strip().lower() for c in rows[hi]]
    col = {k: hdr.index(k) for k in ("person", "month", "capacity hours",
                                     "allocated hours", "projects") if k in hdr}
    out = []
    for r in rows[hi + 1:]:
        def cell(k):
            i = col.get(k)
            return r[i] if i is not None and i < len(r) else ""
        person = str(cell("person")).strip()
        m = month_to_ym(str(cell("month")))
        if not person or not m:
            continue
        projs = [p.strip() for p in str(cell("projects")).split(",") if p.strip()]
        out.append({"person": person, "team": capsheet_team(person),
                    "generic": is_capsheet_generic(person), "month": m,
                    "cap_h": as_num(cell("capacity hours")) or 0.0,
                    "alloc_h": as_num(cell("allocated hours")) or 0.0,
                    "projects": projs})
    return out


def build_capacity_from_sheet(entries):
    """Turn the Capacity-Allocations rows into the dashboard's capacity model.

    Numbers are kept in POINTS (= hours / 32) so the existing front-end (which
    toggles pts/hrs by ×32) renders unchanged; the UI now defaults to the hours
    view. A team's *capacity* counts only NAMED people (an 'Unassigned' slot is
    committed work, not a person with capacity); *scheduled* (allocated) counts
    everyone, and the unassigned share is surfaced separately — mirroring the old
    model's semantics so remaining / % allocated read the same way."""
    PH = POINT_HOURS
    estimated = {s: {} for s in SCOPES}
    scheduled = {s: {} for s in SCOPES}
    unassigned = {t: {} for t in TEAM_ORDER}
    per = {}          # named person -> {team, cap{m}, alloc{m}, projects{m}}
    for e in entries:
        m, team = e["month"], e["team"]
        cap_pts, alloc_pts = e["cap_h"] / PH, e["alloc_h"] / PH
        # scheduled (allocated) counts everyone
        scheduled["Total"][m] = round(scheduled["Total"].get(m, 0) + alloc_pts, 4)
        if team in scheduled:
            scheduled[team][m] = round(scheduled[team].get(m, 0) + alloc_pts, 4)
        if e["generic"]:
            unassigned.setdefault(team, {})
            unassigned[team][m] = round(unassigned[team].get(m, 0) + alloc_pts, 4)
            continue
        # capacity counts only named people
        estimated["Total"][m] = round(estimated["Total"].get(m, 0) + cap_pts, 4)
        if team in estimated:
            estimated[team][m] = round(estimated[team].get(m, 0) + cap_pts, 4)
        if _hidden_person(e["person"]):     # keep team totals, drop the person row
            continue
        p = per.setdefault(e["person"], {"team": team, "cap": {}, "alloc": {}, "projects": {}})
        p["cap"][m] = round(p["cap"].get(m, 0) + cap_pts, 4)
        p["alloc"][m] = round(p["alloc"].get(m, 0) + alloc_pts, 4)
        if e["projects"]:
            p["projects"].setdefault(m, [])
            for pr in e["projects"]:
                if pr not in p["projects"][m]:
                    p["projects"][m].append(pr)

    remaining = {s: {m: (round(estimated[s].get(m, 0) - scheduled[s].get(m, 0), 4)
                         if estimated[s].get(m) is not None else None)
                     for m in sorted(set(estimated[s]) | set(scheduled[s]))} for s in SCOPES}
    pct = {s: {m: (round(scheduled[s].get(m, 0) / estimated[s][m], 4)
                   if estimated[s].get(m) else None)
               for m in sorted(set(estimated[s]) | set(scheduled[s]))} for s in SCOPES}
    person = {}
    for nm, p in per.items():
        est, sch = p["cap"], p["alloc"]
        rem = {m: round(est.get(m, 0) - sch.get(m, 0), 4) for m in sorted(set(est) | set(sch))}
        person[nm] = {"team": p["team"], "has_capacity": bool(est),
                      "estimated": est, "scheduled": sch, "remaining": rem,
                      "projects": p["projects"]}
    months = sorted({e["month"] for e in entries})
    return {"scopes": SCOPES, "estimated": estimated, "scheduled": scheduled,
            "remaining": remaining, "pct": pct, "unassigned": unassigned,
            "person": person, "source": "capacity_allocations_sheet"}, months


def build_people_from_sheet(entries):
    """People roster straight from the Capacity-Allocations sheet (named people
    only, in team order) — so the People tab reflects the current team exactly."""
    seen = {}
    for e in entries:
        if e["generic"] or _hidden_person(e["person"]):
            continue
        seen.setdefault(e["person"], e["team"])
    order = {t: i for i, t in enumerate(TEAM_ORDER)}
    people = [{"name": nm, "team": tm, "is_intern": tm == "Intern"}
              for nm, tm in seen.items()]
    people.sort(key=lambda p: (order.get(p["team"], 9), p["name"]))
    return people


# --------------------------------------------------------------------------- #
#  Per-project estimate accuracy:  budget hours (Asana) vs actual (timesheet)
# --------------------------------------------------------------------------- #
def parse_asana_duration(s):
    """Asana time-tracking duration ('4d 6h', '1w 3d', '13h 30m') -> hours.
    Asana's workday = 8h, work-week = 5d = 40h."""
    if not s:
        return None
    hits = re.findall(r"(\d+(?:\.\d+)?)\s*([wdhm])", str(s).strip().lower())
    if not hits:
        return None
    per = {"w": 40.0, "d": 8.0, "h": 1.0, "m": 1 / 60}
    return round(sum(float(n) * per[u] for n, u in hits), 2)


def _est_hours_by_project():
    """Sum each project's task-level 'Estimated time' from the Asana task dump ->
    a per-project BUDGET in hours, keyed by both project_gid and normalized name."""
    path = os.path.join(ASANA_DIR, "tasks_raw.csv")
    by_gid, by_name = {}, {}
    if not os.path.exists(path):
        return by_gid, by_name
    import csv as _csv
    _csv.field_size_limit(10 ** 7)
    for r in _csv.DictReader(open(path, encoding="utf-8")):
        h = parse_asana_duration(r.get("tcf::Estimated time"))
        if not h:
            continue
        gid = (r.get("project_gid") or "").strip()
        nm = (r.get("project_name") or "").strip()
        if gid:
            by_gid[gid] = round(by_gid.get(gid, 0) + h, 2)
        if nm:
            by_name[norm(nm)] = round(by_name.get(norm(nm), 0) + h, 2)
    return by_gid, by_name


# The even-split assumption, stated verbatim in the UI "How we estimate" panel.
ESTIMATE_METHOD = {
    "planned": ("Planned hours (Capacity Allocations) — the PRIMARY budget. For each "
                "person-month row on the Capacity sheet, that month's Allocated Hours "
                "are split EVENLY across the projects listed in that row, then summed "
                "per project across every month of the sheet. The even split is an "
                "assumption: the sheet doesn't record how a person's hours actually "
                "divide across their several projects that month."),
    "task_est": ("Task-level estimate (Asana) — a SECONDARY reference where present: "
                 "the sum of each project's tasks' 'Estimated time' in Asana."),
    "actual": ("Actual = logged timesheet hours (Asana time-tracking, pulled "
               "workspace-wide so every person's entries are captured, not just the "
               "few on tasks we happened to fetch)."),
    "consequences_over": ("Over ~120% of plan → the work is crowding out a person's other "
                          "projects, hiding overtime, or pushing the next project's start "
                          "later."),
    "consequences_under": ("Under ~60% of plan with real activity → the plan was over-scoped, "
                           "or (most likely right now) time is simply under-logged."),
}


def build_project_estimates(projects, hours_log):
    """Per active project: PLANNED hours (Capacity sheet, even-split — the primary
    budget), TASK-EST hours (Asana task 'Estimated time' — secondary), and ACTUAL
    hours (logged timesheet). Flags projects running over ~120% of plan. This is the
    hours-based est-vs-actual, active projects only, with the methodology carried in
    ``method`` so the UI can show 'How we estimate'."""
    est_gid, est_name = _est_hours_by_project()
    # actual hours per project from the timesheet log (Asana time-tracking).
    actual = {}
    if hours_log and hours_log.get("entries"):
        names = hours_log.get("projects", [])
        for e in hours_log["entries"]:
            nm = names[e[2]] if e[2] < len(names) else ""
            actual[norm(nm)] = round(actual.get(norm(nm), 0) + e[4], 2)

    def lookups(p):
        a = p.get("asana") or {}
        keys = [norm(p["name"])]
        if a.get("name"):
            keys.append(norm(a["name"]))
        return a.get("gid"), keys

    OVER = 1.2   # "over plan" threshold (crowding-out / hidden-overtime signal)
    rows = []
    for p in projects:
        gid, keys = lookups(p)
        planned = p.get("planned_hours")
        task_est = est_gid.get(gid) if gid else None
        if task_est is None:
            task_est = next((est_name[k] for k in keys if k in est_name), None)
        act = next((actual[k] for k in keys if k in actual), None)
        # primary budget = planned hours; fall back to the Asana task estimate.
        budget = planned if planned else task_est
        pct = round(act / planned, 4) if (planned and act is not None) else \
              (round(act / task_est, 4) if (task_est and act is not None) else None)
        p["planned_hours"] = planned
        p["task_est_hours"] = task_est
        p["budget_hours"] = budget            # kept for back-compat (Brief attention line)
        p["actual_hours"] = act
        p["budget_pct"] = pct
        p["over_budget"] = bool(budget and act is not None and act > OVER * budget)
        p["over_plan"] = bool(planned and act is not None and act > OVER * planned)
        p["under_plan"] = bool(planned and act is not None and act < 0.6 * planned)
        if p.get("active") and (planned or task_est or act is not None):
            rows.append({"name": p["name"], "gid": gid,
                         "planned_hours": round(planned, 1) if planned else None,
                         "task_est_hours": task_est, "actual_hours": act,
                         "pct": pct, "over": p["over_plan"] or p["over_budget"],
                         "under": p["under_plan"],
                         "has_budget": budget is not None,
                         "has_planned": planned is not None and planned > 0,
                         "status": p.get("status_display"),
                         "point_person": (p.get("asana") or {}).get("point_person")})
    rows.sort(key=lambda r: (not r["over"], -(r["pct"] or 0), -(r["actual_hours"] or 0), r["name"]))
    n_active = sum(1 for p in projects if p.get("active"))
    n_with_actual = sum(1 for r in rows if r["actual_hours"] is not None)
    return {"projects": rows, "n_active": n_active,
            "n_with_budget": sum(1 for r in rows if r["has_budget"]),
            "n_with_planned": sum(1 for r in rows if r["has_planned"]),
            "n_with_actual": n_with_actual,
            "n_both": sum(1 for r in rows if r["has_budget"] and r["actual_hours"] is not None),
            "n_over": sum(1 for r in rows if r["over"]),
            "method": ESTIMATE_METHOD, "over_ratio": OVER,
            "actuals_end": (hours_log or {}).get("date_max"),
            "point_hours": POINT_HOURS}


# --------------------------------------------------------------------------- #
#  Weekly status updates:  scan each project's latest update for trouble
# --------------------------------------------------------------------------- #
def parse_status_updates():
    """Weekly project status updates pulled from Asana (``status_updates.csv``,
    written by ``asana_pull.py`` — the colored On track / At risk / Off track
    narrative PMs post each week). Returns {project_gid: [updates, newest first]}.
    Returns {} until the pull writes the file, so the detector below stays dormant
    (no fake flags) until the weekly updates are actually available."""
    path = os.path.join(ASANA_DIR, "status_updates.csv")
    if not os.path.exists(path):
        return {}
    by_gid = {}
    for r in csv.DictReader(open(path, encoding="utf-8")):
        gid = (r.get("project_gid") or "").strip()
        if not gid:
            continue
        by_gid.setdefault(gid, []).append({
            "date": (r.get("created_at") or "")[:10],
            "type": (r.get("status_type") or "").strip().lower(),
            "title": (r.get("title") or "").strip(),
            "text": (r.get("text") or "").strip(),
            "author": (r.get("author") or "").strip()})
    for gid in by_gid:
        by_gid[gid].sort(key=lambda u: u["date"], reverse=True)
    return by_gid


# risk phrases we scan the weekly update text for — the "something we don't notice"
STATUS_RISK_PATTERNS = [
    (r"\boff[\s-]?track\b", "off track"),
    (r"\bat[\s-]?risk\b", "flagged at risk"),
    (r"\b(delay|delayed|delays|behind schedule|behind|slipping|slipped|pushed back|running late)\b", "timeline slipping"),
    (r"\b(block(ed|er|ers)?|stuck|stalled|on hold|paused|blocker)\b", "blocked / stalled"),
    (r"\b(waiting on|no response|unresponsive|hasn.?t responded|awaiting|chasing)\b", "waiting on someone"),
    (r"\b(scope creep|out of scope|added scope|more than (expected|planned)|kept adding)\b", "scope creep"),
    (r"\b(over budget|over hours|more hours than|running over)\b", "over budget / hours"),
    (r"\b(short[\s-]?staffed|understaffed|out sick|sick leave|no bandwidth|stretched)\b", "staffing / capacity"),
    (r"\b(concern|concerned|worried|problem|issue|struggl|escalat|red flag|frustrat)\b", "concern raised"),
]
# negation so "no blockers", "not delayed", "no concerns" don't trip the scan
_NEG = re.compile(r"(?:\b(?:no|not|without|zero|never)\b|n['’]t)", re.I)


# the exact rules, surfaced in the UI's "Why we flag things" panel so directors
# know how the risk list is built.
RISK_RULES = [
    "Asana weekly status marked At risk / Off track / On hold.",
    "The latest weekly update's text mentions a risk phrase — delay, blocker, "
    "waiting-on, scope creep, over-hours, short-staffed, or a concern — and the "
    "phrase isn't negated (\"no blockers\" doesn't count).",
    "An active project with no weekly status update in more than 21 days (weekly "
    "cadence expected).",
    "Past its Asana due date and not marked Complete.",
    "Actual logged hours exceed ~120% of the planned (Capacity-sheet) hours.",
]


def detect_status_problems(projects, updates, today):
    """Build the 'Projects at risk' list for ACTIVE projects, from four sources:
    the latest weekly Asana status update (colored status + risk phrases + staleness),
    a past-due Asana due date, and actual hours over ~120% of the plan. Attaches
    ``p['status']`` (the status-update-derived object, for the Brief pill) and returns
    the flagged list with plain-English reasons + Asana links (built in the UI)."""
    pats = [(re.compile(p, re.I), lab) for p, lab in STATUS_RISK_PATTERNS]
    try:
        y, m, d = map(int, today.split("-"))
        today_ord = datetime.date(y, m, d).toordinal()
    except Exception:
        today_ord = None
    flagged, n_with = [], 0
    for p in projects:
        p["status"] = None
        a = p.get("asana") or {}
        gid = a.get("gid")
        reasons, level = [], 0
        ups = updates.get(gid) if gid else None
        if ups:
            n_with += 1
            u = ups[0]
            st = u["type"]
            if st == "off_track":
                reasons.append("Asana status: Off track"); level = max(level, 3)
            elif st == "at_risk":
                reasons.append("Asana status: At risk"); level = max(level, 2)
            elif st == "on_hold":
                reasons.append("Asana status: On hold"); level = max(level, 2)
            blob = (u["title"] + " " + u["text"])
            hits = []
            for rx, lab in pats:
                for mm in rx.finditer(blob):
                    if not _NEG.search(blob[max(0, mm.start() - 22):mm.start()]):
                        hits.append(lab)
                        break
            if hits:
                reasons.append("Update mentions: " + ", ".join(hits[:4]))
                level = max(level, 3 if len(hits) >= 2 else 2)
            if p.get("active") and today_ord is not None and len(u["date"]) == 10:
                try:
                    uy, um, ud = map(int, u["date"].split("-"))
                    stale = today_ord - datetime.date(uy, um, ud).toordinal()
                except Exception:
                    stale = None
                if stale is not None and stale > 21:
                    reasons.append(f"No status update in {stale} days (weekly cadence expected)")
                    level = max(level, 1)
            if reasons:
                p["status"] = {"date": u["date"], "type": st, "level": level,
                               "reasons": list(reasons), "snippet": u["text"][:240],
                               "author": u["author"]}
        # ---- non-update reasons we add so the risk list stands on its own ----
        if not p.get("active"):
            continue
        due = a.get("due_on")
        if due and len(due) == 10 and due < today and a.get("status") != "Complete":
            reasons.append(f"Past its Asana due date ({due}) and not marked Complete")
            level = max(level, 3)
        if p.get("over_plan") and p.get("planned_hours") and p.get("actual_hours") is not None:
            reasons.append(f"Over its planned hours — {round(p['actual_hours'])}h logged vs "
                           f"a {round(p['planned_hours'])}h plan")
            level = max(level, 2)
        if reasons:
            u0 = ups[0] if ups else None
            flagged.append({"name": p["name"], "gid": gid, "level": level,
                            "date": (u0 or {}).get("date", ""), "reasons": reasons,
                            "snippet": (u0 or {}).get("text", "")[:240]})
    flagged.sort(key=lambda x: (-x["level"], x["name"]))
    return {"flagged": flagged, "n": len(flagged), "n_projects_with_updates": n_with,
            "available": bool(updates), "rules": RISK_RULES}


# --------------------------------------------------------------------------- #
#  Update signals — what the weekly status updates reveal BETWEEN the lines:
#  who outside ODL the work leans on (cross-department hand-offs directors care
#  about), dated commitments buried in prose, and updates that have gone green
#  so long they've stopped carrying information. The ODL Podcast is the type
#  case: 22 months of weekly "On track", edits hand-off to a non-ODL editor,
#  production dates living only in the update text.
# --------------------------------------------------------------------------- #
# capitalized tokens that are never people in these updates (weekdays, months,
# seasons, tools, ODL vocabulary, sentence-start words). Compared lowercase.
_UPDATE_STOP_NAMES = {
    "monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday",
    "january", "february", "march", "april", "may", "june", "july", "august",
    "september", "october", "november", "december", "jan", "feb", "mar", "apr",
    "jun", "jul", "aug", "sep", "sept", "oct", "nov", "dec",
    "summer", "fall", "spring", "winter", "break", "holiday", "week", "weeks",
    "this", "next", "last", "the", "our", "new", "all", "everyone", "team",
    "step", "steps", "update", "updates", "status", "notes", "note", "review",
    "notre", "dame", "university", "office", "digital", "learning", "online",
    "odl", "ndl", "tlt", "kaneb", "asana", "canvas", "zoom", "panopto", "drive",
    "google", "youtube", "spotify", "apple", "qualtrics", "adobe", "premiere",
    "course", "courses", "module", "modules", "video", "videos", "series",
    "podcast", "episode", "episodes", "trailer", "production", "recording",
    "filming", "editing", "design", "media", "graphics", "launch", "released",
    "ready", "book", "guide", "norton", "here", "registration", "link", "fee",
    "will", "when", "after", "before", "also", "then", "once", "still",
    # weekly-update template boilerplate ("Summary", "Progress", "Next Steps"…)
    "summary", "progress", "risks", "blockers", "blocker", "milestone",
    "milestones", "accomplishments", "highlights", "overview", "recap",
    "agenda", "action", "items", "goals", "tasks", "key", "kickoff",
    "meeting", "meetings", "email", "emails", "follow", "continue",
    "completed", "complete", "done", "pending", "waiting", "hold",
    "need", "needs", "plan", "plans", "planning", "schedule", "scheduled",
    "there", "thanks", "happy", "since", "aware", "first", "day", "today",
    "tomorrow", "everything", "nothing", "great", "good", "with", "upcoming",
    "events", "event",
}
# a lone capitalized first name only counts when it appears in a person-context
# ("with Ted", "Jim was able to…") — bare capitalized words are too noisy.
_NAME_CTX = re.compile(
    r"\b(?:with|from|by|and|to)\s+([A-Z][a-z]{2,})\b"
    r"|\b([A-Z][a-z]{2,})(?:'s)?\s+(?:will|is|was|has|had|can|working|works|found|"
    r"finds|able|got|gets|back|editing|edits|finalizing|finished|delivered|"
    r"delivering|continues?|recording|records|scheduled|sent|shared|asked)\b")
# literal single spaces only — \s+ let template boilerplate glue across newlines
# ("Summary\n\nProgress") and read as a two-word name
_FULLNAME = re.compile(r"\b([A-Z][a-z]+(?: [A-Z][a-z]+)+)\b")
_UPDATE_MONTH_RX = re.compile(
    r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+"
    r"(\d{1,2})(?:st|nd|rd|th)?\b")


def _upcoming_dates(update, t0):
    """Dates inside the latest update's text that land within the next 30 days —
    commitments living only in prose (e.g. 'Next Production: Wednesday, July 15th')."""
    out, blob = [], (update.get("title", "") + " " + update.get("text", ""))
    mon_i = {m: i + 1 for i, m in enumerate(
        ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"])}
    for m in _UPDATE_MONTH_RX.finditer(blob):
        mo, dd = mon_i[m.group(1)[:3]], int(m.group(2))
        yr = t0.year + (1 if mo < t0.month - 6 else 0)   # "Jan 5" seen in Nov = next year
        try:
            dtx = datetime.date(yr, mo, dd)
        except ValueError:
            continue
        if t0 <= dtx <= t0 + datetime.timedelta(days=30):
            frag = re.sub(r"\s+", " ", blob[max(0, m.start() - 60):m.end() + 60]).strip()
            out.append({"date": dtx.isoformat(), "frag": frag})
    return out[:3]


def build_update_signals(projects, updates, today):
    """Scan ACTIVE projects' weekly updates (last 90 days, latest 8) for
    (a) people outside the ODL roster the work leans on — the cross-department
    hand-offs directors want to know about, (b) dated commitments in prose,
    (c) always-green streaks and cadence/owner hygiene gaps. Feeds the Director
    Brief's Top 5. Dormant ({available: False}) until status_updates.csv exists."""
    if not updates:
        return {"available": False, "crossdept": [], "always_green": [], "hygiene": []}
    t0 = datetime.date.fromisoformat(today)
    cutoff = (t0 - datetime.timedelta(days=90)).isoformat()
    roster_first = {nm.split()[0].lower() for nm in CAPSHEET_TEAM}
    for _t, ns in ROSTER_SEED.items():
        roster_first |= {n.split()[0].lower() for n in ns if n}
    cross, green, hygiene = [], [], []
    for p in projects:
        if not p.get("active"):
            continue
        a = p.get("asana") or {}
        gid = a.get("gid")
        ups = updates.get(gid) or []
        if len(ups) < 3:
            continue
        # the project's own faculty partners are expected in its updates — exclude;
        # ditto every word of the project's own title (updates constantly restate
        # it, and title fragments are not people)
        own = set()
        for part in re.split(r"[,;&/]| and ", a.get("faculty") or ""):
            own |= {t.lower() for t in part.strip().split()}
        own |= {t for t in re.split(r"[^A-Za-z]+", (p.get("name") or "").lower()) if t}
        recent = [u for u in ups if u["date"] >= cutoff][:8]
        ext = {}
        for u in recent:
            blob = u["title"] + " " + u["text"]
            fulls = set()
            for mf in _FULLNAME.finditer(blob):
                toks = [w.lower() for w in mf.group(1).split()]
                if (any(w in _UPDATE_STOP_NAMES for w in toks)
                        or toks[0] in roster_first or any(w in own for w in toks)):
                    continue
                fulls.add(mf.group(1))
            first_of_full = {f.split()[0].lower() for f in fulls}
            singles = set()
            for ms in _NAME_CTX.finditer(blob):
                nm1 = ms.group(1) or ms.group(2)
                lw = (nm1 or "").lower()
                if (not nm1 or lw in _UPDATE_STOP_NAMES or lw in roster_first
                        or lw in own or lw in first_of_full):
                    continue
                singles.add(nm1)
            for nm1 in fulls | singles:
                ext[nm1] = ext.get(nm1, 0) + 1
        # keep repeat mentions (any full name counts once; lone first names need 2+)
        ext = {k: v for k, v in ext.items() if v >= 2 or len(k.split()) >= 2}
        upcoming = _upcoming_dates(ups[0], t0)
        if ext:
            top = sorted(ext.items(), key=lambda kv: (-kv[1], kv[0]))[:4]
            cross.append({"gid": gid, "name": p["name"],
                          "externals": [{"name": k, "n": v} for k, v in top],
                          "upcoming": upcoming, "total": sum(ext.values())})
        # ---- hygiene: cadence gaps on a weekly reporter + owner≠updater ----
        issues = []
        ds = []
        for u in ups[:14]:
            try:
                ds.append(datetime.date.fromisoformat(u["date"]))
            except ValueError:
                pass
        gaps = [(ds[i] - ds[i + 1]).days for i in range(len(ds) - 1)]
        med_gap = sorted(gaps)[len(gaps) // 2] if gaps else None
        if med_gap and med_gap <= 10 and gaps:
            big = max(gaps)
            if big >= max(21, 2.5 * med_gap):
                i = gaps.index(big)
                issues.append(f"~weekly cadence, but a {big}-day silent gap "
                              f"({ds[i + 1].isoformat()} → {ds[i].isoformat()})")
        owner = a.get("owner") or a.get("point_person")
        auths = [u["author"] for u in ups[:8] if u.get("author")]
        if owner and auths:
            dom = max(set(auths), key=auths.count)
            if dom and dom.split()[0].lower() != owner.split()[0].lower():
                issues.append(f"updates written by {dom} while the Asana owner is {owner}")
        if issues:
            hygiene.append({"gid": gid, "name": p["name"], "issues": issues})
        # ---- always-green: a streak so long the color carries no signal ----
        streak = 0
        for u in ups:
            if u["type"] == "on_track":
                streak += 1
            else:
                break
        if streak >= 13:                     # ≈ a quarter of weekly greens
            green.append({"gid": gid, "name": p["name"], "streak": streak,
                          "since": ups[streak - 1]["date"], "n": len(ups),
                          "gap_note": next((i for i in issues if "gap" in i), "")})
    cross.sort(key=lambda x: -x["total"])
    green.sort(key=lambda x: -x["streak"])
    # owner≠updater first (a director-level ownership question), then widest gaps
    hygiene.sort(key=lambda h: (-sum("owner" in i for i in h["issues"]), -len(h["issues"])))
    return {"available": True, "as_of": today, "window_days": 90,
            "crossdept": cross[:8], "always_green": green[:5], "hygiene": hygiene[:12]}


# --------------------------------------------------------------------------- #
#  Timesheet compliance:  who logged time last week / the last 4 weeks
# --------------------------------------------------------------------------- #
def _roster_author_map(authors, team):
    """Map each timesheet-author index to a roster person (exact normalized name,
    then first-name). Authors not on the current roster (e.g. departed Lawrence)
    map to nothing, so they never surface person-level."""
    roster_norm = {norm(nm): nm for nm in team}
    roster_first = {}
    for nm in team:
        roster_first.setdefault(norm(nm.split()[0]) if nm.split() else norm(nm), nm)
    a2p = {}
    for ai, a in enumerate(authors):
        key = roster_norm.get(norm(a)) or roster_first.get(norm((a.split() or [a])[0]))
        if key:
            a2p[ai] = key
    return a2p


def build_timesheet_compliance(hours_log, people, today, since="2026-06"):
    """Did each current-team member log any time last week / the last 4 weeks, AND
    how many hours per person per month from ``since`` onward (vs their capacity)?
    Computed from the timesheet against the current roster (Capacity sheet). NOTE:
    the log ends where Asana's entry feature lapsed (2026-06-24), so recent windows
    read low until entries resume — the UI says so."""
    team = [p["name"] for p in people]
    team_of = {p["name"]: p["team"] for p in people}
    authors = (hours_log or {}).get("people", [])
    ent = (hours_log or {}).get("entries", [])
    a2p = _roster_author_map(authors, team)

    logged = {}                        # roster person -> set of dates with any time
    hours_by = {}                      # roster person -> {month: hours}  (>= since)
    months = set()
    for e in ent:
        who = a2p.get(e[1])
        if not who:
            continue
        logged.setdefault(who, set()).add(e[0])
        mo = e[0][:7]
        if mo >= since:
            months.add(mo)
            hb = hours_by.setdefault(who, {})
            hb[mo] = round(hb.get(mo, 0.0) + e[4], 2)
    months = sorted(months)

    # per-person monthly logged hours vs capacity hours (for the People tab + the
    # Brief's per-person hours table). Zeros are rendered honestly.
    rows = []
    for nm in sorted(team, key=lambda n: (["Design", "Media", "PM", "Intern", "Other"].index(team_of.get(n, "Other")) if team_of.get(n, "Other") in ["Design", "Media", "PM", "Intern", "Other"] else 9, n)):
        by = hours_by.get(nm, {})
        rows.append({"name": nm, "team": team_of.get(nm, "Other"),
                     "by": {m: by.get(m, 0.0) for m in months},
                     "total": round(sum(by.values()), 1)})
    person_hours = {"since": since, "months": months, "rows": rows,
                    "point_hours": POINT_HOURS}

    def _ago(days):
        y, m, d = map(int, today.split("-"))
        return (datetime.date(y, m, d) - datetime.timedelta(days=days)).isoformat()

    def window(days):
        start = _ago(days)
        submitted = sorted(nm for nm in team
                           if any(dt >= start for dt in logged.get(nm, ())))
        missing = sorted(nm for nm in team if nm not in submitted)
        return {"start": start, "end": today, "days": days,
                "submitted": submitted, "missing": missing,
                "pct": round(len(submitted) / len(team), 4) if team else None}
    return {"as_of": today, "team": sorted(team),
            "data_max": (hours_log or {}).get("date_max"),
            "last_week": window(7), "last_4_weeks": window(28),
            "person_hours": person_hours}


# --------------------------------------------------------------------------- #
#  Reflection-report themes: synthesized recurring takeaways + progress
# --------------------------------------------------------------------------- #
def _months_before(ym, n):
    try:
        y, m = map(int, ym.split("-")[:2])
    except Exception:
        return ym
    idx = y * 12 + (m - 1) - n
    return f"{idx // 12:04d}-{idx % 12 + 1:02d}"


def build_reflection_considerations(curated, reflections):
    """Merge the curated 'key considerations' (reflection_key_considerations.json,
    the reviewed override) with what we auto-extracted from each report PDF at build
    time. Curated entries win; any report with auto-extracted considerations that
    ISN'T in the curated list is appended, so a NEW report dropped in the Reflection
    folder flows through automatically on the next rebuild. The Drive folder link is
    forced to the canonical 'Project Reflection Reports' folder."""
    curated = curated if isinstance(curated, dict) else {}
    reports = [dict(r) for r in (curated.get("reports") or []) if isinstance(r, dict)]
    have = {norm(r.get("project", "")) for r in reports if r.get("project")}
    for rf in (reflections or []):
        if rf.get("type") != "report":
            continue
        kc = rf.get("key_considerations") or []
        if not kc:
            continue
        proj = rf.get("project") or rf.get("label")
        if not proj or norm(proj) in have:
            continue                     # curated override wins; no duplicates
        reports.append({"project": proj, "report_title": rf.get("label"),
                        "drive_url": rf.get("drive_url") or rf.get("doc_url"),
                        "date": "", "key_considerations": kc, "auto": True})
        have.add(norm(proj))
    out = dict(curated)
    out["reports"] = reports
    out["folder_url"] = REFLECTION_FOLDER_URL   # canonical "All reflection reports" folder
    return out


def build_reflection_themes(rc):
    """Attach the synthesized recurring THEMES (reflection_themes.json) to the
    reflection considerations: resolve each theme's report list against the
    extracted reports (for dates + Drive links), count them, find the most recent,
    and flag whether the lesson is still recurring (raised in a recent report =
    the team keeps re-learning it, i.e. not yet operationalized). This is the
    'understand the takeaways + are we making progress' view, vs a verbatim dump."""
    if not isinstance(rc, dict):
        return rc
    tf = load_json(REFLECTION_THEMES_FILE, {})
    themes_in = tf.get("themes") if isinstance(tf, dict) else None
    reports = rc.get("reports", [])
    if not themes_in or not reports:
        return rc
    by_norm = {norm(r.get("project")): r for r in reports if r.get("project")}
    all_dates = sorted(d for d in (r.get("date") for r in reports) if d)
    newest = all_dates[-1] if all_dates else ""
    recent_floor = _months_before(newest, 5) if newest else ""
    out = []
    for t in themes_in:
        rs = []
        for nm in t.get("reports", []):
            r = by_norm.get(norm(nm))
            if r:
                rs.append({"project": r["project"], "date": r.get("date") or "",
                           "drive_url": r.get("drive_url")})
        if not rs:
            continue
        dates = sorted(d for d in (x["date"] for x in rs) if d)
        latest = dates[-1] if dates else ""
        out.append({"theme": t.get("theme", ""), "what_it_means": t.get("what_it_means", ""),
                    "n": len(rs), "latest": latest, "earliest": dates[0] if dates else "",
                    "recurring": bool(latest and recent_floor and latest >= recent_floor),
                    "reports": sorted(rs, key=lambda x: x["date"], reverse=True)})
    out.sort(key=lambda x: (-x["n"], x["theme"]))
    rc = dict(rc)
    rc["themes"] = out
    rc["themes_recurring"] = sum(1 for t in out if t["recurring"])
    rc["reports_newest"] = newest
    return rc


# --------------------------------------------------------------------------- #
#  Intake queue — live from the Asana "NDL Project Tracking & Awareness" board
# --------------------------------------------------------------------------- #
def build_intake(brief_inputs, today):
    """Intake requests pulled LIVE from the Asana "NDL Project Tracking & Awareness"
    board (gid NDL_BOARD_GID) in tasks_raw.csv — the sections in INTAKE_SECTIONS,
    skipping completed tasks. Each item: name, section, age in days (from created_at),
    assignee, and its Asana task link. Falls back to brief_inputs.json's manual list
    (with a note) when that board isn't in the snapshot yet."""
    path = os.path.join(ASANA_DIR, "tasks_raw.csv")
    items, board_seen = [], False
    if os.path.exists(path):
        import csv as _csv
        _csv.field_size_limit(10 ** 7)
        want = set(INTAKE_SECTIONS)
        for r in _csv.DictReader(open(path, encoding="utf-8")):
            gid = (r.get("project_gid") or "").strip()
            pname = (r.get("project_name") or "").strip()
            if gid != NDL_BOARD_GID and "ndl project tracking" not in pname.lower():
                continue
            board_seen = True
            section = (r.get("section") or "").strip()
            if section not in want:
                continue
            if str(r.get("completed")).strip().lower() in ("true", "1", "yes"):
                continue
            created = (r.get("created_at") or "")[:10]
            age = None
            if len(created) == 10:
                try:
                    age = (datetime.date.fromisoformat(today) -
                           datetime.date.fromisoformat(created)).days
                except Exception:
                    age = None
            tgid = (r.get("task_gid") or "").strip()
            items.append({"name": (r.get("task_name") or "").strip(), "section": section,
                          "age_days": age, "created": created,
                          "assignee": (r.get("assignee") or "").strip() or None,
                          "gid": tgid,
                          "url": (ASANA_WS + "/task/" + tgid) if tgid else None})
        # order by our section priority, then oldest first
        order = {s: i for i, s in enumerate(INTAKE_SECTIONS)}
        items.sort(key=lambda x: (order.get(x["section"], 9), -(x["age_days"] or 0)))
    if board_seen:
        return {"source": "asana", "sections": INTAKE_SECTIONS, "items": items,
                "board_url": NDL_BOARD_URL, "note": ""}
    # fallback — the tracking board isn't in the snapshot yet
    manual = [x for x in ((brief_inputs or {}).get("intake") or []) if x and x.get("name")]
    for x in manual:
        x.setdefault("section", "Received Requests - Triage Needed")
    return {"source": "manual", "sections": INTAKE_SECTIONS, "items": manual,
            "board_url": NDL_BOARD_URL,
            "note": "The NDL Project Tracking & Awareness board isn't in the Asana "
                    "snapshot yet, so this shows the manual brief_inputs.json list. "
                    "It switches to live board tasks automatically once the pull "
                    "includes that board."}


# --------------------------------------------------------------------------- #
#  Plan-a-project model — slot-based intake planning (how many builds the team
#  sustains in flight), book-by lead times against real semester dates, and the
#  capacity runway cut by the academic year — all from calibrated project data.
#  Replaces the old "free hours ÷ median hours = 206 more single videos" check,
#  which answered a question nobody was asking.
# --------------------------------------------------------------------------- #
ARCHETYPE_LABELS = {"full_course": "Full course (build)",
                    "course_redesign": "Course redesign / update",
                    "video_series": "Video series",
                    "single_video": "Single video",
                    "xr_interactive": "XR / immersive"}
# size tiers echo the faculty guide's S/M/L language so both sites talk the same way
ARCHETYPE_TIERS = {"full_course": "Large", "course_redesign": "Medium",
                   "video_series": "Medium", "single_video": "Small",
                   "xr_interactive": "Varies"}

# ND academic calendar 2026–27 (registrar.nd.edu) — the windows that actually shape
# ODL's year. Free hours get prorated into these; the notes carry the seasonal advice
# (ODL is a service studio for faculty: demand spikes at semester starts, filming
# needs faculty on campus, the sprint cohort pre-books part of every summer).
PLAN_WINDOWS = [
    {"key": "runway", "label": "Now → Fall start", "start": None, "end": "2026-08-23",
     "notes": ["Fall-semester asks land now — triage before classes start Aug 24",
               "2026 sprint-cohort builds wrap up over the summer"]},
    {"key": "fall26", "label": "Fall 2026 semester", "start": "2026-08-24", "end": "2026-12-18",
     "notes": ["Classes Aug 24 – Dec 9 · finals Dec 11–17",
               "First ~2 weeks: faculty-support surge — keep quick-turnaround room",
               "Best filming: Sep – early Nov (midterm break Oct 17–25, Thanksgiving Nov 25–29)"]},
    {"key": "winter", "label": "Winter break", "start": "2026-12-19", "end": "2027-01-10",
     "notes": ["Faculty mostly away — filming and reviews pause",
               "Good window for editing, course build and QA backlog"]},
    {"key": "spring27", "label": "Spring 2027 semester", "start": "2027-01-11", "end": "2027-05-07",
     "notes": ["Classes Jan 11 – Apr 28 · finals May 3–7",
               "Sprint applications due early April (2026 cycle: Apr 6)",
               "Midterm break Mar 6–14 · Easter Mar 26–29"]},
    {"key": "summer27", "label": "Summer 2027 — sprint season", "start": "2027-05-08", "end": "2027-07-31",
     "notes": ["Sprint accelerator ≈May 1; awarded faculty build May–Aug ($3,000 stipends)",
               "Commencement May 14–16",
               "Prime window to start Fall-2027 course builds"]},
]

# launch targets faculty actually ask for — "book-by" dates are computed against these
PLAN_TARGETS = [
    {"key": "fall26", "label": "Fall 2026 (Aug 24)", "date": "2026-08-24"},
    {"key": "spring27", "label": "Spring 2027 (Jan 11)", "date": "2027-01-11"},
    {"key": "fall27", "label": "Fall 2027 (≈Aug 23)", "date": "2027-08-23"},
]

SPRINTS_INFO = {
    "url": "https://learning.nd.edu/projects-partnerships/funding-opportunities/#sprints",
    "blurb": "Digital Learning Sprints fund faculty — a $3,000 stipend plus ODL/TLT team "
             "support — for 1–3-month explorations (AI, flipped classrooms, COIL, async "
             "modules, VR). Annual cycle: applications due early April (2026: Apr 6), "
             "one-day accelerator ≈May 1, builds run over the summer — the cohort "
             "pre-books part of ODL's summer capacity every year.",
}
FACULTY_GUIDE_URL = "https://nd-learning.github.io/FacultyOnboardingGuide/"


def _days_in_month(y, m):
    return ((datetime.date(y + (m == 12), (m % 12) + 1, 1)
             - datetime.date(y, m, 1)).days)


def _window_free_hours(cap, start, end):
    """Free (unallocated) hours per team between two dates, prorating each capacity
    month by the share of its days inside the window."""
    s = datetime.date.fromisoformat(start)
    e = datetime.date.fromisoformat(end)
    out = {}
    for team in ("Design", "Media", "PM", "Intern"):
        rem = (cap.get("remaining") or {}).get(team) or {}
        tot, have = 0.0, False
        for m_, pts in rem.items():
            if pts is None:
                continue
            y_, mo_ = int(m_[:4]), int(m_[5:7])
            dim = _days_in_month(y_, mo_)
            lo = max(s, datetime.date(y_, mo_, 1))
            hi = min(e, datetime.date(y_, mo_, dim))
            if lo > hi:
                continue
            tot += pts * POINT_HOURS * (((hi - lo).days + 1) / dim)
            have = True
        if have:
            out[team] = round(tot, 1)
    return out


def _archetype_by_gid():
    """gid -> archetype from the estimator's reviewed classification
    (data_all/derived/archetypes.csv). {} when the snapshot lacks it."""
    path = os.path.join(ASANA_DIR, "derived", "archetypes.csv")
    if not os.path.exists(path):
        return {}
    out = {}
    for r in csv.DictReader(open(path, encoding="utf-8")):
        g, a = (r.get("gid") or "").strip(), (r.get("archetype") or "").strip()
        if g and a:
            out[g] = a
    return out


def _monthly_in_flight(arch_by_gid, updates, today):
    """How many projects of each archetype showed WORK per month — logged time
    (time_entries.csv) or a posted weekly status update — over the last 12 full
    months. This is the team's demonstrated sustained load, the honest basis for
    'how many more can we take' (raw hours÷median wildly overstates it: long
    builds spend few hours per week; attention, not hours, is what runs out)."""
    sig = {}
    path = os.path.join(ASANA_DIR, "time_entries.csv")
    if os.path.exists(path):
        for r in csv.DictReader(open(path, encoding="utf-8")):
            a = arch_by_gid.get((r.get("project_gid") or "").strip())
            d = (r.get("entry_date") or "")[:7]
            if a and len(d) == 7:
                sig.setdefault((a, d), set()).add(r.get("project_gid"))
    for g, ups in (updates or {}).items():
        a = arch_by_gid.get(g)
        if not a:
            continue
        for u in ups:
            d = (u.get("date") or "")[:7]
            if len(d) == 7:
                sig.setdefault((a, d), set()).add(g)
    y, m = int(today[:4]), int(today[5:7])
    months_used = sorted(f"{(y * 12 + m - 1 - k) // 12:04d}-{(y * 12 + m - 1 - k) % 12 + 1:02d}"
                         for k in range(1, 13))
    series = {}
    for (a, d), gs in sig.items():
        if d in months_used:
            series.setdefault(a, {})[d] = len(gs)
    return ({a: [by.get(d, 0) for d in months_used] for a, by in series.items()},
            months_used)


def _window_capacity_coverage(cap, start, end):
    """How many named people have capacity entered in this window's months, vs the
    roster size — so we can flag windows where the capacity sheet isn't fully filled
    in (e.g. 2027) and avoid presenting a misleadingly low 'free hours' as a plan."""
    persons = (cap.get("person") or {})
    roster = len(persons)
    if not roster:
        return 0, 0
    s = datetime.date.fromisoformat(start)
    e = datetime.date.fromisoformat(end)
    def _in(m):
        y, mo = int(m[:4]), int(m[5:7])
        return not (datetime.date(y, mo, _days_in_month(y, mo)) < s
                    or datetime.date(y, mo, 1) > e)
    wmonths = set()
    for p in persons.values():
        wmonths |= {m for m in (p.get("estimated") or {}) if _in(m)}
    best = 0
    for m in wmonths:
        best = max(best, sum(1 for p in persons.values()
                             if (p.get("estimated") or {}).get(m)))
    return best, roster


def build_plan_model(cap, months, now, projects=None, updates=None, today=None):
    """The Plan tab's data: (1) slot room per archetype — typical/peak in-flight
    load from 12 months of history vs what's running now; (2) book-by dates per
    launch target from calibrated start-to-launch spans; (3) the capacity runway
    prorated into academic-year windows; (4) sprint + faculty-guide routing.
    Returns None when calibration.json is absent (tab hides)."""
    cal = load_json(CALIBRATION_FILE, None)
    if not isinstance(cal, dict) or not cal.get("archetype_effort_hours"):
        return None
    today = today or datetime.date.today().isoformat()
    weeks_cal = (cal.get("calendar") or {}).get("span_weeks_by_archetype") or {}

    def _weeks_stats(key):
        w = weeks_cal.get(key) or {}
        p50, p75 = w.get("p50"), w.get("p75")
        vs = sorted(w.get("values") or [])
        if p50 is None and vs:
            p50 = vs[len(vs) // 2]
        if p75 is None and vs:
            p75 = vs[-1]                     # small n: worst-seen stands in for "safe"
        return {"n": w.get("n"), "p50": p50, "p75": p75,
                "min": w.get("min"), "max": w.get("max")}

    arch = []
    for key, d in cal["archetype_effort_hours"].items():
        ph = d.get("production_hours") or {}
        projects_named = sorted((d.get("projects") or {}).keys())
        arch.append({
            "key": key, "label": ARCHETYPE_LABELS.get(key, key.replace("_", " ").title()),
            "tier": ARCHETYPE_TIERS.get(key, ""),
            "n": ph.get("n"), "p25": ph.get("p25"), "p50": ph.get("p50"),
            "p75": ph.get("p75"), "min": ph.get("min"), "max": ph.get("max"),
            "values": ph.get("values"), "note": ph.get("note", ""),
            "weeks": _weeks_stats(key),
            "examples": projects_named[:6], "n_examples": len(projects_named)})
    # XR has calibrated durations but no effort hours yet — show it (weeks-only)
    if "xr_interactive" not in {a["key"] for a in arch} and weeks_cal.get("xr_interactive"):
        arch.append({"key": "xr_interactive", "label": ARCHETYPE_LABELS["xr_interactive"],
                     "tier": ARCHETYPE_TIERS["xr_interactive"], "n": None, "p25": None,
                     "p50": None, "p75": None, "min": None, "max": None, "values": None,
                     "note": "staff hours not yet calibrated", "weeks": _weeks_stats("xr_interactive"),
                     "examples": [], "n_examples": 0})
    arch.sort(key=lambda a: -(a["p50"] or 0))

    # ---- slot room: demonstrated in-flight load vs what's running now ----
    arch_by_gid = _archetype_by_gid()
    series, conc_months = _monthly_in_flight(arch_by_gid, updates or {}, today)
    running, unclassified = {}, []
    for p in (projects or []):
        if not p.get("active"):
            continue
        g = (p.get("asana") or {}).get("gid")
        a = arch_by_gid.get(g) if g else None
        if a:
            running[a] = running.get(a, 0) + 1
        else:
            unclassified.append(p["name"])
    concurrency = {}
    for a in {x["key"] for x in arch}:
        vals = sorted(series.get(a, []))
        med = vals[len(vals) // 2] if vals else None
        peak = vals[-1] if vals else None
        now_n = running.get(a, 0)
        concurrency[a] = {"typical": med, "peak": peak, "now": now_n,
                          "open": (max(0, med - now_n) if med is not None else None)}

    # ---- book-by dates per launch target ----
    lead_rows = []
    for a in arch:
        wk = a.get("weeks") or {}
        if not wk.get("p50"):
            continue
        by = {}
        for t in PLAN_TARGETS:
            tgt = datetime.date.fromisoformat(t["date"])
            start_by = tgt - datetime.timedelta(weeks=wk["p50"])
            safe_by = tgt - datetime.timedelta(weeks=wk.get("p75") or wk["p50"])
            lead_days = (start_by - datetime.date.fromisoformat(today)).days
            by[t["key"]] = {"start_by": start_by.isoformat(), "safe_by": safe_by.isoformat(),
                            "verdict": ("late" if lead_days < 0 else
                                        "tight" if lead_days < 21 else "open")}
        lead_rows.append({"key": a["key"], "label": a["label"], "tier": a["tier"],
                          "weeks_p50": wk["p50"], "weeks_p75": wk.get("p75"), "by": by})

    # ---- capacity runway cut into academic-year windows ----
    windows = []
    for w in PLAN_WINDOWS:
        if w["end"] < today:
            continue
        start = max(w["start"] or today, today)
        free = _window_free_hours(cap, start, w["end"])
        weeks = round(((datetime.date.fromisoformat(w["end"])
                        - datetime.date.fromisoformat(start)).days + 1) / 7, 1)
        cov, roster = _window_capacity_coverage(cap, start, w["end"])
        windows.append({"key": w["key"], "label": w["label"], "start": start,
                        "end": w["end"], "weeks": weeks, "notes": w["notes"],
                        "free": free,
                        "free_total": round(sum(free.values()), 1),
                        "cap_people": cov, "roster_n": roster,
                        "cap_complete": bool(roster and cov >= 0.6 * roster)})

    # current team availability = remaining capacity hours per team, next 6 months
    # (kept for the per-month detail table under the runway cards)
    fut = [m for m in sorted(set(months)) if m >= now][:6]
    PH = POINT_HOURS
    avail = {}
    for t in ("Design", "Media", "PM", "Intern"):
        rem = cap.get("remaining", {}).get(t, {})
        by = {m: round((rem.get(m) or 0) * PH, 1) for m in fut if rem.get(m) is not None}
        if by:
            avail[t] = {"by": by, "total": round(sum(by.values()), 1)}
    return {"source": "estimator calibration.json",
            "provenance": (cal.get("_provenance") or {}).get("source", ""),
            "archetypes": arch, "availability": avail, "avail_months": fut,
            "point_hours": PH,
            "concurrency": concurrency, "concurrency_months": conc_months,
            "unclassified_active": {"n": len(unclassified), "names": sorted(unclassified)[:8]},
            "lead_times": {"targets": PLAN_TARGETS, "rows": lead_rows},
            "windows": windows, "sprints": SPRINTS_INFO, "guide_url": FACULTY_GUIDE_URL,
            "note": "Archetype hour ranges are the estimator's calibrated ODL-staff "
                    "production hours per project type (full lifecycle). Faculty time "
                    "is not included — the estimator has zero faculty hours logged."}


# --------------------------------------------------------------------------- #
#  Compute the full data payload (shared by the CLI build and the live server)
# --------------------------------------------------------------------------- #
def compute_data(do_recs=True, write_status=True, verbose=False):
    """Parse the Capacity Allocations sheet + Asana snapshot and return the complete
    ``data`` dict the dashboard renders. Re-reads its sources every call, so served
    live (serve.py) the numbers are current. The Excel workbook is no longer read.

      do_recs       also derive/merge recommendations (set False for a fast
                    report-only pass).
      write_status  persist newly-seen recommendation IDs back to statuses.json
                    (the canonical build does; the live server must NOT, so it
                    never clobbers the shared file on every page load).
      verbose       print the build-validation report.
    """
    now_ym = datetime.date.today().isoformat()[:7]
    today = datetime.date.today().isoformat()

    asana_rows, asana_snap = load_asana()
    impact = parse_impact_tracker()
    faculty = parse_faculty_ratings(impact)
    nd = load_json(ND_DEPT_FILE, {})
    if not isinstance(nd, dict):
        nd = {}
    departments = build_departments(impact, nd)
    faculty_years = build_faculty_years(impact, nd)

    # ---- capacity + project list from the Capacity Allocations sheet + Asana ----
    cap_entries = parse_capacity_sheet(CAPACITY_CSV) or []
    people = build_people_from_sheet(cap_entries)
    cap, cap_months = build_capacity_from_sheet(cap_entries)
    months = cap_months
    projects, asana_only = build_projects(asana_rows, cap_entries)

    updates = parse_status_updates()
    enrich_projects(projects, updates, now_ym, today)
    reflections = parse_reflections(projects, faculty, impact)

    # actual logged hours (timesheet), per-project planned-vs-actual, compliance
    hours_log = parse_time_entries()
    est_actual = build_project_estimates(projects, hours_log)   # sets p.over_plan etc.
    timesheet = build_timesheet_compliance(hours_log, people, today)
    # unified "projects at risk": weekly status updates + past-due + over-plan
    status_problems = detect_status_problems(projects, updates, today)
    # between-the-lines update signals (cross-dept names, prose dates, green streaks)
    update_signals = build_update_signals(projects, updates, today)
    brief_inputs = load_json(BRIEF_INPUTS_FILE, {})
    intake = build_intake(brief_inputs, today)
    plan_model = build_plan_model(cap, months, now_ym, projects, updates, today)

    mset = set(months)
    for s in SCOPES:
        mset |= set(cap["estimated"][s]) | set(cap["scheduled"][s])
    all_months = sorted(mset)

    if verbose:
        _print_validation(projects, people, all_months, asana_snap,
                          faculty, reflections, departments, faculty_years,
                          cap, now_ym, status_problems, intake, plan_model)
        ea = est_actual
        print(f"\nestimate accuracy (active projects; planned=Capacity-sheet even-split, "
              f"actual=timesheet): {ea['n_with_planned']}/{ea['n_active']} active have a "
              f"planned budget, {ea['n_with_actual']} have logged time, {ea['n_over']} over plan")
        lw, l4 = timesheet["last_week"], timesheet["last_4_weeks"]
        print(f"timesheet compliance (of {len(timesheet['team'])} on the team, log ends "
              f"{timesheet['data_max']}): last week {len(lw['submitted'])} submitted "
              f"({round((lw['pct'] or 0)*100)}%), last 4 weeks {len(l4['submitted'])} "
              f"({round((l4['pct'] or 0)*100)}%)")

    data = {
        "meta": {"generated_at": datetime.datetime.now().isoformat(timespec="seconds"),
                 "today": today,
                 "point_hours": POINT_HOURS, "full_monthly_points": FULL_MONTHLY_POINTS,
                 "months": all_months, "asana_snapshot_date": asana_snap,
                 "sources": {"capacity_sheet": CAPSHEET_URL,
                             "asana_project_base": ASANA_PROJECT_BASE, "asana_home": ASANA_HOME,
                             "asana_ws": ASANA_WS, "impact_board": ASANA_IMPACT_BOARD,
                             "ndl_board": NDL_BOARD_URL,
                             "reflection_folder": REFLECTION_FOLDER_URL},
                 "status_values": recommend.STATUS_VALUES},
        "teams": TEAM_ORDER, "people": people,
        "capacity": cap,
        # projects = Asana projects.csv ∪ Capacity-sheet names (overhead excluded).
        "projects": projects, "asana_only_projects": asana_only,
        # Plan tab: archetype hour ranges (estimator calibration) + team availability.
        "plan_model": plan_model,
        # actual hours logged (Asana time-tracking) for the People-tab timesheet view.
        "hours_log": hours_log,
        # hours-based estimate accuracy (planned = Capacity sheet even-split budget,
        # task-est = Asana, actual = timesheet) + timesheet compliance, for the Brief.
        "est_actual": est_actual, "timesheet": timesheet,
        # unified projects-at-risk: weekly Asana status updates + past-due + over-plan.
        "status_problems": status_problems,
        # what the weekly updates reveal between the lines (cross-dept reliance,
        # prose-only dates, always-green streaks) — powers the Brief's Top 5.
        "update_signals": update_signals,
        # intake queue — live from the Asana NDL tracking board (fallback: brief_inputs).
        "intake": intake,
        "faculty": faculty, "reflections": reflections, "departments": departments,
        "faculty_years": faculty_years,
        # manual weekly inputs for the Director Brief (running notes, wins, round-up).
        "brief_inputs": brief_inputs,
        # reflection reports: verbatim "key considerations" per report (curated JSON
        # overriding auto-extracted PDF bullets, so new reports auto-flow) + a curated
        # synthesis of the recurring THEMES across them (with a progress read).
        "reflection_considerations": build_reflection_themes(
            build_reflection_considerations(load_json(REFLECTION_KC_FILE, {}), reflections)),
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


def _print_validation(projects, people, all_months, asana_snap,
                      faculty, reflections, departments, faculty_years,
                      cap, now, status_problems, intake, plan_model):
    print("=" * 72)
    print("ODL PM DASHBOARD — BUILD VALIDATION")
    print("=" * 72)
    active = [p for p in projects if p.get("active")]
    with_plan = sum(1 for p in projects if p.get("planned_hours"))
    print(f"projects: {len(projects)} (from Asana ∪ Capacity sheet) | active: {len(active)}"
          f" | with planned hours: {with_plan} | people: {len(people)} "
          f"| months: {all_months[0]}..{all_months[-1]}")
    print(f"asana matched: {sum(1 for p in projects if p.get('asana'))}/{len(projects)}"
          f" | snapshot: {asana_snap} | Other-team people: "
          f"{[p['name'] for p in people if p['team']=='Other']}")
    sp = status_problems or {}
    print(f"projects at risk: {sp.get('n', 0)} flagged "
          f"({sp.get('n_projects_with_updates', 0)} projects have a weekly status update; "
          f"updates available: {sp.get('available')})")
    ik = intake or {}
    print(f"intake queue: {len(ik.get('items', []))} items (source: {ik.get('source')})")
    if plan_model:
        print(f"plan model: {len(plan_model['archetypes'])} archetypes from calibration.json, "
              f"availability for {len(plan_model.get('availability', {}))} teams")
    else:
        print("plan model: calibration.json absent — Plan tab hidden")
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
    cmonths = sorted(cap['estimated']['Total'])
    print(f"capacity source: {cap.get('source','?')} — {len(cap['person'])} people, "
          f"{len(cmonths)} months ({cmonths[0] if cmonths else '-'}..{cmonths[-1] if cmonths else '-'})")
    print(f"\nCAPACITY @ {now} (Capacity Allocations sheet; 1 pt = {POINT_HOURS}h):")
    for s in SCOPES:
        e = cap['estimated'][s].get(now); sc = cap['scheduled'][s].get(now)
        r = cap['remaining'][s].get(now); pc = cap['pct'][s].get(now)
        if e is None and sc is None:
            continue
        h = lambda v: ("-" if v is None else f"{round(v*POINT_HOURS)}h")
        print(f"   {s:<7} cap={h(e)} alloc={h(sc)} remaining={h(r)} util={round(pc*100) if pc else '-'}%")
    over = [(nm, m, pp['remaining'][m]) for nm, pp in cap['person'].items()
            for m in pp['remaining'] if pp['remaining'][m] is not None and pp['remaining'][m] < -0.5 and m >= now]
    print(f"individuals over capacity (remaining<−0.5, {now}+): {len(over)} person-months")


# --------------------------------------------------------------------------- #
#  Public (redacted) build — safe to publish on the open web
# --------------------------------------------------------------------------- #
# People whose names are embedded in project TITLES but not in any Asana faculty
# field (co-PIs, testimonial subjects), so the structured scrub can't find them.
# Add new ones here if they show up in a future refresh (the --public build prints
# a leak check). Common project phrases (Maximizing Mendoza, Virtual Borders, …)
# are deliberately NOT here — only real personal names.
PUBLIC_EXTRA_NAMES = {
    "Seth Berry", "Elizabeth Wood", "Jason Reed", "Whitney James",
    "Fr Nate", "Nate", "Vicki", "Nathaniel Myers", "Nathaniel",
    "Ardea Russo", "Andrea Russo", "Roberto",
}


def redact_for_public(data):
    """Mutate `data` so it carries NO personal identifiers — safe for a public
    site. Teammates become "Design A / Media B / …"; faculty names, satisfaction
    scores and verbatim feedback are dropped; per-person timesheet names, the
    intake queue and internal links (Drive/Asana/workbook) are removed; personal
    names embedded in project titles / recs / allocation lists are scrubbed. Team
    & portfolio aggregates, capacity %s, phases and estimation coverage remain."""
    # stable anonymized labels for teammates ("Design A", "Media B", …)
    people = data.get("people", [])
    seen, label = {}, {}
    for p in sorted(people, key=lambda x: (TEAM_ORDER.index(x["team"]) if x["team"] in TEAM_ORDER else 9, x["name"])):
        seen[p["team"]] = seen.get(p["team"], 0) + 1
        label[p["name"]] = f"{p['team']} {chr(64 + seen[p['team']])}"

    # personal names to scrub out of any free text / project titles. Faculty are
    # stored as "Prof. Josephine Sarpong Akosa" but titles say "Josephine Akosa",
    # so strip honorifics and register first+last and last-name variants too.
    HON = re.compile(r"^(prof|professor|dr|mr|mrs|ms|rev|fr)\.?\s+", re.I)
    fac = set()
    def _addnames(s):
        for part in re.split(r"[,;&/]| and ", s or ""):
            q = HON.sub("", part.strip()).strip()
            t = q.split()
            if 2 <= len(t) <= 4 and len(q) <= 40 and not any(c.isdigit() for c in q):
                fac.add(q)                       # full "Josephine Sarpong Akosa"
                fac.add(t[0] + " " + t[-1])      # "Josephine Akosa" (title form)
                # NB: no bare last-name — a lone surname over-scrubs common words
                # ("Support", "Center"), and the name↔score link is already cut by
                # dropping faculty ratings, so a stray surname is low-risk.
    for r in (data.get("faculty") or {}).get("ratings", []):
        _addnames(r.get("faculty"))
    for p in data.get("projects", []):
        _addnames((p.get("asana") or {}).get("faculty"))
    for a in data.get("asana_only_projects", []):
        _addnames(a.get("faculty"))
    for d in (data.get("departments") or {}).get("departments", []):
        for pr in d.get("projects", []):
            _addnames(pr.get("faculty"))
    fac |= PUBLIC_EXTRA_NAMES                      # title-embedded names not in Asana

    # workbook role rows / freetext use FIRST names ("Annie", "Michael"); map those.
    first_label = {}
    for full, lbl in label.items():
        first_label.setdefault(full.split()[0], lbl)

    def relabel(n):
        return label.get(n) or first_label.get(n) or "a teammate"

    # EVERY teammate name that appears anywhere — the sheet, the seed roster
    # (incl. departed people like Lawrence), workbook role rows, and the logged-
    # hours authors — so the freetext scrub below catches all of them.
    teammate = set(label) | set(first_label)
    for _team, _ns in ROSTER_SEED.items():
        teammate |= set(_ns)
    for _p in data.get("projects", []):
        for _r in _p.get("roles", []):
            if _r.get("person") and not _r.get("generic"):   # skip "Media"/"Design" placeholders
                teammate.add(_r["person"])
    for _n in (data.get("hours_log") or {}).get("people", []):
        teammate.add(_n)
        _t = _n.split()
        if len(_t) >= 2:
            teammate.add(_t[0])

    # never scrub team names / role words (they're not people, and scrubbing them
    # would rewrite the "Media A" labels and the "Media has room" recommendations)
    NOTNAMES = {"design", "media", "pm", "intern", "other", "graphics", "team",
                "staff", "student", "developer", "designer", "tbd", "unassigned", "lead"}
    scrub_map = dict(label)                        # full sheet name -> stable label
    for nm in teammate:
        if nm and nm.lower() not in NOTNAMES:
            scrub_map.setdefault(nm, "a teammate")  # first names / departed -> generic
    for nm in fac:
        scrub_map.setdefault(nm, "[faculty]")
    # word-boundary patterns, longest first (so "Josephine Akosa" beats "Akosa",
    # and \b never clobbers a name that is a substring of a real word)
    _pats = [(re.compile(r"\b" + re.escape(nm) + r"\b"), rep)
             for nm, rep in sorted(scrub_map.items(), key=lambda kv: -len(kv[0])) if nm]

    def scrub(t):
        if not t:
            return t
        for pat, rep in _pats:
            t = pat.sub(rep, t)
        return t

    # people / capacity / logged-hours / timesheet
    for p in people:
        p["name"] = relabel(p["name"])
    cap = data.get("capacity") or {}
    if cap.get("person"):
        cap["person"] = {relabel(k): v for k, v in cap["person"].items()}
        for pp in cap["person"].values():
            pr = pp.get("projects") or {}
            for m in list(pr):
                pr[m] = [scrub(x) for x in pr[m]]
    hl = data.get("hours_log") or {}
    if hl.get("people"):
        hl["people"] = [relabel(n) for n in hl["people"]]
    if hl.get("projects"):
        hl["projects"] = [scrub(n) for n in hl["projects"]]
    ts = data.get("timesheet") or {}
    if ts:
        ts["team"] = [relabel(n) for n in ts.get("team", [])]
        for wk in ("last_week", "last_4_weeks"):
            w = ts.get(wk) or {}
            w["submitted"] = [relabel(n) for n in w.get("submitted", [])]
            w["missing"] = [relabel(n) for n in w.get("missing", [])]

    # projects: scrub titles, blank owner/point-person/faculty
    for p in data.get("projects", []):
        p["name"] = scrub(p.get("name"))
        for role in p.get("roles", []):
            if role.get("person"):
                role["person"] = relabel(role["person"])
        a = p.get("asana")
        if a:
            a["name"] = scrub(a.get("name"))
            for k in ("point_person", "owner", "faculty"):
                if a.get(k):
                    a[k] = "—"
    for a in data.get("asana_only_projects", []):
        a["name"] = scrub(a.get("name"))
        for k in ("point_person", "owner", "faculty"):
            if a.get(k):
                a[k] = "—"
    for r in (data.get("est_actual") or {}).get("projects", []):
        r["name"] = scrub(r.get("name"))
        r["point_person"] = None
    # weekly status-update snippets can quote external partners/SMEs not in our
    # name set — drop the raw text (keep the generic flag reasons) for the public build
    for p in data.get("projects", []):
        if p.get("status"):
            p["status"].pop("snippet", None)
            p["status"].pop("author", None)
    for f in (data.get("status_problems") or {}).get("flagged", []):
        f.pop("snippet", None)

    # drop faculty, departments, reflections entirely (names + scores + links)
    data["faculty"] = {"source": None, "scale": 5, "fsi": None, "nps": None, "distribution": {}, "ratings": []}
    data["departments"] = {"departments": [], "unknown_count": 0, "unknown_projects": [], "nd_loaded": False}
    data["faculty_years"] = {"years": [], "unknown": [], "min_year": None, "max_year": None}
    data["reflections"] = []
    data["reflection_considerations"] = {}

    # recommendations: drop faculty-quoting ones; relabel person scope; scrub text
    recs = []
    for r in data.get("recommendations", []):
        if r.get("category") == "faculty-feedback" or r.get("source") == "manual":
            continue   # faculty-feedback quotes reports; manual recs are internal freetext w/ names
        if r.get("scope_type") == "person" and r.get("scope"):
            r["scope"] = relabel(r["scope"])
        for f in ("title", "detail", "suggested_action", "metric", "scope"):
            if r.get(f):
                r[f] = scrub(r[f])
        r["doc_url"] = None
        recs.append(r)
    data["recommendations"] = recs

    # update signals quote outside collaborators / guests by name — drop for public
    data["update_signals"] = {"available": False, "crossdept": [],
                              "always_green": [], "hygiene": []}

    # strip internal links + manual weekly inputs (intake queue, etc.)
    data.setdefault("meta", {})["sources"] = {}
    data["brief_inputs"] = {}
    # intake carries requestor / assignee names + internal Asana links — drop it.
    data["intake"] = {"source": None, "items": [], "sections": [],
                      "board_url": None, "note": ""}

    # GLOBAL freetext scrub: walk EVERY string in the payload (role labels, phase
    # names, project titles, rec text, logged-hours dimensions…) so no personal
    # name survives anywhere — the catch-all behind the field-specific handling above.
    def _walk(o):
        if isinstance(o, dict):
            return {k: _walk(v) for k, v in o.items()}
        if isinstance(o, list):
            return [_walk(v) for v in o]
        return scrub(o) if isinstance(o, str) else o
    scrubbed = _walk(data)
    data.clear()
    data.update(scrubbed)
    data["_public"] = True
    return data


def _public_leak_names(data):
    """Every personal name that must NOT appear in the public build — used by the
    self-check the --public build runs against the rendered HTML. Call BEFORE
    redact_for_public() (it reads the un-redacted names)."""
    HON = re.compile(r"^(prof|professor|dr|mr|mrs|ms|rev|fr)\.?\s+", re.I)
    names = set(PUBLIC_EXTRA_NAMES)
    for _t, ns in ROSTER_SEED.items():
        for n in ns:
            names.add(n); names.add(n.split()[0])
    for p in data.get("people", []):
        names.add(p["name"]); names.add(p["name"].split()[0])
    for pp in (data.get("capacity") or {}).get("person", {}):
        names.add(pp); names.add(pp.split()[0])
    for n in (data.get("hours_log") or {}).get("people", []):
        names.add(n)
        if len(n.split()) >= 2:
            names.add(n.split()[0])

    def _fac(s):
        for part in re.split(r"[,;&/]| and ", s or ""):
            q = HON.sub("", part.strip()).strip(); t = q.split()
            if 2 <= len(t) <= 4 and not any(c.isdigit() for c in q):
                names.add(q); names.add(t[0] + " " + t[-1])
    for r in (data.get("faculty") or {}).get("ratings", []):
        _fac(r.get("faculty"))
    for p in data.get("projects", []):
        _fac((p.get("asana") or {}).get("faculty"))
    for d in (data.get("departments") or {}).get("departments", []):
        for pr in d.get("projects", []):
            _fac(pr.get("faculty"))
    STOP = {"pm", "the", "and", "xr", "design", "media", "ldp"}
    return {n for n in names if len(n) >= 3 and n.lower() not in STOP}


# --------------------------------------------------------------------------- #
#  Main
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--report", action="store_true")
    ap.add_argument("--public", action="store_true",
                    help="also write a REDACTED data_public.json + index_public.html "
                         "(no personal names / faculty scores / internal links) safe to publish")
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
        if args.public:
            leak_names = _public_leak_names(data)     # names to scan for (pre-redaction)
            redact_for_public(data)   # mutates the in-memory copy (already serialized above)
            pout = os.path.join(HERE, "data_public.json")
            with open(pout, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=1, default=str)
            ppath, pn = render.render("data_public.json", "index_public.html")
            print(f"wrote {ppath}  ({pn // 1024} KB)  [REDACTED — safe to publish]")
            phtml = open(ppath, encoding="utf-8").read()
            leaks = sorted(n for n in leak_names if re.search(r"\b" + re.escape(n) + r"\b", phtml))
            if leaks:
                print(f"  ⚠ PUBLIC LEAK CHECK FAILED — {len(leaks)} personal name(s) still in "
                      f"index_public.html: {leaks[:20]}\n    → add them to PUBLIC_EXTRA_NAMES (or "
                      f"fix a template placeholder) and rebuild before publishing.")
            else:
                print(f"  ✓ public leak check passed — no personal name in index_public.html "
                      f"({len(leak_names)} checked)")
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
