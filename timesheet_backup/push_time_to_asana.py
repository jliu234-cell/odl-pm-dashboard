#!/usr/bin/env python3
"""Push manual Time Log hours INTO Asana as time-tracking entries.

The missing direction of timesheet_bridge.py. After people fill hours in the
planning sheet's "Time Log" tab (which exists because Asana's time-entry UI
lapsed 2026-06-24), this feeds those hours back into Asana as real
time_tracking_entries, so Asana can hold one canonical history once the paid
feature returns — nothing to migrate by hand.

WHAT ASANA ALLOWS  (verified against the API, 2026-07)
  * Endpoint:  POST /tasks/<task_gid>/time_tracking_entries
      body: {"duration_minutes": <int>, "entered_on": "YYYY-MM-DD",
             "description": "<who · phase · note>"}
  * created_by is READ-ONLY. Every entry is attributed to the token owner
    (you), NOT the person who did the work — there is no way around this with a
    Personal Access Token. So we preserve the real person (plus phase + note) in
    the entry's `description`, which IS writable and shows in Asana's timesheet
    UI. Run `probe` and it prints the created_by it sees, so this is visible.
  * The endpoint needs time tracking on the domain's plan / add-on. While the
    feature is down, writes return 402 Payment Required — the same signal
    asana_pull.py already treats as "time tracking unavailable" on the read
    side. This script detects a 402 and stops cleanly, changing nothing. The
    `description` field may additionally need the "Timesheets & Budgets" add-on;
    if so, the script automatically falls back to posting entries without it.

SHAPE
  One "Manual time log — <Project>" task per project holds that project's manual
  entries. The Time Log tab records a project + phase + free-text note, never a
  real Asana task GID, so per-project is the honest granularity: hours roll up
  to the right project, and person/phase/note live in each entry's description.

IDEMPOTENT
  Two local ledgers next to this script mean re-runs never double-post:
    asana_timelog_task_map.json  project (lowercased) -> Manual-time-log task gid
    asana_timelog_pushed.json    hashed row key        -> [entry gid, ...] pushed
  The pushed ledger is keyed by a hash and stores only Asana gids — no names.

USAGE
  # 0. one-time capability check — does the write actually work right now?
  #    Self-cleaning: creates one entry, reads it, deletes it. Leaves no trace.
  python3 push_time_to_asana.py probe --task <ANY_TASK_GID>
  #    (or let it make + delete a throwaway task in a project:)
  python3 push_time_to_asana.py probe --project <PROJECT_GID>

  # 1. see exactly what WOULD be pushed — writes nothing (default):
  python3 push_time_to_asana.py push --xlsx "<fresh Time Log export>.xlsx"

  # 2. actually push (add --limit N for a cautious first batch):
  python3 push_time_to_asana.py push --xlsx "<...>.xlsx" --push

Token: ASANA_TOKEN env, or the macOS keychain entry `asana_token`
(security add-generic-password -s asana_token -a "$USER" -w '0/...'),
exactly like asana_pull.py / asana_push.py.

This is a one-time reconciliation tool for the outage-period manual rows: run it
once when time tracking returns. It de-dupes against its own prior runs (the
ledger), NOT against time entries a person may re-log natively in Asana, so push
the sheet OR re-log in Asana for a given stretch, not both.
"""
import argparse
import csv
import datetime
import hashlib
import json
import os
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(HERE))              # the "PM intern" folder
ASANA_DIR = os.environ.get("ODL_ASANA_DIR") or os.path.join(ROOT, "odl_estimator", "data_all")
PROJECTS_CSV = os.path.join(ASANA_DIR, "projects.csv")
DEFAULT_XLSX = os.path.join(HERE, "ODL Time Log (import into planning sheet).xlsx")

TASK_MAP_FILE = os.path.join(HERE, "asana_timelog_task_map.json")
PUSHED_FILE = os.path.join(HERE, "asana_timelog_pushed.json")

API = "https://app.asana.com/api/1.0"

MANUAL_TASK_NOTES = (
    "Time logged in the planning sheet's \"Time Log\" tab during the Asana "
    "time-entry outage (which began 2026-06-24). Each time entry on this task "
    "is one manual row; the person who did the work, the phase, and any note "
    "are in that entry's description (Asana attributes every API-created entry "
    "to the token owner, so don't read the 'created by' as the worker). "
    "Created by push_time_to_asana.py."
)

# Reuse the Time Log tab reader so there is one source of truth for the schema.
sys.path.insert(0, HERE)
from timesheet_bridge import open_log, read_log, key as row_key  # noqa: E402


# --------------------------------------------------------------------------- #
# Asana HTTP layer — raises so callers can handle 402/403 (unlike asana_push's #
# api(), which sys.exits; we need to catch the "feature down" 402 gracefully). #
# --------------------------------------------------------------------------- #
class ApiError(Exception):
    def __init__(self, code, body, method, path):
        self.code, self.body = code, body
        super().__init__(f"Asana API {code} on {method} {path}: {body[:300]}")


_TOKEN = None


def token():
    global _TOKEN
    if _TOKEN:
        return _TOKEN
    tok = os.environ.get("ASANA_TOKEN", "").strip()
    if not tok:
        try:
            tok = subprocess.check_output(
                ["security", "find-generic-password", "-s", "asana_token", "-w"],
                stderr=subprocess.DEVNULL).decode().strip()
        except Exception:
            pass
    if not tok:
        sys.exit("ERROR: set ASANA_TOKEN (export ASANA_TOKEN='0/...') or store it "
                 "in the keychain (security add-generic-password -s asana_token "
                 "-a \"$USER\" -w '0/...').")
    _TOKEN = tok
    return tok


def api(method, path, data=None, params=None):
    url = path if path.startswith("http") else API + path
    if params:
        url += "?" + urllib.parse.urlencode(params)
    body = json.dumps({"data": data}).encode() if data is not None else None
    req = urllib.request.Request(url, data=body, method=method, headers={
        "Authorization": "Bearer " + token(),
        "Accept": "application/json", "Content-Type": "application/json"})
    for attempt in range(1, 7):
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                txt = r.read().decode()
                return json.loads(txt).get("data") if txt else None
        except urllib.error.HTTPError as e:
            code = e.code
            eb = ""
            try:
                eb = e.read().decode()
            except Exception:
                pass
            if code in (429, 500, 502, 503, 504) and attempt < 6:
                ra = e.headers.get("Retry-After")
                time.sleep(float(ra) if ra else min(30, 2 ** attempt))
                continue
            raise ApiError(code, eb, method, path)
        except OSError as e:
            if attempt < 6:
                time.sleep(min(30, 2 ** attempt))
                continue
            raise ApiError(0, str(e), method, path)


class FeatureDown(Exception):
    """Raised when time-tracking writes are unavailable (402)."""


def create_entry(task_gid, minutes, entered_on, description, include_desc):
    """POST one time entry. On a description-specific 402 (add-on required),
    retry once without it and return (entry, desc_dropped=True). A 402 on the
    bare entry means time tracking itself is off -> FeatureDown."""
    payload = {"duration_minutes": int(minutes), "entered_on": entered_on}
    if include_desc and description:
        payload["description"] = description
    try:
        return api("POST", f"/tasks/{task_gid}/time_tracking_entries", payload), False
    except ApiError as e:
        if e.code == 402 and "description" in payload:
            bare = {"duration_minutes": int(minutes), "entered_on": entered_on}
            try:
                return api("POST", f"/tasks/{task_gid}/time_tracking_entries", bare), True
            except ApiError as e2:
                if e2.code == 402:
                    raise FeatureDown(e2.body)
                raise
        if e.code == 402:
            raise FeatureDown(e.body)
        raise


# --------------------------------------------------------------------------- #
# Local state + lookups                                                        #
# --------------------------------------------------------------------------- #
def load_json(path, default):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return default


def save_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, indent=1, ensure_ascii=False)


def project_gid_map():
    """lowercased project name -> gid, preferring non-archived on name clashes."""
    if not os.path.exists(PROJECTS_CSV):
        sys.exit(f"missing {PROJECTS_CSV} — run the estimator pull first.")
    best = {}
    for r in csv.DictReader(open(PROJECTS_CSV, encoding="utf-8")):
        name = (r.get("project_name") or "").strip()
        gid = (r.get("project_gid") or "").strip()
        if not name or not gid:
            continue
        arch = (r.get("archived") or "").strip().lower() in ("true", "1", "yes")
        lc = name.lower()
        if lc not in best or (best[lc][1] and not arch):   # replace an archived hit
            best[lc] = (gid, arch, name)
    return {lc: v[0] for lc, v in best.items()}, {lc: v[2] for lc, v in best.items()}


def keyhash(k):
    return hashlib.sha1(json.dumps(k, ensure_ascii=False).encode()).hexdigest()[:16]


def entry_description(m):
    parts = [str(m.get(f) or "").strip() for f in ("person", "phase", "task")]
    return " · ".join(p for p in parts if p)[:1024]


def manual_rows(xlsx):
    _wb, ws = open_log(xlsx)
    log, _ = read_log(ws)
    out = []
    for m in log:
        if m["source"].strip().lower() == "asana":
            continue                                   # mirror rows stay in the CSV
        if int(round(m["hours"] * 60)) <= 0:
            continue
        out.append(m)
    return out


# --------------------------------------------------------------------------- #
# probe — is writing time even possible right now? (self-cleaning)            #
# --------------------------------------------------------------------------- #
def cmd_probe(args):
    me = api("GET", "/users/me", params={"opt_fields": "name,email"})
    print(f"authenticated as: {me.get('name')} <{me.get('email')}>")

    scratch = None
    task_gid = args.task
    if not task_gid:
        if not args.project:
            sys.exit("probe needs --task <gid> (an existing task) or "
                     "--project <gid> (a throwaway task is created + deleted).")
        res = api("POST", "/tasks", {
            "name": "ODL time-tracking capability probe — safe to delete",
            "projects": [args.project],
            "notes": "Created by push_time_to_asana.py probe; auto-deleted."})
        scratch = task_gid = res["gid"]
        print(f"created throwaway task {task_gid} in project {args.project}")

    today = datetime.date.today().isoformat()
    try:
        entry, desc_dropped = create_entry(
            task_gid, 1, today, "capability probe (safe to delete)", include_desc=True)
    except FeatureDown as e:
        if scratch:
            _try_delete(f"/tasks/{scratch}")
        print("\nRESULT: time-tracking WRITES are UNAVAILABLE (HTTP 402).")
        print("        The paid feature is still down — nothing can be pushed yet.")
        print("        The Time Log tab stays the canonical record until it returns.")
        print(f"        (server said: {str(e)[:160]})")
        return

    gid = entry["gid"]
    back = api("GET", f"/time_tracking_entries/{gid}",
               params={"opt_fields": "duration_minutes,entered_on,description,created_by.name"})
    _try_delete(f"/time_tracking_entries/{gid}")
    if scratch:
        _try_delete(f"/tasks/{scratch}")

    print("\nRESULT: time-tracking WRITES are AVAILABLE. ✅  (test entry deleted)")
    desc_msg = ("NO — needs the Timesheets add-on; entries will push without it"
                if desc_dropped else "yes")
    print(f"        description field accepted: {desc_msg}")
    cb = (back.get('created_by') or {}).get('name') if back else None
    print(f"        entries will be attributed in Asana to: {cb or me.get('name')} "
          f"(the token owner — the real worker is in each entry's description)")
    print("\nNext: python3 push_time_to_asana.py push --xlsx \"<fresh Time Log export>.xlsx\"")


def _try_delete(path):
    try:
        api("DELETE", path)
    except ApiError as e:
        print(f"  warning: cleanup failed for {path} (HTTP {e.code}) — "
              f"delete it by hand in Asana.", file=sys.stderr)


# --------------------------------------------------------------------------- #
# push — the manual rows -> Asana direction                                   #
# --------------------------------------------------------------------------- #
def cmd_push(args):
    rows = manual_rows(args.xlsx)
    name2gid, name2canon = project_gid_map()
    task_map = load_json(TASK_MAP_FILE, {})
    pushed = load_json(PUSHED_FILE, {})

    # Schedule only occurrences not already pushed in a prior run (dups preserved).
    already = {h: len(v) for h, v in pushed.items()}
    seen = {}
    work, unknown = [], {}
    for m in rows:
        proj_lc = m["project"].strip().lower()
        if proj_lc not in name2gid:
            unknown[m["project"].strip()] = unknown.get(m["project"].strip(), 0) + 1
            continue
        h = keyhash(row_key(m["date"], m["person"], m["project"], m["task"], m["hours"]))
        seen[h] = seen.get(h, 0) + 1
        if seen[h] <= already.get(h, 0):
            continue                                   # this occurrence already pushed
        work.append((h, m, proj_lc))

    # Group the work by project for a readable plan.
    by_proj = {}
    for h, m, proj_lc in work:
        by_proj.setdefault(proj_lc, []).append((h, m))

    total_hours = sum(m["hours"] for _h, m, _ in work)
    print(f"{len(rows)} manual row(s) in the tab · {len(work)} new entr(y/ies) to push "
          f"across {len(by_proj)} project(s) · {total_hours:.1f}h"
          f"{' · DRY RUN' if not args.push else ''}")
    if unknown:
        print("\n  skipped — project name not found in projects.csv (fix the name in the "
              "sheet, or refresh the pull), rows:")
        for name, n in sorted(unknown.items()):
            print(f"    {n:>3} × {name}")

    if not work:
        print("\nNothing to push." if not unknown else "\nNothing pushable.")
        return

    if not args.push:
        print("\nWould create/reuse one \"Manual time log — <project>\" task per project "
              "and post these entries (re-run with --push):")
        for proj_lc, items in sorted(by_proj.items()):
            canon = name2canon[proj_lc]
            has = task_map.get(proj_lc, {}).get("gid")
            hrs = sum(m["hours"] for _h, m in items)
            print(f"\n  {canon}  [{'existing task' if has else 'NEW task'}]  "
                  f"— {len(items)} entr(y/ies), {hrs:.1f}h")
            for _h, m in items[:6]:
                print(f"      {m['date']}  {int(round(m['hours']*60)):>4}m  {entry_description(m)[:70]}")
            if len(items) > 6:
                print(f"      … +{len(items) - 6} more")
        print("\nNo changes made.")
        return

    # ---- live push ---------------------------------------------------------- #
    include_desc = True
    created_tasks = posted = 0
    try:
        for proj_lc, items in sorted(by_proj.items()):
            gid = task_map.get(proj_lc, {}).get("gid")
            if not gid:
                res = api("POST", "/tasks", {
                    "name": f"Manual time log — {name2canon[proj_lc]}",
                    "projects": [name2gid[proj_lc]], "notes": MANUAL_TASK_NOTES})
                gid = res["gid"]
                task_map[proj_lc] = {"gid": gid, "name": name2canon[proj_lc]}
                save_json(TASK_MAP_FILE, task_map)
                created_tasks += 1
                print(f"  + task: Manual time log — {name2canon[proj_lc]} ({gid})")

            for h, m in items:
                if args.limit and posted >= args.limit:
                    raise StopIteration
                entry, dropped = create_entry(
                    gid, round(m["hours"] * 60), m["date"],
                    entry_description(m), include_desc)
                if dropped and include_desc:
                    include_desc = False
                    print("  note: description needs the Timesheets add-on — "
                          "pushing entries without it.", file=sys.stderr)
                pushed.setdefault(h, []).append(entry["gid"])
                posted += 1
                save_json(PUSHED_FILE, pushed)         # persist per entry: crash-safe
    except FeatureDown as e:
        save_json(PUSHED_FILE, pushed)
        print(f"\nSTOPPED: time-tracking writes returned 402 mid-run — the feature is "
              f"unavailable. Pushed {posted} before stopping; the ledger is saved so a "
              f"re-run resumes cleanly.\n(server said: {str(e)[:160]})")
        sys.exit(1)
    except StopIteration:
        pass

    save_json(PUSHED_FILE, pushed)
    print(f"\nDone — created {created_tasks} task(s), posted {posted} time entr(y/ies). "
          f"Ledger: {os.path.basename(PUSHED_FILE)}."
          + (f"  (--limit {args.limit} reached; re-run to continue.)"
             if args.limit and posted >= args.limit else ""))


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    pr = sub.add_parser("probe", help="self-cleaning check: can we write time entries right now?")
    pr.add_argument("--task", help="post the throwaway test entry on this existing task gid")
    pr.add_argument("--project", help="instead, create + delete a throwaway task in this project gid")
    pr.set_defaults(func=cmd_probe)

    pu = sub.add_parser("push", help="post the tab's Manual rows into Asana as time entries")
    pu.add_argument("--xlsx", default=DEFAULT_XLSX,
                    help="workbook with the Time Log tab (default: the kit file; "
                         "for real data point at a fresh export of the live Sheet)")
    pu.add_argument("--push", action="store_true", help="actually write (default: dry run)")
    pu.add_argument("--limit", type=int, default=0, help="stop after N entries (cautious first batch)")
    pu.set_defaults(func=cmd_push)

    args = ap.parse_args()
    try:
        args.func(args)
    except ApiError as e:
        sys.exit(f"\nAPI ERROR: {e}")


if __name__ == "__main__":
    main()
