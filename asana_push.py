#!/usr/bin/env python3
"""Push recommendation statuses back to Asana (the "write to Asana" option).

Mirrors each recommendation as a task in a dedicated Asana project (a
"Capacity Recommendations" board you create once), so the team can also see
and act on them in Asana. Idempotent: a local id->task-gid map (asana_task_map.json)
means re-runs UPDATE the same tasks instead of duplicating them.

Status mapping:
  Achieved / Dismissed -> task completed = true
  In Progress / Not Started -> completed = false (status shown in the task notes)

Setup (once):
  1. In Asana, create a project "ODL Capacity Recommendations"; copy its GID
     from the URL (app.asana.com/0/<GID>/list).
  2. export ASANA_TOKEN="0/..."         (My Settings > Apps > Developer apps)
     export ASANA_REC_PROJECT_GID="<GID>"
     (token also read from keychain: `security add-generic-password -s asana_token -a "$USER" -w '0/...'`)

Run:
  python3 asana_push.py            # DRY RUN — prints the plan, writes nothing
  python3 asana_push.py --push     # actually create/update tasks in Asana
"""
import os, sys, json, time, subprocess, argparse
import urllib.parse, urllib.request, urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))
API = "https://app.asana.com/api/1.0"
MAP_FILE = os.path.join(HERE, "asana_task_map.json")


def token():
    tok = os.environ.get("ASANA_TOKEN", "").strip()
    if not tok:
        try:
            tok = subprocess.check_output(
                ["security", "find-generic-password", "-s", "asana_token", "-w"],
                stderr=subprocess.DEVNULL).decode().strip()
        except Exception:
            pass
    if not tok:
        sys.exit("ERROR: set ASANA_TOKEN (export ASANA_TOKEN='0/...') or store it in the keychain "
                 "(security add-generic-password -s asana_token -a \"$USER\" -w '0/...').")
    return tok


def api(method, path, data=None):
    url = path if path.startswith("http") else API + path
    body = json.dumps({"data": data}).encode() if data is not None else None
    req = urllib.request.Request(url, data=body, method=method, headers={
        "Authorization": "Bearer " + token(),
        "Accept": "application/json", "Content-Type": "application/json"})
    for attempt in range(1, 7):
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.loads(r.read().decode()).get("data")
        except urllib.error.HTTPError as e:
            if e.code == 429 and attempt < 6:
                time.sleep(min(60, 2 ** attempt)); continue
            sys.exit(f"Asana API {e.code} on {method} {path}: {e.read().decode()[:300]}")
        except OSError as e:
            if attempt < 6:
                time.sleep(2 ** attempt); continue
            sys.exit(f"network error: {e}")


def load(path, default):
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return default


def rec_task(rec):
    """Build the Asana task fields for a recommendation."""
    st = rec.get("status", "Not Started")
    done = st in ("Achieved", "Dismissed")
    notes = (f"[{st}] · severity: {rec.get('severity','')} · {rec.get('scope_type','')}: {rec.get('scope','')}\n\n"
             f"{rec.get('detail','')}\n\n"
             f"Suggested: {rec.get('suggested_action','')}\n"
             f"Owner: {rec.get('owner','') or '—'}   Target: {rec.get('target_month','') or '—'}\n"
             f"Notes: {rec.get('notes','') or '—'}\n"
             f"Source: {rec.get('source','auto')} · id: {rec.get('id')}"
             + (f"\nEvidence: {rec['evidence_url']}" if rec.get("evidence_url") else ""))
    name = f"{'✓ ' if st=='Achieved' else ''}{rec.get('title','(untitled)')}"
    return {"name": name, "notes": notes, "completed": done}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--push", action="store_true", help="actually write to Asana (default: dry run)")
    args = ap.parse_args()

    data = load(os.path.join(HERE, "data.json"), None)
    if not data:
        sys.exit("data.json not found — run build.py first.")
    recs = data.get("recommendations", [])
    # overlay the latest tracked statuses (the canonical file the UI exports)
    statuses = load(os.path.join(HERE, "statuses.json"), {})
    for r in recs:
        s = statuses.get(r["id"])
        if s:
            for k in ("status", "owner", "target_month", "notes", "evidence_url"):
                if k in s:
                    r[k] = s[k]

    proj = os.environ.get("ASANA_REC_PROJECT_GID", "").strip()
    task_map = load(MAP_FILE, {})

    print(f"{len(recs)} recommendations · {sum(1 for r in recs if r.get('status')=='Achieved')} achieved · "
          f"target project: {proj or '(ASANA_REC_PROJECT_GID not set)'}")
    if not args.push:
        print("\nDRY RUN — would create/update these tasks (re-run with --push):")
        for r in recs:
            t = rec_task(r)
            action = "update" if r["id"] in task_map else "create"
            print(f"  [{action:>6}] {t['completed'] and '✓' or ' '} {r.get('status',''):<12} {t['name'][:60]}")
        print("\nNo changes made.")
        return

    if not proj:
        sys.exit("ERROR: set ASANA_REC_PROJECT_GID to the target Asana project GID.")

    created = updated = 0
    for r in recs:
        fields = rec_task(r)
        gid = task_map.get(r["id"])
        if gid:
            api("PUT", f"/tasks/{gid}", fields); updated += 1
        else:
            fields["projects"] = [proj]
            res = api("POST", "/tasks", fields)
            task_map[r["id"]] = res["gid"]; created += 1
        print(f"  ok: {fields['name'][:60]}")
    with open(MAP_FILE, "w", encoding="utf-8") as f:
        json.dump(task_map, f, indent=1)
    print(f"\nDone — created {created}, updated {updated}. Map saved to {os.path.basename(MAP_FILE)}.")


if __name__ == "__main__":
    main()
