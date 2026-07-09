# Timesheet backup — keeping time updates flowing while Asana's timesheet is down

**Situation** (verified 2026-07-02): Asana's time-entry (timesheet) feature lapsed —
the last entry anyone could log is dated **2026-06-24**. The *read* API still works
(the nightly pull still succeeds and history is intact); the team just has no way
to **enter** time. Full mechanics are written up in the PM-intern working docs
(`audits/2026-07-02 data-flow-map.md`, kept out of this repo since it's internal
notes rather than code — ask the PM team if you need a copy).

**Design in one sentence**: the team logs time in a new **"Time Log" tab of the
planning Google Sheet they already use** (same people/project/phase names as
Asana, dropdowns included); the dashboard's existing nightly build ingests those
rows automatically; when Asana time tracking returns, the two sources merge
without double-counting and nothing has to be migrated.

Why this shape: the dashboard's CI already exports the live planning Sheet every
night (`fetch_drive.py`) — so a tab in that Sheet is the *only* place manual
entries can live with **zero new infrastructure, credentials, or habits**. And
compliance was already thin (only 4 of ~11 loggers since May), so the backup had
to be *lower* friction than Asana's timer, not just equivalent.

## What's in this folder

| File | What it is |
|---|---|
| `ODL Time Log (import into planning sheet).xlsx` | The kit: 4 tabs ready to import into the live Sheet. **Time Log** (the ledger — Date / Person / Project / Phase / Task / Hours / Source, all dropdowns, pre-seeded with 2026's 177 real Asana entries shown in grey), **Read Me** (team instructions), **Time Summary** (live totals by person × month, by project, "most recent entry" per person), **Lists** (the dropdown vocabularies, pulled from the latest Asana snapshot). |
| `make_time_log_kit.py` | Regenerates the kit from `odl_estimator/data_all/` (refreshes dropdowns + seed). |
| `timesheet_bridge.py` | `sync-from-asana`: mirrors new Asana entries into the Time Log tab (incremental, idempotent, keeps legit duplicates). `export-merged`: Asana CSV + Manual rows → `time_entries_merged.csv` in the canonical schema, deduped — for feeding the estimator's calibration. Both have `--dry-run`. |
| `push_time_to_asana.py` | The **into-Asana** direction (the reverse of the bridge): posts the tab's Manual rows back to Asana as real `time_tracking_entries`, one per row, so Asana holds one canonical history once the paid feature returns. One "Manual time log — <project>" task per project; the worker, phase, and note ride in each entry's **description** (Asana forces `created_by` to the token owner, so the entry author is *you*, not the worker — `probe` prints this). **Dry-run by default**, idempotent via two local ledgers. Run `probe` first to confirm writes are live; run `push` once when time tracking returns. |

Plus one change in the parent folder: **`../build.py`** now parses a "Time Log"
tab when the workbook has one (`parse_manual_time_log` + `parse_time_entries(sheets)`).
Manual rows flow into the same `hours_log` the People-tab charts and estimator
calibration use. A manual row matching an Asana entry on (date, person, project,
hours) is skipped as a duplicate. Without the tab, behavior is byte-identical to
before (verified by full build).

This folder lives inside the `odl-pm-dashboard` repo (private) rather than the
public `odl-estimator` repo, since the kit and bridge scripts carry real
people's names and hours — the same sensitivity level as `../data.json`, which
already lives in this repo.

## Rollout (about 10 minutes)

1. **Add the tabs to the live Sheet**: open the planning Sheet ("ODL Project and
   Capacity Planning") → File → Import → Upload the kit xlsx → **"Insert new
   sheet(s)"**. All four tabs land with dropdowns, formatting, and formulas
   intact. Don't use "Replace spreadsheet".
2. **Sanity-check**: pick the first empty Time Log row, confirm the Person /
   Project / Phase dropdowns offer the Asana names, type a test row, watch Time
   Summary update. Delete the test row.
3. **Ship the dashboard change**: already committed and pushed to `odl-pm-dashboard`
   main alongside this folder — the next nightly build ingests Manual rows
   automatically (CI already fetches the Sheet). Nothing further to deploy.
4. **Tell the team** — suggested note:
   > Asana's timesheet feature is down, so until it's back we're logging time in
   > the planning sheet — new "Time Log" tab, same dropdowns as Asana, one row
   > per chunk of work, under a minute. Rows marked "Asana" in grey are history —
   > don't edit those. Everything still flows to the dashboard nightly. A Friday
   > 5-minute wrap-up is plenty.

## Coordination rhythm (suggested)

- **Team**: log as you go or in a Friday 5-minute wrap-up. One row per chunk.
- **PMs**: the Time Summary tab's "Most recent entry" column shows who's current
  at a glance — a friendly nudge in the Monday meeting beats any compliance
  metric. (This also finally makes the Director Brief's timesheet panel
  computable from real data instead of a hand-typed %, if you want that later.)
- **Weekly (optional)**: `python3 timesheet_bridge.py sync-from-asana --xlsx <fresh export>`
  mirrors any entries that do land in Asana into the sheet. Only useful once
  Asana entries resume; the dashboard doesn't need it either way.

## When Asana time tracking comes back

Switch back to logging in Asana — nothing to migrate. Dedup on both paths
(build.py and the bridge) means overlap doesn't double-count.

To push the outage-period manual rows *into* Asana for one canonical history,
use **`push_time_to_asana.py`** (this is the piece that was a to-do; it's built
now):

```
# 1. Confirm the paid feature is actually back (self-cleaning: creates one
#    time entry, reads it, deletes it — leaves no trace). Needs a task gid to
#    test on, or a project gid to make+delete a throwaway task in:
python3 push_time_to_asana.py probe --task <ANY_TASK_GID>

# 2. Dry-run against a fresh export of the live Sheet — prints exactly what it
#    would create, writes nothing:
python3 push_time_to_asana.py push --xlsx "<fresh Time Log export>.xlsx"

# 3. Push for real (start cautious with --limit, then run the rest):
python3 push_time_to_asana.py push --xlsx "<...>.xlsx" --push --limit 5
python3 push_time_to_asana.py push --xlsx "<...>.xlsx" --push
```

Two things to know going in:

- **Attribution.** Asana attributes every API-created time entry to the token
  owner (`created_by` is read-only), so all pushed entries show *you* as author.
  The real worker + phase + note are preserved in each entry's **description**
  (which shows in Asana's timesheet UI). `probe` prints the author it sees so
  there's no surprise. If per-person authorship in Asana matters more than a
  canonical total, the alternative is to leave history in the Time Log tab (the
  dashboard already reads it) and only *start fresh* in Asana — no push.
- **Feature gating.** The endpoint needs time tracking on the domain's plan, so
  while the feature is down every write returns `402` and the script stops
  cleanly, changing nothing. The `description` field may additionally want the
  "Timesheets & Budgets" add-on; if so the script auto-falls-back to posting
  entries without it. That's why step 1 exists — don't skip it.

Idempotent by design: two local ledgers next to the script
(`asana_timelog_task_map.json`, `asana_timelog_pushed.json`) mean re-runs never
double-post. It de-dupes against its own prior runs, **not** against entries a
person re-logs natively in Asana — so for any given stretch, push the sheet *or*
re-log in Asana, not both.

## Design notes

- **Source column semantics**: "Manual" (or blank) = authored in the sheet →
  ingested by build.py. "Asana" = mirror of an API entry → display-only, never
  re-ingested (the CSV stays canonical).
- **Dedup keys**: exact key (date, person, project, task, hours) inside the
  sheet; loose key (date, person, project, hours) across sources — so re-logging
  the same work under a differently-worded task still collapses to one entry.
- **Names**: Person dropdown uses Asana full names (matches `hours_log`), so
  merged data needs no mapping. The workbook's first-name convention only
  applies to the allocation tabs, which are untouched.
- **Dropdowns are suggestions, not gates** (soft validation): a new project can
  be typed before it exists in Asana; refresh `Lists` later with
  `make_time_log_kit.py` or by hand.
- The estimator's calibration reads `data_all/time_entries.csv` (Asana-only).
  During the outage that's frozen but valid history. To include manual hours in
  a local calibration run: `timesheet_bridge.py export-merged` and point the
  tooling at `time_entries_merged.csv`.
