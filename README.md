# ODL Capacity & Performance Dashboard

An internal, PM-facing website for **Online Digital Learning** to plan and track
multiple projects at once against real team capacity — and to surface concrete
ways to improve the team's workflow and process.

It answers, in one place:

- **Are we over capacity** this month / any month, overall and per team (Design,
  Media, PM, Intern)?
- **Who is over- or under-allocated**, and on what?
- **Which active projects are at risk**, and why? — flagged from status updates,
  past-due dates, or hours over/under plan (the *Projects at risk* panel)
- **Can we take on a new project** starting in a given month without blowing
  past capacity? (the *Plan a project* what-if)
- **How can we improve our process?** — concrete recommendations from capacity
  reviews and faculty reflection reports, with a simple way to mark progress
  (the *Recommendations* tab)

## The capacity model

Hours-based, driven entirely by the **Capacity Allocations** Google Sheet
(per-person, per-month Capacity Hours / Allocated Hours / Projects):

> **1 point = 32 hours ≈ 1 productive week. 4 points/person/month = a full load.**
> (40h/week × ~75% efficiency for meetings/breaks ≈ 32 productive hours.)

Everything is monthly. `remaining = estimated − scheduled`; `allocated % =
scheduled / estimated` — estimated, scheduled, remaining, and % allocated all
come straight from the Capacity Allocations sheet, not a workbook tab. Over
100% = over capacity.

## How it's built

```
capacity_allocations.csv ────────────────┐                          ┌─► index.html  (baked, self-contained,
   (Capacity Allocations Google Sheet)    ├─ build.py ──► data.json ──┤    open from Drive — offline snapshot)
../odl_estimator/data_all/*.csv ─────────┘   (parse +    (render.py   └─► serve.py ──► /api/data  (LIVE: recomputes
   (Asana snapshot)                          recompute)   bakes)                       per request, real-time)
```

The dashboard runs in **two modes from the same code**:

- **Offline snapshot** — open `index.html` directly (Drive/Canvas). It's a
  single self-contained file baked from the last build. Works with no server.
- **Live** — run `python3 serve.py` and open the URL it prints. The page fetches
  `/api/data` on load and recomputes the capacity master view + the
  "how many projects can we take" estimator from the *current* sources every
  time, so when teammates update their time it shows up on **Refresh** — no
  rebuild. If it can't reach the server (e.g. opened as a file) it silently
  falls back to the baked snapshot. The header badge shows which mode you're in
  (🟢 Live · *time* vs 📦 Snapshot · *date*).

- **`build.py`** parses the Capacity Allocations sheet and the Asana snapshot,
  recomputes the capacity model *consistently* from the Capacity sheet, joins
  live project status from Asana, and writes `data.json`. The ETL lives in a
  reusable `compute_data()` that **`serve.py`** also calls to serve live
  numbers. It prints a validation report.
- **`serve.py`** is the local live server (see **Live mode** below).
- **`recommend.py`** auto-derives recommendations and merges in manual ones +
  tracked statuses (stable IDs, so statuses survive a rebuild).
- **`render.py`** bakes `data.json` into **`index.html`** — one file, no
  external assets, works offline. Just open it (or host it on Canvas/Drive).

### Rebuild / refresh

```bash
python3 build.py          # re-parse Capacity sheet + Asana snapshot, regenerate index.html
python3 serve.py          # …or serve live numbers without rebuilding (see below)
```

The Asana snapshot is produced by the existing `../odl_estimator` pipeline
(`asana_pull.py` / nightly `refresh.py`). Re-run that first to pull the latest
Asana state, then `build.py` here to refresh the dashboard — or use the live
server's **Sync Asana** button to pull and recompute in place.

### Keeping GitHub in sync

**Cloud refresh (recommended — no laptop required).** A GitHub Actions workflow
(`.github/workflows/refresh.yml`) rebuilds the dashboard every day entirely on
GitHub's runners and commits the result, so it keeps working after anyone leaves.
It pulls its sources from the cloud:

- the **Capacity Allocations** Google Sheet via `fetch_drive.py` using
  a Google service account;
- the **Asana snapshot** by checking out `data_all` from the `odl-estimator` repo
  (which already self-refreshes nightly in the cloud) — so no separate Asana token;
- the **reflection PDFs** from the Reflection Drive folder (also via the service
  account), so the reflection-grounded recommendations aren't wiped on rebuild.

One-time setup — repo **Settings → Secrets and variables → Actions**:

1. **Create a Google service account** (Google Cloud Console → enable the Drive API
   → make a service account → download its JSON key). Ideally make it under a
   team/ND Google project so it outlives any one person.
2. **Grant the service account read access** to both sources. Easiest: add its
   email as a **Viewer member of the "NDL ODL" shared drive** (where both the
   Capacity Allocations Sheet and the `…/ODL PM Folder/Project Reflection Reports 2025`
   folder live) — then it can read both, and any new reflection reports, automatically.
3. Add the secret **`GDRIVE_SA_KEY`** = the whole JSON key.
4. Add the secret **`ESTIMATOR_REPO_TOKEN`** = a GitHub PAT with read access to
   `odl-estimator` (so the job can check out its `data_all`). Skip this if you make
   `odl-estimator` public and drop the `token:` line in the workflow.
5. **Actions → dashboard-refresh → Run workflow** to test, then it runs daily at
   13:00 UTC (after the estimator's nightly Asana pull). File ids for the Sheet and
   folder are defaults in `fetch_drive.py`; override with env vars if they move.

**Local refresh (optional / fallback).** If you'd rather refresh from your own
machine, `./sync.sh` runs `build.py`, commits the refreshed `data.json` / `index.html`,
and pushes (`--no-push` to commit only); `./install_schedule.sh` schedules it daily
via launchd while you're logged in. Use **one** daily mechanism, not both — pick the
cloud workflow or the local schedule so two jobs don't fight over the same commit.

> Continuity: both data sources — the **Capacity Allocations Sheet** and the **reflection-reports
> folder** — now live in the **NDL ODL shared drive** (org-owned), so they survive
> any one account. What still needs handling so it all keeps running after you leave:
> transfer both GitHub repos to an ND org / successor (so the Actions + secrets
> aren't tied to a personal account), and create the Google service account under a
> team/ND Google project rather than a personal one.

### View it as a live website (GitHub Pages)

So the team opens a URL instead of downloading `index.html`, a second workflow
(`.github/workflows/pages.yml`) publishes the dashboard to **GitHub Pages** and
re-publishes it every time `index.html` changes — so the live page stays current
with the daily refresh on its own. It serves **only** `index.html`, so nothing else
(source, raw `data.json`) is exposed as a separate page.

Turn it on once: repo → **Settings → Pages → Build and deployment → Source =
"GitHub Actions"**. The site is then at `https://nd-learning.github.io/PMdashboard/`.

Visibility: free GitHub Pages serves only **public** repos. Since the dashboard
holds internal data (names, capacity, satisfaction scores), to keep the repo
**private** and have a **login-restricted** site, host it in a GitHub org / paid
plan — e.g. ND's **github.nd.edu** restricted to members. This same workflow works
there unchanged.

## Live mode (real-time from Asana)

```bash
python3 serve.py                       # http://127.0.0.1:8787  (localhost only)
ASANA_TOKEN='0/…' python3 serve.py     # also enables the "Sync Asana" button
```

| Endpoint | What it does |
|---|---|
| `GET /` | serves `index.html` (and sibling assets) |
| `GET /api/data` | the full payload, **recomputed live** from the Capacity Allocations Google Sheet + the current Asana CSV snapshot. Fast, no network. |
| `POST /api/sync` | re-pulls Asana via the `../odl_estimator` refresh pipeline (needs `ASANA_TOKEN`), then recomputes — the literal "pull from Asana" button. Slow. |
| `GET /api/health` | `{ok, asana_snapshot_date, can_sync}` |

- The Asana **token is read from the environment / macOS keychain by the
  pipeline and is never sent to the browser**; the served JSON contains no
  secrets. The server binds to `127.0.0.1` so the token-backed `/api/sync` is
  reachable only from this machine (use `--host 0.0.0.0` only behind a trusted
  proxy).
- Serving **never overwrites `statuses.json`** (only the canonical `build.py`
  does), so the shared recommendation-status file is safe.

## The tabs

The tab bar is split into the **Director Brief** (a glanceable weekly summary for
the Director) and, after a `PM detail` divider, the existing **PM-detail** tabs
(the full working views). The Brief opens by default.

| Tab | What it shows |
|---|---|
| **Director Brief** | A "Monday ODL Brief" organised as three executive questions. **① What needs attention this week?** — *Top 5 things to know* (one-line summary + recommended action each), *Projects at risk* (active projects flagged from their weekly Asana status update — colored status, risk phrases, staleness over 21 days — a past-due Asana date, or actual hours over ~120% of plan, each with a plain reason, the update snippet, an Asana link, and a "Why we flag things" rules list), and *New projects from intake to discuss* (live from the Asana "NDL Project Tracking & Awareness" board — Received Requests / Incentives Requests, Leadership Review or Approval Needed, Under Review by Unit — each linking to its Asana task; falls back to `brief_inputs.json`'s manual intake list, with a note, if that board isn't in the snapshot yet). **② How is the portfolio moving?** — *phase congestion* (active projects per phase and which team they tie up), *year-to-date stats*, *budget vs actual hours* (planned hours from the Capacity sheet vs. logged timesheet hours, active projects only, with a "How we estimate" panel and over/under-plan consequences), and *timesheet compliance* (honest — reads ~0% while Asana time entry is resuming after lapsing 2026-06-24; shows who logged/didn't, plus a computed per-person June hours table). **③ How are we improving?** — *recommendations progress* (are we acting on the tracker?), *wins & momentum* (the team's All-Hands shout-outs + faculty wins), and *ND Learning round-up* draft blurbs (with a Copy button). **✍ Round-up email draft** — a suggested plain-text round-up email assembled from the team's running notes (shout-outs + updates) and live Asana activity (recent completions, in-progress work by phase + point person, faculty wins), in an editable text box with Copy / Reset. Every figure is **derived from the Capacity sheet + Asana data the PM tabs use**; the only hand-entered pieces (the weekly **running-notes** shout-outs/updates pasted from the team agenda Google Doc, curated round-ups, and the timesheet-compliance %) live in **`brief_inputs.json`** and show a clear empty state until filled in — the intake queue itself is now live from Asana, falling back to `brief_inputs.json`'s manual list only if that board isn't in the snapshot yet. |
| **Overview** | This-month KPIs, a team-allocation heatmap by month, top open recommendations. |
| **Capacity** | Two views: **Summary** (estimated vs scheduled vs remaining per scope — Total / Design / Media / PM / Intern — monthly, with % allocated and unstaffed work) and **Master view**, the *full plan behind the summary*: every project's month-by-month staffing, expandable to each role/person, exactly as it stands now (live when served). Shows a **"Source: Capacity Allocations (Google Sheet) — updated nightly ↗"** line and defaults to **hours**. |
| **People** | Primary view: actual **timesheet hours** per person by month (2026-06 onward) vs. their capacity; the allocation / remaining / %-used views are still there too. Click a name for their projects + monthly load. Everyone's monthly allocation stays visible; Lawrence (departed) never appears at the person level. |
| **Projects** | All projects — Asana `projects.csv` **unioned** with the project names on the Capacity sheet; overhead/non-project boards (z-Professional Development, Time Off, Impact Tracker, the tracking board, `Test -*`, templates, etc.) are excluded from project-level panels via an easily-edited `OVERHEAD_PATTERNS` constant in `build.py`, though they still count toward capacity totals. A project counts as **active** if it's planned on the Capacity sheet this month or later, or it's non-archived in Asana with a weekly status update in the last 30 days. Shows project **size** (effort in points/hours), status, phase, timeline — with an explicit **Sort** control plus sortable columns, plus a **Projects at risk** panel (same flagging rules as the Brief). Click for a phase Gantt + per-role monthly staffing + Asana facts. |
| **Faculty** | The projects ODL works with, grouped two ways (toggle): **By department** — Impact Tracker projects per Notre Dame academic department (from Asana; gaps filled from `nd.edu`, tagged with provenance + confidence) with a context blurb — and **By year** — projects per calendar year they were active (Asana Start → End), each annotated with its department, faculty, type, FSI, status and that year's logged hours; projects with no dates are listed separately. Plus the Faculty Satisfaction Index / Net Promoter Score and a per-project ratings table. (Reflection reports back the faculty-feedback recommendations on the **Recommendations** tab. Fuller perspectives live in **Qualtrics**, access pending.) |
| **Plan a project** | No longer a workbook size model. Shows archetype **hour ranges** from the estimator's calibrated quartiles (`../odl_estimator/data_all/derived/calibration.json`) — **Full course, Course redesign, Video series, Single video** — alongside the team's current **available capacity** (remaining hours from the capacity model) and a rough **"how many fit"** check for each archetype. The tab hides gracefully if `calibration.json` is absent. *(Faculty time is not included — this tool measures the ODL team's capacity only.)* |
| **Recommendations** | A list of *ways to improve our workflow and process*, grouped by theme. Items draw on **capacity** signals (incl. who's over/under-utilized), light **Asana-hygiene**, and **reflection-report** items — each report's own *Recommendations / Lessons Learned* section distilled into an actionable item ("Apply lessons from …", "What to do: …"). It reads first: click an item for the detail. As the team acts on one, mark it **To do / In progress / Done** and add a note — a deliberately simple control (no owner/target/evidence clutter). A progress line and **Save our progress** (→ `statuses.json`) sit at the top. The reflection reports are listed at the bottom; fuller student/Qualtrics perspectives are still pending. |

## Recommendations

Recommendations are **auto-derived** from the data (team/person over- or
under-capacity, unstaffed committed work, completed projects with no post-project
report, on-hold projects, budget-vs-actual hours drift, status-update-derived
**project-at-risk** flags, and the **reflection reports** — each report's own
*Recommendations / Lessons Learned*
section distilled into one actionable item per project) **plus manual notes** in
`recommendations_manual.json`. *(Low-signal "Impact-Tracker outdated" items were
intentionally dropped to keep the list focused.)*

The tab **reads first** ("here's what we can improve"), with a **simple progress
control**: as the team acts on an item, mark it **To do / In progress / Done** and
add a note. (This deliberately replaces an earlier, cluttered tracker — four status
states plus owner / target / evidence / notes fields — that nobody used; the new one
is just *mark + note*.) Marks save per-browser; **Save our progress ↓** downloads
`statuses.json` to share/commit next to `build.py` (it only ever *adds* new IDs, never
overwrites edits), and **Load ↑** imports one. Optionally, `asana_push.py` mirrors
each recommendation as a task in a dedicated Asana project (idempotent; **dry-run by
default**, `--push` to write).

## Files

| File | Role |
|---|---|
| `build.py` | ETL: Capacity Allocations CSV + Asana snapshot + reflection PDFs + `nd_departments.json` → `data.json` (+ validation report) → calls `render.py`. ETL exposed as `compute_data()` for the live server. |
| `capacity_allocations.csv` | exported snapshot of the **Capacity Allocations** Google Sheet (via `fetch_drive.py`) — per-person, per-month Capacity/Allocated Hours + Projects; drives both the capacity model and each project's planned-hours budget. |
| `serve.py` | local **live server** — recomputes the payload per request (`/api/data`) and can pull Asana on demand (`/api/sync`). Token from env/keychain, never sent to the client. |
| `recommend.py` | auto-derives recommendations (incl. faculty-feedback), merges manual + tracked statuses |
| `nd_departments.json` | ND academic-department context blurbs + inferred project→department map (sourced from `nd.edu`, with confidence) |
| `reflection_drive_links.json` | filename → Google Drive view URL for the `../Reflection/` PDFs (so links open from Drive/Canvas) |
| `render.py` | bakes `data.json` into the self-contained `index.html` |
| `template.html` | the app (HTML/CSS/JS, hand-rolled SVG charts — no CDN) |
| `index.html` | **the deliverable** — open this |
| `data.json` | generated data (also embedded in `index.html`) |
| `statuses.json` | canonical tracked recommendation statuses (pipeline-maintained) |
| `recommendations_manual.json` | manually-entered recommendations |
| `brief_inputs.json` | manual weekly inputs for the **Director Brief**: a **fallback intake queue** (used only if the Asana "NDL Project Tracking & Awareness" board isn't in the snapshot yet — intake is otherwise live from Asana), the **`running_notes`** (Shout-outs + Project Updates pasted from the weekly team-agenda Google Doc — drives Wins & momentum and the Round-up draft), curated round-up drafts, and the timesheet-compliance %. The Brief's other sections are auto-derived. |
| `asana_push.py` | optional write-back of statuses to Asana |
| `../odl_estimator/data_all/derived/calibration.json` | estimator's calibrated hour-range quartiles (Full course / Course redesign / Video series / Single video) — read by the **Plan a project** tab; that tab hides gracefully if this file is absent. |

## Reflection reports & faculty feedback

`build.py` indexes the PDFs in `../Reflection/` (downloaded from the Asana
Impact Tracker's "Link to Reflection Report" field). Evaluation/Retrospective
reports get a Summary + Takeaways extracted (via `fitz` (PyMuPDF), falling back
to `pypdf`, if installed — links still work without them); faculty survey
responses are linked as-is. Each is matched to a project by name/faculty, and
each matched report also becomes a **faculty-feedback recommendation** on the
Recommendations tab.

Links open the best available target, in order: the live **Google Doc** (the
Asana "Link to Reflection Report"; reports only) → the **Google Drive** copy of
the PDF (from `reflection_drive_links.json`; covers reports *and* surveys) → a
relative `../Reflection/<file>` path (last resort — only resolves when that
folder sits one level up from `odl_pm_dashboard/`). Drop more PDFs in and re-run
`build.py`; refresh `reflection_drive_links.json` (filename → Drive view URL) if
you want newly-added files to open from Drive/Canvas.

## Notes

- No invented numbers — every figure traces to the Capacity sheet or the Asana
  snapshot. No secrets are stored in any file; the Asana token (for the snapshot
  pull and the optional write-back) lives in the environment/keychain.
- This is **internal/PM-facing** and distinct from the faculty-facing
  `../odl_estimator` effort estimator; the two share the Asana snapshot.
