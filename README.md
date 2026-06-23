# ODL Capacity & Performance Dashboard

An internal, PM-facing website for **Online Digital Learning** to plan and track
multiple projects at once against real team capacity — and to surface concrete
ways to improve the team's workflow and process.

It answers, in one place:

- **Are we over capacity** this month / any month, overall and per team (Design,
  Media, PM, Intern)?
- **Who is over- or under-allocated**, and on what?
- **Can we take on a new project** of a given size starting in a given month
  without blowing past capacity? (the *Plan a project* what-if)
- **How can we improve our process?** — concrete recommendations from capacity
  reviews and faculty reflection reports, with a simple way to mark progress
  (the *Recommendations* tab)

## The capacity model

Straight from the workbook's *Explanations* tab:

> **1 point = 32 hours ≈ 1 productive week. 4 points/person/month = a full load.**
> (40h/week × ~75% efficiency for meetings/breaks ≈ 32 productive hours.)

Everything is monthly. `remaining = estimated − scheduled`; `allocated % =
scheduled / estimated`. Over 100% = over capacity.

## How it's built

```
ODL Project and Capacity Planning.xlsx ─┐                          ┌─► index.html  (baked, self-contained,
   (live Google-Drive sheet)            ├─ build.py ──► data.json ──┤    open from Drive — offline snapshot)
../odl_estimator/data_all/projects.csv ─┘   (parse +    (render.py   └─► serve.py ──► /api/data  (LIVE: recomputes
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

- **`build.py`** parses the workbook and the Asana snapshot, recomputes the
  capacity model *consistently* from the `Projects` tab, joins live project
  status from Asana, and writes `data.json`. The ETL lives in a reusable
  `compute_data()` that **`serve.py`** also calls to serve live numbers. It
  prints a validation report (and verifies the per-person estimated capacity
  still sums **exactly** to the workbook's own Total).
- **`serve.py`** is the local live server (see **Live mode** below).
- **`recommend.py`** auto-derives recommendations and merges in manual ones +
  tracked statuses (stable IDs, so statuses survive a rebuild).
- **`render.py`** bakes `data.json` into **`index.html`** — one file, no
  external assets, works offline. Just open it (or host it on Canvas/Drive).

### Rebuild / refresh

```bash
python3 build.py          # re-parse workbook + snapshot, regenerate index.html
python3 serve.py          # …or serve live numbers without rebuilding (see below)
```

The Asana snapshot is produced by the existing `../odl_estimator` pipeline
(`asana_pull.py` / nightly `refresh.py`). Re-run that first to pull the latest
Asana state, then `build.py` here to refresh the dashboard — or use the live
server's **Sync Asana** button to pull and recompute in place.

### Keeping GitHub in sync

Unlike the sibling `odl_estimator` (which self-refreshes **in the cloud** via a
GitHub Actions cron job, because it pulls Asana directly with a token), this
dashboard rebuilds from the **local** capacity workbook (a Google-Drive sheet) and
the Asana snapshot under `../odl_estimator/data_all/`. A cloud runner can't see
those, so the refresh has to run **locally**, then push.

- **Manual:** `./sync.sh` — runs `build.py`, commits the refreshed `data.json` /
  `index.html`, and pushes. `./sync.sh --no-push` commits only.
- **Daily (automatic):** `./install_schedule.sh` installs a launchd job that runs
  `sync.sh` every morning at 08:00, while the Mac is awake and you're logged in (so
  the macOS keychain can serve the GitHub credential). `./install_schedule.sh
  uninstall` removes it; logs land in `sync_launchd.log` (git-ignored).

If the daily push ever stalls on a credential prompt, store a GitHub personal
access token once (`git config credential.helper osxkeychain` is already in use;
re-authenticating from a normal `git push` caches it), or switch the remote to SSH
with a deploy key.

## Live mode (real-time from Asana)

```bash
python3 serve.py                       # http://127.0.0.1:8787  (localhost only)
ASANA_TOKEN='0/…' python3 serve.py     # also enables the "Sync Asana" button
```

| Endpoint | What it does |
|---|---|
| `GET /` | serves `index.html` (and sibling assets) |
| `GET /api/data` | the full payload, **recomputed live** from the capacity workbook (a live Google-Drive sheet) + the current Asana CSV snapshot. Fast, no network. |
| `POST /api/sync` | re-pulls Asana via the `../odl_estimator` refresh pipeline (needs `ASANA_TOKEN`), then recomputes — the literal "pull from Asana" button. Slow. |
| `GET /api/health` | `{ok, asana_snapshot_date, can_sync}` |

- The Asana **token is read from the environment / macOS keychain by the
  pipeline and is never sent to the browser**; the served JSON contains no
  secrets. The server binds to `127.0.0.1` so the token-backed `/api/sync` is
  reachable only from this machine (use `--host 0.0.0.0` only behind a trusted
  proxy).
- Serving **never overwrites `statuses.json`** (only the canonical `build.py`
  does), so the shared recommendation-status file is safe.

> **Why the dashboard recomputes instead of reading the workbook's "Team
> Capacity" tab:** that tab is maintained by hand and has drifted (its Total
> doesn't match its own sub-teams, nor the `Projects` tab). The dashboard
> recomputes from the single internally-consistent source (the `Projects` tab)
> so the numbers always reconcile, and surfaces the drift as a data-hygiene
> recommendation. The workbook's tab is still shown as a *reference* overlay
> (Capacity tab → "show workbook reference").

## The tabs

The tab bar is split into the **Director Brief** (a glanceable weekly summary for
the Director) and, after a `PM detail` divider, the existing **PM-detail** tabs
(the full working views). The Brief opens by default.

| Tab | What it shows |
|---|---|
| **Director Brief** | A "Monday ODL Brief" organised as three executive questions. **① What needs attention this week?** — *Top 5 things to know* (one-line summary + recommended action each), *Projects needing attention* (the top 5 active projects flagging a risk — including ones that look green in Asana but carry a hidden one — each expandable for the problem + suggested action), and *New projects from intake to discuss* (requestor + description). **② How is the portfolio moving?** — *phase congestion* (active projects per phase and which team they tie up), *year-to-date stats*, *estimated vs actual size*, and *timesheet compliance*. **③ How are we improving?** — *recommendations progress* (are we acting on the tracker?), *wins & momentum* (the team's All-Hands shout-outs + faculty wins), and *ND Learning round-up* draft blurbs (with a Copy button). **✍ Round-up email draft** — a suggested plain-text round-up email assembled from the team's running notes (shout-outs + updates) and live Asana activity (recent completions, in-progress work by phase + point person, faculty wins), in an editable text box with Copy / Reset. Every figure is **derived from the same workbook + Asana data the PM tabs use**; the only hand-entered pieces (the intake queue, the weekly **running-notes** shout-outs/updates pasted from the team agenda Google Doc, curated round-ups, and the timesheet-compliance %) live in **`brief_inputs.json`** and show a clear empty state until filled in. |
| **Overview** | This-month KPIs, a team-allocation heatmap by month, top open recommendations. |
| **Capacity** | Two views: **Summary** (estimated vs scheduled vs remaining per scope — Total / Design / Media / PM / Intern — monthly, with % allocated and unstaffed work) and **Master view**, the *full plan behind the summary*: every project's month-by-month staffing, expandable to each role/person, exactly as it stands now (live when served). |
| **People** | Per-person remaining capacity heatmap (the workbook's red/green over/under check), live. Click a name for their projects + monthly load. |
| **Projects** | All projects: project **size** (effort in points/hours), status, phase, timeline — with an explicit **Sort** control plus sortable columns. Click for a phase Gantt + per-role monthly staffing + Asana facts. (The qualitative T-shirt size is no longer shown here; the standard XS/S/M/L staffing profiles live on the Plan tab as a "Staffing template".) |
| **Faculty** | The projects ODL works with, grouped two ways (toggle): **By department** — Impact Tracker projects per Notre Dame academic department (from Asana; gaps filled from `nd.edu`, tagged with provenance + confidence) with a context blurb — and **By year** — projects per calendar year they were active (Asana Start → End), each annotated with its department, faculty, type, FSI, status and that year's logged hours; projects with no dates are listed separately. Plus the Faculty Satisfaction Index / Net Promoter Score and a per-project ratings table. (Reflection reports back the faculty-feedback recommendations on the **Recommendations** tab. Fuller perspectives live in **Qualtrics**, access pending.) |
| **Plan a project** | Two tools: (1) a **"how many projects can we take?"** estimator — pick a **project size** (XS/S/M/L) + a period (Summer / a semester / custom) and it compares the period's **live** surplus capacity, after current + staged work, to the work one project needs per team, and names the binding team; (2) **stage** individual projects at **any** start month to see the month-by-month effect on each team. The per-size footprint is computed from **real recent projects** each build (not a frozen workbook template), so it includes real PM time. Each size's reliable **median total** effort is split across teams by ODL's **pooled team mix** (point-weighted Design/Media/PM share over all sized projects) rather than per-size per-team medians — the latter were noisy on small samples (XL is only 2 projects) and produced a non-monotonic ladder. The "how many fit" count divides each project's **full** team cost by the window's spare capacity, so a long project can't look cheap just because only its early months land inside a short window. Together these keep the estimate sane and monotonic (more small projects fit than large ones). *(Faculty-time estimates are deliberately not used here — this tool measures the ODL team's capacity only.)* |
| **Recommendations** | A list of *ways to improve our workflow and process*, grouped by theme. Items draw on **capacity** signals (incl. who's over/under-utilized), light **Asana-hygiene**, and **reflection-report** items — each report's own *Recommendations / Lessons Learned* section distilled into an actionable item ("Apply lessons from …", "What to do: …"). It reads first: click an item for the detail. As the team acts on one, mark it **To do / In progress / Done** and add a note — a deliberately simple control (no owner/target/evidence clutter). A progress line and **Save our progress** (→ `statuses.json`) sit at the top. The reflection reports are listed at the bottom; fuller student/Qualtrics perspectives are still pending. |

## Recommendations

Recommendations are **auto-derived** from the data (team/person over- or
under-capacity, unstaffed committed work, completed projects with no post-project
report, on-hold projects, estimate-vs-actual size drift, workbook drift, and the
**reflection reports** — each report's own *Recommendations / Lessons Learned*
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
| `build.py` | ETL: workbook + Asana snapshot + reflection PDFs + `nd_departments.json` → `data.json` (+ validation report) → calls `render.py`. ETL exposed as `compute_data()` for the live server. |
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
| `brief_inputs.json` | manual weekly inputs for the **Director Brief**: the intake queue, the **`running_notes`** (Shout-outs + Project Updates pasted from the weekly team-agenda Google Doc — drives Wins & momentum and the Round-up draft), curated round-up drafts, and the timesheet-compliance %. The Brief's other sections are auto-derived. |
| `asana_push.py` | optional write-back of statuses to Asana |

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

- No invented numbers — every figure traces to a workbook cell or the Asana
  snapshot. No secrets are stored in any file; the Asana token (for the snapshot
  pull and the optional write-back) lives in the environment/keychain.
- This is **internal/PM-facing** and distinct from the faculty-facing
  `../odl_estimator` effort estimator; the two share the Asana snapshot.
