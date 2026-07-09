# Putting the PM dashboard behind Notre Dame (Okta) login

Status as of 2026-07-09:

- The public GitHub Pages site (`nd-learning.github.io/PMdashboard`) is **taken down**
  and the `ND-Learning/PMdashboard` repo is **private**.
- Still pending (repo-owner-only toggles, ~1 minute):
  1. github.com/jliu234-cell/odl-pm-dashboard → Settings → Pages → **unpublish**
     (the old public copy is still served there until you do).
  2. Optionally make `jliu234-cell/odl-pm-dashboard` private too — it holds the
     full `data.json` (names + hours). If you do, tell Claude so the
     cross-repo checkouts get a token.
- The nightly build keeps running either way (`jliu234-cell/odl-pm-dashboard`,
  13:00 UTC): fresh `index.html`/`data.json` are committed daily, ready for
  whatever authenticated host OIT provides.

## The ask for OIT (help.nd.edu — copy/paste)

> Subject: Hosting an internal static dashboard behind Okta login
>
> Our team (ODL / Notre Dame Learning) has an internal project-management
> dashboard that must be restricted to Notre Dame accounts (ideally a specific
> team list). Technically it is a single self-contained static `index.html`
> (~800 KB, no server code, no database), regenerated nightly by an automated
> pipeline (GitHub Actions) that can push/upload wherever needed.
>
> What is the recommended OIT service for hosting a static internal page behind
> Okta/ND SSO? Specifically:
> 1. Does ND offer enterprise GitHub (or GitLab) with SSO-protected Pages we
>    could publish to?
> 2. If not, is there an OIT web-hosting option (or Conductor equivalent) that
>    can front a static page with Okta and give us upload/API access for the
>    nightly refresh?
> 3. If neither fits, can an Okta app integration be registered for a small
>    static site we host (e.g., behind Cloudflare Access with Okta as the IdP)?
>
> We need: ND-login-gated viewing for ~15 people, and a way for our nightly
> automation to update one HTML file.

## What happens after OIT answers

Bring the answer back to a Claude session and it wires the nightly deploy:

- **Enterprise GitHub Pages** → add its clone URL + a PAT as secrets on
  `ND-Learning/PMdashboard`; a small workflow pushes the fresh `index.html`
  there nightly. (Confirm with OIT that the instance's Pages require login —
  on some setups Pages are public even when the instance is SSO-gated.)
- **OIT web hosting / Conductor** → whatever upload mechanism they give
  (SFTP/API/rsync) goes into the same nightly workflow as a secret.
- **Okta app + Cloudflare Access** → create the Cloudflare account, OIT
  connects Okta as the IdP; Claude sets up Cloudflare Pages + the Access
  policy and the nightly deploy.

## Interim access (until OIT responds)

The dashboard has no public URL right now. Team options meanwhile:
- Open `index.html` from this repo/Drive directly (self-contained, works offline).
- Or ask Claude to temporarily publish it as a page in a Canvas course
  (Canvas is already behind ND Okta; the faculty guide deploys this way) —
  this is also the fastest *permanent* fallback if OIT is slow.
