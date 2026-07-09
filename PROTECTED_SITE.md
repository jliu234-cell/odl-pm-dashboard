# Password-protected dashboard site

**URL:** https://nd-learning.github.io/PMdashboard-site/ — enter the ODL team
password; the browser remembers it for 30 days. (ND institutional login was
considered and dropped 2026-07-09; the dashboard is internal, not for faculty,
so a shared team password suffices.)

How it works:
- `jliu234-cell/odl-pm-dashboard` (builder) rebuilds `index.html` nightly
  (13:00 UTC) from the capacity sheet + Asana snapshot.
- `ND-Learning/PMdashboard-site` (public, holds ONLY ciphertext) runs
  `encrypt-and-publish` (14:30 UTC): fetches the fresh build, encrypts it with
  StatiCrypt (AES-256 + PBKDF2), commits `index.html`, and GitHub Pages serves
  it. A plaintext-leak guard fails the job if unencrypted content ever appears.
- This repo (`ND-Learning/PMdashboard`, private) is the source of truth for
  the dashboard code; it no longer publishes anywhere itself.

Password rotation: PMdashboard-site → Settings → Secrets and variables →
Actions → `STATICRYPT_PASSWORD` → update, then Actions → encrypt-and-publish →
Run workflow. Everyone gets the new password; "remembered" browsers must
re-enter it.

Residual exposure to keep in mind:
- Anyone WITH the password can save and share the decrypted page — a shared
  password protects against the public, not against leaks by password-holders.
- `jliu234-cell/odl-pm-dashboard` is still a PUBLIC repo carrying the plaintext
  `data.json`/`index.html`, and its old Pages site may still be published —
  making that repo private + unpublishing its Pages (Settings → Pages) closes
  the loop. After privatizing, add a read PAT as `BUILDER_REPO_TOKEN` on
  PMdashboard-site and set `token:` on the builder checkout step in
  `encrypt-and-publish.yml` (a note marks the spot).
