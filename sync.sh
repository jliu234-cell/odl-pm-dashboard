#!/bin/zsh
# sync.sh -- rebuild the dashboard from the live capacity workbook + the current
# Asana snapshot, then commit & push the refreshed data.json / index.html to
# GitHub. Run it whenever the data changes, or schedule it with ./install_schedule.sh.
#
#   ./sync.sh             rebuild + commit + push
#   ./sync.sh --no-push   rebuild + commit only (don't push)
#
# Note: build.py reads ../ODL Project and Capacity Planning.xlsx (a Google-Drive
# sheet) and ../odl_estimator/data_all/ (the Asana snapshot), so this must run on
# a machine that can see those. The git push uses the macOS keychain credential,
# so you must be logged in for it to authenticate non-interactively.
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
cd "$HERE"
echo "===== dashboard sync: $(date) ====="
python3 build.py
git add -A
if git diff --cached --quiet; then
  echo "no changes to commit."
  exit 0
fi
git commit -m "dashboard refresh $(date -u +%F)"
if [[ "$1" == "--no-push" ]]; then
  echo "committed (push skipped)."
else
  git push && echo "pushed to GitHub."
fi
