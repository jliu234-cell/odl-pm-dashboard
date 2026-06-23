#!/bin/zsh
# install_schedule.sh -- schedule the daily dashboard sync on this Mac via launchd.
#   ./install_schedule.sh            install/update (daily at 08:00)
#   ./install_schedule.sh uninstall  remove
# The job runs sync.sh (rebuild + commit + push) and logs to sync_launchd.log.
# It only runs while this Mac is awake and you're logged in (so the keychain can
# serve the GitHub credential); launchd runs a missed job at next wake.
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
LABEL="com.odl.dashboard.sync"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"

if [[ "$1" == "uninstall" ]]; then
  launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
  rm -f "$PLIST"
  echo "uninstalled $LABEL"
  exit 0
fi

chmod +x "$HERE/sync.sh"
mkdir -p "$HOME/Library/LaunchAgents"
/usr/bin/python3 - "$HERE" "$PLIST" "$LABEL" <<'PY'
import plistlib, sys
here, plist_path, label = sys.argv[1], sys.argv[2], sys.argv[3]
plistlib.dump({
    "Label": label,
    # run via a login shell so PATH / python deps (openpyxl) resolve like a terminal
    "ProgramArguments": ["/bin/zsh", "-lc", f"'{here}/sync.sh'"],
    "WorkingDirectory": here,
    "StartCalendarInterval": {"Hour": 8, "Minute": 0},
    "StandardOutPath": f"{here}/sync_launchd.log",
    "StandardErrorPath": f"{here}/sync_launchd.log",
}, open(plist_path, "wb"))
print(f"wrote {plist_path}")
PY

launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
launchctl bootstrap "gui/$(id -u)" "$PLIST"
echo "scheduled: $LABEL daily at 08:00 (logs: sync_launchd.log)"
echo "test now with:  launchctl kickstart gui/$(id -u)/$LABEL"
