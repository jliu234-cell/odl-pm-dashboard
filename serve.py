#!/usr/bin/env python3
"""ODL PM Dashboard — local live server (real-time capacity from Asana).

The dashboard's ``index.html`` works on its own as a self-contained, offline
snapshot (open it from Drive/Canvas). Run THIS server when you want the numbers
to be **live** instead of a frozen bake: the page fetches ``/api/data`` on load
and recomputes the capacity master view + the "how many projects can we take"
estimator from the *current* sources every request — so when teammates update
their time, a Refresh reflects it without rebuilding.

Endpoints
---------
  GET  /                 -> index.html (and any sibling static file)
  GET  /api/data         -> the full dashboard payload, recomputed live from the
                            capacity workbook (a live Google-Drive sheet) + the
                            current Asana CSV snapshot. Fast (no network).
  POST /api/sync         -> re-pull Asana via the odl_estimator refresh pipeline
                            (needs ASANA_TOKEN), then recompute. Slow; this is
                            the literal "pull from Asana" button.
  GET  /api/health       -> {ok, asana_snapshot_date, can_sync}

Security
--------
  * Binds to 127.0.0.1 by default — the Asana-token-backed /api/sync is reachable
    only from this machine. Use --host 0.0.0.0 to expose it (only behind a
    trusted network / auth proxy).
  * The Asana token is read from the environment (ASANA_TOKEN) or the macOS
    keychain by the existing pipeline; it is NEVER sent to the browser. The
    served JSON contains no secrets.

Run
---
  python3 serve.py                 # http://127.0.0.1:8787
  python3 serve.py --port 9000
  ASANA_TOKEN='0/…' python3 serve.py   # enables /api/sync
"""
import argparse
import json
import os
import re
import subprocess
import sys
import threading
import time
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

import build

HERE = os.path.dirname(os.path.abspath(__file__))
ESTIMATOR_DIR = os.path.join(os.path.dirname(HERE), "odl_estimator")
REFRESH_PY = os.path.join(ESTIMATOR_DIR, "refresh.py")

# A short cache so quick successive page loads / navigations don't re-parse the
# 1.4 MB workbook every time; a Refresh click forces a recompute via ?fresh=1.
_CACHE = {"data": None, "at": 0.0}
_CACHE_TTL = 8.0
_LOCK = threading.Lock()
_SYNC = {"running": False, "started": 0.0, "last": None}

STATIC_TYPES = {".html": "text/html; charset=utf-8",
                ".css": "text/css", ".js": "application/javascript",
                ".png": "image/png", ".svg": "image/svg+xml", ".ico": "image/x-icon",
                ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".gif": "image/gif",
                ".woff": "font/woff", ".woff2": "font/woff2"}
# only these extensions are served statically — so the server's own source
# (*.py) and local data/status files (*.json) are NOT downloadable over the port.
SERVE_EXTS = set(STATIC_TYPES)


def live_data(fresh=False):
    """compute_data, lightly cached. Never writes statuses.json (write_status
    False) so serving never clobbers the shared recommendation-status file."""
    now = time.time()
    with _LOCK:
        if not fresh and _CACHE["data"] is not None and (now - _CACHE["at"]) < _CACHE_TTL:
            return _CACHE["data"]
    data = build.compute_data(do_recs=True, write_status=False, verbose=False)
    data["meta"]["served_at"] = build.datetime.datetime.now().isoformat(timespec="seconds")
    data["meta"]["live"] = True
    payload = json.dumps(data, default=str).encode("utf-8")
    with _LOCK:
        _CACHE["data"] = payload
        _CACHE["at"] = time.time()
    return payload


def can_sync():
    return bool(REFRESH_PY and os.path.exists(REFRESH_PY))


def run_sync():
    """Re-pull Asana through the odl_estimator refresh pipeline. Returns
    (ok, message). Token/scope handling lives in refresh.py (env or keychain)."""
    if not can_sync():
        return False, "refresh pipeline not found (odl_estimator/refresh.py)"
    with _LOCK:
        if _SYNC["running"]:
            return False, "a sync is already running"
        _SYNC["running"] = True
        _SYNC["started"] = time.time()
    try:
        env = dict(os.environ)
        proc = subprocess.run([sys.executable, "refresh.py"], cwd=ESTIMATOR_DIR,
                              env=env, capture_output=True, text=True, timeout=900)
        ok = proc.returncode == 0
        tail = ((proc.stdout or "")[-2000:] + (proc.stderr or "")[-2000:]).strip()
        # full output stays in the SERVER log only; the browser never gets raw
        # subprocess stdout/stderr (path / future-token-leak hygiene).
        sys.stderr.write("\n--- /api/sync refresh output ---\n" + tail + "\n--- end ---\n")
        # bust the data cache so the next /api/data reflects the fresh snapshot
        with _LOCK:
            _CACHE["data"] = None
        if ok:
            return True, "synced"
        scrub = re.sub(r"(0/[0-9A-Fa-f]{8,}|Bearer\s+\S+)", "***", tail)
        last = next((l for l in reversed(scrub.splitlines()) if l.strip()), "")
        return False, "Asana sync failed — see server log." + (f" ({last[-160:].strip()})" if last else "")
    except subprocess.TimeoutExpired:
        return False, "sync timed out after 15 min"
    except Exception as e:
        return False, f"sync error: {e}"
    finally:
        with _LOCK:
            _SYNC["running"] = False
            _SYNC["last"] = time.time()


class Handler(BaseHTTPRequestHandler):
    server_version = "ODLDash/1.0"

    def _send(self, code, body, ctype="application/json", extra=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _err(self, code, msg):
        self._send(code, json.dumps({"ok": False, "error": msg}))

    def do_GET(self):
        path = urlparse(self.path).path
        try:
            if path in ("/api/data", "/api/data/"):
                fresh = "fresh=1" in (urlparse(self.path).query or "")
                return self._send(200, live_data(fresh=fresh))
            if path in ("/api/health", "/api/health/"):
                snap = None
                try:
                    _, snap = build.load_asana()
                except Exception:
                    pass
                return self._send(200, json.dumps(
                    {"ok": True, "asana_snapshot_date": snap, "can_sync": can_sync(),
                     "sync_running": _SYNC["running"]}))
            return self._static(path)
        except Exception:
            traceback.print_exc()
            return self._err(500, "compute failed — see server log")

    def do_HEAD(self):
        self.do_GET()

    def do_POST(self):
        path = urlparse(self.path).path
        if path in ("/api/sync", "/api/sync/"):
            ok, msg = run_sync()
            if not ok:
                return self._err(503, msg)
            # return the freshly-recomputed payload so the page updates in one round-trip
            try:
                return self._send(200, live_data(fresh=True))
            except Exception:
                traceback.print_exc()
                return self._err(500, "synced, but recompute failed — reload")
        return self._err(404, "not found")

    def _static(self, path):
        rel = path.lstrip("/") or "index.html"
        # realpath resolves both ".." traversal AND symlinks, then a
        # sep-terminated prefix keeps us inside HERE (a sibling dir that merely
        # shares the name — …/odl_pm_dashboard_x — can't escape). Only allowlisted
        # extensions are served, so .py/.json/dotfiles are never downloadable.
        root = os.path.realpath(HERE)
        full = os.path.realpath(os.path.join(root, rel))
        ext = os.path.splitext(full)[1].lower()
        if ((full != root and not full.startswith(root + os.sep))
                or ext not in SERVE_EXTS or not os.path.isfile(full)):
            return self._err(404, "not found")
        with open(full, "rb") as f:
            body = f.read()
        self._send(200, body, STATIC_TYPES.get(ext, "application/octet-stream"))

    def log_message(self, fmt, *args):  # quieter, single-line
        sys.stderr.write("  %s - %s\n" % (self.address_string(), fmt % args))


def main():
    ap = argparse.ArgumentParser(description="ODL PM Dashboard live server")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=8787)
    args = ap.parse_args()

    # warm the cache + surface any parse error before we start accepting requests
    print("ODL PM Dashboard — live server")
    print("  warming capacity data …")
    try:
        live_data(fresh=True)
        print("  ok.")
    except Exception:
        traceback.print_exc()
        sys.exit("  FAILED to compute capacity data — fix the error above first.")
    print(f"  /api/sync (pull from Asana): {'ENABLED' if can_sync() else 'unavailable'}"
          + ("" if os.environ.get('ASANA_TOKEN') else "  (set ASANA_TOKEN to actually pull)"))
    httpd = ThreadingHTTPServer((args.host, args.port), Handler)
    url = f"http://{args.host}:{args.port}/"
    print(f"\n  serving {url}")
    print("  open that URL for LIVE numbers (the page falls back to the baked\n"
          "  snapshot if it can't reach this server). Ctrl-C to stop.\n")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n  stopped.")


if __name__ == "__main__":
    main()
