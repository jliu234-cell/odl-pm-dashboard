#!/usr/bin/env python3
"""Download the dashboard's external Google-Drive sources -- the Capacity
Allocations Sheet and the reflection PDFs -- with a service-account key, so the
GitHub Actions refresh can run in the cloud with no laptop.

One-time setup (so it keeps working after anyone leaves):
  1. Create a Google service account (Google Cloud Console > IAM > Service
     Accounts), enable the Drive API, and download a JSON key.
  2. Share the Capacity Allocations Sheet AND the Reflection folder with the
     service account's email address (Viewer is enough).
  3. Put the JSON key in the repo as the GitHub secret GDRIVE_SA_KEY.

Usage:  python3 fetch_drive.py <out_dir>      (default out_dir: _sources)
Writes: <out>/capacity_allocations.csv   (Capacity Allocations Sheet -> csv; the capacity source)
        <out>/Reflection/<file>           (all files in the folder)
Then build.py is pointed at them via ODL_CAPACITY_CSV / ODL_REFLECTION_DIR.

The file ids default to the current ODL files; override with env vars if they move.
"""
import os, sys, io, json

# Both sources live in the "NDL ODL" Google **shared drive** (org-owned, so they
# outlast any one account). Easiest access: add the service account as a Viewer
# MEMBER of that shared drive -- then it can read both without per-file shares.
#   capacity    = the live "Capacity Allocations" Google Sheet
#   reflections = Shared drives/NDL ODL/ODL/ODL PM Folder/Project Reflection Reports 2025
REFLECTION_FOLDER_ID = os.environ.get("ODL_REFLECTION_FOLDER_ID", "1R7A1ALplVfXd2XgUH_ojQPOGCHI1soTH")
# the live "Capacity Allocations" sheet — now the source of capacity numbers
CAPACITY_SHEET_ID = os.environ.get("ODL_CAPACITY_SHEET_ID", "1YD9b8vLnglbA5pmFO6HsE-bvMq7wluHv1P1wY4FCujw")
CAPACITY_NAME = "capacity_allocations.csv"


def drive():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build as gbuild
    key = os.environ.get("GDRIVE_SA_KEY")
    if not key:
        sys.exit("ERROR: set GDRIVE_SA_KEY to the service-account JSON key.")
    creds = service_account.Credentials.from_service_account_info(
        json.loads(key), scopes=["https://www.googleapis.com/auth/drive.readonly"])
    return gbuild("drive", "v3", credentials=creds)


def fetch_capacity(svc, out_dir):
    out = os.path.join(out_dir, CAPACITY_NAME)
    # native Google Sheet -> export the first tab to CSV (build.py reads it via
    # ODL_CAPACITY_CSV).
    # NOTE: exports the sheet's first/default tab; the Capacity Allocations data
    # must be the first tab (gid 1025292964). If a second tab is ever added
    # before it, switch to the gviz export URL with &gid=1025292964.
    data = svc.files().export(fileId=CAPACITY_SHEET_ID, mimeType="text/csv").execute()
    with open(out, "wb") as f:
        f.write(data)
    print(f"  capacity -> {out} ({len(data)} bytes)")
    return out


def fetch_reflections(svc, out_dir):
    from googleapiclient.http import MediaIoBaseDownload
    rdir = os.path.join(out_dir, "Reflection")
    os.makedirs(rdir, exist_ok=True)
    n, tok = 0, None
    while True:
        resp = svc.files().list(
            q=f"'{REFLECTION_FOLDER_ID}' in parents and trashed=false",
            fields="nextPageToken, files(id,name,mimeType)", pageSize=200, pageToken=tok,
            supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
        for f in resp.get("files", []):
            if f["mimeType"] == "application/vnd.google-apps.folder":
                continue
            buf = io.BytesIO()
            dl = MediaIoBaseDownload(buf, svc.files().get_media(fileId=f["id"], supportsAllDrives=True))
            done = False
            while not done:
                _, done = dl.next_chunk()
            with open(os.path.join(rdir, f["name"]), "wb") as o:
                o.write(buf.getvalue())
            n += 1
        tok = resp.get("nextPageToken")
        if not tok:
            break
    print(f"  reflections -> {rdir} ({n} files)")
    return rdir


def main():
    out_dir = sys.argv[1] if len(sys.argv) > 1 else "_sources"
    os.makedirs(out_dir, exist_ok=True)
    svc = drive()
    try:
        fetch_capacity(svc, out_dir)
    except Exception as e:   # capacity CSV is committed; a fetch failure isn't fatal
        print(f"  WARNING: capacity fetch failed ({e}); using the committed CSV.")
    try:
        fetch_reflections(svc, out_dir)
    except Exception as e:   # reflections are nice-to-have; never block the refresh
        print(f"  WARNING: reflections fetch failed ({e}); continuing without them.")
    print("done.")


if __name__ == "__main__":
    main()
