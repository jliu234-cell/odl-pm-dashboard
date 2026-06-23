#!/usr/bin/env python3
"""Download the dashboard's external Google-Drive sources -- the capacity workbook
and the reflection PDFs -- with a service-account key, so the GitHub Actions
refresh can run in the cloud with no laptop.

One-time setup (so it keeps working after anyone leaves):
  1. Create a Google service account (Google Cloud Console > IAM > Service
     Accounts), enable the Drive API, and download a JSON key.
  2. Share the workbook Sheet AND the Reflection folder with the service
     account's email address (Viewer is enough).
  3. Put the JSON key in the repo as the GitHub secret GDRIVE_SA_KEY.

Usage:  python3 fetch_drive.py <out_dir>      (default out_dir: _sources)
Writes: <out>/ODL Project and Capacity Planning.xlsx   (Sheet exported to xlsx)
        <out>/Reflection/<file>                          (all files in the folder)
Then build.py is pointed at them via ODL_WORKBOOK / ODL_REFLECTION_DIR.

The file ids default to the current ODL files; override with env vars if they move.
"""
import os, sys, io, json

# Both sources live in the "NDL ODL" Google **shared drive** (org-owned, so they
# outlast any one account). Easiest access: add the service account as a Viewer
# MEMBER of that shared drive -- then it can read both without per-file shares.
#   workbook  = the live capacity Google Sheet
#   reflections = Shared drives/NDL ODL/ODL/ODL PM Folder/Project Reflection Reports 2025
SHEET_ID = os.environ.get("ODL_WORKBOOK_FILE_ID", "1fGHuQqu9iWC3TXjPr0hKBvCqz-8ZB9o73e2r0lbZFFE")
REFLECTION_FOLDER_ID = os.environ.get("ODL_REFLECTION_FOLDER_ID", "1R7A1ALplVfXd2XgUH_ojQPOGCHI1soTH")
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
WORKBOOK_NAME = "ODL Project and Capacity Planning.xlsx"


def drive():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build as gbuild
    key = os.environ.get("GDRIVE_SA_KEY")
    if not key:
        sys.exit("ERROR: set GDRIVE_SA_KEY to the service-account JSON key.")
    creds = service_account.Credentials.from_service_account_info(
        json.loads(key), scopes=["https://www.googleapis.com/auth/drive.readonly"])
    return gbuild("drive", "v3", credentials=creds)


def fetch_workbook(svc, out_dir):
    out = os.path.join(out_dir, WORKBOOK_NAME)
    # native Google Sheet -> export to .xlsx (build.py reads it with data_only=True)
    data = svc.files().export(fileId=SHEET_ID, mimeType=XLSX_MIME).execute()
    with open(out, "wb") as f:
        f.write(data)
    print(f"  workbook -> {out} ({len(data)} bytes)")
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
    fetch_workbook(svc, out_dir)
    try:
        fetch_reflections(svc, out_dir)
    except Exception as e:   # reflections are nice-to-have; never block the refresh
        print(f"  WARNING: reflections fetch failed ({e}); continuing without them.")
    print("done.")


if __name__ == "__main__":
    main()
