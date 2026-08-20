# ============================================================================
# drive_utils.py — thin Google Drive API layer for the GitHub Actions runner.
# ============================================================================
# GitHub Actions has no filesystem mount of your Drive the way Colab does, so
# every read/write goes through the Drive API explicitly via a service
# account. This module keeps that entirely separate from the audit logic in
# run_daily_audit.py, which just calls list_pdfs_recursive/download/upload.
# ============================================================================

import io
import os

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

SCOPES = ['https://www.googleapis.com/auth/drive']


def get_drive_service():
    creds_path = os.environ['GDRIVE_SERVICE_ACCOUNT_FILE']
    creds = service_account.Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds), creds.service_account_email


def check_folder_access(service, folder_id, label, service_account_email):
    """Fetches folder_id's own metadata (not its children) so a bad ID or a
    missing share fails fast with a clear message instead of a raw 404 deep
    inside a list() call."""
    try:
        meta = service.files().get(
            fileId=folder_id,
            fields="id, name, mimeType, driveId",
            supportsAllDrives=True,
        ).execute()
    except Exception as e:
        raise SystemExit(
            f"\n[CONFIG ERROR] Could not access {label} (id: {folder_id!r}).\n"
            f"  Service account: {service_account_email}\n"
            f"  Raw error: {e}\n\n"
            f"  Checklist:\n"
            f"  1. In Drive, right-click the folder -> Share -> add "
            f"{service_account_email} as Editor (this is the #1 cause).\n"
            f"  2. Confirm the ID itself: open the folder in a browser, copy only "
            f"the segment after '/folders/' in the URL — not the whole URL, and "
            f"watch for a trailing '/' or '?usp=sharing' getting pasted along with it.\n"
            f"  3. If this folder lives inside a Shared Drive (Team Drive) rather "
            f"than My Drive, the service account also needs to be added as a "
            f"member of that Shared Drive itself, not just the folder.\n"
        )
    if meta.get('mimeType') != 'application/vnd.google-apps.folder':
        raise SystemExit(f"\n[CONFIG ERROR] {label} (id: {folder_id!r}) is not a folder "
                          f"(mimeType={meta.get('mimeType')}). Double check the ID.\n")
    print(f"  OK: {label} -> '{meta.get('name')}'")
    return meta


def list_children(service, folder_id):
    """Yields (file_id, name, mime_type) for the direct children of folder_id."""
    page_token = None
    while True:
        resp = service.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="nextPageToken, files(id, name, mimeType)",
            pageToken=page_token,
            pageSize=1000,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        for f in resp.get('files', []):
            yield f['id'], f['name'], f['mimeType']
        page_token = resp.get('nextPageToken')
        if not page_token:
            break


def list_pdfs_recursive(service, root_folder_id, path_prefix=""):
    """Yields (file_id, virtual_path) for every PDF under root_folder_id,
    walking subfolders. virtual_path mimics a local relative path
    (e.g. 'Toyota/Land Cruiser/T5_L_Brochure_2026.pdf') so the existing
    tokenize()-based matching logic works unchanged against it."""
    for file_id, name, mime in list_children(service, root_folder_id):
        vpath = f"{path_prefix}/{name}" if path_prefix else name
        if mime == 'application/vnd.google-apps.folder':
            yield from list_pdfs_recursive(service, file_id, vpath)
        elif name.lower().endswith('.pdf'):
            yield file_id, vpath


def find_file_id_by_name(service, folder_id, name):
    for file_id, fname, _mime in list_children(service, folder_id):
        if fname == name:
            return file_id
    return None


def download_bytes(service, file_id) -> bytes:
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _status, done = downloader.next_chunk()
    return buf.getvalue()


def download_to_path(service, file_id, local_path):
    os.makedirs(os.path.dirname(local_path) or '.', exist_ok=True)
    data = download_bytes(service, file_id)
    with open(local_path, 'wb') as f:
        f.write(data)
    return local_path


def upload_or_update(service, folder_id, local_path, remote_name=None, mime_type=None):
    """Creates the file in folder_id if it doesn't exist there yet, otherwise
    overwrites the existing one in place — so daily re-runs replace the same
    file on Drive instead of piling up dated duplicates."""
    remote_name = remote_name or os.path.basename(local_path)
    existing_id = find_file_id_by_name(service, folder_id, remote_name)
    media = MediaFileUpload(local_path, mimetype=mime_type, resumable=True)
    if existing_id:
        service.files().update(fileId=existing_id, media_body=media, supportsAllDrives=True).execute()
        return existing_id
    file_metadata = {'name': remote_name, 'parents': [folder_id]}
    created = service.files().create(body=file_metadata, media_body=media, fields='id', supportsAllDrives=True).execute()
    return created['id']
