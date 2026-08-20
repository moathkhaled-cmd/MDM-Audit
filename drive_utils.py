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
    return build('drive', 'v3', credentials=creds)


def list_children(service, folder_id):
    """Yields (file_id, name, mime_type) for the direct children of folder_id."""
    page_token = None
    while True:
        resp = service.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="nextPageToken, files(id, name, mimeType)",
            pageToken=page_token,
            pageSize=1000,
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
    request = service.files().get_media(fileId=file_id)
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
        service.files().update(fileId=existing_id, media_body=media).execute()
        return existing_id
    file_metadata = {'name': remote_name, 'parents': [folder_id]}
    created = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    return created['id']
