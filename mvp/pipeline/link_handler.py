import os
import re
import time
import inspect
import urllib.parse
import requests

_DRIVE_PATTERNS = [
    r'drive\.google\.com',
    r'docs\.google\.com',
]
_SHAREPOINT_PATTERNS = [
    r'sharepoint\.com',
    r'onedrive\.live\.com',
    r'1drv\.ms',
    r'onedrive\.com',
]

# Kept modest on purpose: a single slow/hung link shouldn't be able to tie up a
# worker slot for minutes when a batch may have hundreds of these downloads to get
# through (see app.py's per-candidate document-link download step).
_DOWNLOAD_TIMEOUT = 30

# Google Drive throttles anonymous downloads when too many requests hit it at once
# (the bounded thread pool in Step 3b fires several gdown calls concurrently) — a
# genuinely public file can still come back as "restricted" under that burst. These
# are the backoff delays (seconds) between retries before treating it as a real failure.
_DRIVE_RETRY_DELAYS = (2, 5, 10)


def _classify_url(url):
    """Classify a URL without making any network request."""
    for pat in _DRIVE_PATTERNS:
        if re.search(pat, url, re.IGNORECASE):
            return 'drive'
    for pat in _SHAREPOINT_PATTERNS:
        if re.search(pat, url, re.IGNORECASE):
            return 'sharepoint'
    path = url.lower().split('?')[0].rstrip('/')
    if any(path.endswith(ext) for ext in ('.zip', '.rar', '.7z')):
        return 'zip_link'
    return 'unknown'


def _filename_from_response(response, url):
    """Derive a filename from Content-Disposition header or the URL path."""
    cd = response.headers.get('Content-Disposition', '')
    m = re.search(r'filename=["\']?([^"\';\r\n]+)', cd)
    if m:
        return m.group(1).strip().strip('"\'')
    name = url.split('?')[0].rstrip('/').split('/')[-1]
    return name if name else 'downloaded_file.bin'


# ── Google Drive helpers ─────────────────────────────────────────────────────

def _is_drive_folder_url(url):
    """Return True for Drive folder links — these need download_folder, not uc?id=."""
    return bool(re.search(r'drive\.google\.com/drive/folders/', url, re.IGNORECASE))


def _is_drive_docs_url(url):
    """
    Return True for Google Docs/Sheets/Slides/Forms URLs.
    These are not downloadable as a single binary file without export negotiation.
    """
    return bool(re.search(
        r'docs\.google\.com/(document|spreadsheets|presentation|forms|drawings)',
        url, re.IGNORECASE,
    ))


def _extract_drive_file_id(url):
    """
    Extract the file ID from common Google Drive file URL formats.
    Returns None if the URL doesn't encode a single downloadable file.

    Handled patterns:
      drive.google.com/file/d/{id}/...
      drive.google.com/open?id={id}
      drive.google.com/uc?id={id}
    """
    m = re.search(r'/file/d/([a-zA-Z0-9_-]+)', url)
    if m:
        return m.group(1)
    qs = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
    if 'id' in qs:
        return qs['id'][0]
    return None


def _looks_like_html(path):
    """Peek at the first 512 bytes — return True if the file is an HTML page."""
    try:
        with open(path, 'rb') as f:
            head = f.read(512).lower()
        return head.lstrip().startswith(b'<!doctype') or b'<html' in head[:64]
    except Exception:
        return False


def _try_drive_file_download(url, file_id, dest_dir):
    """
    Attempt an anonymous download of a Google Drive *file* via gdown.
    Works for "Anyone with the link" files — no credentials required.

    Returns (local_path, None) on success.
    Returns (None, reason_str) on any failure.

    ─── Future: account-specific access ────────────────────────────────────────
    Once IT provisions domain-wide delegation or the client shares the file with
    our service account email, replace this with:

        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload
        import io
        creds = service_account.Credentials.from_service_account_file(
            SA_KEY_PATH, scopes=['https://www.googleapis.com/auth/drive.readonly'])
        drive = build('drive', 'v3', credentials=creds)
        req   = drive.files().get_media(fileId=file_id)
        buf   = io.BytesIO()
        dl    = MediaIoBaseDownload(buf, req)
        done  = False
        while not done:
            _, done = dl.next_chunk()
        # write buf.getvalue() to dest_dir
    ─────────────────────────────────────────────────────────────────────────────
    """
    try:
        import gdown
    except ImportError:
        return None, 'gdown is not installed — run: pip install gdown'

    uc_url = f'https://drive.google.com/uc?id={file_id}'

    _file_params = {}
    try:
        _file_params = inspect.signature(gdown.download).parameters
    except Exception:
        pass
    file_kwargs = {'quiet': False}
    if 'fuzzy' in _file_params:
        file_kwargs['fuzzy'] = True

    last_reason = None
    for attempt, delay in enumerate((0,) + _DRIVE_RETRY_DELAYS):
        if delay:
            time.sleep(delay)

        try:
            result = gdown.download(uc_url, output=dest_dir + os.sep, **file_kwargs)
        except Exception as e:
            last_reason = f'gdown error: {e}'
            continue

        if result is None:
            last_reason = 'gdown returned None — file likely requires login or is restricted'
            continue

        if not os.path.exists(result) or os.path.getsize(result) == 0:
            last_reason = 'gdown produced an empty file'
            continue

        if _looks_like_html(result):
            try:
                os.remove(result)
            except Exception:
                pass
            last_reason = 'gdown returned an HTML page — file requires account-specific access'
            continue

        return result, None

    return None, last_reason


def _try_drive_folder_download(url, dest_dir):
    """
    Download an entire Google Drive *folder* using gdown.download_folder.
    Works for "Anyone with the link" public folders — no credentials required.

    The original folder structure is preserved inside dest_dir so the grouping
    agent can use candidate sub-folder names (ARS, emp_id, name) as matching signals.

    Returns (list_of_new_top_level_paths, None) on success.
    Returns ([], reason_str) on failure.

    ─── Future: account-specific / restricted folders ───────────────────────────
    Once IT provisions a service account with drive.readonly scope, replace the
    gdown.download_folder call with:

        from google.oauth2 import service_account
        from googleapiclient.discovery import build

        creds  = service_account.Credentials.from_service_account_file(
            SA_KEY_PATH, scopes=['https://www.googleapis.com/auth/drive.readonly'])
        drive  = build('drive', 'v3', credentials=creds)

        # Extract folder ID from URL, then enumerate:
        results = drive.files().list(
            q=f"'{folder_id}' in parents",
            fields='files(id, name, mimeType)',
        ).execute()
        for file in results.get('files', []):
            # download each file with get_media(fileId=file['id'])
    ─────────────────────────────────────────────────────────────────────────────
    """
    try:
        import gdown
    except ImportError:
        return [], 'gdown is not installed — run: pip install gdown'

    # Snapshot existing items in dest_dir so we can find what gdown creates
    before = set(os.listdir(dest_dir)) if os.path.isdir(dest_dir) else set()

    # Build kwargs based on what this installed gdown version actually accepts.
    # Avoids TypeError when running gdown < 5.x (no remaining_ok) or future versions
    # where the signature changes again.
    _params = {}
    try:
        _params = inspect.signature(gdown.download_folder).parameters
    except Exception:
        pass
    folder_kwargs = {'quiet': False}
    if 'use_cookies' in _params:
        folder_kwargs['use_cookies'] = False
    if 'remaining_ok' in _params:
        folder_kwargs['remaining_ok'] = True

    try:
        result = gdown.download_folder(url, output=dest_dir, **folder_kwargs)
    except Exception as e:
        err = str(e).lower()
        if any(x in err for x in ('permission', 'access denied', '403', 'cannot retrieve', 'forbidden')):
            return [], (
                'Folder requires account-specific Google access — not publicly shared. '
                'Ask the client to set sharing to "Anyone with the link", '
                'or share with our service account once provisioned.'
            )
        return [], f'gdown folder download error: {e}'

    if not result:
        return [], (
            'gdown returned empty — folder may require sign-in, be empty, '
            'or contain more files than gdown can enumerate anonymously.'
        )

    # Find the new top-level items gdown created (folders or files)
    after     = set(os.listdir(dest_dir)) if os.path.isdir(dest_dir) else set()
    new_items = [os.path.join(dest_dir, item) for item in (after - before)]

    if not new_items:
        return [], 'No new files or folders found after download — folder may be empty'

    print(f'[LinkHandler] Drive folder downloaded: {len(result)} file(s) in {len(new_items)} top-level item(s)')
    return new_items, None


# ── SharePoint / OneDrive helpers ────────────────────────────────────────────

def _try_sharepoint_download(url, dest_dir):
    """
    Attempt an anonymous download of a SharePoint / OneDrive file by appending
    the download=1 query parameter. If the server returns HTML (login redirect),
    the attempt is treated as a failure.

    Returns (local_path, None) on success.
    Returns (None, reason_str) on failure.

    ─── Future: Microsoft Graph API ────────────────────────────────────────────
    Once an Azure AD app registration with Files.Read.All (or Sites.Read.All)
    has admin consent from the client tenant, replace this with:

        import msal, requests
        app   = msal.ConfidentialClientApplication(
            CLIENT_ID,
            authority=f'https://login.microsoftonline.com/{TENANT_ID}',
            client_credential=CLIENT_SECRET,
        )
        token = app.acquire_token_for_client(
            scopes=['https://graph.microsoft.com/.default'])
        headers  = {'Authorization': f'Bearer {token["access_token"]}'}
        response = requests.get(graph_download_url, headers=headers, stream=True)
    ─────────────────────────────────────────────────────────────────────────────
    """
    separator    = '&' if '?' in url else '?'
    download_url = url + separator + 'download=1'

    try:
        resp = requests.get(download_url, timeout=_DOWNLOAD_TIMEOUT, stream=True, allow_redirects=True)
        resp.raise_for_status()

        ct = resp.headers.get('Content-Type', '')
        if ct.startswith('text/html'):
            return None, (
                'Server returned HTML — tenant requires login or blocks anonymous sharing. '
                'Ask the client to set the link to "Anyone with the link can view", '
                'or request Graph API credentials from the tenant admin.'
            )

        fname = _filename_from_response(resp, url)
        dest  = os.path.join(dest_dir, fname)

        with open(dest, 'wb') as f:
            for chunk in resp.iter_content(chunk_size=65536):
                if chunk:
                    f.write(chunk)

        if os.path.getsize(dest) == 0:
            os.remove(dest)
            return None, 'Downloaded file is empty — SharePoint link may require authentication'

        return dest, None

    except Exception as e:
        return None, f'Request failed: {e}'


# ── Public API ───────────────────────────────────────────────────────────────

def handle_links(links, run_upload):
    """
    Process links produced by email_parser.

    Drive folder ("anyone with the link") : gdown.download_folder — full tree downloaded.
    Drive file   ("anyone with the link") : gdown.download.
    Drive folder/file (restricted)        : flagged pending with clear reason for CST.
    Google Docs/Sheets/Slides             : flagged immediately (no single-file download).
    SharePoint/OneDrive ("anyone")        : requests GET with ?download=1.
    SharePoint (restricted / tenant-only) : flagged pending with clear reason for CST.
    zip_link / unknown                    : plain HTTP GET.

    Args:
        links      : list of {url, type} dicts from email_parser output
        run_upload : per-run upload directory — downloaded files/folders land here

    Returns:
        downloaded_paths : list of local file or directory paths ready for extractor
        pending_entries  : list of {url, type, status, reason} for assembler manifest
        errors           : list of {url, reason} for failed plain-HTTP downloads
    """
    downloaded_paths = []
    pending_entries  = []
    errors           = []

    for entry in links:
        url       = (entry.get('url') or '').strip()
        link_type = entry.get('type') or _classify_url(url)

        if not url:
            continue

        # ── Google Drive ─────────────────────────────────────────────────────
        if link_type == 'drive':

            # Google Docs/Sheets/Slides — cannot be downloaded as a single binary
            if _is_drive_docs_url(url):
                pending_entries.append({
                    'url':    url,
                    'type':   link_type,
                    'status': 'link_pending_manual_fetch',
                    'reason': (
                        'Google Docs/Sheets/Slides link — cannot be downloaded as a single file. '
                        'Ask the client to export as PDF/XLSX and re-share.'
                    ),
                })
                print(f'[LinkHandler] Pending (Google Docs/Sheets/Slides): {url}')
                continue

            # Drive folder — use download_folder (works for public "anyone with the link" folders)
            if _is_drive_folder_url(url):
                new_items, fail_reason = _try_drive_folder_download(url, run_upload)
                if new_items:
                    downloaded_paths.extend(new_items)
                    print(f'[LinkHandler] Drive folder downloaded: {url} → {len(new_items)} item(s)')
                else:
                    pending_entries.append({
                        'url':    url,
                        'type':   link_type,
                        'status': 'link_pending_manual_fetch',
                        'reason': f'Drive folder download failed — CST to collect manually. Detail: {fail_reason}',
                    })
                    print(f'[LinkHandler] Drive folder failed ({fail_reason}): {url}')
                continue

            # Drive single file
            file_id = _extract_drive_file_id(url)
            if not file_id:
                pending_entries.append({
                    'url':    url,
                    'type':   link_type,
                    'status': 'link_pending_manual_fetch',
                    'reason': 'Could not extract Google Drive file ID from URL — CST to collect manually',
                })
                print(f'[LinkHandler] Pending (no file ID extractable): {url}')
                continue

            local_path, fail_reason = _try_drive_file_download(url, file_id, run_upload)
            if local_path:
                downloaded_paths.append(local_path)
                print(f'[LinkHandler] Drive file downloaded: {url} → {os.path.basename(local_path)}')
            else:
                pending_entries.append({
                    'url':    url,
                    'type':   link_type,
                    'status': 'link_pending_manual_fetch',
                    'reason': f'Anonymous download failed — CST to collect manually. Detail: {fail_reason}',
                })
                print(f'[LinkHandler] Drive file failed ({fail_reason}): {url}')

        # ── SharePoint / OneDrive ────────────────────────────────────────────
        elif link_type == 'sharepoint':
            local_path, fail_reason = _try_sharepoint_download(url, run_upload)
            if local_path:
                downloaded_paths.append(local_path)
                print(f'[LinkHandler] SharePoint downloaded: {url} → {os.path.basename(local_path)}')
            else:
                pending_entries.append({
                    'url':    url,
                    'type':   link_type,
                    'status': 'link_pending_manual_fetch',
                    'reason': f'Anonymous download failed — CST to collect manually. Detail: {fail_reason}',
                })
                print(f'[LinkHandler] SharePoint failed ({fail_reason}): {url}')

        # ── Direct ZIP / unknown — plain HTTP GET ────────────────────────────
        else:
            try:
                resp = requests.get(url, timeout=_DOWNLOAD_TIMEOUT, stream=True)
                resp.raise_for_status()

                # A webpage/social link (e.g. an email signature's Facebook or
                # company-website URL) returns text/html, not a document. Without
                # this check the raw page markup gets saved and mistaken for a file.
                ct = resp.headers.get('Content-Type', '')
                if ct.startswith('text/html'):
                    raise ValueError(
                        'Server returned an HTML page, not a downloadable file — '
                        'likely a webpage or social-media link rather than a document.'
                    )

                fname = _filename_from_response(resp, url)
                dest  = os.path.join(run_upload, fname)

                with open(dest, 'wb') as f:
                    for chunk in resp.iter_content(chunk_size=65536):
                        if chunk:
                            f.write(chunk)

                if os.path.getsize(dest) == 0:
                    os.remove(dest)
                    raise ValueError('Downloaded file is empty — link may require authentication.')

                downloaded_paths.append(dest)
                print(f'[LinkHandler] Downloaded: {url} → {fname}')

            except Exception as e:
                errors.append({'url': url, 'reason': str(e)})
                print(f'[LinkHandler] Failed to download {url}: {e}')

    return downloaded_paths, pending_entries, errors
