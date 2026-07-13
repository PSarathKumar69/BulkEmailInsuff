import os
import re
import shutil
import tarfile
import zipfile
import email as _email_stdlib
import email.policy

# Extensions that trigger recursive unpacking
_ARCHIVE_EXTS       = {'.zip', '.rar', '.7z', '.tar', '.tgz', '.tbz2', '.txz'}
_EMAIL_EXTS         = {'.eml', '.msg'}

# Outlook embeds signature logos and HTML body images as attachments named image001.png etc.
# These are never candidate documents — skip them during .msg/.eml extraction.
_INLINE_IMAGE_RE = re.compile(r'^image\d+\.(png|jpg|jpeg|gif|bmp|tif|tiff)$', re.IGNORECASE)


def _is_tar(path):
    """True for .tar and compound extensions like .tar.gz / .tar.bz2 / .tar.xz."""
    name = os.path.basename(path).lower()
    return name.endswith(('.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz2', '.tar.xz', '.txz'))
# Extensions copied as-is (no unpacking)
_LOOSE_EXTS         = {
    '.pdf', '.png', '.jpg', '.jpeg',
    '.docx', '.tiff', '.tif', '.heic',
    '.doc', '.xlsx', '.xls', '.xlsb', '.csv',
    '.pptx', '.ppt', '.txt', '.xml',
}

_MAX_DEPTH = 10  # guard against pathological nesting

# ── unrar availability check (done once) ────────────────────────────────────

_UNRAR_CHECKED   = False
_UNRAR_AVAILABLE = False


def _check_unrar():
    global _UNRAR_CHECKED, _UNRAR_AVAILABLE
    if not _UNRAR_CHECKED:
        _UNRAR_AVAILABLE = (
            shutil.which('unrar') is not None
            or shutil.which('unrar-free') is not None
        )
        _UNRAR_CHECKED = True
    return _UNRAR_AVAILABLE


# ── public API ───────────────────────────────────────────────────────────────

def extract_all(attachment_paths, extract_root):
    """
    Process a list of file paths (or directories) into extract_root.
      Directory          → tree is copied preserving structure; archives within are unpacked
      .zip / .rar / .7z → recursively unpacked into a named sub-folder
      .eml / .msg       → body saved as email_body.txt + attachments extracted
      anything else     → copied flat into extract_root (no unpacking)

    Returns list of error dicts: [{identifier, path, reason}]
    """
    os.makedirs(extract_root, exist_ok=True)
    errors = []
    for path in attachment_paths:
        if os.path.isdir(path):
            _handle_directory(path, extract_root, errors, depth=0)
        else:
            _handle_file(path, extract_root, errors, depth=0)
    return errors


# ── internal helpers ─────────────────────────────────────────────────────────

def _handle_file(path, dest_dir, errors, depth):
    """Route a single file to the right handler."""
    if depth > _MAX_DEPTH:
        errors.append({
            'identifier': os.path.basename(path),
            'path':       path,
            'reason':     f'Skipped — maximum nesting depth ({_MAX_DEPTH}) exceeded',
        })
        return

    ext = os.path.splitext(path)[1].lower()

    if ext == '.zip':
        _extract_zip(path, dest_dir, errors, depth)
    elif ext == '.rar':
        _extract_rar(path, dest_dir, errors, depth)
    elif ext == '.7z':
        _extract_7z(path, dest_dir, errors, depth)
    elif _is_tar(path):
        _extract_tar(path, dest_dir, errors, depth)
    elif ext == '.eml':
        _extract_eml(path, dest_dir, errors, depth)
    elif ext == '.msg':
        _extract_msg(path, dest_dir, errors, depth)
    else:
        # Loose document — copy flat; never unpack
        os.makedirs(dest_dir, exist_ok=True)
        _copy_safe(path, dest_dir)


def _handle_directory(directory, dest_dir, errors, depth):
    """
    Copy a pre-downloaded directory tree into dest_dir preserving folder structure.
    Archives found within are recursively unpacked; loose files are copied flat
    into their containing sub-folder (not into dest_dir root).

    Used for Drive folder downloads where gdown preserves the original structure:
        run_upload/ClientFolder/6197-001269_Arun/aadhaar.pdf
    becomes:
        run_extract/ClientFolder/6197-001269_Arun/aadhaar.pdf
    so the grouping agent can use folder names as matching signals.
    """
    if depth > _MAX_DEPTH:
        errors.append({
            'identifier': os.path.basename(directory),
            'path':       directory,
            'reason':     f'Skipped — maximum nesting depth ({_MAX_DEPTH}) exceeded',
        })
        return

    folder_name = os.path.basename(directory.rstrip(os.sep)) or 'download'
    out_dir     = os.path.join(dest_dir, folder_name)
    os.makedirs(out_dir, exist_ok=True)

    try:
        entries = [e for e in os.listdir(directory) if not e.startswith('.')]
    except Exception as e:
        errors.append({'identifier': folder_name, 'path': directory, 'reason': str(e)})
        return

    for name in entries:
        if name in ('__MACOSX', '__pycache__'):
            continue
        full_path = os.path.join(directory, name)
        if os.path.isdir(full_path):
            _handle_directory(full_path, out_dir, errors, depth + 1)
        else:
            _handle_file(full_path, out_dir, errors, depth + 1)


def _subdir(dest_dir, filename):
    """Return a sub-folder path named after the file (without extension)."""
    name = os.path.splitext(os.path.basename(filename))[0]
    return os.path.join(dest_dir, name)


def _copy_safe(src, dest_dir):
    """Copy src into dest_dir, adding a suffix on collision."""
    os.makedirs(dest_dir, exist_ok=True)
    fname = os.path.basename(src)
    dest  = os.path.join(dest_dir, fname)
    if os.path.abspath(src) == os.path.abspath(dest):
        return  # already in place
    if os.path.exists(dest):
        base, ext = os.path.splitext(fname)
        n = 2
        while os.path.exists(os.path.join(dest_dir, f'{base}_{n}{ext}')):
            n += 1
        dest = os.path.join(dest_dir, f'{base}_{n}{ext}')
    shutil.copy2(src, dest)


# ── archive handlers ─────────────────────────────────────────────────────────

def _extract_zip(path, dest_dir, errors, depth):
    name    = os.path.splitext(os.path.basename(path))[0]
    out_dir = _subdir(dest_dir, path)
    os.makedirs(out_dir, exist_ok=True)
    try:
        with zipfile.ZipFile(path, 'r') as zf:
            abs_out = os.path.abspath(out_dir)
            for member in zf.namelist():
                if member.startswith('__MACOSX'):
                    continue
                dest_path = os.path.abspath(os.path.join(abs_out, member))
                if not (dest_path == abs_out or dest_path.startswith(abs_out + os.sep)):
                    errors.append({
                        'identifier': name,
                        'path':       path,
                        'reason':     f'Skipped unsafe ZIP entry (path traversal): {member}',
                    })
                    continue
                zf.extract(member, out_dir)
        _recurse_dir(out_dir, errors, depth + 1)
    except RuntimeError as e:
        reason = 'password_protected' if 'password' in str(e).lower() else str(e)
        errors.append({'identifier': name, 'path': path, 'reason': reason})
    except zipfile.BadZipFile:
        errors.append({'identifier': name, 'path': path, 'reason': 'not a valid ZIP file'})
    except Exception as e:
        errors.append({'identifier': name, 'path': path, 'reason': str(e)})


def _extract_rar(path, dest_dir, errors, depth):
    name = os.path.splitext(os.path.basename(path))[0]
    if not _check_unrar():
        errors.append({
            'identifier': name,
            'path':       path,
            'reason': (
                'RAR extraction requires the system unrar binary which was not found. '
                'Install it with: apt install unrar  (Ubuntu/Debian) or '
                'brew install rar  (macOS). '
                'This file has been skipped — extract it manually and re-run.'
            ),
        })
        return

    out_dir = _subdir(dest_dir, path)
    os.makedirs(out_dir, exist_ok=True)
    try:
        import rarfile
        with rarfile.RarFile(path, 'r') as rf:
            abs_out = os.path.abspath(out_dir)
            for info in rf.infolist():
                dest_path = os.path.abspath(os.path.join(abs_out, info.filename))
                if not (dest_path == abs_out or dest_path.startswith(abs_out + os.sep)):
                    errors.append({
                        'identifier': name,
                        'path':       path,
                        'reason':     f'Skipped unsafe RAR entry (path traversal): {info.filename}',
                    })
                    continue
                rf.extract(info, out_dir)
        _recurse_dir(out_dir, errors, depth + 1)
    except Exception as e:
        reason = 'password_protected' if 'password' in str(e).lower() else str(e)
        errors.append({'identifier': name, 'path': path, 'reason': reason})


def _extract_7z(path, dest_dir, errors, depth):
    name    = os.path.splitext(os.path.basename(path))[0]
    out_dir = _subdir(dest_dir, path)
    os.makedirs(out_dir, exist_ok=True)
    try:
        import py7zr
        with py7zr.SevenZipFile(path, mode='r') as sz:
            sz.extractall(path=out_dir)
        _recurse_dir(out_dir, errors, depth + 1)
    except Exception as e:
        reason = 'password_protected' if 'password' in str(e).lower() else str(e)
        errors.append({'identifier': name, 'path': path, 'reason': reason})


def _extract_tar(path, dest_dir, errors, depth):
    fname = os.path.basename(path)
    # Strip compound extensions cleanly for folder naming
    stem = fname
    for suffix in ('.tar.gz', '.tar.bz2', '.tar.xz', '.tgz', '.tbz2', '.txz', '.tar'):
        if stem.lower().endswith(suffix):
            stem = stem[:-len(suffix)]
            break
    out_dir = os.path.join(dest_dir, stem or fname)
    os.makedirs(out_dir, exist_ok=True)
    try:
        with tarfile.open(path, 'r:*') as tf:
            abs_out = os.path.abspath(out_dir)
            for member in tf:
                dest_path = os.path.abspath(os.path.join(abs_out, member.name))
                if not (dest_path == abs_out or dest_path.startswith(abs_out + os.sep)):
                    errors.append({
                        'identifier': fname,
                        'path':       path,
                        'reason':     f'Skipped unsafe TAR entry (path traversal): {member.name}',
                    })
                    continue
                try:
                    tf.extract(member, out_dir)
                except Exception as ex:
                    errors.append({
                        'identifier': fname,
                        'path':       path,
                        'reason':     f'Failed to extract TAR entry {member.name}: {ex}',
                    })
        _recurse_dir(out_dir, errors, depth + 1)
    except Exception as e:
        reason = 'password_protected' if 'password' in str(e).lower() else str(e)
        errors.append({'identifier': fname, 'path': path, 'reason': reason})


# ── email container handlers ─────────────────────────────────────────────────

def _extract_eml(path, dest_dir, errors, depth):
    name    = os.path.splitext(os.path.basename(path))[0]
    out_dir = _subdir(dest_dir, path)
    os.makedirs(out_dir, exist_ok=True)
    try:
        with open(path, 'rb') as f:
            msg = _email_stdlib.message_from_binary_file(f, policy=_email_stdlib.policy.default)

        # Save plain-text body so the grouping agent can see any inline candidate data
        body = _eml_body_text(msg)
        if body.strip():
            with open(os.path.join(out_dir, 'email_body.txt'), 'w', encoding='utf-8', errors='replace') as bf:
                bf.write(body)

        # Extract and recurse into each attachment
        for part in msg.iter_attachments():
            fname   = os.path.basename(part.get_filename() or 'attachment') or 'attachment'
            payload = part.get_payload(decode=True)
            if not payload:
                continue
            # Skip Outlook inline embedded images (signature logos, HTML body images)
            if _INLINE_IMAGE_RE.match(fname):
                continue
            # Skip if the same filename already exists in the parent folder —
            # the client re-attached files they already submitted as direct ZIP entries.
            if os.path.exists(os.path.join(dest_dir, fname)):
                continue
            att_path = os.path.join(out_dir, fname)
            with open(att_path, 'wb') as af:
                af.write(payload)
            _handle_file(att_path, out_dir, errors, depth + 1)
            # If the attachment was an archive/email it has already been unpacked
            # into its own sub-folder; remove the raw file to avoid duplication.
            _try_remove_if_container(att_path)

    except Exception as e:
        errors.append({'identifier': name, 'path': path, 'reason': str(e)})


def _extract_msg(path, dest_dir, errors, depth):
    name    = os.path.splitext(os.path.basename(path))[0]
    out_dir = _subdir(dest_dir, path)
    os.makedirs(out_dir, exist_ok=True)
    try:
        import extract_msg as _extract_msg_lib
        msg = _extract_msg_lib.Message(path)

        body = (msg.body or '').strip()
        if body:
            with open(os.path.join(out_dir, 'email_body.txt'), 'w', encoding='utf-8', errors='replace') as bf:
                bf.write(body)

        for att in msg.attachments:
            fname = os.path.basename((att.longFilename or att.shortFilename or 'attachment').strip()) or 'attachment'
            if not fname:
                continue
            # Skip Outlook inline embedded images (signature logos, HTML body images)
            if _INLINE_IMAGE_RE.match(fname):
                continue
            # Skip if the same filename already exists in the parent folder —
            # the client re-attached files they already submitted as direct ZIP entries.
            if os.path.exists(os.path.join(dest_dir, fname)):
                continue
            att_path = os.path.join(out_dir, fname)
            # extract_msg saves via att.save(); fall back to writing bytes directly
            try:
                att.save(customPath=out_dir, customFilename=fname)
            except Exception:
                data = getattr(att, 'data', None)
                if data:
                    with open(att_path, 'wb') as f:
                        f.write(data)
                else:
                    continue
            if os.path.exists(att_path):
                _handle_file(att_path, out_dir, errors, depth + 1)
                _try_remove_if_container(att_path)

        msg.close()
    except Exception as e:
        errors.append({'identifier': name, 'path': path, 'reason': str(e)})


# ── recursive directory scan ─────────────────────────────────────────────────

def _recurse_dir(directory, errors, depth):
    """
    Walk an already-extracted directory.
    Any archives or email containers found inside are unpacked in place
    (into a same-named sub-folder) and the original container file is removed.
    Loose files are already in their correct location and are left untouched.
    """
    if depth > _MAX_DEPTH:
        return

    for dirpath, dirs, filenames in os.walk(directory):
        # Skip macOS resource-fork dirs and Python cache dirs
        dirs[:] = [
            d for d in dirs
            if not d.startswith('.') and d not in ('__MACOSX', '__pycache__')
        ]
        for fname in filenames:
            if fname.startswith('.'):
                continue
            ext   = os.path.splitext(fname)[1].lower()
            fpath = os.path.join(dirpath, fname)
            if ext in _ARCHIVE_EXTS or ext in _EMAIL_EXTS or _is_tar(fpath):
                _handle_file(fpath, dirpath, errors, depth)
                _try_remove_if_container(fpath)


# ── utilities ────────────────────────────────────────────────────────────────

def _eml_body_text(msg):
    """Extract all text/plain parts from an email.message object."""
    parts = []
    for part in msg.walk():
        if part.get_content_type() == 'text/plain':
            payload = part.get_payload(decode=True)
            if payload:
                charset = part.get_content_charset() or 'utf-8'
                parts.append(payload.decode(charset, errors='replace'))
    return '\n'.join(parts)


def _try_remove_if_container(path):
    """Remove a file if it is an archive or email container (already unpacked)."""
    ext = os.path.splitext(path)[1].lower()
    if ext in _ARCHIVE_EXTS or ext in _EMAIL_EXTS or _is_tar(path):
        try:
            os.remove(path)
        except Exception:
            pass
