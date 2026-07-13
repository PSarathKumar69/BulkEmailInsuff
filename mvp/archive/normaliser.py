import os
import re

GENERIC_NAMES = {
    'screenshot', 'scan', 'document', 'file', 'image', 'untitled',
    'img', 'page', 'photo', 'picture', 'attachment', 'copy',
    'temp', 'test', 'sample', 'new', 'unknown',
}


def _is_generic(name):
    base = name.lower().strip()
    stripped = re.sub(r'\d+$', '', base)
    return base in GENERIC_NAMES or stripped in GENERIC_NAMES


def _norm_id(s):
    s = re.sub(r'[\s\-_\.]', '', str(s).lower())
    return s.lstrip('0') or '0'


def _find_real_root(root):
    """
    Detect and unwrap single-folder wrapper layers so the normaliser operates
    on the true candidate level, not a client-named parent folder.

    Rules (applied recursively until stable):
      - Root has exactly 1 folder and 0 files:
          * Inner has only FILES (≥1) and no sub-folders
            → files-per-candidate pattern wrapped in one folder → unwrap
          * Inner has only FOLDERS and no files
            → folder-per-candidate pattern wrapped in one folder → unwrap + recurse
          * Inner has both files AND folders
            → mixed content (one candidate's docs + sub-archives) → stop, don't unwrap
      - Root has multiple entries or already has files → return as-is
    """
    try:
        entries = [e for e in os.listdir(root) if not e.startswith('.')]
    except Exception:
        return root

    dirs  = [e for e in entries if os.path.isdir(os.path.join(root, e))]
    files = [e for e in entries if os.path.isfile(os.path.join(root, e))]

    if len(dirs) == 1 and len(files) == 0:
        inner = os.path.join(root, dirs[0])
        try:
            inner_entries = [e for e in os.listdir(inner) if not e.startswith('.')]
        except Exception:
            return root

        inner_dirs  = [e for e in inner_entries if os.path.isdir(os.path.join(inner, e))]
        inner_files = [e for e in inner_entries if os.path.isfile(os.path.join(inner, e))]

        if inner_files and not inner_dirs:
            # Files-per-candidate in a wrapper folder → unwrap one level
            return inner

        if inner_dirs and not inner_files:
            # Folder-per-candidate in a wrapper folder → unwrap and recurse
            return _find_real_root(inner)

        # Mixed content (files + folders) → treat as one candidate's folder, stop

    return root


def _iter_files(folder):
    for dirpath, _, filenames in os.walk(folder):
        for f in filenames:
            yield os.path.join(dirpath, f)


def _flag_duplicates(units):
    """
    Flag units whose normalised identifiers are identical.
    Two folders named EMP001_RahulSharma and EMP001_RahulSharma_updated both
    normalise to the same key and would map to the same Excel row — flag both.
    """
    seen = {}
    for unit in units:
        if unit.get('flag'):
            continue
        key = _norm_id(unit['identifier'])
        seen.setdefault(key, []).append(unit)

    for key, group in seen.items():
        if len(group) > 1:
            for unit in group:
                unit['flag'] = 'duplicate_id'
                unit['notes'] = (
                    f'{len(group)} units share the same normalised identifier — '
                    'manual review needed'
                )


def build_candidate_units(extract_root, extraction_errors=None):
    units = []

    # Pre-seed units for ZIPs that failed to extract (password-protected, corrupt, etc.)
    for err in (extraction_errors or []):
        reason = err.get('reason', 'extraction_failed')
        flag = 'password_protected' if reason == 'password_protected' else 'extraction_failed'
        note = (
            'Password protected — cannot extract without password'
            if flag == 'password_protected'
            else f'ZIP extraction failed: {reason}'
        )
        units.append({
            'identifier': err.get('identifier', 'unknown'),
            'type': 'file',
            'path': err.get('path', ''),
            'flag': flag,
            'notes': note,
        })

    real_root = _find_real_root(extract_root)

    try:
        entries = sorted(e for e in os.listdir(real_root) if not e.startswith('.'))
    except FileNotFoundError:
        return units

    for entry in entries:
        full_path = os.path.join(real_root, entry)

        if os.path.isdir(full_path):
            inner_files = list(_iter_files(full_path))
            if not inner_files:
                units.append({
                    'identifier': entry,
                    'type': 'folder',
                    'path': full_path + os.sep,
                    'flag': 'empty_folder',
                    'notes': 'Folder contains no documents',
                })
            else:
                units.append({
                    'identifier': entry,
                    'type': 'folder',
                    'path': full_path + os.sep,
                })

        elif os.path.isfile(full_path):
            name_no_ext = os.path.splitext(entry)[0]
            unit = {
                'identifier': name_no_ext,
                'type': 'file',
                'path': full_path,
            }
            if _is_generic(name_no_ext):
                unit['flag'] = 'non_identifiable'
                unit['notes'] = 'Generic or non-identifiable filename'
            units.append(unit)

    _flag_duplicates(units)
    return units
