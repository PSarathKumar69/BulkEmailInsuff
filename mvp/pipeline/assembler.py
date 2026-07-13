import os
import re
import json
import shutil
import csv
from datetime import datetime


def _safe(s):
    return re.sub(r'[<>:"/\\|?*\s]', '_', str(s).strip())


def _folder_name(key, name):
    safe_name = _safe(name)
    return f"{_safe(key)}_{safe_name}" if safe_name else _safe(key)


def _norm_for_match(s):
    return re.sub(r'[\s\-_/]', '', (s or '').lower())


def _match_check_from_path(rel_path, checks):
    """
    Detect which check a file belongs to from intermediate folder names in its path.
    Only the folder components are inspected — the filename itself is ignored.

    Matching priority:
      1. check_id substring found in a folder name
      2. A significant keyword from check_type found in a folder name

    Returns the matching check dict, or None if no signal found (caller routes flat).
    """
    parts = rel_path.replace('\\', '/').split('/')
    folder_parts = parts[:-1]   # skip the filename
    if not folder_parts:
        return None

    for part in folder_parts:
        pn = _norm_for_match(part)
        if not pn:
            continue
        # 1. check_id match
        for check in checks:
            cid = _norm_for_match(check.get('check_id', ''))
            if cid and cid in pn:
                return check
        # 2. check_type keyword match — skip generic words that appear in every check name
        _SKIP = {'check', 'verification', 'verif', 'the', 'and', 'bgv', 'background'}
        for check in checks:
            words = [
                w for w in re.findall(r'\b\w{3,}\b', (check.get('check_type') or '').lower())
                if w not in _SKIP
            ]
            for word in words:
                if word in pn:
                    return check
    return None


def _check_subfolder_name(check):
    """
    Build the check sub-folder name: CheckType_CheckID.
    Falls back to whichever part is available.
    """
    ctype = _safe(check.get('check_type') or '')
    cid   = _safe(check.get('check_id')   or '')
    if ctype and cid:
        return f'{ctype}_{cid}'
    return cid or ctype or 'check'


def _unique_filename(folder, filename, collision_prefix=None):
    """Return (final_filename, was_collision).

    On collision, tries <collision_prefix>_<filename> first (when prefix given),
    then falls back to <base>_2, _3, … so something always works.
    """
    if not os.path.exists(os.path.join(folder, filename)):
        return filename, False
    # Try ARS-prefixed name first
    if collision_prefix:
        prefixed = f"{_safe(str(collision_prefix))}_{filename}"
        if not os.path.exists(os.path.join(folder, prefixed)):
            return prefixed, True
    # Numeric fallback
    base, ext = os.path.splitext(filename)
    n = 2
    while True:
        candidate = f'{base}_{n}{ext}'
        if not os.path.exists(os.path.join(folder, candidate)):
            return candidate, True
        n += 1


def _copy_flat(src, dest_folder, collision_prefix=None):
    """Copy src into dest_folder (flat). Returns (dest_filename, was_collision)."""
    os.makedirs(dest_folder, exist_ok=True)
    fname, collision = _unique_filename(dest_folder, os.path.basename(src), collision_prefix)
    shutil.copy2(src, os.path.join(dest_folder, fname))
    return fname, collision


def assemble_output(
    groups,
    unassigned,
    ambiguous,
    extract_root,
    output_root,
    all_candidates=None,
    pending_links=None,
    direct_files=None,
    skipped_candidates=None,
    files_in_system=False,
    run_id=None,
):
    """
    Build the output folder tree and manifest from grouping agent results.

    Args:
        groups          : list of group dicts from grouping_agent
        unassigned      : list of unassigned file entries
        ambiguous       : list of ambiguous file entries
        extract_root    : per-run extracted directory
        output_root     : base output directory (output/)
        all_candidates  : full candidate list — used to build the check sub-folder lookup
                          for multi-check ARS rows and to write extracted_info.txt for
                          every candidate (see below)
        pending_links   : list of {url, type, status} from link_handler — manifest-only rows.
                          May also carry folder_key/candidate_name/emp_id when the link came
                          from a specific candidate's Excel row (see direct_files).
        direct_files    : {folder_key: [(absolute_path, relative_path, source_url), ...]} — files
                          downloaded from links found anywhere in a candidate's own Excel rows
                          (column-name-agnostic detection — see excel_reader.py). relative_path
                          is relative to that link's own isolated extract dir (used for check-
                          subfolder detection, e.g. a ZIP that unpacked into Employment/doc1.pdf).
                          source_url is the specific link that produced that file — a candidate
                          can have more than one real link (one per check type). Assigned directly
                          to that candidate at auto-matched confidence, bypassing grouping_agent
                          entirely since the Excel row already deterministically identifies them.
        skipped_candidates : candidates beyond this MVP's per-batch cap (see app.py's
                          MAX_CANDIDATES) — NOT in all_candidates, so they get no folder,
                          no download, no extracted_info.txt. Still written to the manifest
                          (status skipped_capacity_limit) so nothing is silently dropped.
        files_in_system : True when document files were present in this run (ZIP/loose files).
                          Only changes the manifest reason wording for candidates with zero
                          documents — every candidate gets extracted_info.txt regardless (see
                          the per-candidate info.txt loop below).
        run_id          : UUID hex for the current run — appended to batch folder name to
                          prevent collisions when two requests land in the same second.

    Confidence routing:
        >= 95  → ARS_Name/         (Auto Matched)
        80–94  → _REVIEW/ARS_Name/ (Needs Review)
        < 80   → _UNASSIGNED/      (low confidence)
        Unassigned / Ambiguous → _UNASSIGNED/
        Cross-candidate duplicate (same path in 2+ groups) → _UNASSIGNED/
        direct_files → ARS_Name/ always (Auto Matched, confidence 100)

    Every candidate in all_candidates also gets extracted_info.txt with their merged
    Excel/email metadata — written into wherever their real documents (if any) already
    landed, or into a fresh ARS_Name/ folder if they have none.

    Returns (manifest_rows, stats_dict, batch_folder_path).
    """
    all_candidates     = all_candidates     or []
    pending_links      = pending_links      or []
    direct_files       = direct_files       or {}
    skipped_candidates = skipped_candidates or []

    ts         = datetime.now().strftime('%Y%m%d_%H%M%S')
    uid_suffix = f'_{run_id[:8]}' if run_id else ''
    batch      = os.path.join(output_root, f'output_batch_{ts}{uid_suffix}')
    os.makedirs(batch, exist_ok=True)

    # Build a lookup: folder_key → checks list (only for multi-check candidates).
    # Used to decide whether a file should go into a check sub-folder or flat.
    cand_check_lookup = {
        c['folder_key']: c['checks']
        for c in all_candidates
        if c.get('multi_check') and c.get('checks') and c.get('folder_key')
    }

    # folder_key → cross-Excel-file merge note (from excel_reader.parse_excel_multi),
    # surfaced in the manifest reason so CST can see when a candidate's identity was
    # backfilled or reconciled across two attachments.
    cand_merge_notes = {
        c['folder_key']: c['_merge_note']
        for c in all_candidates
        if c.get('_merge_note') and c.get('folder_key')
    }

    # folder_key → candidate dict — used to resolve name/emp_id for direct_files entries.
    fk_to_cand = {c['folder_key']: c for c in all_candidates if c.get('folder_key')}

    # folder_key → top-level ARS_Name (or _REVIEW/ARS_Name) folder actually used for that
    # candidate's real documents — so extracted_info.txt lands alongside them instead of
    # in a separate, redundant folder.
    folder_dest_for_fk = {}

    # Detect cross-candidate duplicates
    path_fk = {}
    for group in groups:
        fk = group.get('folder_key', 'UNKNOWN')
        for fe in group.get('files', []):
            p = fe.get('path', '')
            path_fk.setdefault(p, []).append(fk)
    cross_dups = {p for p, fks in path_fk.items() if len(fks) > 1}

    manifest = []
    stats = {'auto_matched': 0, 'review': 0, 'unassigned': 0, 'info_only': 0, 'skipped_capacity_limit': 0}

    # Track which folder_keys have at least one file (to detect info-only candidates)
    folder_keys_with_files = set()

    # ── Assigned files ───────────────────────────────────────────────────────
    for group in groups:
        fk    = group.get('folder_key', 'UNKNOWN')
        cname = group.get('candidate_name', 'Unknown')
        emp   = group.get('emp_id', '')

        for fe in group.get('files', []):
            rel_path   = fe.get('path', '')
            confidence = fe.get('confidence', 0)
            reason     = fe.get('reason', '')
            merge_note = cand_merge_notes.get(fk, '')
            if merge_note:
                reason = f'{reason} | {merge_note}' if reason else merge_note
            src        = os.path.join(extract_root, rel_path.replace('/', os.sep))

            if not os.path.exists(src):
                manifest.append({
                    'original_path':   rel_path,
                    'original_folder': os.path.dirname(rel_path),
                    'filename':        os.path.basename(rel_path),
                    'source_type':     'attachment',
                    'link_url':        '',
                    'folder_key':      fk,
                    'candidate_name':  cname,
                    'emp_id':          emp,
                    'confidence':      confidence,
                    'reason':          reason,
                    'status':          'ERROR — source file not found',
                    'output_path':     '',
                })
                continue

            folder_keys_with_files.add(fk)

            # Detect check sub-folder — only for multi-check ARS candidates
            # where the client's files are also organised by check in sub-folders.
            check_sub = None
            candidate_checks = cand_check_lookup.get(fk)
            if candidate_checks:
                matched_check = _match_check_from_path(rel_path, candidate_checks)
                if matched_check:
                    check_sub = _check_subfolder_name(matched_check)

            if rel_path in cross_dups:
                dest_folder = os.path.join(batch, '_UNASSIGNED')
                fname, _    = _copy_flat(src, dest_folder)
                status      = 'Unassigned — cross-candidate duplicate'
                out         = os.path.join(f'output_batch_{ts}', '_UNASSIGNED', fname)
                stats['unassigned'] += 1

            elif confidence >= 95:
                cand_folder = os.path.join(batch, _folder_name(fk, cname))
                folder_dest_for_fk.setdefault(fk, cand_folder)
                dest_folder = os.path.join(cand_folder, check_sub) if check_sub else cand_folder
                fname, collision = _copy_flat(src, dest_folder, collision_prefix=fk)
                status = 'Auto Matched' + (' — filename collision' if collision else '')
                rel_parts = [_folder_name(fk, cname)] + ([check_sub] if check_sub else []) + [fname]
                out = os.path.join(f'output_batch_{ts}', *rel_parts)
                stats['auto_matched'] += 1

            elif confidence >= 80:
                cand_folder = os.path.join(batch, '_REVIEW', _folder_name(fk, cname))
                folder_dest_for_fk.setdefault(fk, cand_folder)
                dest_folder = os.path.join(cand_folder, check_sub) if check_sub else cand_folder
                fname, collision = _copy_flat(src, dest_folder, collision_prefix=fk)
                status = 'Needs Review' + (' — filename collision' if collision else '')
                rel_parts = ['_REVIEW', _folder_name(fk, cname)] + ([check_sub] if check_sub else []) + [fname]
                out = os.path.join(f'output_batch_{ts}', *rel_parts)
                stats['review'] += 1

            else:
                dest_folder = os.path.join(batch, '_UNASSIGNED')
                fname, _    = _copy_flat(src, dest_folder)
                status      = 'Unassigned — low confidence'
                out         = os.path.join(f'output_batch_{ts}', '_UNASSIGNED', fname)
                stats['unassigned'] += 1

            manifest.append({
                'original_path':   rel_path,
                'original_folder': os.path.dirname(rel_path),
                'filename':        os.path.basename(src),
                'source_type':     'attachment',
                'link_url':        '',
                'folder_key':      fk,
                'candidate_name':  cname,
                'emp_id':          emp,
                'confidence':      confidence,
                'reason':          reason,
                'status':          status,
                'output_path':     out,
            })

    # ── Direct candidate-link files ──────────────────────────────────────────
    # Files downloaded from links found anywhere in a candidate's own Excel rows
    # (column-name-agnostic — see excel_reader.py's per-cell URL detection). The
    # Excel row already deterministically identifies the candidate, so these
    # bypass grouping_agent entirely and go straight into that candidate's folder
    # at auto-matched confidence — no AI matching needed or wanted here. A
    # candidate can have more than one real link (one per check type); each file
    # keeps track of exactly which URL produced it.
    for fk, file_entries in direct_files.items():
        cand = fk_to_cand.get(fk)
        if not cand:
            continue
        cname = cand.get('name', '')
        emp   = cand.get('emp_id', '')

        cand_folder = os.path.join(batch, _folder_name(fk, cname))
        folder_dest_for_fk.setdefault(fk, cand_folder)

        candidate_checks = cand_check_lookup.get(fk)

        for src, rel_path, source_url in file_entries:
            if not os.path.exists(src):
                continue
            check_sub = None
            if candidate_checks:
                matched_check = _match_check_from_path(rel_path, candidate_checks)
                if matched_check:
                    check_sub = _check_subfolder_name(matched_check)

            dest_folder = os.path.join(cand_folder, check_sub) if check_sub else cand_folder
            fname, collision = _copy_flat(src, dest_folder, collision_prefix=fk)
            folder_keys_with_files.add(fk)
            status = 'Auto Matched' + (' — filename collision' if collision else '')
            rel_parts = [_folder_name(fk, cname)] + ([check_sub] if check_sub else []) + [fname]
            out = os.path.join(f'output_batch_{ts}', *rel_parts)
            stats['auto_matched'] += 1

            manifest.append({
                'original_path':   rel_path,
                'original_folder': os.path.dirname(rel_path),
                'filename':        fname,
                'source_type':     'link',
                'link_url':        source_url,
                'folder_key':      fk,
                'candidate_name':  cname,
                'emp_id':          emp,
                'confidence':      100,
                'reason':          'Downloaded from a document link found in this candidate\'s Excel row',
                'status':          status,
                'output_path':     out,
            })

    # ── Unassigned ───────────────────────────────────────────────────────────
    for entry in unassigned:
        rel_path    = entry.get('path', '')
        reason      = entry.get('reason', 'No candidate signal')
        src         = os.path.join(extract_root, rel_path.replace('/', os.sep))
        dest_folder = os.path.join(batch, '_UNASSIGNED')
        out = ''
        if os.path.exists(src):
            fname, _ = _copy_flat(src, dest_folder)
            out = os.path.join(f'output_batch_{ts}', '_UNASSIGNED', fname)

        # Still lands in _UNASSIGNED (that candidate has no folder this run), but the
        # rule-based fallback identified exactly whose file this is — see
        # grouping_agent._rule_based_fallback's capped_match tagging. Labeled distinctly
        # so this doesn't read as an unresolved matching failure.
        capped = entry.get('capped_match')
        status = 'unassigned_capacity_capped' if capped else 'Unassigned'

        manifest.append({
            'original_path':   rel_path,
            'original_folder': os.path.dirname(rel_path),
            'filename':        os.path.basename(rel_path),
            'source_type':     'attachment',
            'link_url':        '',
            'folder_key':      capped.get('folder_key', '') if capped else '',
            'candidate_name':  capped.get('candidate_name', '') if capped else '',
            'emp_id':          capped.get('emp_id', '') if capped else '',
            'confidence':      0,
            'reason':          reason,
            'status':          status,
            'output_path':     out,
        })
        stats['unassigned'] += 1

    # ── Ambiguous ────────────────────────────────────────────────────────────
    for entry in ambiguous:
        rel_path    = entry.get('path', '')
        reason      = entry.get('reason', 'Ambiguous — multiple candidates equally likely')
        src         = os.path.join(extract_root, rel_path.replace('/', os.sep))
        dest_folder = os.path.join(batch, '_UNASSIGNED')
        out = ''
        if os.path.exists(src):
            fname, _ = _copy_flat(src, dest_folder)
            out = os.path.join(f'output_batch_{ts}', '_UNASSIGNED', fname)
        manifest.append({
            'original_path':   rel_path,
            'original_folder': os.path.dirname(rel_path),
            'filename':        os.path.basename(rel_path),
            'source_type':     'attachment',
            'link_url':        '',
            'folder_key':      '',
            'candidate_name':  '',
            'emp_id':          '',
            'confidence':      0,
            'reason':          reason,
            'status':          'Ambiguous',
            'output_path':     out,
        })
        stats['unassigned'] += 1

    # ── extracted_info.txt for every candidate ───────────────────────────────
    #
    # Every candidate in all_candidates gets extracted_info.txt with their merged
    # Excel/email metadata, EVERY time — regardless of whether they also have real
    # documents (from AI grouping or a direct Excel document-link download).
    # Otherwise the Excel/email details behind a candidate who does have documents
    # would simply be lost. The txt is written into wherever their real documents
    # already live (folder_dest_for_fk), or a fresh ARS_Name/ folder if they have none.
    for cand in all_candidates:
        fk = cand.get('folder_key', '')
        if not fk or cand.get('_flag'):
            continue

        cname       = cand.get('name', '')
        emp         = cand.get('emp_id', '')
        source_type = 'body_text' if cand.get('_source') == 'body' else 'attachment'
        merge_note  = cand_merge_notes.get(fk, '')
        has_docs    = fk in folder_keys_with_files

        dest_folder = folder_dest_for_fk.get(fk) or os.path.join(batch, _folder_name(fk, cname))
        os.makedirs(dest_folder, exist_ok=True)

        # Write exactly what the client sent — original column names and row values.
        # Internal pipeline fields (folder_key, ars, check_id, etc.) are NOT written here;
        # they are used only for folder naming and are already captured in the manifest.
        info = dict(cand.get('raw_fields') or {})
        with open(os.path.join(dest_folder, 'extracted_info.txt'), 'w', encoding='utf-8') as f:
            json.dump(info, f, ensure_ascii=False, indent=2)

        out = os.path.join(
            f'output_batch_{ts}',
            os.path.relpath(dest_folder, batch),
            'extracted_info.txt',
        )

        if has_docs:
            reason = 'Excel/email metadata record for this candidate — real document(s) also present in this batch'
            status = 'info_txt_with_documents'
        elif files_in_system:
            reason = (
                'Candidate listed in Excel/email but no documents were matched by AI — '
                'info-only record created from Excel/email data'
            )
            status = 'info_only_no_documents'
            stats['info_only'] += 1
        else:
            reason = 'Candidate fully resolved from email — no document files in this batch'
            status = 'info_only_no_documents'
            stats['info_only'] += 1
        if merge_note:
            reason = f'{reason} | {merge_note}'

        manifest.append({
            'original_path':   '',
            'original_folder': '',
            'filename':        'extracted_info.txt',
            'source_type':     source_type,
            'link_url':        '; '.join(cand.get('document_links') or []),
            'folder_key':      fk,
            'candidate_name':  cname,
            'emp_id':          emp,
            'confidence':      100,
            'reason':          reason,
            'status':          status,
            'output_path':     out,
        })

    # ── Pending links (no folder — manifest row only) ────────────────────────
    for link in pending_links:
        default_reason = f'Link type: {link.get("type", "unknown")} — manual fetch required'
        manifest.append({
            'original_path':   '',
            'original_folder': '',
            'filename':        '',
            'source_type':     'link',
            'link_url':        link.get('url', ''),
            'folder_key':      link.get('folder_key', ''),
            'candidate_name':  link.get('candidate_name', ''),
            'emp_id':          link.get('emp_id', ''),
            'confidence':      0,
            'reason':          link.get('reason', default_reason),
            'status':          'link_pending_manual_fetch',
            'output_path':     '',
        })

    # ── Candidates skipped by the MVP per-batch cap ──────────────────────────
    # No folder, no download, no extracted_info.txt — but still on record in the
    # manifest so CST knows to process them in a follow-up run rather than losing
    # them silently.
    for cand in skipped_candidates:
        manifest.append({
            'original_path':   '',
            'original_folder': '',
            'filename':        '',
            'source_type':     'attachment',
            'link_url':        '; '.join(cand.get('document_links') or []),
            'folder_key':      cand.get('folder_key', ''),
            'candidate_name':  cand.get('name', ''),
            'emp_id':          cand.get('emp_id', ''),
            'confidence':      0,
            'reason':          (
                'Skipped — this MVP processes at most a fixed number of candidates per batch. '
                'Re-run this candidate in a separate, smaller batch.'
            ),
            'status':          'skipped_capacity_limit',
            'output_path':     '',
        })
        stats['skipped_capacity_limit'] += 1

    # ── manifest.csv ─────────────────────────────────────────────────────────
    fieldnames = [
        'original_path', 'original_folder', 'filename',
        'source_type', 'link_url',
        'folder_key', 'candidate_name', 'emp_id',
        'confidence', 'reason', 'status', 'output_path',
    ]
    manifest_path = os.path.join(batch, 'manifest.csv')
    with open(manifest_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(manifest)

    return manifest, stats, batch
