import os
import re
import json
import time
from google import genai
from google.genai import types

_SYSTEM = (
    "You are a document grouping assistant for a background verification company called AuthBridge.\n\n"
    "You receive:\n"
    "1. Email body context (optional but VERY important when present) — the text from the original "
    "email and any forwarded emails or attached emails. This often contains ARS numbers, candidate "
    "names, Employee IDs, UAN numbers, and explicit descriptions of what each attachment is for. "
    "Treat any ARS / ID / name mentioned in the body as a strong signal for nearby or listed files.\n"
    "2. A list of file paths extracted from the client's attachments (ZIP, Drive folder, loose files, "
    "or files extracted from a forwarded email).\n"
    "3. A list of candidates from an Excel sheet or email body. Each candidate has a folder_key "
    "(the unique identifier — ARS number, Check ID, Employee ID, or Candidate Name), "
    "plus emp_id, candidate_name, and optionally a checks list. "
    "The checks list contains the individual background checks for this candidate "
    "(e.g. Employment, Education, Address) — each check has a check_id and check_type.\n\n"
    "Your job is to group every file to the correct candidate using any available signal:\n"
    "- ARS number, Employee ID, UAN, or candidate name mentioned in the email body context\n"
    "- Folder name or file name containing the folder_key value (ARS number, check_id, etc.)\n"
    "- Folder name or file name containing the emp_id\n"
    "- Partial or full candidate name match in the path\n"
    "- Check type or check_id found in a folder name (e.g. 'Employment/', 'CHK-001/') — "
    "these are strong signals that this file belongs to that candidate\n"
    "- Proximity — files inside a named candidate folder belong to that candidate\n"
    "- When the body says 'please find attached documents for ARS 6197-XXXX' and there is only "
    "one candidate, ALL files likely belong to that candidate — assign at high confidence\n\n"
    "Rules:\n"
    "- Every file must be accounted for. No file left ungrouped without a reason.\n"
    "- A file can only belong to one candidate.\n"
    "- If a file has no identifiable signal (e.g. screenshot.pdf, scan.pdf) AND the body gives "
    "no clue, mark as UNASSIGNED. But if the body references a single ARS or candidate and files "
    "are generic-named, they likely belong to that candidate — assign with moderate confidence.\n"
    "- If a file could equally belong to two or more candidates, mark as AMBIGUOUS.\n"
    "- Assign a confidence score 0-100 to each file assignment.\n"
    "- Files inside a clearly named candidate folder inherit high confidence from the folder name.\n\n"
    "Return only valid JSON. No explanation outside the JSON."
)


_JUNK_DIRS  = {'__MACOSX', '__pycache__'}
_JUNK_FILES = {'Thumbs.db', 'desktop.ini', 'ehthumbs.db', 'thumbs.db'}


def build_file_tree(extract_root):
    """
    Return list of relative file paths (forward slashes) under extract_root.
    email_body.txt files are excluded — they are context for the grouping agent,
    not candidate documents. Use collect_body_context() to read them separately.
    """
    file_paths = []
    for dirpath, dirs, filenames in os.walk(extract_root):
        dirs[:] = [
            d for d in dirs
            if not d.startswith('.') and d not in _JUNK_DIRS
        ]
        for fname in filenames:
            if fname.startswith('.') or fname in _JUNK_FILES:
                continue
            if fname == 'email_body.txt':
                continue  # context file — collected separately by collect_body_context()
            full = os.path.join(dirpath, fname)
            rel  = os.path.relpath(full, extract_root).replace('\\', '/')
            file_paths.append(rel)
    return file_paths


def collect_body_context(extract_root):
    """
    Scan extract_root for email_body.txt files written by the extractor when
    processing .eml / .msg containers. Returns their combined text so the
    grouping agent can use inline candidate references as grouping signals.
    Capped at 6000 characters per file, 20000 total — enough for any real email.
    """
    _PER_FILE_CAP = 6000
    _TOTAL_CAP    = 20000
    parts         = []
    total         = 0

    for dirpath, dirs, filenames in os.walk(extract_root):
        dirs[:] = [d for d in dirs if not d.startswith('.') and d not in _JUNK_DIRS]
        for fname in filenames:
            if fname != 'email_body.txt':
                continue
            full_path = os.path.join(dirpath, fname)
            rel_dir   = os.path.relpath(dirpath, extract_root).replace('\\', '/')
            label     = f'[Forwarded email body — from: {rel_dir}]' if rel_dir != '.' else '[Email body]'
            try:
                with open(full_path, 'r', encoding='utf-8', errors='replace') as f:
                    text = f.read().strip()
                if not text:
                    continue
                if len(text) > _PER_FILE_CAP:
                    text = text[:_PER_FILE_CAP] + '\n[… truncated]'
                parts.append(f'{label}\n{text}')
                total += len(text)
                if total >= _TOTAL_CAP:
                    break
            except Exception:
                pass
        if total >= _TOTAL_CAP:
            break

    return '\n\n'.join(parts)


def _strip_fences(text):
    text = text.strip()
    if text.startswith('```'):
        parts = text.split('```')
        text = parts[1] if len(parts) > 1 else text
        if text.startswith('json'):
            text = text[4:]
    return text.strip()


def _parse_response(text):
    return json.loads(_strip_fences(text))


def _validate_and_ensure_coverage(result, file_paths):
    """
    1. Remove LLM-hallucinated paths (not in actual file list).
    2. Deduplicate: if a file is in both groups and unassigned/ambiguous, keep the group assignment.
    3. Ensure every real path is accounted for.
    """
    valid = set(file_paths)

    # Strip hallucinated paths from groups
    for group in result.get('groups', []):
        real_files = []
        for f in group.get('files', []):
            p = f.get('path', '')
            if p in valid:
                real_files.append(f)
            else:
                print(f'[GroupingAgent] Dropping hallucinated path: {p}')
        group['files'] = real_files

    # Build set of paths already assigned to a group (group assignment wins)
    in_group = set()
    for group in result.get('groups', []):
        for f in group.get('files', []):
            in_group.add(f.get('path', ''))

    # Remove duplicates from unassigned/ambiguous that are already in a group
    result['unassigned'] = [
        u for u in result.get('unassigned', [])
        if u.get('path', '') not in in_group and u.get('path', '') in valid
    ]
    result['ambiguous'] = [
        a for a in result.get('ambiguous', [])
        if a.get('path', '') not in in_group and a.get('path', '') in valid
    ]

    # Build covered set
    covered = set(in_group)
    for u in result.get('unassigned', []):
        covered.add(u.get('path', ''))
    for a in result.get('ambiguous', []):
        covered.add(a.get('path', ''))

    # Any real file not mentioned at all → unassigned
    for path in file_paths:
        if path not in covered:
            result.setdefault('unassigned', []).append({
                'path': path,
                'reason': 'Not mentioned in grouping response — defaulted to unassigned',
            })


def _normalize_id_for_search(s):
    """
    Collapse whitespace/hyphen/underscore runs to a single '#' marker (instead of
    deleting them) before substring-searching an ID in a filename.

    Deleting separators outright breaks the '(?<!\\d)...(?!\\d)' boundary guard in
    _match_by_id: a real-world filename like "6197-003526-1.jpg" (multiple photos
    of the same document, sequence-numbered) strips to "61970035261.jpg" — the "1"
    from "-1" lands immediately after the ARS digits with nothing between them, so
    the trailing-digit guard falsely rejects a genuine match. Collapsing to a
    non-digit placeholder instead of deleting preserves that boundary while still
    treating "-", "_", and " " as equivalent separators on both sides of the match.
    """
    return re.sub(r'[\s\-_]+', '#', (s or '')).lower()


def _build_id_maps(candidates):
    """Build {normalised folder_key: candidate} and {normalised numeric emp_id: candidate} maps."""
    folder_key_map = {}
    empid_map = {}
    for c in candidates:
        fk  = _normalize_id_for_search(c.get('folder_key') or '')
        emp = re.sub(r'[\s\-_]', '', (c.get('emp_id') or '')).lower()
        if fk:
            folder_key_map[fk] = c
        if emp and emp.isdigit():
            empid_map[emp] = c
    return folder_key_map, empid_map


def _match_by_id(filename, norm_path, folder_key_map, empid_map):
    """Try folder_key (ARS/check_id/emp_id) then bare emp_id match. Returns (candidate, reason) or (None, None)."""
    for fk_key, cand in folder_key_map.items():
        if fk_key and re.search(r'(?<!\d)' + re.escape(fk_key) + r'(?!\d)', norm_path):
            reason = (
                f'{cand.get("folder_key_type", "ID")} {cand["folder_key"]} '
                f'found in filename (rule-based fallback)'
            )
            return cand, reason

    for num in re.findall(r'\b\d{4,10}\b', filename):
        norm_num = re.sub(r'[\s\-_]', '', num).lower()
        if norm_num in empid_map:
            return empid_map[norm_num], f'Employee ID {num} found in filename (rule-based fallback)'

    return None, None


def _rule_based_fallback(unassigned, groups, candidates, skipped_candidates=None):
    """
    For files the LLM missed, try exact ID matching from the filename/path:
    - ARS pattern (digits-digits, e.g. 6197-002204)
    - Employee ID (standalone number that matches emp_id exactly)
    Moves matched files out of unassigned and into the correct group.

    Files that don't match any of the active `candidates` are then checked against
    `skipped_candidates` (the ones cut by the MVP's per-batch cap — see app.py's
    MAX_CANDIDATES). A match there can't be placed in a group (that candidate has no
    output folder this run), so the entry stays in still_unassigned but is tagged with
    `capped_match` so the assembler can label it honestly instead of a generic
    "Unassigned" that reads like a matching failure.
    """
    folder_key_map, empid_map = _build_id_maps(candidates)
    capped_folder_key_map, capped_empid_map = _build_id_maps(skipped_candidates or [])

    groups_by_fk = {g['folder_key']: g for g in groups}

    still_unassigned = []
    for entry in unassigned:
        path = entry.get('path', '')
        filename = os.path.basename(path)
        norm_path = _normalize_id_for_search(path + filename)

        matched_candidate, match_reason = _match_by_id(filename, norm_path, folder_key_map, empid_map)

        if matched_candidate:
            fk = matched_candidate['folder_key']
            print(f'[RuleBasedFallback] Matched "{filename}" → {fk} ({matched_candidate.get("name", "")}): {match_reason}')
            if fk in groups_by_fk:
                groups_by_fk[fk]['files'].append({
                    'path': path, 'confidence': 96, 'reason': match_reason,
                })
            else:
                new_group = {
                    'folder_key':     fk,
                    'candidate_name': matched_candidate.get('name', ''),
                    'emp_id':         matched_candidate.get('emp_id', ''),
                    'files': [{'path': path, 'confidence': 96, 'reason': match_reason}],
                }
                groups.append(new_group)
                groups_by_fk[fk] = new_group
            continue

        capped_candidate, capped_reason = _match_by_id(filename, norm_path, capped_folder_key_map, capped_empid_map)
        if capped_candidate:
            print(f'[RuleBasedFallback] "{filename}" matches capacity-capped candidate '
                  f'{capped_candidate["folder_key"]} ({capped_candidate.get("name", "")}) — not this run\'s output')
            entry = dict(entry)
            entry['reason'] = (
                f'{capped_reason} — but that candidate was skipped by this batch\'s '
                f'{"MAX_CANDIDATES"} cap. Will be picked up when they are re-run.'
            )
            entry['capped_match'] = {
                'folder_key':     capped_candidate['folder_key'],
                'candidate_name': capped_candidate.get('name', ''),
                'emp_id':         capped_candidate.get('emp_id', ''),
            }
            still_unassigned.append(entry)
            continue

        still_unassigned.append(entry)

    return still_unassigned


# Gemini 2.5 Flash output cap is ~64K tokens.
# Each file entry in the JSON response costs ~50 tokens (path + confidence + reason).
# 500 files × 50 ≈ 25K tokens — safely under the limit.
_MAX_FILES_PER_CALL = 500

MODELS      = ['gemini-2.5-flash', 'gemini-2.5-flash-lite']
RETRY_WAITS = [0, 5, 15, 30]


def _is_transient(e):
    s = str(e)
    return any(x in s for x in ('503', '429', 'UNAVAILABLE', 'RESOURCE_EXHAUSTED'))


def _chunk_files(file_paths, chunk_size):
    """
    Split file_paths into chunks of up to chunk_size, keeping files that share
    the same top-level folder together so the LLM retains proximity signals.
    """
    from collections import defaultdict
    folder_groups = defaultdict(list)
    for path in file_paths:
        top = path.split('/')[0]
        folder_groups[top].append(path)

    chunks, current = [], []
    for files in folder_groups.values():
        if len(files) > chunk_size:
            for i in range(0, len(files), chunk_size):
                chunks.append(files[i:i + chunk_size])
        elif len(current) + len(files) > chunk_size:
            if current:
                chunks.append(current)
            current = list(files)
        else:
            current.extend(files)
    if current:
        chunks.append(current)
    return chunks


def _merge_chunk_results(results):
    """Merge {groups, unassigned, ambiguous} dicts from multiple LLM calls."""
    groups_by_fk = {}
    unassigned, ambiguous = [], []
    for r in results:
        for g in r.get('groups', []):
            fk = g.get('folder_key', '')
            if fk in groups_by_fk:
                groups_by_fk[fk]['files'].extend(g.get('files', []))
            else:
                groups_by_fk[fk] = {**g, 'files': list(g.get('files', []))}
        unassigned.extend(r.get('unassigned', []))
        ambiguous.extend(r.get('ambiguous', []))
    return {
        'groups':     list(groups_by_fk.values()),
        'unassigned': unassigned,
        'ambiguous':  ambiguous,
    }


def _call_chunk(chunk, candidate_list, client, body_context=None):
    """
    Generator. Yields status strings before each retry sleep so the caller can
    forward them to the UI. Finally yields {'result': <dict>} on success or
    {'failed': True} if all models and retries are exhausted.
    """
    _JSON_SCHEMA = (
        "{\n"
        '  "groups": [\n'
        '    {\n'
        '      "folder_key": "...",\n'
        '      "candidate_name": "...",\n'
        '      "emp_id": "...",\n'
        '      "files": [\n'
        '        {"path": "...", "confidence": 99, "reason": "..."}\n'
        '      ]\n'
        '    }\n'
        '  ],\n'
        '  "unassigned": [\n'
        '    {"path": "...", "reason": "..."}\n'
        '  ],\n'
        '  "ambiguous": [\n'
        '    {"path": "...", "reason": "..."}\n'
        '  ]\n'
        '}'
    )

    context_section = ''
    if body_context and body_context.strip():
        context_section = (
            "Email body context — read this carefully before grouping. "
            "It often contains the ARS number, candidate name, or explicit description "
            "of what each file is for:\n"
            f"{body_context.strip()}\n\n"
            "---\n\n"
        )

    base_msg = (
        f"{context_section}"
        f"File paths:\n{json.dumps(chunk, indent=2)}\n\n"
        f"Candidates:\n{json.dumps(candidate_list, indent=2)}\n\n"
        f"Return JSON in exactly this structure:\n{_JSON_SCHEMA}"
    )
    retry_msg = base_msg + '\n\nCRITICAL: Return only valid JSON. No text outside the JSON object.'

    for model in MODELS:
        for attempt, wait in enumerate(RETRY_WAITS):
            if wait:
                yield f'AI grouping: API busy — retrying with {model} in {wait}s (attempt {attempt + 1}/{len(RETRY_WAITS)})…'
                time.sleep(wait)
            msg = base_msg if attempt == 0 else retry_msg
            try:
                response = client.models.generate_content(
                    model=model,
                    contents=msg,
                    config=types.GenerateContentConfig(system_instruction=_SYSTEM),
                )
                result = _parse_response(response.text)
                print(f'[GroupingAgent] Success with {model} on attempt {attempt + 1}')
                yield {'result': result}
                return
            except Exception as e:
                if _is_transient(e):
                    print(f'[GroupingAgent] {model} attempt {attempt + 1} — transient error: {e}')
                    if attempt == len(RETRY_WAITS) - 1:
                        print(f'[GroupingAgent] All retries for {model} exhausted, trying next model…')
                        yield f'AI grouping: switching to fallback model ({MODELS[-1]})…'
                else:
                    print(f'[GroupingAgent] {model} attempt {attempt + 1} — error: {e}')
                    break  # non-transient — skip remaining retries for this model

    yield {'failed': True}


def group_files(file_paths, candidates, extract_root, body_context=None, skipped_candidates=None):
    """
    Generator. Yields status strings during retries so the caller can stream
    them to the UI. Finally yields {'result': (groups, unassigned, ambiguous)}.

    Args:
        body_context: Combined email body text (main email + any inner .eml/.msg bodies)
                      passed verbatim into the AI prompt so it can use inline candidate
                      references (ARS numbers, names, file descriptions) as grouping signals.
        skipped_candidates: Candidates cut by the MVP's MAX_CANDIDATES cap (see app.py).
                      Not sent to the AI (they're out of scope for this run's output) but
                      still checked by the rule-based fallback so a leftover file that
                      matches one of them gets labeled honestly instead of a generic
                      "Unassigned" — see _rule_based_fallback's capped_match tagging.

    Usage:
        for event in group_files(...):
            if isinstance(event, str):
                emit_to_ui(event)
            else:
                groups, unassigned, ambiguous = event['result']
    """
    api_key = os.environ.get('GEMINI_API_KEY', '').strip()

    if not api_key:
        yield {'result': (
            [],
            [{'path': p, 'reason': 'No GEMINI_API_KEY configured — all files unassigned'} for p in file_paths],
            [],
        )}
        return

    client = genai.Client(api_key=api_key)

    candidate_list = [
        {
            'folder_key':      c.get('folder_key', ''),
            'folder_key_type': c.get('folder_key_type', 'unknown'),
            'emp_id':          c.get('emp_id', ''),
            'candidate_name':  c.get('name', ''),
            'checks':          c.get('checks', []),
        }
        for c in candidates
        if not c.get('_flag')
    ]

    chunks = (
        _chunk_files(file_paths, _MAX_FILES_PER_CALL)
        if len(file_paths) > _MAX_FILES_PER_CALL
        else [file_paths]
    )

    if len(chunks) > 1:
        yield f'Large batch — splitting {len(file_paths)} files into {len(chunks)} chunks of ≤{_MAX_FILES_PER_CALL}…'

    chunk_results = []
    failed_paths  = []

    for i, chunk in enumerate(chunks):
        if len(chunks) > 1:
            yield f'AI grouping chunk {i + 1}/{len(chunks)} ({len(chunk)} files)…'

        chunk_result = None
        for event in _call_chunk(chunk, candidate_list, client, body_context=body_context):
            if isinstance(event, str):
                yield event          # forward retry status to UI
            elif 'result' in event:
                chunk_result = event['result']
            # {'failed': True} — chunk_result stays None

        if chunk_result is not None:
            chunk_results.append(chunk_result)
        else:
            failed_paths.extend(chunk)

    if not chunk_results:
        all_unassigned = [
            {'path': p, 'reason': 'LLM grouping failed after all retries — manual review required'}
            for p in file_paths
        ]
        groups = []
        remaining = _rule_based_fallback(all_unassigned, groups, candidates, skipped_candidates)
        yield {'result': (groups, remaining, [])}
        return

    result = _merge_chunk_results(chunk_results) if len(chunk_results) > 1 else chunk_results[0]

    for p in failed_paths:
        result['unassigned'].append({
            'path': p,
            'reason': 'LLM chunk failed after all retries — manual review required',
        })

    _validate_and_ensure_coverage(result, file_paths)

    groups    = result.get('groups', [])
    ambiguous = result.get('ambiguous', [])

    remaining_unassigned = _rule_based_fallback(
        result.get('unassigned', []), groups, candidates, skipped_candidates
    )

    yield {'result': (groups, remaining_unassigned, ambiguous)}
