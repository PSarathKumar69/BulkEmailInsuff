import os
import sys
import uuid
import json
import shutil
import concurrent.futures
from flask import Flask, request, jsonify, render_template, Response, stream_with_context
from dotenv import load_dotenv

# Pipeline modules log with characters like → and — in their print() statements.
# On a Windows console using a legacy codepage (cp1252, not UTF-8), printing those
# raises UnicodeEncodeError — and because several of those prints sit right after a
# successful operation inside the same try/except (e.g. link_handler.py logging a
# successful download), that crash gets caught by the surrounding except and
# mis-recorded as a failure for something that actually succeeded. Forcing UTF-8
# here, once, at process start fixes it for every print() everywhere in the app.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

load_dotenv(os.path.join(os.path.dirname(__file__), '..', '.env'))

from pipeline.email_parser   import parse_email
from pipeline.extractor      import extract_all
from pipeline.excel_reader   import parse_excel, parse_excel_multi
from pipeline.grouping_agent import build_file_tree, group_files, collect_body_context
from pipeline.assembler      import assemble_output
from pipeline.discovery      import discover, scan_new_content

# Candidate document-link downloads (Step 3b) run in a bounded thread pool — these
# are I/O-bound network calls (Drive/SharePoint/HTTP), and a batch can legitimately
# have 1000+ candidates each with their own link, so downloading them one at a time
# can take long enough for the browser's SSE connection to be dropped as idle.
_LINK_DOWNLOAD_WORKERS = 8

def _merge_note_from_rows(rows):
    """
    Collect any cross-Excel-file merge conflicts/warnings attached to a candidate's
    row(s) by excel_reader.parse_excel_multi() into one human-readable note for the
    manifest. Empty string if the candidate wasn't touched by cross-file merging.
    """
    notes = []
    flagged_name_only = False
    for r in rows:
        for c in r.get('_merge_conflict', []) or []:
            if c not in notes:
                notes.append(c)
        if r.get('_merge_basis') == 'name_only' and not flagged_name_only:
            flagged_name_only = True
            notes.append('Matched across Excel files by name only (no shared ID) — verify carefully')
    return ' | '.join(notes)


def _collect_links(rows):
    """
    Union every row's document_links across all rows contributing to a candidate
    (e.g. multiple check-rows in one Excel), de-duplicated, order preserved.
    A candidate can genuinely have more than one real document link (one per
    check type), so — unlike a single "first non-blank wins" value — every link
    found anywhere in the candidate's own rows is kept and downloaded.
    """
    seen, links = set(), []
    for r in rows:
        for url in r.get('document_links') or []:
            if url not in seen:
                seen.add(url)
                links.append(url)
    return links


def _merge_raw_fields(rows):
    """
    Union every raw_fields dict across all rows contributing to one candidate —
    e.g. multiple check-rows in one Excel, or rows matched across multiple Excel
    attachments. Used to build extracted_info.txt for info-only candidates so it
    contains every unique column from every source, not just the primary row's.

    If a column header has the same value everywhere it appears, it's stored once
    under its own name. If the same header genuinely disagrees across rows, every
    distinct value is kept, disambiguated by source (filename when the row came
    from parse_excel_multi, otherwise its check_type/check_id) — nothing is
    silently dropped in favour of a single winner.
    """
    by_header = {}
    for r in rows:
        source = r.get('_source_file') or r.get('check_type') or r.get('check_id') or ''
        for header, value in (r.get('raw_fields') or {}).items():
            by_header.setdefault(header, []).append((value, source))

    merged = {}
    for header, entries in by_header.items():
        distinct_values = {v for v, _ in entries}
        if len(distinct_values) == 1:
            merged[header] = entries[0][0]
            continue
        for value, source in entries:
            key = f'{header} ({source})' if source else header
            base_key, n = key, 2
            while key in merged and merged[key] != value:
                key = f'{base_key}_{n}'
                n += 1
            merged[key] = value
    return merged


def _merge_dicts_prefer_existing(existing_dict, new_dict, new_source_label):
    """
    Union new_dict into existing_dict. Existing values win on conflict — used
    when a candidate's raw_fields were already built from their primary Excel
    row(s) and a later-discovered source (e.g. an Excel found deep inside
    their own document-link chain) adds more fields. A genuinely new header
    is added as-is; a header that already exists with a different value is
    kept under both, the new one disambiguated by source, so nothing is
    silently overwritten.
    """
    merged = dict(existing_dict)
    for key, value in new_dict.items():
        if key not in merged:
            merged[key] = value
        elif merged[key] != value:
            new_key, n = f'{key} ({new_source_label})', 2
            while new_key in merged and merged[new_key] != value:
                new_key = f'{key} ({new_source_label})_{n}'
                n += 1
            merged[new_key] = value
    return merged


def _build_candidates_from_excel_rows(excel_rows, candidates):
    """
    Group excel_rows by ARS (multi-check candidates) or fall back to
    check_id/emp_id/name, then merge additively into `candidates` (mutated
    in place). A row whose folder_key already has a candidate in the list
    extends that candidate's checks/raw_fields/document_links instead of
    creating a duplicate folder — this is what lets an Excel discovered
    anywhere (top-level, nested in a ZIP, nested behind a link at any hop)
    be folded into the same candidate list no matter when it's found,
    without ever splitting one ARS across two folders.
    """
    by_fk = {c['folder_key']: c for c in candidates if c.get('folder_key')}

    ars_groups, no_ars_rows = {}, []
    for r in excel_rows:
        ars = (r.get('ars') or '').strip()
        if ars:
            ars_groups.setdefault(ars, []).append(r)
        elif r.get('folder_key'):
            no_ars_rows.append(r)

    def _merge_or_append(fk, fk_type, ars, rows):
        new_checks = [
            {
                'check_id':   (row.get('check_id')   or '').strip(),
                'check_type': (row.get('check_type') or '').strip(),
            }
            for row in rows
        ]
        existing = by_fk.get(fk)
        if existing:
            existing['checks'].extend(new_checks)
            existing['multi_check'] = len(existing['checks']) > 1
            existing['raw_fields'] = _merge_dicts_prefer_existing(
                existing['raw_fields'], _merge_raw_fields(rows), 'later-discovered Excel'
            )
            new_note = _merge_note_from_rows(rows)
            if new_note:
                existing['_merge_note'] = ' | '.join(n for n in (existing.get('_merge_note', ''), new_note) if n)
            for url in _collect_links(rows):
                if url not in existing['document_links']:
                    existing['document_links'].append(url)
            return

        primary = rows[0]
        new_cand = {
            'ars':             ars,
            'check_id':        primary.get('check_id', ''),
            'name':            primary.get('name', ''),
            'emp_id':          primary.get('emp_id', ''),
            'folder_key':      fk,
            'folder_key_type': fk_type,
            'checks':          new_checks,
            'multi_check':     len(rows) > 1,
            'raw_fields':      _merge_raw_fields(rows),
            '_merge_note':     _merge_note_from_rows(rows),
            'document_links':  _collect_links(rows),
        }
        candidates.append(new_cand)
        by_fk[fk] = new_cand

    for ars, rows in ars_groups.items():
        _merge_or_append(ars, 'ars', ars, rows)
    for row in no_ars_rows:
        fk = row.get('folder_key', '')
        _merge_or_append(fk, row.get('folder_key_type', 'unknown'), '', [row])


def _merge_body_candidates_into(candidates, body_candidate_items, default_source=''):
    """
    Merge candidate leads pulled from email body text (outer email or any
    nested .eml/.msg found during discovery) into `candidates` (mutated in
    place). Matched against existing candidates by identifier first (ARS,
    then emp_id/UAN, then name) — Excel stays authoritative per Rule 9: a
    match backfills raw_fields as supplementary context rather than
    overwriting identity fields. No match means a genuinely new candidate
    not covered by any Excel — added as a new auto-matched entry, the same
    way a candidate only present in a second Excel attachment already is.

    body_candidate_items: list of either a bare body_candidate_data dict, or
    a (dict, source_label) tuple when the item came from discovery (so the
    manifest can note which forwarded email it was found in).
    """
    by_fk = {c['folder_key']: c for c in candidates if c.get('folder_key')}

    for item in body_candidate_items:
        if isinstance(item, tuple):
            bc, source_label = item
        else:
            bc, source_label = item, default_source

        id_type = bc.get('identifier_type', 'other')
        ident   = bc.get('identifier', '')
        name    = bc.get('name', '')
        if id_type == 'ars':
            fk, fk_type = ident, 'ars'
        elif id_type in ('emp_id', 'uan'):
            fk, fk_type = ident or name, 'emp_id' if ident else 'name'
        else:
            fk, fk_type = ident or name, 'name'
        if not fk:
            continue

        existing = by_fk.get(fk)
        note = f'Also mentioned in {source_label}' if source_label else ''
        if existing:
            existing['raw_fields'] = _merge_dicts_prefer_existing(
                existing['raw_fields'], bc.get('extra_fields') or {}, source_label or 'email body'
            )
            if note:
                existing['_merge_note'] = ' | '.join(n for n in (existing.get('_merge_note', ''), note) if n)
            continue

        new_cand = {
            'ars':             ident if id_type == 'ars' else '',
            'check_id':        '',
            'name':            name,
            'emp_id':          ident if id_type in ('emp_id', 'uan') else '',
            'folder_key':      fk,
            'folder_key_type': fk_type,
            'checks':          [],
            'multi_check':     False,
            '_source':         'body',
            '_merge_note':     note,
            'raw_fields':      bc.get('extra_fields') or {},
            'document_links':  [],
        }
        candidates.append(new_cand)
        by_fk[fk] = new_cand


def _copy_raw_into_extract(paths, run_extract):
    """
    Copy files into run_extract flat, skipping any name collision by suffixing.
    Used only when a whole pipeline STEP fails outright (see _degrade() in
    process()) — anything not yet unpacked/classified still needs to reach CST
    for manual triage, so it's pulled into the same pool build_file_tree() scans,
    instead of vanishing with the run_upload/run_extract temp dirs on cleanup.
    """
    os.makedirs(run_extract, exist_ok=True)
    for p in paths:
        if not os.path.isfile(p):
            continue
        fname = os.path.basename(p)
        dest  = os.path.join(run_extract, fname)
        if os.path.exists(dest):
            base, ext = os.path.splitext(fname)
            n = 2
            while os.path.exists(dest):
                dest = os.path.join(run_extract, f'{base}_{n}{ext}')
                n += 1
        try:
            shutil.copy2(p, dest)
        except Exception:
            pass


app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = int(os.environ.get('MAX_UPLOAD_MB', 200)) * 1024 * 1024

UPLOAD_FOLDER  = os.path.join(os.path.dirname(__file__), 'uploads')
EXTRACT_FOLDER = os.path.join(os.path.dirname(__file__), 'extracted')
OUTPUT_FOLDER  = os.path.join(os.path.dirname(__file__), 'output')

@app.route('/')
def index():
    from flask import make_response
    r = make_response(render_template('index.html'))
    r.headers['Cache-Control'] = 'no-store'
    return r


@app.route('/delete-output', methods=['POST'])
def delete_output():
    data         = request.get_json(force=True, silent=True) or {}
    batch_folder = data.get('batch_folder', '').strip()

    # Reject anything that isn't a bare folder name (no path separators, no traversal)
    if (
        not batch_folder
        or '/' in batch_folder
        or '\\' in batch_folder
        or '..' in batch_folder
        or not batch_folder.startswith('output_batch_')
    ):
        return jsonify({'error': 'Invalid folder name'}), 400

    full_path = os.path.join(OUTPUT_FOLDER, batch_folder)
    if not os.path.isdir(full_path):
        return jsonify({'error': 'Folder not found'}), 404

    shutil.rmtree(full_path, ignore_errors=True)
    return jsonify({'ok': True, 'deleted': batch_folder})


@app.route('/process', methods=['POST'])
def process():
    subject     = request.form.get('subject', '').strip()
    body        = request.form.get('body', '').strip()
    attachments = request.files.getlist('attachments')

    if not subject and not body and not attachments:
        return jsonify({'error': 'Provide at least a subject, body, or one attachment.'}), 400

    # Stream files straight to disk before the SSE generator starts — the request
    # context is still live here, and f.save() never buffers the whole file in RAM.
    # Dirs are created now so the generator's finally block can always clean them up.
    run_id      = uuid.uuid4().hex
    run_upload  = os.path.join(UPLOAD_FOLDER,  run_id)
    run_extract = os.path.join(EXTRACT_FOLDER, run_id)
    # Sibling to run_extract, NOT nested inside it — per-candidate document-link
    # downloads land here so they never get picked up by build_file_tree(run_extract)
    # (which feeds the general AI-grouping pool). Kept isolated so these files can be
    # assigned directly to their known candidate without going through grouping_agent.
    run_link_extract = os.path.join(EXTRACT_FOLDER, run_id + '_candlinks')
    os.makedirs(run_upload)
    os.makedirs(run_extract)
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    saved_paths = []
    try:
        for f in attachments:
            if not f.filename:
                continue
            safe_name = os.path.basename(f.filename)
            path = os.path.join(run_upload, safe_name)
            f.save(path)  # Werkzeug streams directly to disk
            saved_paths.append(path)
    except Exception as e:
        shutil.rmtree(run_upload,  ignore_errors=True)
        shutil.rmtree(run_extract, ignore_errors=True)
        return jsonify({'error': f'Failed to save attachments: {e}'}), 500

    def pipeline():
        def ev(data):
            return f"data: {json.dumps(data, ensure_ascii=False)}\n\n"

        try:
            attachment_filenames = [os.path.basename(p) for p in saved_paths]

            # ── Shared state, initialised up-front ─────────────────────────
            # If a whole pipeline STEP fails outright below (not a single file/link
            # — those already degrade gracefully throughout, see extraction_errors/
            # pending_entries/_UNASSIGNED routing), _degrade() hands CST whatever's
            # accumulated in these so far instead of aborting with nothing. They're
            # initialised here so every one of them is always safe to read/append,
            # no matter how early in the pipeline the failure happens.
            extraction_errors    = []
            pending_entries      = []
            candidates            = []
            skipped_candidates    = []
            candidate_link_files  = {}
            groups = unassigned = ambiguous = []

            def _degrade(reason):
                """
                A whole pipeline STEP failed (AI retries exhausted, unreadable Excel,
                an unexpected exception) — not a single file/link. Rather than discard
                the run, dump everything accumulated so far — plus any attachment that
                never got a chance to be extracted/classified — into _UNASSIGNED, so
                CST still gets a batch + manifest to work from, and can re-run just the
                corrected input afterward instead of losing the whole batch.
                """
                nonlocal groups, unassigned, ambiguous
                extraction_errors.append({'identifier': 'pipeline_failure', 'path': '', 'reason': reason})
                _copy_raw_into_extract(saved_paths, run_extract)
                fps = build_file_tree(run_extract)
                groups, unassigned, ambiguous = [], [
                    {
                        'path':   p,
                        'reason': f'Pipeline step failed before this file could be classified — manual triage required: {reason}',
                    }
                    for p in fps
                ], []

            def _do_assembly():
                """
                Step 5 — build output folders + manifest.csv, then emit the final SSE
                event. Shared by the normal end-of-pipeline path and every _degrade()
                fallback, so both produce an identically-shaped batch + manifest.
                """
                yield ev({'step': 'Assembling output folders…'})
                files_in_system = (
                    bool(build_file_tree(run_extract))
                    or bool(candidate_link_files)
                    or bool(pending_entries)
                )
                manifest_rows, stats, batch_folder = assemble_output(
                    groups,
                    unassigned,
                    ambiguous,
                    run_extract,
                    OUTPUT_FOLDER,
                    all_candidates=candidates,
                    pending_links=pending_entries,
                    direct_files=candidate_link_files,
                    skipped_candidates=skipped_candidates,
                    files_in_system=files_in_system,
                    run_id=run_id,
                )
                yield ev({'step': f'Output ready in {os.path.basename(batch_folder)}'})

                _unassigned_statuses = {
                    'Unassigned',
                    'Ambiguous',
                    'Unassigned — low confidence',
                    'Unassigned — cross-candidate duplicate',
                    'ERROR — source file not found',
                    'link_pending_manual_fetch',
                    'skipped_capacity_limit',
                }
                folder_buckets = {}
                for row in manifest_rows:
                    fk = row['folder_key']
                    s  = row['status']
                    if not fk or s in _unassigned_statuses:
                        continue
                    if fk not in folder_buckets:
                        op = row.get('output_path', '')
                        folder_buckets[fk] = {
                            'folder_key':     fk,
                            'candidate_name': row['candidate_name'],
                            'emp_id':         row['emp_id'],
                            'output_folder':  os.path.dirname(op).replace('\\', '/'),
                            'file_count':     0,
                            'status':         'Auto Matched',
                            'files':          [],
                        }
                    bucket = folder_buckets[fk]
                    bucket['file_count'] += 1
                    bucket['files'].append({
                        'filename':    row['filename'],
                        'confidence':  row['confidence'],
                        'reason':      row['reason'],
                        'status':      row['status'],
                        'source_type': row['source_type'],
                    })
                    if 'Needs Review' in s:
                        bucket['status'] = 'Needs Review'
                    elif s == 'info_only_no_documents' and bucket['status'] == 'Auto Matched':
                        bucket['status'] = 'Info Only'

                results = list(folder_buckets.values())

                unassigned_out = [
                    {
                        'filename':       row['filename'],
                        'link_url':       row.get('link_url', ''),
                        'folder_key':     row.get('folder_key', ''),
                        'candidate_name': row.get('candidate_name', ''),
                        'reason':         row['reason'],
                        'status':         row['status'],
                        'source_type':    row['source_type'],
                    }
                    for row in manifest_rows
                    if not row['folder_key'] or row['status'] in _unassigned_statuses
                ]

                yield ev({
                    'done':       True,
                    'results':    results,
                    'unassigned': unassigned_out,
                    'manifest':   manifest_rows,
                    'extraction_errors': [
                        {'path': e.get('path', e.get('identifier', '')), 'reason': e.get('reason', '')}
                        for e in extraction_errors
                    ],
                    'stats': {
                        'total_files':            len(manifest_rows),
                        'auto_matched':           stats['auto_matched'],
                        'review':                 stats['review'],
                        'unassigned':             stats['unassigned'],
                        'info_only':              stats['info_only'],
                        'skipped_capacity_limit': stats['skipped_capacity_limit'],
                        'failed_extractions':     len(extraction_errors),
                        'pending_links':          len(pending_entries),
                    },
                    'batch_folder': os.path.basename(batch_folder),
                })

            # ── Step -1: Pre-extract top-level .eml/.msg attachments FIRST ────
            # A forwarded email attachment usually carries the real context for what
            # the rest of the batch is about — read its body text BEFORE the main
            # email-classification call, not just later as grouping context. This is
            # a deterministic extension check (no AI needed to know a file is .eml/.msg).
            # Nested .eml/.msg found INSIDE a ZIP/archive can't be pre-read this way —
            # those are only discovered once Step 1 unpacks the archive.
            eml_msg_paths = [
                p for p in saved_paths
                if os.path.splitext(p)[1].lower() in ('.eml', '.msg')
            ]
            if eml_msg_paths:
                yield ev({'step': f'Reading {len(eml_msg_paths)} forwarded email attachment(s) first…'})
                try:
                    extraction_errors = extract_all(eml_msg_paths, run_extract)
                except Exception as e:
                    yield ev({'step': f'⚠️ Reading forwarded email attachment(s) failed — falling back to manual triage: {e}'})
                    _degrade(str(e))
                    yield from _do_assembly()
                    return

            inner_email_context = collect_body_context(run_extract)

            # ── Step 0: Email parsing ─────────────────────────────────────
            yield ev({'step': 'Parsing email with AI…'})
            try:
                parsed = parse_email(
                    subject, body, attachment_filenames,
                    inner_context=inner_email_context,
                )
            except Exception as e:
                yield ev({'step': f'⚠️ AI email classification failed — falling back to manual triage: {e}'})
                _degrade(str(e))
                yield from _do_assembly()
                return

            excel_fnames   = set(parsed.get('excel_attachments',           []))
            links          = parsed.get('links', [])
            body_candidates = parsed.get('body_candidate_data')

            n_excel   = len(excel_fnames)
            n_archive = len(parsed.get('archive_attachments', []))
            n_email   = len(parsed.get('email_container_attachments', []))
            n_loose   = len(parsed.get('loose_file_attachments', []))
            n_links   = len(links)
            yield ev({
                'step': (
                    f'Email classified — {n_excel} Excel, {n_archive} archive(s), '
                    f'{n_email} email container(s), {n_loose} loose file(s), {n_links} link(s)'
                ),
            })

            # ── Step 1: Extract remaining non-Excel attachments ───────────
            # .eml/.msg were already extracted in Step -1 — excluded here so they
            # aren't processed twice.
            excel_paths     = [p for p in saved_paths if os.path.basename(p) in excel_fnames]
            non_excel_paths = [
                p for p in saved_paths
                if os.path.basename(p) not in excel_fnames and p not in eml_msg_paths
            ]

            if non_excel_paths:
                yield ev({'step': f'Extracting {len(non_excel_paths)} attachment(s)…'})
                try:
                    extraction_errors.extend(extract_all(non_excel_paths, run_extract))
                except Exception as e:
                    yield ev({'step': f'⚠️ Extracting attachment(s) failed — falling back to manual triage: {e}'})
                    _degrade(str(e))
                    yield from _do_assembly()
                    return

            # ── Step 2 (recursive discovery): links + nested emails, no matter ──
            # ── how deep they're buried ──────────────────────────────────────
            # Replaces the old single-pass "download the outer body's links once"
            # step. A link can lead to a ZIP containing a forwarded .eml whose own
            # body has another link, mentions another candidate, etc. — this keeps
            # chasing that chain (each nested .eml/.msg gets its own real AI
            # classification pass, not just a text dump) until nothing new turns
            # up or the hop cap (pipeline.discovery.MAX_LINK_HOPS) is reached.
            # Excel/CSV files found anywhere in that chain — top-level, nested in
            # a ZIP, nested behind a link — are collected here and parsed once,
            # together, in Step 3 below.
            seen_files = set()
            init_excels, init_email_items, _init_loose = scan_new_content(run_extract, seen_files)
            discovered_excels   = list(init_excels)
            pending_email_texts = [(text, rel_dir, 1) for text, rel_dir in init_email_items]
            stage1_body_candidates = []

            if links or pending_email_texts:
                yield ev({
                    'step': (
                        f'Resolving {len(links)} link(s)'
                        + (f' and {len(pending_email_texts)} nested email(s)' if pending_email_texts else '')
                        + '…'
                    ),
                })
                try:
                    stage1_result = None
                    for event in discover(links, pending_email_texts, run_extract, run_upload, seen_files):
                        if isinstance(event, str):
                            yield ev({'step': event})
                        else:
                            stage1_result = event['result']
                    discovered_excels.extend(stage1_result['excel_paths'])
                    stage1_body_candidates = stage1_result['body_candidates']
                    pending_entries.extend(stage1_result['pending_entries'])
                    extraction_errors.extend(stage1_result['extraction_errors'])
                    if pending_entries:
                        yield ev({'step': f'{len(pending_entries)} link(s) flagged for manual fetch'})
                except Exception as e:
                    yield ev({'step': f'⚠️ Link/email discovery failed — falling back to manual triage: {e}'})
                    _degrade(str(e))
                    yield from _do_assembly()
                    return

            if discovered_excels:
                yield ev({
                    'step': (
                        f'Found {len(discovered_excels)} Excel/CSV file(s) nested in extracted '
                        f'content — merging into the candidate list'
                    ),
                })
                excel_paths = excel_paths + discovered_excels

            # ── Step 3: Parse Excel / determine candidate list ────────────
            if excel_paths:
                if len(excel_paths) == 1:
                    yield ev({'step': 'Parsing Excel / CSV spreadsheet…'})
                    try:
                        excel_rows = parse_excel(excel_paths[0])
                    except Exception as e:
                        yield ev({'step': f'⚠️ Excel/CSV parsing failed — falling back to manual triage: {e}'})
                        _degrade(str(e))
                        yield from _do_assembly()
                        return
                else:
                    yield ev({
                        'step': f'Parsing {len(excel_paths)} Excel/CSV files and cross-referencing…'
                    })
                    try:
                        excel_rows = parse_excel_multi(excel_paths)
                    except Exception as e:
                        yield ev({'step': f'⚠️ Excel/CSV parsing failed — falling back to manual triage: {e}'})
                        _degrade(str(e))
                        yield from _do_assembly()
                        return

                # Group rows by ARS so that one candidate = one folder_key.
                # Multiple rows per ARS = multiple checks (Employment, Education, etc.)
                # → stored as a checks list; multi_check=True triggers check sub-folders
                # in the assembler when the client's files are also check-organised.
                # Rows with no ARS use check_id / emp_id / name as their folder_key
                # and each get their own separate folder (one row = one folder).
                _build_candidates_from_excel_rows(excel_rows, candidates)

                n_checks = sum(len(c['checks']) for c in candidates)
                yield ev({
                    'step': (
                        f'Excel ready — {len(candidates)} candidate(s), '
                        f'{n_checks} check row(s) found'
                    ),
                })

            # Merge in candidate leads from body text — the outer email/subject
            # (Step 0) and every nested .eml/.msg discovered above (Step 2).
            # Excel stays authoritative (Rule 9): a match backfills supplementary
            # context only; a genuinely new ARS/ID not covered by any Excel
            # becomes its own auto-matched candidate, exactly like a candidate
            # who only exists in a second Excel attachment already does.
            if body_candidates:
                _merge_body_candidates_into(candidates, body_candidates, default_source='email body')
            if stage1_body_candidates:
                _merge_body_candidates_into(candidates, stage1_body_candidates)
                yield ev({
                    'step': f'{len(stage1_body_candidates)} candidate lead(s) found in nested forwarded email(s)',
                })

            # ── Step 3b (recursive discovery, per candidate): document links ──
            # A link found anywhere in an Excel row belongs to ONE specific candidate
            # (unlike a body-text link, whose owner has to be inferred by AI grouping).
            # Each one is resolved the same recursive way as Step 2 — chasing nested
            # links/emails found inside it too — but isolated per candidate so the
            # resulting files can be assigned directly, bypassing group_files()
            # entirely. A candidate can have more than one real link (one per check
            # type), so this works per-LINK, not per-candidate.
            #
            # Run downloads in a bounded thread pool, not one-at-a-time — a batch can
            # legitimately have 1000+ links to fetch (photos, certificates, etc.), and
            # downloading those sequentially can run long enough that the browser's SSE
            # connection gives up with no progress ever shown. Progress is streamed
            # periodically (not per-link) so large batches don't flood the UI, while
            # still keeping the connection alive throughout.
            # folder_key -> [(absolute_path, relative_path, source_url), ...]
            _link_work_items = [
                (cand, url) for cand in candidates for url in (cand.get('document_links') or [])
            ]
            if _link_work_items:
                yield ev({
                    'step': f'Downloading {len(_link_work_items)} candidate document link(s) from Excel…'
                })

                def _download_one(i, cand, url):
                    fk = cand['folder_key']
                    cand_upload_dir  = os.path.join(run_upload, '_candidate_links', str(i))
                    cand_extract_dir = os.path.join(run_link_extract, str(i))
                    os.makedirs(cand_upload_dir,  exist_ok=True)
                    os.makedirs(cand_extract_dir, exist_ok=True)
                    cand_seen = set()
                    disc_excels, disc_body_cands, disc_pending, disc_errors = [], [], [], []
                    try:
                        for event in discover(
                            [{'url': url, 'type': None}], [], cand_extract_dir, cand_upload_dir,
                            cand_seen, owner_label=f'{fk}: ',
                        ):
                            if not isinstance(event, str):
                                result         = event['result']
                                disc_excels     = result['excel_paths']
                                disc_body_cands = result['body_candidates']
                                disc_pending    = result['pending_entries']
                                disc_errors     = result['extraction_errors']
                    except Exception as e:
                        disc_errors = [{'identifier': url, 'path': url, 'reason': str(e)}]

                    # Keep the absolute path (to copy from), the path relative to this
                    # link's own extract dir (for check-subfolder detection in
                    # assembler.py — e.g. a ZIP that unpacked into Employment/doc1.pdf),
                    # and the source URL (so each file's manifest row shows which link
                    # produced it).
                    linked_files = [
                        (os.path.join(cand_extract_dir, rel), rel, url)
                        for rel in build_file_tree(cand_extract_dir)
                    ]
                    return fk, cand, url, linked_files, disc_pending, disc_errors, disc_excels, disc_body_cands

                _link_ok = _link_pending = _link_failed = 0
                _completed = 0
                _total = len(_link_work_items)
                _progress_every = max(1, min(25, _total // 20 or 1))
                _late_excels, _late_body_candidates = [], []

                with concurrent.futures.ThreadPoolExecutor(max_workers=_LINK_DOWNLOAD_WORKERS) as pool:
                    futures = [
                        pool.submit(_download_one, i, cand, url)
                        for i, (cand, url) in enumerate(_link_work_items)
                    ]
                    for future in concurrent.futures.as_completed(futures):
                        _completed += 1
                        try:
                            fk, cand, url, linked_files, pending, dl_errors, disc_excels, disc_body_cands = future.result()
                        except Exception as e:
                            _link_failed += 1
                            extraction_errors.append({
                                'identifier': 'candidate document link',
                                'path':       '',
                                'reason':     f'Unexpected error downloading link: {e}',
                            })
                            continue

                        extraction_errors.extend(dl_errors)
                        _late_excels.extend(disc_excels)
                        if disc_body_cands:
                            _late_body_candidates.extend(disc_body_cands)
                        if linked_files:
                            candidate_link_files.setdefault(fk, []).extend(linked_files)
                            _link_ok += 1
                        for p in pending:
                            p['folder_key']     = fk
                            p['candidate_name'] = cand.get('name', '')
                            p['emp_id']         = cand.get('emp_id', '')
                            pending_entries.append(p)
                            _link_pending += 1

                        if _completed % _progress_every == 0 or _completed == _total:
                            yield ev({
                                'step': (
                                    f'Candidate document links — {_completed}/{_total} processed '
                                    f'({_link_ok} downloaded, {_link_pending} pending, {_link_failed} failed)'
                                ),
                            })

                # Rare, but "no matter where found" applies here too: an Excel or a
                # new candidate lead surfacing deep inside one candidate's own
                # document-link chain (e.g. a nested forwarded email mentioning a
                # second candidate not in anyone's Excel). Merged after the pool
                # completes, not inside worker threads, to keep the Gemini calls
                # and shared candidate-list mutation single-threaded.
                if _late_excels:
                    yield ev({
                        'step': f'Found {len(_late_excels)} Excel/CSV file(s) inside candidate document links — merging…',
                    })
                    try:
                        late_rows = (
                            parse_excel(_late_excels[0]) if len(_late_excels) == 1
                            else parse_excel_multi(_late_excels)
                        )
                        _build_candidates_from_excel_rows(late_rows, candidates)
                    except Exception as e:
                        extraction_errors.append({
                            'identifier': 'nested excel (candidate link)', 'path': '', 'reason': str(e),
                        })
                if _late_body_candidates:
                    _merge_body_candidates_into(candidates, _late_body_candidates)
                    yield ev({
                        'step': f'{len(_late_body_candidates)} candidate lead(s) found inside candidate document links',
                    })

            # ── Step 4: AI grouping (conditional) ────────────────────────
            # Exclude the nested Excel/CSV file(s) already consumed above as a
            # candidate-identity source — they're metadata, not a document to group.
            excel_rel_exclude = {
                os.path.relpath(p, run_extract).replace('\\', '/')
                for p in discovered_excels
            }
            file_paths = [p for p in build_file_tree(run_extract) if p not in excel_rel_exclude]
            has_files  = bool(file_paths)

            # Build body context: main email body + any inner .eml/.msg body texts.
            # Passed verbatim into the grouping prompt as supplementary raw text —
            # candidate leads and links from this content are already extracted
            # structurally above, but the raw wording can still help the grouping
            # AI disambiguate an otherwise-generic filename.
            _body_ctx_parts = []
            if subject or body:
                _body_ctx_parts.append(
                    f'Subject: {subject}\n\n{body}'.strip()
                )
            _inner_ctx = collect_body_context(run_extract)
            if _inner_ctx:
                _body_ctx_parts.append(_inner_ctx)
            body_context = '\n\n---\n\n'.join(_body_ctx_parts) or None

            if not has_files and not pending_entries and candidates:
                # Fully resolved — no documents at all, just candidate data from body/Excel
                yield ev({'step': 'No document files — all candidate(s) will be info-only'})
                groups, unassigned, ambiguous = [], [], []

            elif not has_files and not pending_entries:
                yield ev({'error': 'No files and no candidate data found. Nothing to process.'})
                return

            else:
                if not candidates:
                    yield ev({
                        'step': (
                            f'No candidate list found — all {len(file_paths)} file(s) '
                            f'will go to _UNASSIGNED'
                        ),
                    })
                    groups     = []
                    unassigned = [
                        {'path': p, 'reason': 'No candidate list — manual assignment required'}
                        for p in file_paths
                    ]
                    ambiguous  = []
                else:
                    yield ev({
                        'step': f'Sending {len(file_paths)} file(s) to AI grouping agent…',
                        'agent_input': {
                            'file_count':      len(file_paths),
                            'candidate_count': len(candidates),
                            'files':           file_paths,
                            'candidates': [
                                {
                                    'folder_key':      c.get('folder_key'),
                                    'folder_key_type': c.get('folder_key_type'),
                                    'name':            c.get('name'),
                                    'emp_id':          c.get('emp_id'),
                                    'checks':          c.get('checks', []),
                                }
                                for c in candidates
                            ],
                        },
                    })
                    groups = unassigned = ambiguous = None
                    for _ge in group_files(file_paths, candidates, run_extract, body_context=body_context, skipped_candidates=skipped_candidates):
                        if isinstance(_ge, str):
                            yield ev({'step': _ge})
                        else:
                            groups, unassigned, ambiguous = _ge['result']

                    assigned_count = sum(len(g.get('files', [])) for g in groups)
                    yield ev({
                        'step': (
                            f'AI responded — {assigned_count} file(s) assigned to '
                            f'{len(groups)} candidate(s), {len(unassigned)} unassigned, '
                            f'{len(ambiguous)} ambiguous'
                        ),
                        'agent_output': {
                            'groups':     groups,
                            'unassigned': unassigned,
                            'ambiguous':  ambiguous,
                        },
                    })

            # ── Step 5: Assemble output ───────────────────────────────────
            yield from _do_assembly()

        except Exception as e:
            # Last-resort safety net for anything not already caught by a specific
            # step above (e.g. a bug inside grouping_agent/assembler themselves).
            # Still try to hand CST a real batch via the same _degrade() path rather
            # than just reporting an error with nothing produced.
            try:
                yield ev({'step': f'⚠️ Unexpected pipeline error — falling back to manual triage: {e}'})
                _degrade(str(e))
                yield from _do_assembly()
            except Exception as e2:
                yield ev({'error': f'Unexpected error: {e}; fallback also failed: {e2}'})

        finally:
            shutil.rmtree(run_upload,      ignore_errors=True)
            shutil.rmtree(run_extract,     ignore_errors=True)
            shutil.rmtree(run_link_extract, ignore_errors=True)

    return Response(stream_with_context(pipeline()), content_type='text/event-stream')


if __name__ == '__main__':
    debug = os.environ.get('FLASK_DEBUG', 'false').lower() == 'true'
    port  = int(os.environ.get('PORT', 5000))
    app.run(debug=debug, port=port, threaded=True)
