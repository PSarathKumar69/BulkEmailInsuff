import os

from .email_parser import parse_email
from .link_handler import handle_links
from .extractor import extract_all

# A link found inside content that itself came from a previous link counts as
# the next hop. Bounds worst-case cost/time on a batch with long forwarding
# chains (email -> zip -> email -> zip -> ...) now that MAX_CANDIDATES no
# longer caps how many candidates (and therefore how many chains) a run has.
MAX_LINK_HOPS = 3

_SPREADSHEET_EXTS = {'.xlsx', '.xls', '.xlsb', '.csv', '.ods'}
_JUNK_DIRS  = {'__MACOSX', '__pycache__'}
_JUNK_FILES = {'Thumbs.db', 'desktop.ini', 'ehthumbs.db', 'thumbs.db'}


def scan_new_content(root, seen):
    """
    Walk `root`, returning content not already present in `seen` — a set this
    function mutates so repeated calls over a growing directory tree only
    ever report what's new since the last call:

      excel_paths      : newly found .xlsx/.xls/.xlsb/.csv/.ods files
      email_body_items : (text, rel_dir) for newly found email_body.txt files
                         written by extractor.py when it unpacks a .eml/.msg
      loose_paths       : everything else newly found (not returned to
                         candidate-grouping here — callers that need the full
                         file tree still use grouping_agent.build_file_tree)

    Generalises the two one-off scans this pipeline used to do exactly once
    (a single pass for nested Excel files, a single pass for nested email
    body text) into something that can be called repeatedly as new content
    keeps appearing from chased links.
    """
    excel_paths, email_body_items, loose_paths = [], [], []

    for dirpath, dirs, filenames in os.walk(root):
        dirs[:] = [d for d in dirs if not d.startswith('.') and d not in _JUNK_DIRS]
        for fname in filenames:
            if fname.startswith('.') or fname in _JUNK_FILES:
                continue
            full = os.path.join(dirpath, fname)
            if full in seen:
                continue
            seen.add(full)

            if fname == 'email_body.txt':
                try:
                    with open(full, 'r', encoding='utf-8', errors='replace') as f:
                        text = f.read().strip()
                except Exception:
                    text = ''
                if text:
                    rel_dir = os.path.relpath(dirpath, root).replace('\\', '/')
                    email_body_items.append((text, rel_dir))
                continue

            ext = os.path.splitext(fname)[1].lower()
            if ext in _SPREADSHEET_EXTS:
                excel_paths.append(full)
            else:
                loose_paths.append(full)

    return excel_paths, email_body_items, loose_paths


def discover(links, email_texts, extract_root, upload_root, seen_files, owner_label=''):
    """
    Generator. Recursively resolves `links` (each starting at hop 1) plus
    `email_texts` (already-extracted email bodies waiting to be classified —
    e.g. from a top-level attachment Step 1 already unpacked), and anything
    newly found while doing so:

      - A downloaded link is extracted (extract_all already recurses through
        nested archives/emails on disk). Any new email_body.txt or Excel/CSV
        file that appears is picked up via scan_new_content.
      - Each not-yet-classified email body gets its own AI classification
        pass (email_parser.parse_email, called with subject='' since this is
        nested content, not the outer email) — not just dumped as text
        context — so links and candidate leads buried in a forwarded chain
        are reliably surfaced instead of getting lost in a giant text blob.
      - Links found that way are queued at hop+1. Once a link's hop exceeds
        MAX_LINK_HOPS, it is not fetched — it's recorded as a pending entry
        for manual CST fetch instead (never silently dropped).

    All downloaded/extracted content lands under `extract_root` — pass the
    shared run_extract for unowned/batch-level discovery (Step 2 today), or
    one isolated per-link directory for candidate-owned discovery (Step 3b
    today), exactly as the existing callers already isolate that content.

    Yields progress strings (prefixed with `owner_label` when given), then
    finally yields {'result': {...}}:
      excel_paths       : newly discovered spreadsheet paths (not yet parsed)
      body_candidates   : [(body_candidate_data_item, source_label), ...]
      pending_entries   : [{url, type, status, reason}, ...] for the manifest
      extraction_errors : [{identifier, path, reason}, ...]
    """
    discovered_excels          = []
    discovered_body_candidates = []
    pending_entries            = []
    extraction_errors          = []

    pending_links = [(link, 1) for link in (links or [])]
    pending_email_texts = [
        (text, rel_dir, hop) for (text, rel_dir, hop) in (email_texts or [])
    ]

    while pending_links or pending_email_texts:
        this_round_links, pending_links = pending_links, []
        for link, hop in this_round_links:
            url = (link.get('url') or '').strip()
            if not url:
                continue

            if hop > MAX_LINK_HOPS:
                pending_entries.append({
                    'url':    url,
                    'type':   link.get('type'),
                    'status': 'link_pending_manual_fetch',
                    'reason': (
                        f'Link-chase depth limit ({MAX_LINK_HOPS} hops) reached — '
                        'this link was found inside previously-chased content. '
                        'CST to fetch manually.'
                    ),
                })
                continue

            yield f'{owner_label}Fetching link (hop {hop}/{MAX_LINK_HOPS}): {url}'
            try:
                downloaded, dl_pending, dl_errors = handle_links([link], upload_root)
            except Exception as e:
                downloaded, dl_pending, dl_errors = [], [], [{'url': url, 'reason': str(e)}]

            pending_entries.extend(dl_pending)
            for err in dl_errors:
                extraction_errors.append({
                    'identifier': url, 'path': url, 'reason': err.get('reason', ''),
                })

            if downloaded:
                extraction_errors.extend(extract_all(downloaded, extract_root))
                new_excels, new_email_items, _new_loose = scan_new_content(extract_root, seen_files)
                discovered_excels.extend(new_excels)
                for text, rel_dir in new_email_items:
                    pending_email_texts.append((text, rel_dir, hop))

        this_round_emails, pending_email_texts = pending_email_texts, []
        for text, rel_dir, hop in this_round_emails:
            yield f'{owner_label}Reading forwarded email content ({rel_dir})…'
            try:
                nested = parse_email('', text, [])
            except Exception as e:
                extraction_errors.append({
                    'identifier': rel_dir, 'path': rel_dir,
                    'reason': f'Nested email classification failed: {e}',
                })
                continue

            for link in (nested.get('links') or []):
                pending_links.append((link, hop + 1))
            for bc in (nested.get('body_candidate_data') or []):
                discovered_body_candidates.append((bc, f'forwarded email ({rel_dir})'))

    yield {'result': {
        'excel_paths':       discovered_excels,
        'body_candidates':   discovered_body_candidates,
        'pending_entries':   pending_entries,
        'extraction_errors': extraction_errors,
    }}
