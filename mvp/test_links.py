"""
Standalone test harness for link_handler.py.

Usage:
    1. Paste your sample links into SAMPLE_LINKS below.
    2. Run from the mvp/ directory:
           python test_links.py

For each link the script reports:
    OK       — file downloaded successfully, shows filename and byte size
    PENDING  — anonymous access failed or not applicable; shows the reason
    FAIL     — plain-HTTP download errored (zip_link / unknown type only)

No pipeline steps run — only link_handler is exercised.
Downloaded files are written to a temp directory and deleted when the script finishes.
"""

import os
import sys
import shutil
import tempfile

# ─── Paste your sample links here ────────────────────────────────────────────
#
# Each entry is a dict with:
#   url  — the full link as it would appear in the email body
#   type — one of: "drive" | "sharepoint" | "zip_link" | "unknown"
#          If omitted, link_handler will classify it automatically.
#
SAMPLE_LINKS = [
    # Google Drive — public file (should succeed)
    # {"url": "https://drive.google.com/file/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms/view", "type": "drive"},

    # Google Drive — restricted/private file (should fail → PENDING)
    # {"url": "https://drive.google.com/file/d/XXXX_PRIVATE_ID/view", "type": "drive"},

    # Google Drive — folder link (should go PENDING immediately, no attempt)
    # {"url": "https://drive.google.com/drive/folders/1234ABCD", "type": "drive"},

    # Google Drive — Google Sheet (should go PENDING immediately, no attempt)
    # {"url": "https://docs.google.com/spreadsheets/d/1ABC.../edit", "type": "drive"},

    # SharePoint — public/anonymous link (should succeed with ?download=1)
    # {"url": "https://tenant.sharepoint.com/:b:/g/personal/user/ABC123?e=XYZ", "type": "sharepoint"},

    # SharePoint — restricted link (should fail → PENDING)
    # {"url": "https://tenant.sharepoint.com/sites/Team/Docs/file.pdf", "type": "sharepoint"},

    # Direct ZIP link (plain HTTP, no auth)
    # {"url": "https://example.com/files/sample.zip", "type": "zip_link"},
]
# ─────────────────────────────────────────────────────────────────────────────


def _bar(char='─', width=70):
    return char * width


def run():
    if not SAMPLE_LINKS:
        print('\nNo sample links configured.')
        print('Edit the SAMPLE_LINKS list in this file and re-run.\n')
        sys.exit(0)

    # Ensure imports resolve from the mvp/ directory
    mvp_dir = os.path.dirname(os.path.abspath(__file__))
    if mvp_dir not in sys.path:
        sys.path.insert(0, mvp_dir)

    from pipeline.link_handler import handle_links

    tmp_dir = tempfile.mkdtemp(prefix='link_test_')
    print(f'\nTemp download dir: {tmp_dir}')
    print(_bar())

    try:
        downloaded, pending, errors = handle_links(SAMPLE_LINKS, tmp_dir)

        rows = []

        for path in downloaded:
            size  = os.path.getsize(path) if os.path.exists(path) else 0
            fname = os.path.basename(path)
            rows.append(('OK', fname, f'{size:,} bytes', ''))

        for p in pending:
            rows.append(('PENDING', p['url'], '', p.get('reason', '')))

        for e in errors:
            rows.append(('FAIL', e['url'], '', e.get('reason', '')))

        # Print results table
        print(f'\n{"RESULT":<10}  {"FILE / URL":<45}  {"SIZE":<12}  REASON')
        print(_bar())
        for result, target, size, reason in rows:
            label = target if len(target) <= 45 else '…' + target[-43:]
            print(f'{result:<10}  {label:<45}  {size:<12}  {reason}')

        print(_bar())
        print(
            f'\nSummary: {len(downloaded)} downloaded | '
            f'{len(pending)} pending (manual) | '
            f'{len(errors)} failed\n'
        )

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == '__main__':
    run()
