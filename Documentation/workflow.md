# CST Bulk Insufficiency — End-to-End Workflow

How a bulk insufficiency email goes from the input UI to finished output folders. Reflects the current state of `mvp/`, including everything built beyond the original `CLAUDE.md` spec (multi-Excel merging, per-candidate document links, unconditional metadata records, etc.)

---

## 0. Input

The user pastes into the UI (`templates/index.html` → `POST /process` in `app.py`):

| Field | Notes |
|---|---|
| Subject line | plain text |
| Email body | full thread as-is, including any Drive/SharePoint/ZIP links inline as plain text |
| Attachments | any number of files, any format (see **Supported Formats** below) |

Each request gets a fresh `run_id` (UUID) and three working directories:
- `uploads/<run_id>/` — raw attachments as uploaded
- `extracted/<run_id>/` — fully unpacked content (feeds AI grouping)
- `extracted/<run_id>_candlinks/` — **isolated** sibling directory for files downloaded from per-candidate Excel links (kept separate so they never get double-processed by AI grouping)

All three are wiped in the `finally:` block once the run completes, whether it succeeds or fails. Output batches in `output/` are never touched by this cleanup — they persist.

---

## Step 0 — Email Parsing (`pipeline/email_parser.py`)

One Gemini call reads: subject, body/thread text, and attachment **filenames only** (never file content). It classifies every attachment by extension into one of four buckets, extracts any URLs from the body text, and — if the body itself clearly states candidate-level data (e.g. "attaching UAN for these 2 candidates: ...") — extracts that as `body_candidate_data`.

Output:
```json
{
  "excel_attachments": [...],
  "archive_attachments": [...],
  "email_container_attachments": [...],
  "loose_file_attachments": [...],
  "links": [{"url": "...", "type": "drive|sharepoint|zip_link|unknown"}],
  "body_candidate_data": [...] | null,
  "has_files_or_links": true|false
}
```
Retries once on invalid JSON; raises to the UI on a second failure (no API key → immediate error, since every AI step requires one).

---

## Step 1 — Container Unpacking (`pipeline/extractor.py`)

Every attachment **except** those classified as Excel is routed through `extract_all()`:

- **Archive** (`.zip`, `.rar`, `.7z`, `.tar`/`.tgz`/`.tbz2`/`.txz`) → recursively unpacked into `extracted/<run_id>/<archive_name>/`, including archives nested inside archives (up to 10 levels deep). Password-protected or corrupt archives are skipped and logged — the rest of the batch continues.
- **Email container** (`.eml`, `.msg`) → body text saved as `email_body.txt` (used later as extra grouping context), attachments extracted and recursively processed the same way.
- **Everything else** (PDF, images, DOCX, etc.) → copied flat into `extracted/`, untouched.

Excel/CSV files are deliberately **excluded** from this step — they're data sources, not documents, and are handled separately in Step 3.

---

## Step 2 — Email-Body Link Handling (`pipeline/link_handler.py`)

For every link found in the email body text (Step 0), `handle_links()` attempts a real download (this goes further than a stub — despite links historically being described as detection-only, the code actually downloads):

| Link type | Behavior |
|---|---|
| Drive file, "anyone with the link" | Downloaded via `gdown` |
| Drive folder, "anyone with the link" | Entire folder tree downloaded, structure preserved |
| Drive file/folder, restricted | Flagged `link_pending_manual_fetch` with a reason (needs proper sharing or service-account access) |
| Google Docs/Sheets/Slides | Always flagged pending — can't be downloaded as one file |
| SharePoint/OneDrive, anonymous access | Downloaded via HTTP GET with `?download=1` |
| SharePoint/OneDrive, tenant-restricted | Flagged pending |
| Generic ZIP link / unknown URL | Plain HTTP GET attempt |

Anything successfully downloaded is fed back through `extract_all()` into the same shared `extracted/<run_id>/` pool. Anything that fails or requires manual fetch becomes a `pending_entries` row that surfaces later in the manifest — with **no** candidate attached, since a body-level link isn't inherently tied to one person.

---

## Step 3 — Excel Parsing / Candidate List (`pipeline/excel_reader.py`)

This is where the actual candidate list comes from. Three sources, in priority order:

1. **One Excel/CSV attachment** → `parse_excel()`. AI reads the first ~10–20 sample rows, identifies the header row, and maps each column to a standard field: `ars_number`, `check_id`, `check_type`, `candidate_name`, `emp_id`, **`document_link`** (a Drive/SharePoint/ZIP URL column pointing to that candidate's own documents), or `unknown`. No rule-based/fuzzy mapping — AI only. Supports `.xlsx` (openpyxl), `.xls` (xlrd), `.xlsb` (pyxlsb), `.ods` (odfpy), `.csv` (stdlib).

2. **Multiple Excel/CSV attachments** → `parse_excel_multi()`. One combined AI call maps columns across *all* files at once (still capped at ~10 sample rows per file) and may additionally flag a custom shared column as a join key if the standard fields alone don't connect the sheets. Rows describing the same real candidate across files are linked with a deterministic cascading exact match — `ars → check_id → emp_id → name → AI-detected custom key` — and identity fields (`ars`, `emp_id`, `name`) are backfilled across every matched row. A match resolved only by name (no shared ID) is flagged for extra scrutiny; genuine conflicts (two different ARS numbers for what looks like the same person) are logged rather than silently dropped, with "first non-blank wins."

3. **No Excel, but the body clearly stated candidate data** → that `body_candidate_data` list is used directly, same schema.

If none of the three exist and documents *were* uploaded, there's no structured list to match against — every file routes to `_UNASSIGNED` for manual handling rather than guessing.

Rows sharing the same ARS are grouped into one candidate with a `checks` list (Employment, Education, etc. — supports multi-check candidates and later check-type subfolders). Every raw column value from every contributing row/file is preserved and merged (`_merge_raw_fields`) — if the same header appears more than once with genuinely different values, both are kept, tagged by source, rather than one silently overwriting the other.

---

## Step 3b — Per-Candidate Document Links (Excel "Link" column)

If a candidate's row has a `document_link`, it's downloaded **immediately and independently** of the general grouping pool:

1. `handle_links()` (same function as Step 2) downloads it.
2. `extract_all()` unpacks it if it's an archive.
3. The resulting files go into `extracted/<run_id>_candlinks/<n>/` — a directory kept **outside** the main `extracted/<run_id>/` tree specifically so these files are never picked up by the general AI-grouping file scan in Step 4.
4. These files are handed to the assembler tagged with that candidate's `folder_key` for **direct, high-confidence assignment** — no AI grouping needed, since the Excel row already deterministically identifies whose documents they are.
5. A failed/restricted link becomes a `pending_entries` row — but unlike a body-level link, this one *is* attributed to the specific candidate (folder_key, name, emp_id populated).

---

## Step 4 — AI Grouping (`pipeline/grouping_agent.py`) — conditional

Skipped entirely when there are no leftover files to group (e.g. an Excel-only or body-only email with nothing to unpack, or every file was already claimed by Step 3b). Otherwise:

- One Gemini call per batch of ≤500 files (chunked and merged if there are more), given: the full extracted file-path list, the candidate list (`folder_key`, name, emp_id, checks), and the combined email body context (outer email + any `.eml`/`.msg` body text found during extraction).
- The AI groups every file to a candidate using ARS/emp_id/name found in folder names, filenames, or body text; check-type folder names as signals for multi-check candidates; and proximity within a named candidate folder. Every file must be accounted for — unmatched or ambiguous files are explicitly listed, not dropped.
- After the AI responds: hallucinated paths are stripped, every real file is confirmed covered (anything missed defaults to unassigned), and a **regex-based rule fallback** does one more pass — extracting ARS patterns and employee IDs directly from filenames — to rescue anything the AI missed, at confidence 96.
- Retries `[0, 5, 15, 30]`s per model, falls back from `gemini-2.5-flash` to `gemini-2.5-flash-lite` on exhaustion. If everything fails, all files go to `_UNASSIGNED` rather than blocking the run.

---

## Step 5 — Assembly (`pipeline/assembler.py`)

Builds `output/output_batch_<timestamp>_<run_id prefix>/` and `manifest.csv`. Routing:

| Source | Destination | Status |
|---|---|---|
| AI-grouped, confidence ≥ 95 | `ARS_CandidateName/` | Auto Matched |
| AI-grouped, confidence 80–94 | `_REVIEW/ARS_CandidateName/` | Needs Review |
| AI-grouped, confidence < 80 | `_UNASSIGNED/` | Unassigned — low confidence |
| Unassigned / Ambiguous / cross-candidate duplicate | `_UNASSIGNED/` | Unassigned / Ambiguous / duplicate |
| Step 3b candidate document link | `ARS_CandidateName/` | Auto Matched (confidence 100, bypasses grouping) |
| Every candidate, regardless of the above | `ARS_CandidateName/extracted_info.txt` (or `_REVIEW/...` if that's where their real files landed) | `info_txt_with_documents` (has real docs too) or `info_only_no_documents` (none) |
| Link that couldn't be resolved | manifest row only, no folder | `link_pending_manual_fetch` |

Key behaviors:
- **Every candidate always gets `extracted_info.txt`** with their full merged Excel/email metadata — whether or not they also have real documents — so that information is never lost just because a document happened to exist elsewhere.
- Whether a candidate has zero matched documents is decided **per candidate**, not for the whole batch — one unrelated stray file elsewhere in the batch no longer suppresses everyone else's info record.
- Files are always copied flat (`os.path.basename` only) — no nested subfolders except optional check-type subfolders (`Employment_CHK1/`) for multi-check candidates.
- Filename collisions within the same folder get an ARS-prefixed name first, then `_2`, `_3`, etc.

`manifest.csv` columns: `original_path, original_folder, filename, source_type, link_url, folder_key, candidate_name, emp_id, confidence, reason, status, output_path`. Every file, every link, and every candidate appears somewhere — nothing is silently dropped.

---

## Output

```
output_batch_20260706_120000_a1b2c3d4/
├── 6197-001269_Arun_Kumar_S/
│   ├── aadhaar.pdf                 ← from AI grouping or a ZIP
│   ├── offer_letter.pdf            ← from this candidate's own Excel document-link
│   └── extracted_info.txt          ← always present: merged Excel/email metadata
├── 6197-003311_Divya_R/
│   └── extracted_info.txt          ← no documents at all for this candidate
├── _REVIEW/
│   └── 6197-002204_Hemalatha_S/
│       ├── degree.pdf
│       └── extracted_info.txt
├── _UNASSIGNED/
│   └── screenshot.pdf
└── manifest.csv
```

---

## Supported Formats

**Attachments** (classified purely by extension in Step 0):
- Excel/CSV: `.xlsx`, `.xls`, `.xlsb`, `.csv`, `.ods`
- Archives: `.zip`, `.rar`* , `.7z`, `.tar`, `.tar.gz`/`.tgz`, `.tar.bz2`/`.tbz2`, `.tar.xz`/`.txz`
- Email containers: `.eml`, `.msg`
- Loose documents (copied as-is): PDF, PNG, JPG/JPEG, DOCX, DOC, TIFF/TIF, HEIC, PPTX, PPT, TXT, XML

*`.rar` requires the system `unrar` binary to be present — not just a pip install.

**Links** (email body text or an Excel "Link"-style column):
- Google Drive — single file or folder, "anyone with the link"
- SharePoint / OneDrive
- Direct ZIP links
- Google Docs/Sheets/Slides — detected but always flagged pending (not downloadable as one file)
- Anything else — attempted via plain HTTP GET, flagged pending on failure

**Not yet handled:**
- Legacy binary formats beyond what's listed above (e.g. `.wpd`, `.rtf` are not explicitly loose-file types, though unrecognized extensions still get copied flat rather than rejected)
- Password-protected archives (skipped and logged, not cracked)
- Drive/SharePoint links requiring tenant-specific or account-based authentication (flagged pending for manual CST fetch — no service-account integration yet)
