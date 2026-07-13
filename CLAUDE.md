# CST Bulk Insufficiency — ARS Mapping

## What It Does

User pastes a bulk insuff email (subject, body/thread, attachments) into an input box that mirrors the real email — a stand-in for the future ERS integration. The system reads the email content, figures out what kind of inputs the client actually sent (Excel, ZIP, RAR, 7z, .eml/.msg, loose docs, Drive/SharePoint/ZIP links, or plain body text), unpacks and parses accordingly, uses AI to group every document to the right candidate, and produces one flat folder per candidate ready for the next pipeline step (Bridge lookup, Doc Classification, OCR).

Discovery is **recursive, not just top-level**: a link, an Excel file, or a forwarded `.eml`/`.msg` can turn up anywhere — a top-level attachment, nested inside a ZIP, or nested behind another link entirely (a Drive link whose ZIP contains a forwarded email whose body has another link, mentions another candidate, etc.). The pipeline keeps chasing that chain — bounded by `pipeline/discovery.py`'s `MAX_LINK_HOPS` (3) — until nothing new turns up, so nothing found anywhere in the email/attachment tree is silently missed. See "Recursive Discovery" under Pipeline Step 2 below.

**Out of scope (unchanged):** ERS integration itself, and any reading of document *content* (that's Doc Classification's job, downstream of this system). Live Drive/SharePoint download is **no longer out of scope** — see Rule 8, updated below.

---

## File Structure

```
mvp/
├── app.py                  # Flask server — GET / , POST /process (SSE streaming)
├── pipeline/
│   ├── email_parser.py     # AI reads subject+body+thread (+ any forwarded .eml/.msg content, read first)+attachment names+links
│   ├── extractor.py        # Recursive unpacking: ZIP, RAR, 7z, .eml, .msg
│   ├── link_handler.py     # Real download: Drive (file/folder), SharePoint/OneDrive, generic ZIP/HTTP — flags pending only on failure
│   ├── discovery.py        # Recursive discovery loop: chases links/nested .eml/.msg/Excel up to MAX_LINK_HOPS (3), re-classifying every nested email found
│   ├── excel_reader.py     # AI column mapping for identity fields + column-agnostic URL scan for document links + multi-file merge
│   ├── grouping_agent.py   # AI batch grouping + regex ID fallback (conditional — skipped if no files)
│   └── assembler.py        # Build output folders + manifest.csv — every candidate gets extracted_info.txt; direct link-file assignment
├── uploads/                # Raw pasted email + attachments — wiped on each run
├── extracted/              # Fully unpacked content — wiped on each run
├── output/                 # Final output batches — never overwritten
├── templates/
│   └── index.html          # Email-replica input UI (see below), vanilla JS, SSE streaming
└── .env                    # GEMINI_API_KEY
```

---

## Tech Stack

```
pip install flask openpyxl google-genai gdown requests python-dotenv py7zr rarfile extract-msg
```

- **Flask** — two routes: `GET /` and `POST /process` (SSE streaming)
- **openpyxl** — read-only streaming mode (`read_only=True`) for `.xlsx`. `.xls`/`.xlsb`/`.csv`/`.ods` handled via dedicated fallback libraries (`xlrd`, `pyxlsb`, `odfpy`) or stdlib `csv`.
- **google-genai** — `from google import genai` (NOT `google-generativeai`)
- **Model** — `gemini-2.5-flash` primary, `gemini-2.5-flash-lite` fallback on 503/429
- **gdown** — anonymous Google Drive file/folder downloads ("anyone with the link")
- **requests** — SharePoint/OneDrive and generic HTTP(S) downloads; also the underlying transport `gdown` uses
- **py7zr** — `.7z` extraction (pure Python, no system dependency)
- **rarfile** — `.rar` extraction. ⚠️ **Requires the system `unrar` binary to be installed separately** — this is an infra dependency, not just a pip install. Confirm it's available in the target environment before relying on this.
- **extract-msg** — `.msg` (Outlook native) parsing
- **email** (stdlib) — `.eml` parsing, no extra install needed
- **shutil / zipfile / os / concurrent.futures** — file operations + the bounded thread pool used for parallel link downloads (see Step 3b)

---

## Input — Email Replica UI

Since ERS integration isn't built yet, the input box replicates what a real bulk insuff email looks like. Sarath copies directly from the actual email into these fields:

| Field        | Type              | Notes                                                                                                                                                 |
| ------------ | ----------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Subject line | text input        |                                                                                                                                                       |
| Email body   | textarea          | Paste the full thread as-is, including any Drive/SharePoint/ZIP links inline as plain text — links are detected from this text, not a separate field |
| Attachments  | multi-file upload | Any format — Excel/CSV, ZIP/RAR/7z, .eml/.msg, PDF/PNG/JPG/JPEG, DOCX, TIFF, HEIC                                                                    |

No separate "link" field — links are expected to appear naturally inside the pasted body text, same as in a real email. **A candidate's own document link can also live inside an Excel cell** (any column — see Step 3b), independent of the body text link detection above.

---

## Pipeline

### 0 — Email Parsing (`email_parser.py`)

Before this runs, any **top-level `.eml`/`.msg` attachment** is pre-extracted (by file extension, no AI needed to know it's an email container) and its body text is read into memory. That text is fed into the same classification call below as extra context — a forwarded email's content usually explains what the rest of the attachments actually are, so it needs to inform classification from the start, not just show up later as grouping context. (A `.eml`/`.msg` nested *inside* a ZIP can't be pre-read this way — it's only discovered once the ZIP is unpacked in Step 1.)

One Gemini call. Reads: subject, body/thread text, the pre-extracted forwarded-email text (if any), and attachment **filenames and types only**. **Never opens attachment content** — that boundary is unchanged from Doc Classification's job.

Outputs structured JSON:

```json
{
  "excel_attachments": ["mapping.xlsx"],
  "archive_attachments": ["docs.zip"],
  "email_container_attachments": ["forwarded_candidate.eml"],
  "loose_file_attachments": ["aadhaar_arun.pdf"],
  "links": [{"url": "...", "type": "drive|sharepoint|zip_link|unknown"}],
  "body_candidate_data": [{"name": "...", "identifier": "...", "identifier_type": "uan|emp_id|ars|other"}] | null,
  "has_files_or_links": true
}
```

`body_candidate_data` is populated when the outer body OR the forwarded-email text clearly states candidate-level info. If Excel is *also* present, Excel is authoritative — `body_candidate_data` is retained only as supplementary context for the grouping agent, not the primary source.

`has_files_or_links = false` with non-null `body_candidate_data` is the **fully-resolved case**: nothing to unpack, nothing to group — forwarded straight to Assembly.

Retry once on JSON parse failure. Raise to UI on second failure.

### 1 — Container Unpacking (`extractor.py`)

For each attachment (except `.eml`/`.msg` already pre-extracted in Step 0, and Excel attachments, handled separately in Step 3):

- **Archive** (`.zip`, `.rar`, `.7z`, `.tar`/`.tgz`/`.tbz2`/`.txz`) → recursively unpack to `extracted/`, including nested archives, until none remain
- **Email container** (`.eml`, `.msg`) → parse and extract its own attachments + body text recursively (treated as another container, not "read for content")
- **Loose document** (PDF/PNG/JPG/JPEG/DOCX/TIFF/HEIC) → copied as-is into `extracted/`, no unpacking needed

Preserve folder structure exactly as extracted. Password-protected or corrupt archives (any type) are skipped and logged — the rest of the batch continues.

### 2 — Link Handling + Recursive Discovery (`link_handler.py` + `discovery.py`)

For each link detected in the email/forwarded-email body text: **a real download is attempted** (this has evolved beyond a stub — see Rule 8, updated below):

| Link type | Behavior |
|---|---|
| Drive file, "anyone with the link" | Downloaded via `gdown` |
| Drive folder, "anyone with the link" | Entire folder tree downloaded, structure preserved |
| Drive file/folder, restricted | Flagged `link_pending_manual_fetch` with a clear reason |
| Google Docs/Sheets/Slides | Always flagged pending — can't be downloaded as one file |
| SharePoint/OneDrive, anonymous access | Downloaded via HTTP GET with `?download=1` |
| SharePoint/OneDrive, tenant-restricted | Flagged pending |
| Generic ZIP link / unknown URL | Plain HTTP GET attempt |

Successful downloads are fed back through `extractor.py` into the same shared `extracted/` pool for AI grouping (Step 4).

**Recursive discovery (`discovery.py`).** This step doesn't stop at one download. A downloaded ZIP can contain a forwarded `.eml`/`.msg` whose own body has another link, mentions another candidate not in any Excel, or is itself just another ZIP — none of that is visible from the outer email alone. So every link is resolved through a loop, not a single pass:

- Each downloaded link is extracted (`extractor.py` already recurses through nested archives/emails on disk). Any new `email_body.txt` or Excel/CSV file that appears as a result is picked up immediately.
- Every not-yet-classified nested email body gets its **own real AI classification pass** — `email_parser.parse_email(subject='', body=<nested text>, ...)` — not just a text dump into the final grouping prompt. This is what reliably surfaces links/candidates buried in a forwarded chain instead of losing them in one giant concatenated blob.
- Links found that way are queued at `hop + 1`. Once a link's hop exceeds `MAX_LINK_HOPS` (3), it is **not** fetched further — it's recorded as a `link_pending_manual_fetch` entry with a "depth limit reached" reason instead of chasing forever. (A link found directly in the outer body, or directly in an Excel cell, starts at hop 1.)
- Excel/CSV files found this way are **not** parsed immediately — every one found anywhere in the whole discovery pass is collected and parsed together, once, in Step 3 (reuses `parse_excel_multi`'s existing cross-file matching instead of running it piecemeal).
- Candidate leads (`body_candidate_data`) found in any nested email are merged into the candidate list the same way Excel-derived candidates are (see Step 3) — Excel still wins on conflict (Rule 9); a genuinely new ARS/ID becomes its own auto-matched candidate.

This same recursive loop runs twice per batch: once **unowned** (Step 2, seeded from the outer email's links, writing into the shared `extracted/` pool for AI grouping) and once **per candidate** (Step 3b, seeded from each candidate's own Excel-cell links, writing into an isolated per-link directory so the result can be assigned directly). Failures/depth-limited links become `pending_entries` — manifest-only rows; unowned ones carry no candidate (a body-level link isn't inherently tied to one person), Step 3b ones carry the owning candidate's `folder_key`/name/emp_id.

### 3 — Candidate Data Extraction (`excel_reader.py` + `app.py`)

- **Every Excel/CSV found anywhere** — top-level attachment, nested inside a ZIP, or discovered via Step 2's recursive link-chasing — is combined into one list before parsing, not handled piecemeal as each is found.
- **One Excel/CSV total** → `parse_excel()`. AI maps each column to a standard identity field: `ars_number`, `check_id`, `check_type`, `candidate_name`, `emp_id`, or `unknown`. No aliases, no rule-based/fuzzy mapping for these — AI only. Retry once on failure. Streams remaining rows, stops at 5 consecutive empty rows. Dedup by ARS number.
- **Multiple Excel/CSV files** → `parse_excel_multi()`. One combined AI call maps columns across all files at once and may flag a custom shared column as a join key if the standard identity fields alone can't connect the sheets. Rows describing the same candidate across files are linked with a deterministic cascading exact match — `ars → check_id → emp_id → name → AI-detected custom key` — and identity fields are backfilled across every matched row. A match resolved only by name (no shared ID) is flagged for extra scrutiny; genuine conflicts between matched rows are logged (first non-blank wins) rather than silently dropped.
- **Candidate leads from body text always merge in** — the outer email/subject's `body_candidate_data` *and* every nested `.eml`/`.msg` discovered during Step 2/3b's recursive discovery — via `app.py`'s `_merge_body_candidates_into()`. Matched against existing candidates by identifier first; Excel stays authoritative on conflict (Rule 9): a match backfills `raw_fields` as supplementary context, never overwrites identity fields. No match (a genuinely new ARS/ID no Excel mentions) becomes its own new auto-matched candidate — this is what closes the gap where a candidate was only ever mentioned in a forwarded email buried inside a ZIP.
- **If nothing (no Excel, no body/email candidate data) exists** and files are present → no structured list to match against. Do not guess — all files route to `_UNASSIGNED` for manual CST handling.
- **No per-batch candidate cap.** Every candidate found is processed in the same run, regardless of volume. The `skipped_candidates` plumbing (`assembler.py`, `grouping_agent.py`) and the `skipped_capacity_limit` manifest/UI status remain in place as dead-but-harmless code paths (always passed an empty list) in case a cap needs to be reintroduced later.

Rows sharing the same ARS are grouped into one candidate with a `checks` list (Employment, Education, etc.) via `app.py`'s `_build_candidates_from_excel_rows()` — supports multi-check candidates and check-type output subfolders. This same function is reused additively for any Excel discovered *after* the candidate list already exists (e.g. one found deep inside a candidate's own document-link chain in Step 3b): a row whose ARS already has a candidate merges into it instead of creating a duplicate folder. Every raw column value from every contributing row/file is preserved; if the same header appears more than once with genuinely different values, both are kept, tagged by source, rather than one silently overwriting the other.

### 3b — Per-Candidate Document Links (`excel_reader.py` + `app.py` + `discovery.py`)

Detection is **column-name-agnostic and deterministic** — not an AI-mapped field. Confirmed in real production data: the *same* column header (e.g. "Link") can hold a genuine document URL for one check-type row and unrelated free text (a postal address, an Aadhaar-card excerpt) for another row of the *same candidate*. So every cell, in every row, regardless of header name, is checked against a strict URL parse (`scheme` is `http`/`https` **and** a real `netloc`) — anything that doesn't parse as an actual URL is never attempted as a download. A candidate's `document_links` is the de-duplicated union of every real link found across all of their rows (a candidate can have more than one — e.g. one per check type).

Each detected link runs through the **same recursive discovery loop as Step 2** (bounded thread pool, default 8 concurrent workers — a batch can have 1000+ links, and downloading them one at a time can run long enough for the browser's SSE connection to be dropped as idle; progress is streamed periodically, not per-link, so large batches don't flood the UI), isolated per top-level link so every file found at any hop — including a nested `.eml`'s own attachments, or a file two hops deep — still lands in the same isolated directory and gets assigned **directly** to that candidate's output folder at confidence 100, bypassing `grouping_agent.py` entirely. Each downloaded file's manifest row records the top-level link that started its chain. A new Excel or a new candidate lead surfacing deep inside one candidate's own link chain (rare, but "no matter where found" applies here too) is merged into the shared candidate list after all candidates' link pools finish downloading — not inside the worker threads — to keep the Gemini calls and the shared candidate list single-threaded.

### 4 — AI Grouping (`grouping_agent.py`) — conditional

Skipped entirely when there are no leftover files to group (Excel/body-only email, or every file already claimed by Step 3b). Otherwise unchanged: one Gemini batch call (chunked at 500 files) sends all remaining file paths from `extracted/` (excluding any Excel/CSV already consumed as a candidate-identity source in Step 3) + the candidate list + combined email body context (outer email + any inner `.eml`/`.msg` body text found during extraction, now supplementary raw text — the structured leads/links from that text were already extracted in Step 2/3b). Gemini groups every file to a candidate using folder names, filenames, IDs, partial names, proximity. Returns `{groups, unassigned, ambiguous}`.

After the LLM response:

1. `_validate_and_ensure_coverage` — strips hallucinated paths, ensures every real file is accounted for. Anything the LLM omitted goes to unassigned.
2. `_rule_based_fallback` — for files still unassigned, extract ARS patterns and employee IDs from filenames via regex, exact lookup against candidate list. Confidence 96 (auto folder).

Retry schedule per model: `[0, 5, 15, 30]` seconds. On exhaustion, try fallback model. If both fail, all files go to `_UNASSIGNED`.

### 5 — Assembly (`assembler.py`)

Routing by confidence:

| Confidence             | Destination                                           |
| ----------------------- | ------------------------------------------------------ |
| ≥ 95                   | `output_batch_TIMESTAMP/ARS_CandidateName/`           |
| 80–94                  | `output_batch_TIMESTAMP/_REVIEW/ARS_CandidateName/`   |
| < 80                    | `output_batch_TIMESTAMP/_UNASSIGNED/`                 |
| Unassigned / Ambiguous  | `output_batch_TIMESTAMP/_UNASSIGNED/`                 |
| Same path in 2+ groups  | `output_batch_TIMESTAMP/_UNASSIGNED/`                 |
| Step 3b document link   | `output_batch_TIMESTAMP/ARS_CandidateName/` — always Auto Matched, confidence 100, bypasses AI grouping |

**Every candidate always gets `extracted_info.txt`** (their full merged Excel/email metadata as JSON) — this is unconditional, not just for candidates with zero documents:

- **Has real documents too** (from AI grouping or a Step 3b link) → status `info_txt_with_documents`, written alongside those documents in whichever folder they landed in (`ARS_Name/` or `_REVIEW/ARS_Name/`).
- **No documents, but other files existed elsewhere in the batch** → status `info_only_no_documents`, reason notes AI found no match — this is decided **per candidate**, not for the whole batch, so one unrelated stray file elsewhere never suppresses this for every other candidate.
- **No documents anywhere in the batch** → status `info_only_no_documents`, reason notes no files were ever attached.
- **Skipped by the MVP cap** → status `skipped_capacity_limit`, no folder, no `.txt` — manifest row only (see Step 3).

Files are always copied **flat** (`os.path.basename(src)` only) except for optional check-type subfolders (`Employment_CHK1/`) under multi-check candidates. Filename collisions within the same ARS folder: keep both, suffix `_2`, `_3`, etc. A candidate with no name in the source data gets a folder named by identifier alone (no trailing separator).

**`manifest.csv` fields:** `original_path`, `original_folder`, `filename`, `source_type` (`attachment` / `link` / `body_text`), `link_url`, `folder_key`, `candidate_name`, `emp_id`, `confidence`, `reason`, `status`, `output_path`. For a candidate with multiple real document links, each downloaded file's `link_url` reflects the specific link that produced it, not one candidate-wide value.

---

## Output Structure

```
output_batch_20260701_101500/
├── 6197-001269_Arun_Kumar_S/          ← real documents (AI-grouped or Step 3b link) + always extracted_info.txt
│   ├── aadhaar.pdf
│   ├── offer_letter.pdf               ← from this candidate's own Excel document link
│   └── extracted_info.txt             ← status info_txt_with_documents
├── 6197-003311_Divya_R/               ← info-only, no documents anywhere
│   └── extracted_info.txt             ← status info_only_no_documents
├── _REVIEW/
│   └── 6197-002204_Hemalatha_S/       ← confidence 80–94, CST confirms
│       ├── degree.pdf
│       └── extracted_info.txt
├── _UNASSIGNED/                       ← no match, low confidence, or no candidate list at all
│   └── screenshot.pdf
└── manifest.csv                       ← includes pending-link and skipped-capacity-limit rows even though no file/folder exists for them
```

---

## Rules — Never Violate

1. **Never read file or link content.** Email parsing uses subject/body/attachment filenames/types/link URLs only. Grouping uses filenames and paths only. Document-link detection (Step 3b) checks only whether a cell's *text* parses as a URL — it never opens what the link points to before downloading. Opening documents to read what's inside is Doc Classification's job — strictly out of scope here. **Clarification (recursive discovery):** a nested `.eml`/`.msg` *body* found anywhere — however deep — is fair game to read and classify, same category as the outer email body always was; that's still email metadata/text, not a document. This does **not** extend to `.txt` files or any other generic document type — a `.txt` is copied as a plain attachment like a PDF, its content is never opened or scanned for links.
2. **Never rename files.** Only the output folder is named (`ARS_CandidateName`). Files keep their original names.
3. **Output folders are always flat**, except for optional check-type subfolders under multi-check candidates (`Employment_CHK1/`, etc.) when the client's own files/links are organized that way.
4. **Every file, link, candidate, and skipped-by-cap candidate appears in the manifest.** Nothing is silently dropped.
5. **Never overwrite prior output.** Always create a new timestamped batch folder.
6. **One file belongs to one candidate.** No file copied to two ARS folders.
7. **No rule-based or fuzzy mapping for Excel *identity* fields** (ARS/check_id/candidate_name/emp_id) — those are understood by AI only. Document-link detection is the one deliberate exception: it's a plain, deterministic URL-format check (not identity classification), specifically because relying on the AI to spot a column by name proved unreliable when the same column's meaning varies row-to-row in real client data.
8. **Link resolution is real, not stubbed, and recursive.** Links found in the email body/forwarded-email text (Step 2) and links found in any Excel cell regardless of column name (Step 3b) are actually downloaded — Drive (file/folder), SharePoint/OneDrive, and generic HTTP(S) — and so is any further link found inside what they lead to (a nested `.eml`'s body, another archive), up to `discovery.py`'s `MAX_LINK_HOPS` (3). A link is only left as `link_pending_manual_fetch` when the download itself fails (restricted access, requires login, etc.) or the hop cap is reached — never as a blanket policy, and never silently dropped either way (Rule 4).
9. **Excel is authoritative over body text** when both exist for the same candidate; body-extracted data is supplementary context only. This applies uniformly regardless of *where* the body text was found — the outer email, a top-level forwarded attachment, or a `.eml`/`.msg` discovered several hops deep in a link chain.
10. ~~This MVP caps candidate volume per batch~~ — **removed.** All candidates in a batch are now processed regardless of volume; no `MAX_CANDIDATES` limit applies.
11. **Discovery is recursive but bounded.** Link-chasing stops at `MAX_LINK_HOPS` (3 hops) and archive nesting stops at `extractor.py`'s `_MAX_DEPTH` (10) — both deliberate, named limits so "no matter where found" can't turn into an unbounded loop on a bad-case batch (see Open Items).

---

## Environment

```
GEMINI_API_KEY=your_key_here
```

If `GEMINI_API_KEY` is missing: email parsing, Excel parsing, and grouping all raise an error (AI required at every AI-dependent step). Document-link detection itself needs no API key (it's a deterministic URL check), but it only runs after Excel parsing has succeeded.

---

## Running

```
cd mvp
python app.py
# → http://localhost:5000
```

---

## Open Items / To Confirm

- **`unrar` binary availability** in the target deployment environment — `rarfile` won't work without it.
- **Legacy Excel formats** (`.xls`, `.xlsb`, `.ods`) — now handled via `xlrd`/`pyxlsb`/`odfpy` fallbacks; confirm real-world frequency justifies keeping all three as dependencies long-term.
- **ERS integration timeline** — this UI is a manual stand-in; once ERS API/webhook access is confirmed, `app.py`'s input route will need to change from manual paste to programmatic ingestion.
- **Large-batch performance** — the `MAX_CANDIDATES` cap has been removed, so a 1000+ candidate single batch now runs end-to-end uncapped. Per-candidate link downloads and AI grouping cost/latency at that volume are not yet optimized — watch for slow runs or SSE timeouts on very large batches.
- **Recursive discovery adds real AI/network call volume.** Every nested `.eml`/`.msg` found anywhere now gets its own `parse_email` call, and link-chasing can add several more downloads per candidate (up to `MAX_LINK_HOPS` = 3 hops each). Combined with the removed `MAX_CANDIDATES` cap, a bad-case batch (long forwarding chains, many candidate-linked ZIPs each containing another email) will run noticeably longer/costlier than a simple one. Not yet load-tested at scale — watch for this alongside the large-batch item above.
- **Account-specific Drive/SharePoint access** — restricted links still require manual CST fetch; service-account/Graph-API integration for authenticated access is future work (see inline comments in `link_handler.py`).
