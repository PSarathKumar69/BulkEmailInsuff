# Pre-Production Test Checklist — CST Bulk Insuff Auto

> Run every test case below before going live. Mark each row ✅ Pass / ❌ Fail / ⚠ Partial.
> For every test: check the UI result cards, open the output batch folder, and open `manifest.csv`.

---

## How to Read Each Test Case

| Field            | Meaning                                         |
| ---------------- | ----------------------------------------------- |
| **Input**  | What to paste / upload                          |
| **Expect** | What you should see in the UI and output folder |
| **Verify** | Specific checks to perform manually             |

---

## Section 1 — Excel Format Coverage

### TC-01 · Excel `.xlsx` + ZIP of loose docs

**Input**

- Subject: `Insuff - Batch Jan 2025`
- Body: *(leave empty)*
- Attachments: `mapping.xlsx` (2–3 candidates with ARS numbers) + `docs.zip` containing PDFs named generically (e.g. `scan001.pdf`, `photo.jpg`)

**Expect**

- AI reads Excel, maps ARS column correctly
- Files from ZIP grouped to correct ARS folders
- `manifest.csv` has `source_type = attachment` for all rows

**Verify**

- [X] No `_REVIEW` or `_UNASSIGNED` for clearly named files
- [X] ARS folders created at top level (confidence ≥ 95)
- [X] `manifest.csv` columns are populated (no blanks in `folder_key`, `candidate_name`)

---

### TC-02 · Excel `.csv` + loose PDF attachments

**Input**

- Subject: `Insuff resubmission`
- Body: *(leave empty)*
- Attachments: `candidates.csv` (comma-separated, with header row) + 2–3 PDFs named after the candidates

**Expect**

- CSV parsed correctly, header row identified by AI
- PDFs grouped by name signal from filenames

**Verify**

- [ ] Correct ARS / name used as folder
- [ ] No duplicate rows in manifest

---

### TC-03 · Excel `.ods` (OpenDocument format)

**Input**

- Subject: `ODS insuff test`
- Body: *(leave empty)*
- Attachments: `mapping.ods` (create in LibreOffice — save as .ods; include ARS, Name, UAN columns) + one PDF

**Expect**

- ODS parsed without error
- Candidate list extracted from ODS correctly

**Verify**

- [ ] Step log shows "Excel ready — N candidate(s)"
- [ ] No "ODS" or "parse" errors in the extraction errors panel

---

### TC-04 · Excel with multi-check rows (same ARS, multiple check types)

**Input**

- Subject: `Multi-check insuff`
- Body: *(leave empty)*
- Attachments: Excel where ARS `6197-001001` appears on 3 rows — Employment Check, Education Check, Reference Check — + a ZIP with files in sub-folders named after each check type

**Expect**

- One folder per ARS (not three)
- Check sub-folders created inside the ARS folder: `Employment_Check_CHK001/`, `Education_Check_CHK002/`

**Verify**

- [ ] ARS folder has check sub-folders, not flat
- [ ] `manifest.csv` `output_path` reflects sub-folder structure

---

### TC-05 · Excel with no ARS column (only Employee ID)

**Input**

- Subject: `EMP insuff batch`
- Body: *(leave empty)*
- Attachments: Excel where ARS column is missing; folder_key derived from Employee ID + loose PDFs named with EMP IDs

**Expect**

- Folder named `EMP10421_Arun_Kumar_S/` (emp_id as key)
- No crash — graceful fallback to emp_id as folder_key

**Verify**

- [ ] Step log shows "N candidate(s)" without errors
- [ ] Folder exists and files are inside it

---

## Section 2 — Archive Format Coverage

### TC-06 · `.zip` — flat (files at root)

**Input**

- Attachments: Excel with 2 candidates + `batch.zip` containing PDFs flat at ZIP root (no sub-folders)

**Expect**

- All PDFs extracted, grouped to candidates by filename/AI signal

**Verify**

- [ ] `extracted/` folder (during run) has files flat — confirmed by grouping step log count

---

### TC-07 · `.zip` — nested (folders inside ZIP)

**Input**

- Attachments: Excel + `batch.zip` where the ZIP has folders `Arun_Kumar/` and `Divya_R/` each containing PDFs

**Expect**

- Folder structure preserved in `extracted/` and used as grouping signal
- AI uses folder names to assign confidently (≥ 95)

**Verify**

- [ ] Both candidates get their own output folder
- [ ] Confidence 95+ for folder-named files

---

### TC-08 · `.zip` inside a `.zip` (nested archive)

**Input**

- Attachments: Excel + outer `batch.zip` which contains `inner.zip` which contains PDFs

**Expect**

- Both layers extracted recursively
- Final PDFs available for grouping

**Verify**

- [ ] Step log shows extraction with no errors
- [ ] Inner PDF appears in manifest — not dropped

---

### TC-09 · `.7z` archive

**Input**

- Attachments: Excel + `docs.7z` (create with 7-Zip; put 2–3 PDFs inside)

**Expect**

- 7z extracted without error (uses `py7zr`)

**Verify**

- [ ] No extraction error in UI error panel
- [ ] Files grouped correctly

---

### TC-10 · `.rar` archive *(requires `unrar` binary)*

**Input**

- Attachments: Excel + `docs.rar` (create with WinRAR)

**Expect**

- RAR extracted correctly if `unrar` binary is installed
- If not installed: extraction error logged, rest of batch continues

**Verify**

- [ ] If `unrar` present: files extracted and grouped
- [ ] If `unrar` absent: error chip shows "unrar" message; other files still processed
- [ ] Confirm `unrar` availability in the target deployment environment before prod

---

## Section 3 — Email Container Formats

### TC-11 · `.eml` attachment — ARS in body, documents as attachments inside

**Input**

- Subject: *(outer email)* `Fwd: Insuff for Hemalatha`
- Body: *(outer email body)* `Please find the forwarded insuff email`
- Attachments: `forwarded_hemalatha.eml`
  - Inner `.eml` body must contain: `ARS: 6197-003311 | Candidate: Hemalatha S`
  - Inner `.eml` attachments: `aadhaar.pdf`, `pan.pdf`

**Expect**

- Step log: "No candidate list found — scanning forwarded email body…"
- Step log: "Found 1 candidate(s) from forwarded email body"
- Folder `6197-003311_Hemalatha_S/` created with `aadhaar.pdf` and `pan.pdf`

**Verify**

- [ ] ARS folder created (not `_UNASSIGNED`)
- [ ] Both PDFs inside the ARS folder
- [ ] `manifest.csv` `source_type = attachment` for both files

---

### TC-12 · `.msg` attachment — Outlook forwarded email

**Input**

- Attachments: `forwarded.msg` (save a real Outlook email as .msg; it should have candidate ARS or name in its body and document attachments inside)

**Expect**

- Same behaviour as TC-11 — inner body scanned, ARS folder created

**Verify**

- [ ] `.msg` body text extracted to `email_body.txt` by extractor
- [ ] Candidate identified from inner body
- [ ] Documents grouped into ARS folder

---

### TC-13 · `.eml` inside a `.zip`

**Input**

- Attachments: Excel + `bundle.zip` containing `candidate_hemalatha.eml` (with ARS in body + PDF inside)

**Expect**

- ZIP extracted → `.eml` found → `.eml` parsed → PDF extracted
- PDF grouped to the ARS from the `.eml` body

**Verify**

- [ ] Multi-level extraction works without errors
- [ ] ARS folder created for the inner candidate

---

## Section 4 — Body-Only Candidate Data

### TC-14 · ARS and candidate data directly in email body — no attachments

**Input**

- Subject: `Insuff — 2 candidates, no docs`
- Body:
  ```
  Hi Team,

  Please find below candidate details for insuff:

  1. Candidate: Arun Kumar S
     ARS No: 6197-001269
     UAN: 100412356789
     Pending: Aadhaar, PAN

  2. Candidate: Divya R
     ARS No: 6197-003311
     Employee ID: EMP20452
     Pending: Offer Letter
  ```
- Attachments: *(none)*

**Expect**

- Both candidates resolved from body
- `info_only_no_documents` status in manifest
- Each ARS folder contains `extracted_info.txt` with original field names (not internal fields)

**Verify**

- [ ] `extracted_info.txt` for Arun has keys `Candidate`, `ARS No`, `UAN`, `Pending` — NOT `folder_key`, `ars`, `emp_id`
- [ ] No grouping agent called (step log skips grouping)
- [ ] Stat chip shows 0 Auto Matched, 2 Info Only

---

### TC-15 · ARS in body + loose file attachments (no Excel)

**Input**

- Subject: `Insuff — Hemalatha`
- Body:
  ```
  Please find documents for below candidate:
  ARS: 6197-003311
  Candidate Name: Hemalatha S
  ```
- Attachments: `degree.pdf`, `experience_letter.pdf` (loosely named — no ARS in filename)

**Expect**

- Body parsed → 1 candidate from body
- Grouping agent called with body context
- Both PDFs grouped to `6197-003311_Hemalatha_S/`

**Verify**

- [ ] Step log: "Using 1 candidate(s) from email body"
- [ ] Both files in ARS folder (confidence ≥ 80)
- [ ] `_UNASSIGNED` is empty

---

## Section 5 — Link Handling

### TC-16 · Google Drive link in body (publicly accessible)

**Input**

- Subject: `Insuff — Drive link batch`
- Body:
  ```
  Hi,
  Please find candidate documents here:
  https://drive.google.com/file/d/REAL_FILE_ID/view?usp=sharing

  ARS: 6197-001001 | Candidate: Priya M
  ```
- Attachments: *(none)*

**Expect**

- Link detected and download attempted via `gdown`
- If download succeeds: file extracted and grouped to ARS folder
- If download fails: manifest row with `link_pending_manual_fetch`, Pending Links chip shows 1

**Verify**

- [ ] No crash — link error is caught and shown in UI
- [ ] If downloaded: file appears in ARS folder
- [ ] If failed: `link_url` populated in manifest row
- [ ] Pending Links chip clickable → shows the link row in UI

---

### TC-17 · SharePoint link in body (stub — expect pending)

**Input**

- Subject: `SharePoint docs`
- Body: `Documents: https://company.sharepoint.com/sites/hr/Documents/insuff_batch.zip`
- Attachments: *(none or Excel)*

**Expect**

- Link detected as `sharepoint` type
- No download attempted (stub)
- Manifest row with `status = link_pending_manual_fetch`

**Verify**

- [ ] Pending Links chip count = 1
- [ ] `link_url` in manifest = the SharePoint URL
- [ ] No error thrown

---

## Section 6 — Edge Cases and Robustness

### TC-18 · Password-protected ZIP

**Input**

- Attachments: Excel + `protected.zip` (set a password in WinRAR or 7-Zip)

**Expect**

- Extraction fails gracefully for the protected ZIP
- Error appears in extraction errors panel (red chip)
- Other attachments still processed normally

**Verify**

- [ ] Extraction Errors chip count ≥ 1
- [ ] Error message mentions "password" or "extraction failed"
- [ ] Non-protected files still grouped

---

### TC-19 · Filename collision — two different candidates send a file named `aadhaar.pdf`

**Input**

- Attachments: Excel with 2 candidates (ARS-A and ARS-B) + `batch.zip` containing:
  - `ARS-A/aadhaar.pdf`
  - `ARS-B/aadhaar.pdf`

**Expect**

- Each ARS folder gets its own `aadhaar.pdf`
- No file overwrites or silent drops

**Verify**

- [ ] Both ARS folders contain `aadhaar.pdf`
- [ ] Manifest has 2 separate rows for the two files, different `output_path`

---

### TC-20 · Duplicate files across candidates (same file listed in two groups)

**Input**

- Manufacture a case where the grouping AI might assign the same file path to two candidates
  (hard to force; instead: verify the cross-duplicate detection logic by checking a batch where filenames are ambiguous)

**Expect**

- Cross-candidate duplicate file goes to `_UNASSIGNED/`
- Manifest row has `status = Unassigned — cross-candidate duplicate`

**Verify**

- [ ] File not copied into any ARS folder
- [ ] Only one copy exists in `_UNASSIGNED/`

---

### TC-21 · No subject, no body — only Excel + ZIP

**Input**

- Subject: *(empty)*
- Body: *(empty)*
- Attachments: `mapping.xlsx` + `docs.zip`

**Expect**

- Pipeline runs normally — subject and body optional
- Candidates from Excel, files from ZIP

**Verify**

- [ ] No error on empty subject/body
- [ ] Results same as TC-01

---

### TC-22 · Completely empty submit (no subject, no body, no attachments)

**Input**

- Subject: *(empty)*, Body: *(empty)*, Attachments: *(none)*

**Expect**

- UI shows validation error immediately: "Provide at least a subject, body, or one attachment."
- No API call made

**Verify**

- [ ] Error toast / message visible
- [ ] No output batch folder created

---

### TC-23 · Very large ZIP (100+ files, multiple candidates)

**Input**

- Attachments: Excel with 10 candidates + a ZIP containing 100+ files

**Expect**

- Pipeline completes without timeout
- Grouping agent handles batch chunking internally
- All candidates get folders; unmatched files go to `_UNASSIGNED`

**Verify**

- [ ] Step log shows correct file count
- [ ] `manifest.csv` has all 100+ rows
- [ ] No memory / timeout error in Flask console

---

## Section 7 — UI and Output Verification

### TC-24 · Stat chip filters work correctly

After any successful run:

**Verify**

- [ ] **Auto Matched chip** (green) — click → only auto-matched candidate cards shown; click again → all shown
- [ ] **Needs Review chip** (amber) — click → only review cards shown
- [ ] **Unassigned chip** (red) — click → only unassigned rows shown
- [ ] **Pending Links chip** (blue) — click → only link-pending rows shown
- [ ] **Extraction Errors chip** (gray) — click → errors panel expands
- [ ] Counts on each chip match the actual rows shown

---

### TC-25 · `manifest.csv` completeness

After any run with a mix of assigned, unassigned, pending, and info-only rows:

**Verify**

- [ ] Every attachment has at least one row in manifest
- [ ] Every link has a manifest row (`source_type = link`, `link_url` populated)
- [ ] Info-only candidates have a row (`status = info_only_no_documents`, `filename = extracted_info.txt`)
- [ ] No row has both `output_path` empty AND `status = Auto Matched`
- [ ] `folder_key` never contains internal pipeline names like "folder_key_type"

---

### TC-26 · `extracted_info.txt` content — only original client fields

After TC-14 or TC-15 (body-only candidates):

**Verify**

- [ ] Open `extracted_info.txt` — keys are the EXACT labels from the email body (e.g. `"ARS No"`, `"Candidate"`, `"UAN"`, `"Pending"`)
- [ ] Keys do NOT include: `folder_key`, `folder_key_type`, `ars`, `check_id`, `emp_id`
- [ ] JSON is valid and readable

---

### TC-27 · Output folder never overwritten

Run the pipeline twice on the same input.

**Verify**

- [ ] Two separate `output_batch_TIMESTAMP_RUNID/` folders exist
- [ ] First run's output is intact and unchanged

---

## Section 8 — Known Infra Dependencies to Confirm Before Prod

| Item                                 | Status | Notes                                                                               |
| ------------------------------------ | ------ | ----------------------------------------------------------------------------------- |
| `unrar` binary installed on server | ☐     | Required for`.rar` extraction. `rarfile` pip package alone is not enough.       |
| `GEMINI_API_KEY` set in `.env`   | ☐     | All AI steps fail without it.                                                       |
| `MAX_UPLOAD_MB` set correctly      | ☐     | Default 200 MB. Increase if batches are larger.                                     |
| `output/` folder writable          | ☐     | Flask process needs write permission.                                               |
| Python packages installed            | ☐     | `flask openpyxl google-genai gdown python-dotenv py7zr rarfile extract-msg odfpy` |
| Port 5000 open / proxied             | ☐     | Or change`PORT` env var.                                                          |

---

## Quick Sanity Sequence (run this first on a new machine)

1. TC-01 — Excel + ZIP (baseline happy path)
2. TC-14 — Body-only candidates, no attachments
3. TC-11 — `.eml` with ARS in body
4. TC-22 — Empty submit (UI validation)
5. TC-18 — Password-protected ZIP (error handling)
6. TC-24 — All 5 filter chips

If these 6 pass cleanly, run the full checklist above.
