# Gemini API — Cost Reference for Insuff Auto Pipeline

> **Pricing source:** [Google AI Developer Pricing](https://ai.google.dev/gemini-api/docs/pricing) — verified July 2026.  
> All figures are estimates based on typical email shapes. Actual usage varies with body length, file count, and candidate count.

---

## Models in Use

| Role | Model | Triggered by |
|---|---|---|
| **Primary** | `gemini-2.5-flash` | Every run — email parser, Excel mapper, grouping agent |
| **Fallback** | `gemini-2.5-flash-lite` | Only on 503 / 429 (API overload). Rare — not included in normal estimates |

---

## Pricing Table

| | gemini-2.5-flash | gemini-2.5-flash-lite |
|---|---|---|
| **Input tokens** | $0.30 / 1M | $0.10 / 1M |
| **Output tokens** | $2.50 / 1M | $0.40 / 1M |
| **Thinking tokens** | $2.50 / 1M (same as output) | Minimal / none |
| **Default thinking mode** | Dynamic (-1) — model decides depth | Off |
| **Max thinking budget** | 24,576 tokens | — |

> **Thinking tokens explained:** Gemini 2.5 Flash reasons internally before replying. Those "reasoning steps" are invisible in the response but billed as output tokens. The model uses more thinking for harder tasks (grouping) and less for simpler ones (email classification). You can disable it with `thinking_budget=0` to cut output cost roughly in half — see Section 6.

---

## Step-by-Step Breakdown

Three pipeline steps call Gemini. The rest (extraction, assembly, manifest) run in Python at zero API cost.

### Step 0 — Email Parser `email_parser.py`

**Called:** Once per run, always.

| What goes in | Approximate tokens |
|---|---|
| System prompt | ~600 |
| Email subject | ~15 |
| Email body / thread | **200 – 2,500** (biggest variable) |
| Attachment filenames list | ~30 per file |
| Output schema template | ~150 |
| **Total input** | **~1,000 – 3,500** |

| What comes out | Approximate tokens |
|---|---|
| JSON: file classifications, links, body candidates | ~300 – 700 |
| Thinking (estimated, simple task) | ~200 – 600 |
| **Total output + thinking** | **~500 – 1,300** |

**Cost range:** $0.0004 – $0.0016 per email  
**What drives it:** Length of the pasted email thread. A one-liner email is ~$0.0004. A full forwarded thread with three replies is ~$0.0015.

---

### Step 3 — Excel Column Mapper `excel_reader.py`

**Called:** Once per Excel / CSV attachment. Skipped if no spreadsheet.

| What goes in | Approximate tokens |
|---|---|
| System prompt | ~500 |
| First 20 rows of spreadsheet (JSON) | ~400 – 1,200 |
| Output schema instruction | ~100 |
| **Total input** | **~1,000 – 1,800** |

| What comes out | Approximate tokens |
|---|---|
| JSON: header row index + column map | ~150 – 400 |
| Thinking (estimated, simple mapping) | ~150 – 500 |
| **Total output + thinking** | **~300 – 900** |

**Cost range:** $0.0004 – $0.0010 per Excel file  
**What drives it:** Number of columns (more columns = more header cells to map). Wide spreadsheets (15+ columns) push toward the upper end.

> **Note:** Only the first 20 rows are sent to the API regardless of how many candidate rows the sheet has. Streaming the rest is pure Python — no extra API calls.

---

### Step 4 — Grouping Agent `grouping_agent.py`

**Called:** Once per run when files are present. Chunked at 500 files per call — large batches trigger multiple calls.

| What goes in | Approximate tokens |
|---|---|
| System prompt | ~400 |
| Candidate list (from Excel) | ~50 per candidate |
| File paths list | **~20 per file path** |
| JSON schema template | ~250 |
| **Total input** | **~1,500 – 15,000+** |

| What comes out | Approximate tokens |
|---|---|
| JSON: groups + unassigned + ambiguous | **~50 – 80 per file assigned** |
| Thinking (complex task — estimated) | ~1,000 – 5,000 |
| **Total output + thinking** | **~1,500 – 30,000+** |

**This is the dominant cost step** — output scales linearly with file count.

| File count | Est. input | Est. output+thinking | Est. cost |
|---|---|---|---|
| 10 files | ~1,800 tokens | ~1,500 tokens | **~$0.0043** |
| 50 files | ~2,700 tokens | ~5,000 tokens | **~$0.013** |
| 100 files | ~3,700 tokens | ~10,000 tokens | **~$0.026** |
| 200 files | ~5,700 tokens | ~20,000 tokens | **~$0.052** |
| 500 files | ~11,500 tokens | ~45,000 tokens | **~$0.116** |
| 500 files × 2 chunks | doubled | doubled | **~$0.232** |

---

## Cost Per Email — Scenarios

| Scenario | Email Parse | Excel Map | Grouping | **Total** |
|---|---|---|---|---|
| **A — Body only** (candidates listed in email, no files) | $0.0012 | — | — | **~$0.001** |
| **B — Excel + small ZIP** (20 candidates, 30 files) | $0.0008 | $0.0007 | $0.010 | **~$0.012** |
| **C — Excel + medium ZIP** (30 candidates, 100 files) | $0.0010 | $0.0007 | $0.026 | **~$0.028** |
| **D — Excel + Drive folder** (50 candidates, 200 files) | $0.0015 | $0.0010 | $0.052 | **~$0.055** |
| **E — Excel + large Drive folder** (50 candidates, 500 files) | $0.0020 | $0.0010 | $0.116 | **~$0.119** |
| **F — Long thread + Excel + ZIP** (long body, 50 files) | $0.0015 | $0.0010 | $0.013 | **~$0.016** |

---

## Format Breakdown — Which Attachment Type Costs More

| Attachment Type | API Cost Impact | Why |
|---|---|---|
| **No files, body-only** | ⬤ Cheapest (~$0.001) | Grouping step is skipped entirely |
| **Excel / CSV alone** | ⬤⬤ Very low (~$0.001–0.002) | Only mapper call; no grouping if no doc files |
| **Single small ZIP** (< 30 files) | ⬤⬤ Low (~$0.010–0.015) | File count is small → small grouping call |
| **Multiple ZIPs** | ⬤⬤⬤ Moderate | All files are pooled before grouping — total file count matters, not ZIP count |
| **RAR / 7z** | Same as ZIP | Extracted identically; cost depends on file count inside |
| **Drive folder link** (public) | ⬤⬤⬤ Moderate–High | Depends on how many files are in the folder; folder structure adds path tokens |
| **Large Drive folder** (200–500+ files) | ⬤⬤⬤⬤ High (~$0.07–0.23) | Grouping chunks; each chunk = one full API call |
| **`.eml` / `.msg` containers** | +10–30% vs ZIP | Each email container extracts its own attachments, inflating the file count fed to grouping |
| **Long forwarded thread body** | ⬤⬤ Low–Moderate (+$0.001) | Only affects email parser input — adds ~500–1,000 tokens to step 0 |
| **Wide Excel (15+ columns)** | +$0.0003 vs narrow sheet | More header cells → slightly larger Excel mapper input |

**Key rule:** the total number of document files sent to the grouping agent is the single biggest cost driver. Everything else is small.

---

## Monthly Cost Projection

| Volume | Avg files/email | Monthly total |
|---|---|---|
| 100 emails / month (small batches) | 30 files | **~$1.20** |
| 100 emails / month (medium batches) | 100 files | **~$2.80** |
| 300 emails / month (medium batches) | 100 files | **~$8.40** |
| 500 emails / month (mixed) | 80 files | **~$11.50** |
| 1,000 emails / month (heavy) | 200 files | **~$55** |

At current AuthBridge CST volumes, monthly Gemini cost is expected to stay in the **₹100 – ₹1,000 range** (< $12 / month) unless volume scales significantly.

---

## Cost Optimisation Options

### Option 1 — Disable thinking (biggest saving)

The grouping agent benefits least from thinking (it's a structured matching task, not reasoning). Add `thinking_budget=0` to the grouping call:

```python
# in grouping_agent.py — _call_chunk()
config=types.GenerateContentConfig(
    system_instruction=_SYSTEM,
    thinking_config=types.ThinkingConfig(thinking_budget=0),   # disables thinking
)
```

Estimated saving: **30–50% off grouping cost** (the dominant step). Accuracy impact: minimal for filename/folder matching. The email parser and Excel mapper already use simple enough prompts that thinking overhead is low.

### Option 2 — Use Flash Lite for email parsing and Excel mapping

Flash Lite is 3–6× cheaper and these two steps are straightforward classification tasks, not complex reasoning:

| Step | Flash cost | Flash Lite cost | Saving |
|---|---|---|---|
| Email parser | ~$0.0012 | ~$0.0004 | ~67% |
| Excel mapper | ~$0.0008 | ~$0.0003 | ~63% |

Change `_MODELS` in `email_parser.py` and `excel_reader.py` to `['gemini-2.5-flash-lite', 'gemini-2.5-flash']` (swap primary/fallback). Keep Flash as primary for grouping — accuracy matters more there.

### Option 3 — Cap thinking budget instead of disabling

A budget of 1,024 thinking tokens gives the model room to reason on ambiguous filenames without running up cost:

```python
thinking_config=types.ThinkingConfig(thinking_budget=1024)
```

This sits between full dynamic mode and zero — a good middle ground for the grouping agent.

### Option 4 — Filter junk files before grouping

Files like `desktop.ini`, `Thumbs.db`, `__MACOSX/` entries are already skipped by `build_file_tree()`. Make sure client ZIPs don't contain large numbers of system/metadata files — each one adds tokens even if filtered.

---

## Quick Reference

```
1 API call cost = (input_tokens × $0.30/1M) + (output_tokens × $2.50/1M)
                  thinking tokens billed at output rate

Rough shortcuts:
  1,000 input tokens  ≈ $0.0003
  1,000 output tokens ≈ $0.0025
  1,000 thinking toks ≈ $0.0025

  Per file in grouping output ≈ $0.00015–0.00025 (output) + thinking overhead
```

---

*Sources: [ai.google.dev/gemini-api/docs/pricing](https://ai.google.dev/gemini-api/docs/pricing) · [pricepertoken.com/gemini-2.5-flash](https://pricepertoken.com/pricing-page/model/google-gemini-2.5-flash) · [pricepertoken.com/gemini-2.5-flash-lite](https://pricepertoken.com/pricing-page/model/google-gemini-2.5-flash-lite)*
