# AI Audit Testing Tool: Build Plan

Draft 1 · 2026-10-20 · for Ti

## 1. What we're building

An Excel add-in that lets an auditor:

1. **Load client evidence** (PDFs, scans, Excel files) into an engagement.
2. **Ask for information in plain language** ("find the invoice total, date and customer for each sample item") and get answers back in cells.
3. **See every answer tagged to its source**, DataSnipper-style: click a cell and the document opens at the exact page and highlighted region (or the exact source cell for Excel evidence).
4. **Use an agent that edits workpapers** (fills test columns, adds tickmarks, writes conclusions) only through proposals the auditor approves.
5. **Run a reviewer agent** that independently re-checks the AI's work against the source and flags anything that doesn't match.

The design rule behind everything: **no number reaches a workpaper without a verifiable pointer to its source, and no AI change is final without a human sign-off.**

## 2. Recommended architecture

```
┌──────────────── Excel (Windows / Mac / Web) ────────────────┐
│  Task pane add-in (Office.js, React + TypeScript)           │
│   • Prompt / chat panel        • Evidence list              │
│   • Tag inspector              • Edit-proposal review (diff)│
│  Document viewer (Office dialog window, PDF.js + overlays)  │
│  Workbook storage: hidden "_Evidence" sheet + custom XML    │
└──────────────────────────────┬──────────────────────────────┘
                               │ HTTPS (Entra ID / SSO token)
┌──────────────────────────────▼──────────────────────────────┐
│  Backend API (Python, FastAPI)                              │
│   • Ingestion pipeline   • Extraction service (LLM)         │
│   • Verification engine  • Agent orchestrator               │
│   • Audit log (append-only)                                 │
├─────────────────────────────────────────────────────────────┤
│  Postgres (+pgvector) · Object storage (encrypted docs)     │
│  OCR / layout service · LLM API (Claude)                    │
└─────────────────────────────────────────────────────────────┘
```

### Why an Office.js add-in

| Option | Pros | Cons |
|---|---|---|
| **Office.js task pane (recommended)** | Runs on Windows, Mac and Excel on the web; modern web stack; deploys centrally through Microsoft 365 admin; no installer | Sandboxed: no direct file-system access; task pane is narrow, so the document viewer lives in a dialog window; some APIs need newer Excel builds |
| VSTO / .NET COM add-in (what DataSnipper uses) | Deep Excel control, dockable custom panes, works offline, fast local PDF rendering | Windows desktop only; MSI installs that firm IT must approve; older stack |
| Standalone web app + Excel export | Fastest to build | Fails the core requirement: auditors want to stay in Excel |

Office.js gets us to a working product fastest and reaches every platform. If a large firm later demands offline or deeper Windows integration, the backend stays the same and only the client changes.

## 3. Document ingestion

Every document becomes a set of **anchored chunks**: text plus an exact location.

| Source | How | Anchor stored |
|---|---|---|
| Digital PDF | PyMuPDF: words with bounding boxes, tables via layout detection | `doc_id, page, bbox[x0,y0,x1,y1]` |
| Scanned PDF / image | OCR + layout (Azure Document Intelligence or AWS Textract; Tesseract as a cheap fallback) | same, plus OCR confidence |
| Excel / CSV evidence | openpyxl / pandas, keeping sheet names and addresses; formulas and values both kept | `doc_id, sheet, cell_range` |
| Word / email (later) | docx / msg parsers, converted to PDF for viewing | page + bbox of the rendered PDF |

Also on ingest:

- **SHA-256 hash** of each original file, so we can prove the evidence wasn't altered.
- Page images rendered for the viewer, so highlights line up exactly with what the auditor sees.
- Chunks indexed for search (keyword + vector embeddings), so prompts can find the right page across hundreds of documents.

## 4. Prompt-driven extraction

Flow for a request like "for each sample row, find the invoice and pull amount, date, customer":

1. **Plan:** turn the prompt into a structured spec (fields, types, which rows, which documents).
2. **Retrieve:** search anchored chunks for candidate pages per row (match on invoice number, vendor, amount).
3. **Extract:** LLM returns structured JSON. Every field must include the chunk ID and the exact quoted text it came from. Claude's document citations feature fits this well.
4. **Verify (deterministic, no AI):**
   - The quoted text must actually appear in that chunk (exact or tight fuzzy match), otherwise reject.
   - Numbers and dates are parsed from the quote and must equal the extracted value.
   - The quote's bounding box becomes the highlight.
5. **Write:** values go into cells as *proposed* (see section 6), each with a tag.

Anything that fails verification is returned as "not found / needs manual review", never guessed.

## 5. Source tagging and tracing (the DataSnipper layer)

**Tag record** (one per linked cell):

```
tag_id, workbook_cell (sheet!A1), doc_id, doc_hash, page, bbox | sheet!range,
quoted_text, extracted_value, method (AI | manual), status (proposed | verified | rejected | exception),
created_by, created_at, verified_by, verified_at, model_version, prompt_id
```

**Where it lives:** in the workbook itself (hidden `_Evidence` sheet plus a custom XML part) so the workpaper stays self-contained when emailed or archived, mirrored to the backend for the audit log.

**In the grid:** tagged cells get a visual marker (colored border or tickmark symbol by status). Selecting a cell shows the source snippet in the task pane, with an "Open source" button that opens the viewer at the page with the region highlighted.

**Manual tagging too:** the auditor drags a box on the document in the viewer and links it to the selected cell. This is core DataSnipper behavior, and it's the fallback whenever the AI misses.

**Refining:** the auditor can redraw the box, pick a different document, or edit the value; the tag keeps its history.

**Per-cell fixes (requested by Ti, 2026-10-20):**
- **Remove tag:** clears a cell's tag (recorded as a "removed" row in the history, so the trail survives) and asks whether to clear the value too.
- **Re-ask AI for this cell:** reruns extraction for just the selected cell against the same document, with an optional hint ("use the invoice total, not the subtotal"). The new answer goes through the same verification and replaces the tag as *proposed*.
- **Bulk verify:** sign off several selected cells at once.

## 6. Agent layer

### Editor agent

- Reads workpaper structure (headers, sample lists, test columns) and the auditor's instruction.
- Tools: read range, extract from evidence, propose cell write, propose tickmark, propose comment / conclusion text.
- **Never writes directly.** Every change is an *edit proposal*: a cell-level diff (old → new, with source tag) shown in the task pane. The auditor accepts all, accepts some, or rejects. Only then does Office.js write to the sheet.
- Keeps its own undo log so a batch can be reverted.

### Reviewer agent

- Runs after extraction or on demand, independently of the editor: separate prompt, ideally a different retrieval path, no access to the editor's reasoning.
- For every tag: re-reads the source region, re-extracts, compares; checks footing / cross-footing and recalculations with deterministic code, not the LLM.
- Outputs a review list: agreed, mismatch (with both values and both locations), missing support, low OCR confidence.
- Results show as review notes; the auditor clears each one.

## 7. Accuracy, audit trail and compliance

- **Grounding over generation:** every value must trace to a verified quote; math is done in code, not by the model.
- **Human sign-off:** AI outputs start as *proposed*; only an auditor action makes them *verified*. This matches how standards treat tools: the auditor stays responsible for the evidence (PCAOB AS 1105, AICPA SAS 142, ISA 500).
- **Audit trail:** append-only log of who did what and when, the prompt, model and version, source hash, and before/after values. Exportable alongside the workpaper (PCAOB AS 1215 / ISA 230 documentation, 7-year retention).
- **Reproducibility:** store prompts, model versions and retrieval results so a reviewer can see why the AI produced an answer.
- **Measuring accuracy:** build an internal test set of real-looking audit documents with known answers; track field-level precision/recall per release; block releases that regress.
- **Client data security:** encryption at rest and in transit, per-firm / per-engagement isolation, no model training on client data (use LLM API terms that guarantee this), SSO via Entra ID, data residency options, and a path to SOC 2. Firms will ask about all of this before a pilot.
- **Keeping client data away from AI (requested by Ti, 2026-10-20):** offer a per-engagement AI mode:
  - *Off:* nothing leaves the auditor's computer. Manual tagging, Find, matching and sign-off still work.
  - *Firm's own cloud:* AI calls go to Claude through the firm's own AWS Bedrock, Google Vertex AI or Microsoft Foundry account, so documents stay inside the firm's cloud and contracts.
  - *Direct API with zero data retention:* for firms that accept it, under an Anthropic zero-data-retention agreement (confirm current terms).
  - Optional redaction of names and identifiers before any AI call, and a per-call log of exactly what was sent.

## 8. Tech stack (proposed)

| Layer | Choice |
|---|---|
| Add-in | TypeScript, React, Office.js, Fluent UI, PDF.js for the viewer |
| Backend | Python 3.12, FastAPI, Celery or a job queue for ingestion |
| Data | Postgres + pgvector, S3 / Azure Blob for documents |
| OCR | Azure Document Intelligence (fits Microsoft-centric firms) |
| LLM | Claude via API, with citations and structured output |
| Auth | Microsoft Entra ID (single sign-on from Excel) |
| Hosting | Azure (closest to where audit firms already are) |

## 9. Phased MVP

**Phase 0: Feasibility spike**
- Office.js add-in skeleton that loads in Excel.
- Open a PDF in a dialog, draw a highlight box, click a cell to jump to it.
- Store and read a tag from the hidden sheet.
- *Exit:* the click-to-trace loop works end to end on one document.

**Phase 1: Core tagging MVP**
- Upload digital PDFs and Excel files to an engagement.
- Manual tagging (draw box, link to cell) and AI extraction from a single prompt, with verification.
- Tag statuses and grid markers; tag inspector.
- Per-cell fixes: remove tag, re-ask AI for one cell, bulk verify.
- Sample selection from a client listing (random or monetary unit, key items, recorded seed and memo), feeding straight into whole-sample testing. Selection is done in code, never by the AI, so it is unbiased and reproducible.
- *Exit:* an auditor can vouch a 25-item sample against digital invoices, with every value traced.

**Phase 2: Testing workflows**
- OCR for scans, table extraction, bulk processing across a sample.
- Templates for common tests: vouching / test of details, bank confirmation tie-out, payroll, fixed asset additions.
- Automatic compare-and-tickmark (agreed / exception).

**Phase 3: Agents**
- Editor agent with edit proposals and approve / reject UI.
- Reviewer agent with the review list.
- Full audit-trail export.

**Phase 4: Firm-ready**
- Centralized M365 deployment, SSO, admin console, retention controls.
- Security review, SOC 2 readiness, pilot with a real audit team.

## 10. Decisions needed from Ti

1. **Platforms:** Office.js (Windows, Mac and web), or Windows-only VSTO like DataSnipper? *Recommended: Office.js.*
2. **Hosting:** our cloud (Azure), or deployable inside a firm's own tenant? *Recommended: our cloud first, firm-tenant later.*
3. **First test to support:** ✅ Decided 2026-10-20: vouching revenue invoices (test of details).
4. **Repository:** create a GitHub repo for the code so building can start with Phase 0.
