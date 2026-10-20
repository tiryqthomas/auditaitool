# Audit AI Tool

An Excel add-in for audit teams. It reads client evidence (PDFs, Excel files), finds information from plain-language prompts, and tags every workpaper value to the exact place in the source it came from, so any number can be traced back in one click. An agent layer will propose workpaper edits for auditor approval and independently review the AI's work.

See [docs/build-plan.md](docs/build-plan.md) for the architecture and phased roadmap.

## Status

**Phase 1: AI-assisted invoice vouching.** Click-to-trace tags, AI extraction for one row or a whole sample (checked against the PDF text), and auditor sign-off with a review history. Documents are still kept in the browser.

## Layout

| Path | What |
|---|---|
| `addin/` | Office.js task pane add-in (React + TypeScript + Vite) |
| `docs/` | Build plan and design notes |
| `samples/` | Fake evidence for demos and tests (never real client data) |
