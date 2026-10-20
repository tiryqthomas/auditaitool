# Excel add-in

The task pane lets you add PDFs, select a cell, drag a box around a value in the document, and link the two. Selecting a tagged cell later shows the quoted source text, and **Open source** reopens the document at that page with the region highlighted.

Tags are stored in the workbook on a hidden `_Evidence` sheet, so they travel with the file. Tagged cells get a coloured border by status (amber = proposed, green = verified, red = rejected or exception).

## Install it (no setup)

The add-in is hosted at https://tiryqthomas.github.io/auditaitool/ and redeploys on every merge to `main`.

1. Download [`manifest.xml`](https://tiryqthomas.github.io/auditaitool/manifest.xml) (right-click, Save link as).
2. Load it into Excel:
   - **Excel on the web** (easiest): open a workbook, then Home > Add-ins > More Add-ins > My Add-ins > Upload My Add-in, and pick `manifest.xml`.
   - **Excel desktop (Microsoft 365):** Home > Add-ins > More Add-ins > My Add-ins > Upload My Add-in. If that option isn't there, use Excel on the web or the developer steps below.

An **Evidence** button appears on the Home tab.

## Run it locally (developers)

Requires Node 20+.

```bash
cd addin
npm install
npx office-addin-dev-certs install   # one time: trusts the localhost HTTPS certificate
npm run dev                          # serves https://localhost:3000
```

Then sideload `manifest.local.xml` instead of `manifest.xml`:

- **Excel on the web:** as above, picking `manifest.local.xml`.
- **Excel on Windows:** share a folder containing `manifest.local.xml`, add it as a trusted add-in catalog (File > Options > Trust Center > Trusted Add-in Catalogs), restart Excel, then Home > Add-ins > Shared Folder.
- **Excel on Mac:** copy `manifest.local.xml` to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`, restart Excel, then Home > Add-ins.

## Try the loop

1. Open the Evidence pane and click **Add PDF**. Use [`samples/sample-invoice-10442.pdf`](../samples/sample-invoice-10442.pdf) from this repo.
2. Select a cell, click **Tag from document**, drag a box around `$12,480.00`, and click **Use this value**.
3. The cell gets `12480` with an amber border. Select it again and click **Open source** to jump back to the highlighted total.

## Extract with AI (Phase 1)

1. Open [`samples/revenue-testing-sample.xlsx`](../samples/revenue-testing-sample.xlsx) (headings: Invoice #, Date, Customer, Amount) and add the sample invoice in the Evidence pane.
2. Select A2:D2 and click **Extract from document**. The column headings are the fields the AI looks for.
3. Each answer is checked against the PDF's own text before it can be used: the AI's quote must appear in the document and the value must appear in the quote. Anything that fails is listed as needing manual review.
4. Click **Write proposed values**. Each cell gets an amber *proposed* tag marked AI, traceable with **Open source**.

The AI call runs in the Supabase Edge Function in `../supabase/functions/extract`, which holds the Anthropic API key as the `ANTHROPIC_API_KEY` secret. The add-in only knows the function's URL and the project's public anon key.

## Tests

```bash
npm test         # unit tests for selection geometry and tag storage
npm run build    # typecheck + production build
```

## Current limits

- Documents are stored in the browser's IndexedDB on this computer only. Phase 1 moves them to the backend, so a workbook opened elsewhere will say the document isn't available.
- Digital PDFs only. Scanned pages have no text layer until OCR lands in Phase 2.
- Tags are always created as *proposed*; there is no verify / reject action yet.
- AI extraction reads one PDF for one row at a time.
