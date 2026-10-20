# Excel add-in

The task pane lets you add client evidence (PDFs, and Excel or CSV files), select a cell, pick a value in the document, and link the two. Selecting a tagged cell later shows the quoted source text, and **Open source** reopens the document at that page with the region highlighted, or at the exact source cell for a spreadsheet.

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

1. Open the Evidence pane and click **Add documents**. Use [`samples/sample-invoice-10442.pdf`](../samples/sample-invoice-10442.pdf) from this repo.
2. Select a cell, click **Tag from document**, drag a box around `$12,480.00`, and click **Use this value**.
3. The cell gets `12480` with an amber border. Select it again and click **Open source** to jump back to the highlighted total.

## Fill a whole sample

1. Open [`samples/revenue-sample-4-items.xlsx`](../samples/revenue-sample-4-items.xlsx). It lists four sampled invoices with the recorded amounts.
2. Click **Add documents** and pick all four `samples/sample-invoice-*.pdf` files.
3. Under **Fill whole sample**, check **Match rows by** is *Invoice #* and click **Fill sample**. Each document is read once and matched to the row with the same invoice number (prefixes like `INV-` and leading zeros are ignored).
4. Review the summary, then click **Write results**:
   - empty cells are filled from the document (amber, proposed);
   - cells that already agree with the document are tagged to it without changing (amber, proposed);
   - cells that differ are left as they are and marked red as exceptions;
   - rows with no matching document, and documents not in the sample, are listed for you.

With the sample files you should see 10442 agree, an amount exception on 10443 (7,350.00 recorded, 7,530.00 on the invoice), 10444 filled, 10445 with no document, and invoice 10450 reported as not in the sample.

## Extract with AI (one row)

1. Open [`samples/revenue-testing-sample.xlsx`](../samples/revenue-testing-sample.xlsx) (headings: Invoice #, Date, Customer, Amount) and add the sample invoice in the Evidence pane.
2. Select A2:D2 and click **Extract from document**. The column headings are the fields the AI looks for.
3. Each answer is checked against the PDF's own text before it can be used: the AI's quote must appear in the document and the value must appear in the quote. Anything that fails is listed as needing manual review.
4. Click **Write proposed values**. Each cell gets an amber *proposed* tag marked AI, traceable with **Open source**.

The AI call runs in the Supabase Edge Function in `../supabase/functions/extract`, which holds the Anthropic API key as the `ANTHROPIC_API_KEY` secret. The add-in only knows the function's URL and the project's public anon key.

## Search a client spreadsheet

1. Click **Add documents** and pick [`samples/client-sales-ledger.xlsx`](../samples/client-sales-ledger.xlsx) (a client sales ledger and customer list) or [`samples/client-ar-aging.csv`](../samples/client-ar-aging.csv). Select it in the list.
2. Under **Search**, type text or an amount and click **Find**, e.g. `10443` or `15,320`. Every matching cell is listed with its column heading and row label.
3. Or type a question and click **Ask AI**, e.g. *What is the amount for invoice 10443?* or *Which invoices are over $10,000?* Each answer names its source cell and is checked against the file: if that cell doesn't hold the value, it is shown as unusable.
4. Select a cell in your workpaper and click **Use in …** on a result. The value is written with an amber *proposed* tag pointing at the source cell. **Show** and **Open source** open the client's sheet with that cell highlighted.

**Tag from document** works for spreadsheets too: click the source cell, then **Use this value**. Spreadsheets are read in the browser (formatted values, and formulas' saved results). Only the AI question sends the file's contents to the AI service, through the Supabase function in `../supabase/functions/ask-sheet`.

## Sign off

Every value the add-in writes starts out awaiting review (amber *proposed*, or red when the AI flagged an exception). Nothing counts as reviewed until an auditor signs it off.

1. Select a tagged cell and check it with **Open source**.
2. Enter your name or initials under **Reviewer** (remembered on this computer).
3. Click **Verify**, **Reject** or **Mark exception**. Reject and exception need a note, for example "Invoice shows 7,530.00; recorded 7,350.00".

The cell's border turns green (verified) or red (rejected, exception), and its value is never changed. Each sign-off is added as a new row on the hidden `_Evidence` sheet with the reviewer, time and note, so **History** on the tag shows every decision on that cell. The **Review** section counts signed-off cells and **Next to review** jumps to the next cell still waiting. Re-running **Fill sample** skips cells that are already signed off.

## Tests

```bash
npm test         # unit tests for selection geometry and tag storage
npm run build    # typecheck + production build
```

## Current limits

- Documents are stored in the browser's IndexedDB on this computer only. Phase 1 moves them to the backend, so a workbook opened elsewhere will say the document isn't available.
- Digital PDFs only. Scanned pages have no text layer until OCR lands in Phase 2.
- The reviewer name is typed in, not tied to a Microsoft account, so sign-offs are only as reliable as the name entered.
- **Fill whole sample** and **Fill one row** read PDFs only; spreadsheets are searched with **Search**.
- **Ask AI** sends up to about 400,000 characters of a spreadsheet; beyond that it only sees the first part (Find still searches everything).
