# Excel add-in (Phase 0)

The task pane lets you add PDFs, select a cell, drag a box around a value in the document, and link the two. Selecting a tagged cell later shows the quoted source text, and **Open source** reopens the document at that page with the region highlighted.

Tags are stored in the workbook on a hidden `_Evidence` sheet, so they travel with the file. Tagged cells get a coloured border by status (amber = proposed, green = verified, red = rejected or exception).

## Run it

Requires Node 20+ and desktop Excel (Windows or Mac) or Excel on the web.

```bash
cd addin
npm install
npx office-addin-dev-certs install   # one time: trusts the localhost HTTPS certificate
npm run dev                          # serves https://localhost:3000
```

Then sideload `manifest.xml`:

- **Excel on the web:** Home > Add-ins > More Add-ins > Upload My Add-in, and pick `manifest.xml`.
- **Excel on Windows:** share a folder containing `manifest.xml`, add it as a trusted add-in catalog (File > Options > Trust Center > Trusted Add-in Catalogs), restart Excel, then Home > Add-ins > Shared Folder.
- **Excel on Mac:** copy `manifest.xml` to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`, restart Excel, then Home > Add-ins.

An **Evidence** button appears on the Home tab.

## Try the loop

1. Open the Evidence pane and click **Add PDF**. Use `../samples/sample-invoice-10442.pdf`.
2. Select a cell, click **Tag from document**, drag a box around `$12,480.00`, and click **Use this value**.
3. The cell gets `12480` with an amber border. Select it again and click **Open source** to jump back to the highlighted total.

## Tests

```bash
npm test         # unit tests for selection geometry and tag storage
npm run build    # typecheck + production build
```

## Phase 0 limits

- Documents are stored in the browser's IndexedDB on this computer only. Phase 1 moves them to the backend, so a workbook opened elsewhere will say the document isn't available.
- Digital PDFs only. Scanned pages have no text layer until OCR lands in Phase 2.
- Tags are always created as *proposed*; there is no verify / reject action yet.
