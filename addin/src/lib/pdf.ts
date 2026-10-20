import * as pdfjs from "pdfjs-dist";
import workerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import type { TextRun } from "./geometry";

pdfjs.GlobalWorkerOptions.workerSrc = workerUrl;

export type PdfDoc = pdfjs.PDFDocumentProxy;
export type PdfPage = pdfjs.PDFPageProxy;

export function loadPdf(bytes: ArrayBuffer): Promise<PdfDoc> {
  // pdf.js takes ownership of the buffer, so hand it a copy.
  return pdfjs.getDocument({ data: bytes.slice(0) }).promise;
}

/** Text runs on a page with boxes in page units at scale 1, top-left origin. */
export async function pageTextRuns(page: PdfPage): Promise<TextRun[]> {
  const viewport = page.getViewport({ scale: 1 });
  const content = await page.getTextContent();
  const runs: TextRun[] = [];
  for (const item of content.items) {
    if (!("str" in item) || !item.str.trim()) continue;
    const tx = pdfjs.Util.transform(viewport.transform, item.transform);
    const height = Math.hypot(tx[2], tx[3]);
    runs.push({ text: item.str, box: { x: tx[4], y: tx[5] - height, width: item.width, height } });
  }
  return runs;
}
