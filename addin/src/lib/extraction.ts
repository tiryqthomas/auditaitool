import type { TextRun } from "./geometry";
import { locateQuote, valueMatchesQuote } from "./locate";
import type { Box } from "./types";

/** What the AI returns for one requested field. */
export interface ExtractedField {
  field: string;
  value: string | null;
  quote: string | null; // exact text copied from the document that shows the value
  page: number | null;
}

/** Anything that can answer "find these fields in this PDF". */
export interface Extractor {
  extract(pdfBase64: string, fields: string[], instructions?: string): Promise<ExtractedField[]>;
}

export type CheckOutcome =
  | { status: "supported"; page: number; box: Box }
  | { status: "not-found" | "quote-not-in-document" | "value-not-in-quote" };

export type CheckedField = ExtractedField & CheckOutcome;

export const OUTCOME_LABEL: Record<CheckOutcome["status"], string> = {
  supported: "Found in document",
  "not-found": "AI could not find it",
  "quote-not-in-document": "AI quote isn't in the document",
  "value-not-in-quote": "Value doesn't match the quote",
};

/**
 * Checks each AI answer against the document's own text, without any AI:
 * the quote must appear in the PDF and the value must appear in the quote.
 * Only "supported" answers are offered for writing into the workbook.
 */
export function checkAgainstDocument(results: ExtractedField[], pages: TextRun[][]): CheckedField[] {
  return results.map((r) => {
    if (!r.value || !r.quote) return { ...r, status: "not-found" };
    if (!valueMatchesQuote(r.value, r.quote)) return { ...r, status: "value-not-in-quote" };
    const match = locateQuote(pages, r.quote, r.page ?? undefined);
    if (!match) return { ...r, status: "quote-not-in-document" };
    return { ...r, status: "supported", page: match.page, box: match.box };
  });
}
