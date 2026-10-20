import type { TextRun } from "./geometry";
import type { Box } from "./types";

export interface QuoteMatch {
  page: number;
  box: Box;
}

const normalise = (s: string) => s.toLowerCase().replace(/\s+/g, " ").trim();

function union(boxes: Box[]): Box {
  const x1 = Math.min(...boxes.map((b) => b.x));
  const y1 = Math.min(...boxes.map((b) => b.y));
  const x2 = Math.max(...boxes.map((b) => b.x + b.width));
  const y2 = Math.max(...boxes.map((b) => b.y + b.height));
  return { x: x1, y: y1, width: x2 - x1, height: y2 - y1 };
}

/**
 * Finds where a quote appears in a document's text layer and returns the box
 * around it. Matching ignores case and whitespace differences only, so an AI
 * answer is accepted only if its quote really is in the document.
 *
 * `pages` holds each page's text runs (index 0 = page 1). `preferPage` is
 * searched first when the AI said which page the quote is on.
 */
export function locateQuote(pages: TextRun[][], quote: string, preferPage?: number): QuoteMatch | null {
  const target = normalise(quote);
  if (!target) return null;
  const order = pages.map((_, i) => i + 1);
  if (preferPage && preferPage >= 1 && preferPage <= pages.length) {
    order.splice(order.indexOf(preferPage), 1);
    order.unshift(preferPage);
  }
  for (const pageNo of order) {
    const runs = pages[pageNo - 1];
    // Concatenate the page's runs, remembering which run each character came from.
    let text = "";
    const owner: number[] = [];
    runs.forEach((run, i) => {
      const piece = normalise(run.text);
      if (!piece) return;
      if (text) {
        text += " ";
        owner.push(-1);
      }
      for (const _ of piece) owner.push(i);
      text += piece;
    });
    const at = text.indexOf(target);
    if (at === -1) continue;
    const hit = new Set(owner.slice(at, at + target.length).filter((i) => i >= 0));
    return { page: pageNo, box: union([...hit].map((i) => runs[i].box)) };
  }
  return null;
}

/** Digits and decimal point only, so "$12,480.00" and "12480" compare equal. */
function numeric(s: string): number | null {
  // Only plain amounts count; dates and IDs like "2026-09-14" are compared as text.
  if (!/^\s*\(?-?[$€£]?\s*\d[\d,]*(\.\d+)?\)?\s*$/.test(s)) return null;
  const negative = /^\s*\(.*\)\s*$/.test(s) || /^\s*-/.test(s);
  const digits = s.replace(/[^\d.]/g, "");
  if (!/^\d+(\.\d+)?$/.test(digits)) return null;
  return (negative ? -1 : 1) * Number(digits);
}

/**
 * Checks that an extracted value is actually supported by its quote: numbers
 * must equal a number in the quote, anything else must appear in it.
 */
export function valueMatchesQuote(value: string, quote: string): boolean {
  const v = value.trim();
  if (!v) return false;
  const asNumber = numeric(v);
  if (asNumber !== null) {
    const numbersInQuote = quote.match(/\(?-?[$€£]?\d[\d,]*(\.\d+)?\)?/g) ?? [];
    return numbersInQuote.some((n) => numeric(n) === asNumber);
  }
  return normalise(quote).includes(normalise(v));
}
