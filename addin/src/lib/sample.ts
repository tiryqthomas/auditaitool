import type { CheckedField } from "./extraction";

/** One cell of the testing table as Excel holds it. */
export interface CellState {
  address: string; // Sheet!C5
  value: unknown; // raw value (dates are serial numbers)
  text: string; // what the auditor sees
}

export interface SampleRow {
  rowIndex: number;
  key: string; // the matching column's displayed value, e.g. the invoice number
  cells: Record<string, CellState>; // by column heading
}

/** Extraction results for one document, already checked against its text. */
export interface DocResult {
  docId: string;
  docName: string;
  fields: CheckedField[];
}

type Supported = Extract<CheckedField, { status: "supported" }>;

export type CellOutcome =
  | { kind: "fill"; header: string; cell: CellState; field: Supported } // empty cell, document value goes in
  | { kind: "agree"; header: string; cell: CellState; field: Supported } // workpaper already matches the document
  | { kind: "exception"; header: string; cell: CellState; field: Supported } // workpaper differs from the document
  | { kind: "unsupported"; header: string; cell: CellState; field: CheckedField }; // AI answer failed the source check

export interface RowPlan {
  row: SampleRow;
  doc?: { docId: string; docName: string };
  problem?: "no-document" | "several-documents";
  outcomes: CellOutcome[];
}

export interface SamplePlan {
  rows: RowPlan[];
  unmatchedDocs: DocResult[];
}

/** Letters and digits only, lower case: "INV-010442" -> "inv010442". */
export function normaliseKey(s: string): string {
  return s.toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * Whether a sample row's key and a document's key refer to the same item.
 * Exact after normalising, or the same number once prefixes like "INV" and
 * leading zeros are dropped ("INV-010442" and "10442").
 */
export function keysMatch(a: string, b: string): boolean {
  const na = normaliseKey(a);
  const nb = normaliseKey(b);
  if (!na || !nb) return false;
  if (na === nb) return true;
  const digits = (s: string) => s.replace(/\D/g, "").replace(/^0+/, "");
  const da = digits(na);
  const db = digits(nb);
  return da.length >= 3 && da === db;
}

function asNumber(v: unknown): number | null {
  if (typeof v === "number") return v;
  if (typeof v !== "string") return null;
  const t = v.trim();
  if (!/^\(?-?[$€£]?\s*\d[\d,]*(\.\d+)?\)?$/.test(t)) return null;
  const negative = t.startsWith("(") || t.startsWith("-");
  return (negative ? -1 : 1) * Number(t.replace(/[^\d.]/g, ""));
}

/** Calendar date as YYYY-MM-DD, from text or an Excel date serial. */
function asDate(v: unknown): string | null {
  if (typeof v === "number" && v > 20000 && v < 80000) {
    return new Date(Math.round((v - 25569) * 86400000)).toISOString().slice(0, 10);
  }
  if (typeof v !== "string" || !/[a-z]|[-/]/i.test(v)) return null;
  const t = Date.parse(v);
  if (Number.isNaN(t)) return null;
  const d = new Date(t);
  // Date.parse reads ISO dates as UTC and others as local time; use the matching fields.
  const iso = /^\d{4}-\d{2}-\d{2}$/.test(v.trim());
  const y = iso ? d.getUTCFullYear() : d.getFullYear();
  const m = (iso ? d.getUTCMonth() : d.getMonth()) + 1;
  const day = iso ? d.getUTCDate() : d.getDate();
  return `${y}-${String(m).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

const normaliseText = (s: string) => s.toLowerCase().replace(/\s+/g, " ").trim();

/** Compares a workpaper cell with a document value: as amounts, dates, then text. */
export function valuesAgree(cell: CellState, docValue: string): boolean {
  const a = asNumber(cell.value);
  const b = asNumber(docValue);
  if (a !== null && b !== null && !(typeof cell.value === "number" && asDate(docValue))) {
    return Math.abs(a - b) < 0.005;
  }
  const da = asDate(cell.value) ?? asDate(cell.text);
  const db = asDate(docValue);
  if (da && db) return da === db;
  return normaliseText(cell.text) === normaliseText(docValue);
}

const isEmpty = (c: CellState) => c.value === "" || c.value === null || c.value === undefined;

/**
 * Matches each sample row to the document whose verified key field equals the
 * row's key, then decides per cell whether to fill it, confirm it, or flag an
 * exception. Never guesses: a row with no document or more than one is left
 * for the auditor.
 */
export function planSample(rows: SampleRow[], keyHeader: string, docs: DocResult[]): SamplePlan {
  const keyOf = (d: DocResult) => {
    const f = d.fields.find((x) => x.field === keyHeader);
    return f && f.status === "supported" ? f.value! : null;
  };
  const used = new Set<string>();

  const planned = rows.map((row): RowPlan => {
    const matches = docs.filter((d) => {
      const k = keyOf(d);
      return k !== null && keysMatch(row.key, k);
    });
    if (matches.length === 0) return { row, problem: "no-document", outcomes: [] };
    if (matches.length > 1) {
      matches.forEach((d) => used.add(d.docId));
      return { row, problem: "several-documents", outcomes: [] };
    }
    const doc = matches[0];
    used.add(doc.docId);
    const outcomes = Object.entries(row.cells).flatMap(([header, cell]): CellOutcome[] => {
      const field = doc.fields.find((f) => f.field === header);
      if (!field) return [];
      if (field.status !== "supported") return isEmpty(cell) ? [{ kind: "unsupported", header, cell, field }] : [];
      if (isEmpty(cell)) return [{ kind: "fill", header, cell, field }];
      // The key already matched (allowing for prefixes like "INV-"), so it agrees by definition.
      if (header === keyHeader) return [{ kind: "agree", header, cell, field }];
      return [{ kind: valuesAgree(cell, field.value!) ? "agree" : "exception", header, cell, field }];
    });
    return { row, doc: { docId: doc.docId, docName: doc.docName }, outcomes };
  });

  return { rows: planned, unmatchedDocs: docs.filter((d) => !used.has(d.docId)) };
}
