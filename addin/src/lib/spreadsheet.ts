import { columnLetter } from "./sheet";

// Client spreadsheets (.xlsx, .xlsm, .csv) used as evidence. They are read in
// the browser, so a tag points at an exact source cell rather than a region.

export interface SourceCell {
  address: string; // e.g. D14 (no sheet)
  value: string | number | boolean | null;
  text: string; // what the auditor sees
}

export interface SourceSheet {
  name: string;
  rows: (SourceCell | null)[][]; // rows[r][c], 0-based
}

export type SourceWorkbook = SourceSheet[];

const SPREADSHEET = /\.(xlsx|xlsm|csv)$/i;

export function isSpreadsheet(name: string): boolean {
  return SPREADSHEET.test(name);
}

export const EVIDENCE_ACCEPT = ".pdf,application/pdf,.xlsx,.xlsm,.csv";

/** Parses RFC 4180 CSV: quoted fields, doubled quotes, commas and newlines inside quotes. */
export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  const src = text.replace(/^﻿/, "");
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"' && src[i + 1] === '"') {
        field += '"';
        i++;
      } else if (ch === '"') quoted = false;
      else field += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === ",") {
      row.push(field);
      field = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && src[i + 1] === "\n") i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else field += ch;
  }
  if (field !== "" || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

/** Builds a sheet from CSV text, keeping each field as written and reading plain numbers as numbers. */
export function sheetFromCsv(name: string, text: string): SourceSheet {
  return {
    name,
    rows: parseCsv(text).map((r, ri) =>
      r.map((field, ci) => {
        if (field.trim() === "") return null;
        const t = field.trim();
        const value = /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : field;
        return { address: `${columnLetter(ci)}${ri + 1}`, value, text: field };
      }),
    ),
  };
}

const isoDate = (d: Date) => d.toISOString().slice(0, 10);

/** Turns an exceljs cell value (formula, rich text, date, hyperlink…) into a plain value. */
function plainValue(v: unknown): string | number | boolean | null {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return isoDate(v);
  if (typeof v === "object") {
    const o = v as Record<string, unknown>;
    if ("result" in o) return plainValue(o.result);
    if ("richText" in o) return (o.richText as { text: string }[]).map((t) => t.text).join("");
    if ("text" in o) return plainValue(o.text);
    if ("error" in o) return String(o.error);
    return null;
  }
  return v as string | number | boolean;
}

/** Shows a number the way it is formatted in the client's file, where the format is a common one. */
function formatNumber(n: number, numFmt?: string): string {
  if (!numFmt || numFmt === "General") return String(n);
  const decimals = /0\.(0+)/.exec(numFmt)?.[1].length ?? 0;
  const grouped = numFmt.includes(",");
  const body = Math.abs(n).toLocaleString("en-US", {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
    useGrouping: grouped,
  });
  const currency = /[$€£]/.exec(numFmt)?.[0] ?? "";
  if (n < 0) return numFmt.includes("(") ? `(${currency}${body})` : `-${currency}${body}`;
  return `${currency}${body}${numFmt.includes("%") ? "%" : ""}`;
}

/** Reads an evidence spreadsheet into memory. */
export async function loadWorkbook(bytes: ArrayBuffer, fileName: string): Promise<SourceWorkbook> {
  if (/\.csv$/i.test(fileName)) {
    return [sheetFromCsv(fileName.replace(/\.csv$/i, ""), new TextDecoder().decode(bytes))];
  }
  // Loaded on demand: the library is large and only needed for Excel evidence.
  const { default: ExcelJS } = await import("exceljs");
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  return wb.worksheets
    .filter((ws) => ws.state === "visible")
    .map((ws) => {
      const rows: (SourceCell | null)[][] = [];
      ws.eachRow({ includeEmpty: false }, (row, r) => {
        const cells: (SourceCell | null)[] = [];
        row.eachCell({ includeEmpty: false }, (cell, c) => {
          const value = plainValue(cell.value);
          if (value === null || value === "") return;
          const text = typeof value === "number" ? formatNumber(value, cell.numFmt) : String(value);
          cells[c - 1] = { address: `${columnLetter(c - 1)}${r}`, value, text };
        });
        rows[r - 1] = Array.from(cells, (x) => x ?? null);
      });
      return { name: ws.name, rows: Array.from(rows, (x) => x ?? []) };
    });
}

/** "D14" → { row: 13, col: 3 } (0-based). */
export function parseAddress(address: string): { row: number; col: number } | null {
  const m = /^([A-Z]+)(\d+)$/.exec(address.replace(/\$/g, "").toUpperCase());
  if (!m) return null;
  return { row: Number(m[2]) - 1, col: [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 };
}

export function getCell(wb: SourceWorkbook, sheetName: string, address: string): SourceCell | null {
  const sheet = wb.find((s) => s.name === sheetName);
  const at = parseAddress(address);
  if (!sheet || !at) return null;
  return sheet.rows[at.row]?.[at.col] ?? null;
}

/** Splits "Sales Ledger!D14" or "'Sales Ledger'!$D$14" into sheet and address. */
export function splitSourceCell(ref: string): [string, string] {
  const i = ref.lastIndexOf("!");
  return [ref.slice(0, i).replace(/^'(.*)'$/, "$1"), ref.slice(i + 1).replace(/\$/g, "").toUpperCase()];
}

/**
 * Describes a cell by its column heading and the first text in its row, e.g.
 * "Amount · INV-10443", so a hit or a tag can be read without opening the file.
 */
/**
 * The heading row is the first row made only of text with at least two cells,
 * which skips report titles above the table. -1 when there is none.
 */
export function findHeaderRow(sheet: SourceSheet, beforeRow = Infinity): number {
  return sheet.rows.findIndex((r, i) => {
    const cells = r.filter((c): c is SourceCell => c !== null);
    return i < beforeRow && cells.length >= 2 && cells.every((c) => typeof c.value === "string");
  });
}

export function cellContext(sheet: SourceSheet, rowIndex: number, colIndex: number): string {
  const headerRow = findHeaderRow(sheet, rowIndex);
  const h = headerRow >= 0 ? sheet.rows[headerRow][colIndex] : null;
  const header = h && typeof h.value === "string" ? h.text : "";
  const label = sheet.rows[rowIndex]?.find(
    (c, i) => c && i !== colIndex && typeof c.value === "string" && !/^\d{4}-\d{2}-\d{2}$/.test(c.text),
  );
  return [header, label?.text].filter(Boolean).join(" · ");
}

export interface SearchHit {
  sheet: string;
  address: string;
  text: string;
  context: string;
}

const MAX_HITS = 200;

export function asAmount(s: string): number | null {
  const t = s.trim();
  const negative = /^\(.*\)$/.test(t) || t.startsWith("-");
  const digits = t.replace(/[()$€£,\s-]/g, "");
  return /^\d+(\.\d+)?$/.test(digits) ? (negative ? -1 : 1) * Number(digits) : null;
}

/** Finds cells containing the text, or equal to the amount, in every sheet. */
export function searchWorkbook(wb: SourceWorkbook, query: string): SearchHit[] {
  const q = query.trim().toLowerCase();
  if (!q) return [];
  const amount = asAmount(q);
  const hits: SearchHit[] = [];
  for (const sheet of wb) {
    sheet.rows.forEach((row, r) =>
      row.forEach((cell, c) => {
        if (!cell || hits.length >= MAX_HITS) return;
        const numeric = typeof cell.value === "number" && amount !== null && Math.abs(cell.value - amount) < 0.005;
        if (numeric || cell.text.toLowerCase().includes(q)) {
          hits.push({ sheet: sheet.name, address: cell.address, text: cell.text, context: cellContext(sheet, r, c) });
        }
      }),
    );
  }
  return hits;
}

/**
 * Writes the workbook as text for the AI, one line per row with each cell's
 * address, so every answer can name the exact cell it came from.
 */
export function workbookAsText(wb: SourceWorkbook, maxChars = 400_000): { text: string; truncated: boolean } {
  const lines: string[] = [];
  let size = 0;
  for (const sheet of wb) {
    lines.push(`=== Sheet: ${sheet.name} ===`);
    for (const row of sheet.rows) {
      const cells = row.filter((c): c is SourceCell => c !== null);
      if (cells.length === 0) continue;
      const line = cells.map((c) => `${c.address}=${JSON.stringify(c.text)}`).join(" | ");
      size += line.length + 1;
      if (size > maxChars) return { text: lines.join("\n"), truncated: true };
      lines.push(line);
    }
  }
  return { text: lines.join("\n"), truncated: false };
}

export interface SheetAnswer {
  answer: string;
  sheet: string;
  cell: string;
  reason: string;
}

export type CheckedAnswer = SheetAnswer & { found: SourceCell | null; supported: boolean; context: string };

/** An AI answer counts only when the named cell exists and holds that value. */
export function checkSheetAnswers(wb: SourceWorkbook, answers: SheetAnswer[]): CheckedAnswer[] {
  return answers.map((a) => {
    const [sheetName, address] = a.cell.includes("!") ? splitSourceCell(a.cell) : [a.sheet, a.cell.replace(/\$/g, "").toUpperCase()];
    const found = getCell(wb, sheetName, address);
    const sheet = wb.find((s) => s.name === sheetName);
    let supported = false;
    let context = "";
    if (found && sheet) {
      const n = asAmount(a.answer);
      supported =
        typeof found.value === "number" && n !== null
          ? Math.abs(found.value - n) < 0.005
          : found.text.trim().toLowerCase() === a.answer.trim().toLowerCase();
      const at = parseAddress(address)!;
      context = cellContext(sheet, at.row, at.col);
    }
    return { ...a, sheet: sheetName, cell: address, found, supported, context };
  });
}
