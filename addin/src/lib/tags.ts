import type { EvidenceTag } from "./types";

// Tags live inside the workbook on a hidden sheet, so a workpaper stays
// self-contained when it is emailed or archived.
export const EVIDENCE_SHEET = "_Evidence";

const COLUMNS = [
  "tagId", "cell", "docId", "docName", "page", "x", "y", "width", "height",
  "quotedText", "method", "status", "createdAt",
] as const;

const STATUS_COLOR: Record<EvidenceTag["status"], string> = {
  proposed: "#C27C0E",
  verified: "#107C10",
  rejected: "#A4262C",
  exception: "#A4262C",
};

export function tagToRow(t: EvidenceTag): (string | number)[] {
  return [
    t.tagId, t.cell, t.docId, t.docName, t.page, t.box.x, t.box.y, t.box.width, t.box.height,
    t.quotedText, t.method, t.status, t.createdAt,
  ];
}

export function rowToTag(row: unknown[]): EvidenceTag {
  const [tagId, cell, docId, docName, page, x, y, width, height, quotedText, method, status, createdAt] = row;
  return {
    tagId: String(tagId),
    cell: String(cell),
    docId: String(docId),
    docName: String(docName),
    page: Number(page),
    box: { x: Number(x), y: Number(y), width: Number(width), height: Number(height) },
    quotedText: String(quotedText),
    method: method as EvidenceTag["method"],
    status: status as EvidenceTag["status"],
    createdAt: String(createdAt),
  };
}

/** Normalises an address so "Sheet1!$B$4" and "Sheet1!B4" match. */
export function normaliseAddress(address: string): string {
  return address.replace(/\$/g, "").replace(/^'(.*)'!/, "$1!");
}

async function evidenceSheet(context: Excel.RequestContext): Promise<Excel.Worksheet> {
  const existing = context.workbook.worksheets.getItemOrNullObject(EVIDENCE_SHEET);
  await context.sync();
  if (!existing.isNullObject) return existing;
  const sheet = context.workbook.worksheets.add(EVIDENCE_SHEET);
  sheet.getRangeByIndexes(0, 0, 1, COLUMNS.length).values = [[...COLUMNS]];
  sheet.visibility = Excel.SheetVisibility.hidden;
  return sheet;
}

export async function readTags(): Promise<EvidenceTag[]> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItemOrNullObject(EVIDENCE_SHEET);
    await context.sync();
    if (sheet.isNullObject) return [];
    const used = sheet.getUsedRange();
    used.load("values");
    await context.sync();
    return used.values.slice(1).filter((r) => r[0] !== "").map(rowToTag);
  });
}

export async function findTag(cell: string): Promise<EvidenceTag | undefined> {
  const target = normaliseAddress(cell);
  const tags = await readTags();
  // Latest tag wins if a cell has been re-tagged.
  return tags.reverse().find((t) => normaliseAddress(t.cell) === target);
}

/** Writes the value into the cell, records the tag and marks the cell by status. */
export async function addTag(tag: EvidenceTag, value: string): Promise<void> {
  await Excel.run(async (context) => {
    const sheet = await evidenceSheet(context);
    const used = sheet.getUsedRange();
    used.load("rowCount");
    await context.sync();
    sheet.getRangeByIndexes(used.rowCount, 0, 1, COLUMNS.length).values = [tagToRow(tag)];

    const [sheetName, address] = splitAddress(tag.cell);
    const target = context.workbook.worksheets.getItem(sheetName).getRange(address);
    target.values = [[toCellValue(value)]];
    for (const edge of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"] as const) {
      const border = target.format.borders.getItem(edge);
      border.style = "Continuous";
      border.weight = "Medium";
      border.color = STATUS_COLOR[tag.status];
    }
    await context.sync();
  });
}

export function splitAddress(cell: string): [string, string] {
  const normal = normaliseAddress(cell);
  const i = normal.lastIndexOf("!");
  return [normal.slice(0, i), normal.slice(i + 1)];
}

/** Turns extracted text like "$1,234.50" or "(500)" into a number when it is one. */
export function toCellValue(text: string): string | number {
  const trimmed = text.trim();
  const negative = /^\(.*\)$/.test(trimmed) || trimmed.startsWith("-");
  const digits = trimmed.replace(/[()$€£,\s-]/g, "");
  if (/^\d+(\.\d+)?$/.test(digits)) return (negative ? -1 : 1) * Number(digits);
  return trimmed;
}
