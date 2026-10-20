import type { EvidenceTag } from "./types";

// Tags live inside the workbook on a hidden sheet, so a workpaper stays
// self-contained when it is emailed or archived.
export const EVIDENCE_SHEET = "_Evidence";

const COLUMNS = [
  "tagId", "cell", "docId", "docName", "page", "x", "y", "width", "height",
  "quotedText", "method", "status", "createdAt", "reviewedBy", "reviewedAt", "note", "sourceCell",
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
    t.quotedText, t.method, t.status, t.createdAt, t.reviewedBy ?? "", t.reviewedAt ?? "", t.note ?? "", t.sourceCell ?? "",
  ];
}

export function rowToTag(row: unknown[]): EvidenceTag {
  const [tagId, cell, docId, docName, page, x, y, width, height, quotedText, method, status, createdAt, reviewedBy, reviewedAt, note, sourceCell] = row;
  // Rows written by older versions lack the later columns.
  const optional = (v: unknown) => (v === undefined || v === "" ? undefined : String(v));
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
    reviewedBy: optional(reviewedBy),
    reviewedAt: optional(reviewedAt),
    note: optional(note),
    sourceCell: optional(sourceCell),
  };
}

/** Normalises an address so "Sheet1!$B$4" and "Sheet1!B4" match. */
export function normaliseAddress(address: string): string {
  return address.replace(/\$/g, "").replace(/^'(.*)'!/, "$1!");
}

async function evidenceSheet(context: Excel.RequestContext): Promise<Excel.Worksheet> {
  const existing = context.workbook.worksheets.getItemOrNullObject(EVIDENCE_SHEET);
  await context.sync();
  if (!existing.isNullObject) {
    // Keeps the header current on workbooks tagged by an older version.
    existing.getRangeByIndexes(0, 0, 1, COLUMNS.length).values = [[...COLUMNS]];
    return existing;
  }
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
  return latestForCell(await readTags(), cell);
}

/** Latest row for a cell: a re-tag or a sign-off supersedes earlier rows. */
export function latestForCell(tags: EvidenceTag[], cell: string): EvidenceTag | undefined {
  const target = normaliseAddress(cell);
  for (let i = tags.length - 1; i >= 0; i--) if (normaliseAddress(tags[i].cell) === target) return tags[i];
  return undefined;
}

/** Every row recorded for a cell, oldest first. */
export function historyForCell(tags: EvidenceTag[], cell: string): EvidenceTag[] {
  const target = normaliseAddress(cell);
  return tags.filter((t) => normaliseAddress(t.cell) === target);
}

/** The current tag of every tagged cell, in workbook order (sheet, then row, then column). */
export function currentTags(tags: EvidenceTag[]): EvidenceTag[] {
  const latest = new Map<string, EvidenceTag>();
  for (const t of tags) latest.set(normaliseAddress(t.cell), t);
  const sheetOrder = new Map<string, number>();
  for (const t of latest.values()) {
    const [sheet] = splitAddress(t.cell);
    if (!sheetOrder.has(sheet)) sheetOrder.set(sheet, sheetOrder.size);
  }
  const key = (t: EvidenceTag): [number, number, number] => {
    const [sheet, address] = splitAddress(t.cell);
    const m = /^([A-Z]+)(\d+)/.exec(address) ?? ["", "A", "0"];
    const col = [...m[1]].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
    return [sheetOrder.get(sheet)!, Number(m[2]), col];
  };
  return [...latest.values()].sort((a, b) => {
    const [ka, kb] = [key(a), key(b)];
    return ka[0] - kb[0] || ka[1] - kb[1] || ka[2] - kb[2];
  });
}

/** A cell awaits review until an auditor signs off its current tag (AI-flagged exceptions included). */
export function awaitingReview(t: EvidenceTag): boolean {
  return !t.reviewedBy;
}

export interface ReviewSummary {
  total: number;
  awaiting: number;
  verified: number;
  exception: number;
  rejected: number;
}

/** Counts tagged cells by their current sign-off. */
export function summarise(tags: EvidenceTag[]): ReviewSummary {
  const summary: ReviewSummary = { total: 0, awaiting: 0, verified: 0, exception: 0, rejected: 0 };
  for (const t of currentTags(tags)) {
    summary.total++;
    if (awaitingReview(t)) summary.awaiting++;
    else if (t.status !== "proposed") summary[t.status]++;
  }
  return summary;
}

/** The next cell still awaiting review after `fromCell`, wrapping to the start. */
export function nextUnreviewed(tags: EvidenceTag[], fromCell?: string): EvidenceTag | undefined {
  const pending = currentTags(tags);
  const from = fromCell ? normaliseAddress(fromCell) : "";
  const start = pending.findIndex((t) => normaliseAddress(t.cell) === from) + 1;
  const ordered = [...pending.slice(start), ...pending.slice(0, start)];
  return ordered.find(awaitingReview);
}

export type ReviewDecision = "verified" | "rejected" | "exception";

/** Builds the sign-off row: same tag and evidence, new status, reviewer, time and note. */
export function reviewedCopy(tag: EvidenceTag, decision: ReviewDecision, reviewer: string, note: string, at = new Date()): EvidenceTag {
  const trimmedNote = note.trim();
  if (!reviewer.trim()) throw new Error("Enter your name or initials before signing off.");
  if (decision !== "verified" && !trimmedNote) throw new Error("Add a note explaining the rejection or exception.");
  return {
    ...tag,
    status: decision,
    reviewedBy: reviewer.trim(),
    reviewedAt: at.toISOString(),
    note: trimmedNote || undefined,
  };
}

/** Records a sign-off. The cell's value is never changed; only its border colour. */
export function reviewTag(tag: EvidenceTag, decision: ReviewDecision, reviewer: string, note: string): Promise<void> {
  return addTags([{ tag: reviewedCopy(tag, decision, reviewer, note) }]);
}

/** Moves the selection to a tagged cell. */
export async function selectCell(cell: string): Promise<void> {
  await Excel.run(async (context) => {
    const [sheetName, address] = splitAddress(cell);
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.activate();
    sheet.getRange(address).select();
    await context.sync();
  });
}

/** Writes the value into the cell, records the tag and marks the cell by status. */
export function addTag(tag: EvidenceTag, value: string): Promise<void> {
  return addTags([{ tag, value }]);
}

/**
 * Records several tags in one batch. A write without a value leaves the cell's
 * existing content alone and only marks it (used for exceptions, where the
 * workpaper value must not be overwritten by what the document says).
 */
export async function addTags(writes: { tag: EvidenceTag; value?: string }[]): Promise<void> {
  if (writes.length === 0) return;
  await Excel.run(async (context) => {
    const sheet = await evidenceSheet(context);
    const used = sheet.getUsedRange();
    used.load("rowCount");
    await context.sync();
    sheet.getRangeByIndexes(used.rowCount, 0, writes.length, COLUMNS.length).values = writes.map((w) => tagToRow(w.tag));

    for (const { tag, value } of writes) {
      const [sheetName, address] = splitAddress(tag.cell);
      const target = context.workbook.worksheets.getItem(sheetName).getRange(address);
      if (value !== undefined) target.values = [[toCellValue(value)]];
      for (const edge of ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"] as const) {
        const border = target.format.borders.getItem(edge);
        border.style = "Continuous";
        border.weight = "Medium";
        border.color = STATUS_COLOR[tag.status];
      }
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
