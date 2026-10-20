import { describe, expect, it } from "vitest";
import { columnLetter } from "./sheet";
import {
  historyForCell,
  latestForCell,
  nextUnreviewed,
  normaliseAddress,
  reviewedCopy,
  rowToTag,
  splitAddress,
  summarise,
  tagToRow,
  toCellValue,
} from "./tags";
import type { EvidenceTag } from "./types";

const tag: EvidenceTag = {
  tagId: "t1",
  cell: "Revenue Testing!D7",
  docId: "abc",
  docName: "INV-10442.pdf",
  page: 2,
  box: { x: 1, y: 2, width: 3, height: 4 },
  quotedText: "$12,480.00",
  method: "manual",
  status: "proposed",
  createdAt: "2026-10-20T00:00:00.000Z",
};

describe("tag rows", () => {
  it("round-trips a tag through a sheet row", () => {
    expect(rowToTag(tagToRow(tag))).toEqual(tag);
  });
});

describe("addresses", () => {
  it("strips absolute markers and quotes", () => {
    expect(normaliseAddress("'Revenue Testing'!$D$7")).toBe("Revenue Testing!D7");
  });
  it("splits sheet and cell", () => {
    expect(splitAddress("'Revenue Testing'!$D$7")).toEqual(["Revenue Testing", "D7"]);
  });
});

describe("toCellValue", () => {
  it("parses currency amounts", () => {
    expect(toCellValue("$12,480.00")).toBe(12480);
  });
  it("treats parentheses as negative", () => {
    expect(toCellValue("(500.25)")).toBe(-500.25);
  });
  it("leaves text alone", () => {
    expect(toCellValue("Acme Corp")).toBe("Acme Corp");
  });
});

describe("columnLetter", () => {
  it("converts zero-based indexes to Excel letters", () => {
    expect([0, 25, 26, 51, 701, 702].map(columnLetter)).toEqual(["A", "Z", "AA", "AZ", "ZZ", "AAA"]);
  });
});

describe("sign-off", () => {
  const at = new Date("2026-10-21T15:00:00.000Z");
  const other: EvidenceTag = { ...tag, tagId: "t2", cell: "Revenue Testing!D5" };
  const lateSheet: EvidenceTag = { ...tag, tagId: "t3", cell: "Other!A1" };

  it("reads rows written before review columns existed", () => {
    const oldRow = tagToRow(tag).slice(0, 13);
    expect(rowToTag(oldRow)).toEqual(tag);
  });

  it("round-trips the reviewer, time and note", () => {
    const reviewed = reviewedCopy(tag, "exception", "TT", "Invoice shows 7,530", at);
    expect(rowToTag(tagToRow(reviewed))).toEqual(reviewed);
    expect(reviewed).toMatchObject({ status: "exception", reviewedBy: "TT", reviewedAt: at.toISOString(), note: "Invoice shows 7,530" });
    expect(reviewed.tagId).toBe(tag.tagId);
  });

  it("needs a reviewer, and a note to reject or mark an exception", () => {
    expect(() => reviewedCopy(tag, "verified", " ", "")).toThrow(/name/);
    expect(() => reviewedCopy(tag, "rejected", "TT", "  ")).toThrow(/note/);
    expect(reviewedCopy(tag, "verified", "TT", "", at).note).toBeUndefined();
  });

  it("treats the latest row as the cell's current tag and keeps the history", () => {
    const rows = [tag, other, reviewedCopy(tag, "verified", "TT", "", at)];
    expect(latestForCell(rows, "'Revenue Testing'!$D$7")?.status).toBe("verified");
    expect(historyForCell(rows, "Revenue Testing!D7").map((t) => t.status)).toEqual(["proposed", "verified"]);
  });

  it("counts cells by current status", () => {
    const aiFlagged: EvidenceTag = { ...other, status: "exception" };
    const rows = [tag, aiFlagged, lateSheet, reviewedCopy(tag, "verified", "TT", "", at)];
    expect(summarise(rows)).toEqual({ total: 3, awaiting: 2, verified: 1, exception: 0, rejected: 0 });
  });

  it("finds the next unreviewed cell in workbook order, wrapping around", () => {
    const rows = [lateSheet, tag, { ...other, status: "exception" as const }];
    expect(nextUnreviewed(rows)?.cell).toBe("Other!A1");
    expect(nextUnreviewed(rows, "Other!A1")?.cell).toBe("Revenue Testing!D5");
    expect(nextUnreviewed(rows, "Revenue Testing!D7")?.cell).toBe("Other!A1");
    const allDone = rows.map((t) => reviewedCopy(t, "verified", "TT", "", at));
    expect(nextUnreviewed(allDone)).toBeUndefined();
  });
});
