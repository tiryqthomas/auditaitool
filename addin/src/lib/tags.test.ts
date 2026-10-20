import { describe, expect, it } from "vitest";
import { columnLetter } from "./sheet";
import { normaliseAddress, rowToTag, splitAddress, tagToRow, toCellValue } from "./tags";
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
