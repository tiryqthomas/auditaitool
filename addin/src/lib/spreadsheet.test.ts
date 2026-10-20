import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import {
  checkSheetAnswers,
  getCell,
  isSpreadsheet,
  loadWorkbook,
  parseCsv,
  searchWorkbook,
  sheetFromCsv,
  splitSourceCell,
  workbookAsText,
} from "./spreadsheet";

const bytes = (path: string) => {
  const b = readFileSync(new URL(path, import.meta.url));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
};

describe("parseCsv", () => {
  it("handles quotes, doubled quotes, commas and CRLF", () => {
    expect(parseCsv('a,"b, c","say ""hi"""\r\n1,2,3\n')).toEqual([
      ["a", "b, c", 'say "hi"'],
      ["1", "2", "3"],
    ]);
  });
});

describe("file types", () => {
  it("recognises spreadsheets by extension", () => {
    expect(["a.xlsx", "b.CSV", "c.xlsm", "d.pdf"].map(isSpreadsheet)).toEqual([true, true, true, false]);
  });
  it("splits quoted sheet references", () => {
    expect(splitSourceCell("'Sales Ledger'!$E$9")).toEqual(["Sales Ledger", "E9"]);
  });
});

describe("sample client ledger", async () => {
  const wb = await loadWorkbook(bytes("../../../samples/client-sales-ledger.xlsx"), "client-sales-ledger.xlsx");

  it("reads every visible sheet with formatted values", () => {
    expect(wb.map((s) => s.name)).toEqual(["Sales Ledger", "Customers"]);
    expect(getCell(wb, "Sales Ledger", "E9")).toMatchObject({ value: 7530, text: "7,530.00" });
    expect(getCell(wb, "Sales Ledger", "B9")?.text).toBe("2026-10-02");
    expect(getCell(wb, "Sales Ledger", "E15")?.text).toBe("(1,200.00)");
    expect(getCell(wb, "Customers", "B3")?.text).toBe("$20,000");
  });

  it("uses a formula's saved result", () => {
    expect(getCell(wb, "Sales Ledger", "E17")).toMatchObject({ value: 57055.24, text: "57,055.24" });
  });

  it("finds text and amounts with their row and column labels", () => {
    expect(searchWorkbook(wb, "10443")).toEqual([
      { sheet: "Sales Ledger", address: "A9", text: "INV-10443", context: "Invoice # · Brightline Retail Inc." },
    ]);
    const amount = searchWorkbook(wb, "$7,530.00");
    expect(amount.map((h) => `${h.address} ${h.context}`)).toEqual(["E9 Amount · INV-10443"]);
  });

  it("checks AI answers against the named cell", () => {
    const [good, wrongCell, missing] = checkSheetAnswers(wb, [
      { answer: "7,530.00", sheet: "Sales Ledger", cell: "E9", reason: "" },
      { answer: "7,530.00", sheet: "Sales Ledger", cell: "E10", reason: "" },
      { answer: "1", sheet: "Nope", cell: "A1", reason: "" },
    ]);
    expect(good).toMatchObject({ supported: true, context: "Amount · INV-10443" });
    expect(wrongCell.supported).toBe(false);
    expect(missing).toMatchObject({ supported: false, found: null });
  });

  it("writes the workbook as addressed text for the AI", () => {
    const { text, truncated } = workbookAsText(wb);
    expect(truncated).toBe(false);
    expect(text).toContain("=== Sheet: Customers ===");
    expect(text).toContain('A9="INV-10443" | B9="2026-10-02"');
    expect(workbookAsText(wb, 200).truncated).toBe(true);
  });
});

describe("CSV evidence", () => {
  it("keeps numbers as written", () => {
    const sheet = sheetFromCsv("aging", readFileSync(new URL("../../../samples/client-ar-aging.csv", import.meta.url), "utf8"));
    expect(sheet.rows[4][6]).toEqual({ address: "G5", value: 14120, text: "14120.00" });
    expect(searchWorkbook([sheet], "14,120")[0]).toMatchObject({ address: "B5", context: "Current · Delta Hardware" });
  });
});
