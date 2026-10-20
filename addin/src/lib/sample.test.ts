import { describe, expect, it } from "vitest";
import type { CheckedField } from "./extraction";
import { keysMatch, planSample, valuesAgree, type CellState, type DocResult, type SampleRow } from "./sample";

const cell = (address: string, value: unknown, text = String(value ?? "")): CellState => ({ address, value, text });
const ok = (field: string, value: string): CheckedField => ({
  field, value, quote: value, page: 1, status: "supported", box: { x: 0, y: 0, width: 1, height: 1 },
});
const missing = (field: string): CheckedField => ({ field, value: null, quote: null, page: null, status: "not-found" });

describe("keysMatch", () => {
  it("matches the same invoice written differently", () => {
    expect(keysMatch("10442", "INVOICE #10442")).toBe(true);
    expect(keysMatch("INV-010442", "10442")).toBe(true);
    expect(keysMatch("A-1", "a1")).toBe(true);
  });
  it("does not match different invoices or short numbers", () => {
    expect(keysMatch("10442", "10443")).toBe(false);
    expect(keysMatch("A-12", "B-12")).toBe(false);
  });
});

describe("valuesAgree", () => {
  it("compares amounts numerically", () => {
    expect(valuesAgree(cell("A", 12480), "$12,480.00")).toBe(true);
    expect(valuesAgree(cell("A", 12480), "$12,840.00")).toBe(false);
  });
  it("compares an Excel date serial with a written date", () => {
    // 46279 is 2026-09-14 in Excel's date system.
    expect(valuesAgree(cell("A", 46279, "9/14/2026"), "2026-09-14")).toBe(true);
    expect(valuesAgree(cell("A", 46279, "9/14/2026"), "September 14, 2026")).toBe(true);
    expect(valuesAgree(cell("A", 46279, "9/14/2026"), "2026-09-15")).toBe(false);
  });
  it("compares text ignoring case and spacing", () => {
    expect(valuesAgree(cell("A", "Acme  Manufacturing llc"), "Acme Manufacturing LLC")).toBe(true);
  });
});

describe("planSample", () => {
  const row = (rowIndex: number, key: string, amount: unknown): SampleRow => ({
    rowIndex,
    key,
    cells: {
      "Invoice #": cell(`S!A${rowIndex + 1}`, key),
      Customer: cell(`S!B${rowIndex + 1}`, ""),
      Amount: cell(`S!C${rowIndex + 1}`, amount),
    },
  });
  const doc = (docId: string, key: string, customer: CheckedField, amount: string): DocResult => ({
    docId, docName: `${docId}.pdf`, fields: [ok("Invoice #", key), customer, ok("Amount", amount)],
  });

  it("fills empty cells, confirms matches and flags differences", () => {
    const plan = planSample(
      [row(1, "10442", 12480), row(2, "10443", 500), row(3, "10444", "")],
      "Invoice #",
      [
        doc("a", "10442", ok("Customer", "Acme"), "$12,480.00"),
        doc("b", "INV-10443", missing("Customer"), "$550.00"),
        doc("c", "99999", ok("Customer", "Other"), "$1.00"),
      ],
    );
    const kinds = plan.rows.map((r) => r.problem ?? r.outcomes.map((o) => `${o.header}:${o.kind}`).join(","));
    expect(kinds).toEqual([
      "Invoice #:agree,Customer:fill,Amount:agree",
      "Invoice #:agree,Customer:unsupported,Amount:exception",
      "no-document",
    ]);
    expect(plan.unmatchedDocs.map((d) => d.docId)).toEqual(["c"]);
  });

  it("refuses to pick when two documents claim the same row", () => {
    const plan = planSample([row(1, "10442", "")], "Invoice #", [
      doc("a", "10442", ok("Customer", "Acme"), "$1"),
      doc("b", "10442", ok("Customer", "Acme"), "$1"),
    ]);
    expect(plan.rows[0].problem).toBe("several-documents");
    expect(plan.unmatchedDocs).toEqual([]);
  });
});
