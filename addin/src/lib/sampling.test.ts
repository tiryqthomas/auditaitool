import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { guessAmountHeader, listingHeaders, readPopulation, seededRandom, selectSample } from "./sampling";
import { loadWorkbook, sheetFromCsv } from "./spreadsheet";

const file = readFileSync(new URL("../../../samples/client-sales-ledger.xlsx", import.meta.url));
const wb = await loadWorkbook(file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer, "ledger.xlsx");
const ledger = wb[0];

describe("readPopulation", () => {
  const pop = readPopulation(ledger, "Invoice #", "Amount");

  it("reads every item below the headings and skips the total row", () => {
    expect(listingHeaders(ledger)).toEqual(["Invoice #", "Date", "Customer", "Description", "Amount", "Status"]);
    expect(guessAmountHeader(listingHeaders(ledger))).toBe("Amount");
    expect(pop.items).toHaveLength(13);
    expect(pop.items[0]).toMatchObject({ key: "INV-10438", keyAddress: "A4", amount: 1384.85 });
    expect(pop.skipped).toEqual([{ row: 16, reason: "total row" }]);
  });

  it("reports rows with no key", () => {
    const sheet = sheetFromCsv("l", "Ref,Amount\nA1,10\n,20\nA3,30\n");
    const p = readPopulation(sheet, "Ref", "Amount");
    expect(p.items.map((i) => i.key)).toEqual(["A1", "A3"]);
    expect(p.skipped).toEqual([{ row: 2, reason: "no Ref" }]);
  });
});

describe("selectSample", () => {
  const pop = readPopulation(ledger, "Invoice #", "Amount");

  it("repeats exactly with the same seed", () => {
    const a = selectSample(pop, { method: "random", size: 5, seed: 42 });
    const b = selectSample(pop, { method: "random", size: 5, seed: 42 });
    expect(a.items.map((i) => i.key)).toEqual(b.items.map((i) => i.key));
    expect(new Set(a.items.map((i) => i.key)).size).toBe(5);
    expect(selectSample(pop, { method: "random", size: 5, seed: 43 }).items.map((i) => i.key)).not.toEqual(a.items.map((i) => i.key));
  });

  it("always includes key items and adds the sample on top", () => {
    const s = selectSample(pop, { method: "random", size: 3, seed: 7, keyItemThreshold: 10000 });
    expect(s.keyItems).toBe(2);
    expect(s.items.filter((i) => i.reason === "key item").map((i) => i.key)).toEqual(["INV-10442", "INV-10446"]);
    expect(s.items).toHaveLength(5);
    expect(s.items.map((i) => i.row)).toEqual([...s.items.map((i) => i.row)].sort((x, y) => x - y));
  });

  it("caps random selection at the population size", () => {
    expect(selectSample(pop, { method: "random", size: 50, seed: 1 }).items).toHaveLength(13);
  });

  it("selects monetary units systematically and leaves out credits", () => {
    const s = selectSample(pop, { method: "mus", size: 4, seed: 9 });
    expect(s.notSelectable).toBe(1); // the credit memo
    expect(s.interval).toBeCloseTo(58255.24 / 4, 2);
    expect(s.items.length).toBeGreaterThan(0);
    expect(s.items.length).toBeLessThanOrEqual(4);
    expect(s.items.every((i) => i.amount! > 0 && i.reason === "monetary unit")).toBe(true);
    // Every item at least as big as the interval is certain to be picked.
    const big = pop.items.filter((i) => i.amount! >= s.interval!).map((i) => i.key);
    for (const k of big) expect(s.items.map((i) => i.key)).toContain(k);
  });

  it("gives a stable sequence for a seed", () => {
    const r = seededRandom(123);
    const first = [r(), r(), r()];
    const again = seededRandom(123);
    expect([again(), again(), again()]).toEqual(first);
    expect(first.every((x) => x >= 0 && x < 1)).toBe(true);
  });
});
