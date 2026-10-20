import { describe, expect, it } from "vitest";
import { checkAgainstDocument } from "./extraction";
import type { TextRun } from "./geometry";

const run = (text: string, x: number, y: number, width = 40): TextRun => ({ text, box: { x, y, width, height: 10 } });
const pages = [[run("INVOICE #10442", 72, 90, 90), run("Total due", 72, 250, 60), run("$12,480.00", 400, 250, 60)]];

describe("checkAgainstDocument", () => {
  it("accepts an answer whose quote and value check out, with its location", () => {
    const [r] = checkAgainstDocument([{ field: "Amount", value: "12480", quote: "Total due $12,480.00", page: 1 }], pages);
    expect(r.status).toBe("supported");
    expect(r.status === "supported" && r.page).toBe(1);
  });

  it("rejects each kind of unsupported answer", () => {
    const results = checkAgainstDocument(
      [
        { field: "Customer", value: null, quote: null, page: null },
        { field: "Amount", value: "12840", quote: "Total due $12,480.00", page: 1 },
        { field: "Invoice", value: "10443", quote: "INVOICE #10443", page: 1 },
      ],
      pages,
    );
    expect(results.map((r) => r.status)).toEqual(["not-found", "value-not-in-quote", "quote-not-in-document"]);
  });
});
