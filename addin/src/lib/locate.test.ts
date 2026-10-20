import { describe, expect, it } from "vitest";
import type { TextRun } from "./geometry";
import { locateQuote, valueMatchesQuote } from "./locate";

const run = (text: string, x: number, y: number, width = 40, height = 10): TextRun => ({
  text,
  box: { x, y, width, height },
});

const page1 = [run("INVOICE #10442", 72, 90, 90), run("Date:", 72, 106, 30), run("2026-09-14", 105, 106, 60)];
const page2 = [run("Total", 72, 250), run("due", 115, 250, 20), run("$12,480.00", 400, 250, 60)];

describe("locateQuote", () => {
  it("finds a quote spanning several runs and boxes all of them", () => {
    expect(locateQuote([page1, page2], "Total due $12,480.00")).toEqual({
      page: 2,
      box: { x: 72, y: 250, width: 388, height: 10 },
    });
  });

  it("ignores case and extra whitespace", () => {
    expect(locateQuote([page1], "invoice  #10442")?.page).toBe(1);
  });

  it("rejects a quote that is not in the document", () => {
    expect(locateQuote([page1, page2], "$12,840.00")).toBeNull();
  });

  it("searches the suggested page first", () => {
    const dup = [run("Total", 72, 50)];
    expect(locateQuote([dup, page2], "Total", 2)?.page).toBe(2);
  });
});

describe("valueMatchesQuote", () => {
  it("matches numbers regardless of formatting", () => {
    expect(valueMatchesQuote("12480", "Total due $12,480.00")).toBe(true);
    expect(valueMatchesQuote("12480.5", "Total due $12,480.00")).toBe(false);
  });

  it("compares dates as text, not numbers", () => {
    expect(valueMatchesQuote("2026-09-14", "Date: 2026-09-14")).toBe(true);
    expect(valueMatchesQuote("2026-09-15", "Date: 2026-09-14")).toBe(false);
  });

  it("matches text values contained in the quote", () => {
    expect(valueMatchesQuote("Acme Manufacturing LLC", "Bill to: Acme Manufacturing LLC")).toBe(true);
    expect(valueMatchesQuote("Acme Corp", "Bill to: Acme Manufacturing LLC")).toBe(false);
  });
});
