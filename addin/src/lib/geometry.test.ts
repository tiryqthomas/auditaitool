import { describe, expect, it } from "vitest";
import { boxFromPoints, textInBox, type TextRun } from "./geometry";

const run = (text: string, x: number, y: number, width = 40, height = 10): TextRun => ({
  text,
  box: { x, y, width, height },
});

describe("boxFromPoints", () => {
  it("normalises a box dragged up and to the left", () => {
    expect(boxFromPoints(100, 80, 20, 30)).toEqual({ x: 20, y: 30, width: 80, height: 50 });
  });
});

describe("textInBox", () => {
  const invoice = [
    run("Invoice", 50, 50),
    run("#10442", 95, 50),
    run("Total", 50, 300),
    run("$12,480.00", 100, 300, 60),
    run("Tax", 50, 320),
    run("$998.40", 100, 320),
  ];

  it("returns only the runs whose centre is inside the box", () => {
    expect(textInBox(invoice, { x: 95, y: 295, width: 70, height: 18 })).toBe("$12,480.00");
  });

  it("keeps reading order across lines", () => {
    expect(textInBox(invoice, { x: 40, y: 290, width: 150, height: 45 })).toBe("Total $12,480.00\nTax $998.40");
  });

  it("returns an empty string when nothing is selected", () => {
    expect(textInBox(invoice, { x: 400, y: 400, width: 10, height: 10 })).toBe("");
  });
});
