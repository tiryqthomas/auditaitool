import type { Box } from "./types";

/** A positioned run of text on a PDF page, in the same units as Box. */
export interface TextRun {
  text: string;
  box: Box;
}

/** Normalises a box drawn by dragging in any direction. */
export function boxFromPoints(x1: number, y1: number, x2: number, y2: number): Box {
  return {
    x: Math.min(x1, x2),
    y: Math.min(y1, y2),
    width: Math.abs(x2 - x1),
    height: Math.abs(y2 - y1),
  };
}

function containsPoint(box: Box, x: number, y: number): boolean {
  return x >= box.x && x <= box.x + box.width && y >= box.y && y <= box.y + box.height;
}

/**
 * Returns the text inside a selection box, in reading order. A run counts as
 * inside when its centre falls within the box, so a box drawn loosely around a
 * figure picks it up without grabbing half of the neighbouring line.
 */
export function textInBox(runs: TextRun[], selection: Box): string {
  const inside = runs.filter((r) =>
    containsPoint(selection, r.box.x + r.box.width / 2, r.box.y + r.box.height / 2),
  );
  // Group into lines: runs whose vertical centres are within half a line height.
  const sorted = [...inside].sort((a, b) => a.box.y - b.box.y || a.box.x - b.box.x);
  const lines: TextRun[][] = [];
  for (const run of sorted) {
    const line = lines[lines.length - 1];
    const centre = run.box.y + run.box.height / 2;
    if (line) {
      const ref = line[0].box;
      if (Math.abs(ref.y + ref.height / 2 - centre) <= Math.max(ref.height, run.box.height) / 2) {
        line.push(run);
        continue;
      }
    }
    lines.push([run]);
  }
  return lines
    .map((line) =>
      line
        .sort((a, b) => a.box.x - b.box.x)
        .map((r) => r.text.trim())
        .filter(Boolean)
        .join(" "),
    )
    .filter(Boolean)
    .join("\n");
}

/** Scales a box between page units and on-screen pixels. */
export function scaleBox(box: Box, factor: number): Box {
  return {
    x: box.x * factor,
    y: box.y * factor,
    width: box.width * factor,
    height: box.height * factor,
  };
}
