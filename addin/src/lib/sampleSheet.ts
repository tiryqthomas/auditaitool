import type { Population, SampleSettings, Selection } from "./sampling";
import { columnLetter } from "./sheet";
import { addTags } from "./tags";
import type { EvidenceTag } from "./types";

/** Column on a generated Sample sheet that traces each item to its population row. Not a field to test. */
export const POPULATION_REF = "Population ref";

const AMOUNT_FORMAT = "#,##0.00;(#,##0.00)";

async function uniqueSheetName(context: Excel.RequestContext, base: string): Promise<string> {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();
  const taken = new Set(sheets.items.map((s) => s.name.toLowerCase()));
  for (let n = 1; ; n++) {
    const name = n === 1 ? base : `${base} (${n})`;
    if (!taken.has(name.toLowerCase())) return name;
  }
}

export interface WrittenSample {
  sheetName: string;
  memoName: string;
}

/**
 * Writes the selected items to a new Sample sheet (the population's columns,
 * then a reference tagged to each item's row in the source file) and a memo
 * sheet recording how the sample was picked, so it can be re-performed.
 */
export async function writeSampleSheet(
  pop: Population,
  selection: Selection,
  settings: SampleSettings,
  source: { docId: string; docName: string },
): Promise<WrittenSample> {
  const headers = pop.headers.filter(Boolean);
  const cols = headers.map((h) => pop.headers.indexOf(h));
  const refCol = headers.length;
  const amountCol = headers.indexOf(pop.amountHeader);

  const written = await Excel.run(async (context) => {
    const sheetName = await uniqueSheetName(context, "Sample");
    const memoName = await uniqueSheetName(context, "Sample memo");
    const sheet = context.workbook.worksheets.add(sheetName);
    const rows = selection.items.map((item) => [
      ...cols.map((c) => {
        const cell = item.cells[c];
        if (!cell) return "";
        // Amounts go in as numbers so Fill sample compares them as amounts.
        return typeof cell.value === "number" ? cell.value : cell.text;
      }),
      `${pop.sheet}!${item.keyAddress}`,
    ]);
    sheet.getRangeByIndexes(0, 0, 1, refCol + 1).values = [[...headers, POPULATION_REF]];
    sheet.getRangeByIndexes(0, 0, 1, refCol + 1).format.font.bold = true;
    if (rows.length) {
      sheet.getRangeByIndexes(1, 0, rows.length, refCol + 1).values = rows;
      if (amountCol >= 0) sheet.getRangeByIndexes(1, amountCol, rows.length, 1).numberFormat = rows.map(() => [AMOUNT_FORMAT]);
    }

    const memo = context.workbook.worksheets.add(memoName);
    const m = (n: number) => n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const lines: [string, string | number][] = [
      ["Sample selection memo", ""],
      ["Population file", source.docName],
      ["Population sheet", pop.sheet],
      ["Items in population", pop.items.length],
      ["Population total", m(selection.populationTotal)],
      ["Rows left out", pop.skipped.map((s) => `row ${s.row + 1} (${s.reason})`).join("; ") || "none"],
      ["Matched by", pop.keyHeader],
      ["Amount column", pop.amountHeader || "none"],
      ["Method", settings.method === "mus" ? "Monetary unit sampling (systematic, random start)" : "Random selection"],
      ["Sample size requested", settings.size],
      ["Key items (always tested)", settings.keyItemThreshold === undefined ? "none" : `${selection.keyItems} at or over ${m(settings.keyItemThreshold)}`],
      ...(selection.interval !== undefined
        ? ([
            ["Sampling interval", m(selection.interval)],
            ["Random start", m(selection.randomStart!)],
            ["Not selectable by MUS", `${selection.notSelectable} zero, negative or blank items; consider testing separately`],
          ] as [string, string][])
        : []),
      ["Seed", settings.seed],
      ["Items selected", selection.items.length],
      ["Sample sheet", sheetName],
      ["Selected", new Date().toLocaleString()],
      ["Note", "Selection is done in code and is reproducible: the same file, settings and seed pick the same items."],
    ];
    memo.getRangeByIndexes(0, 0, lines.length, 2).values = lines;
    memo.getRange("A1").format.font.bold = true;
    memo.getRangeByIndexes(0, 0, lines.length, 2).format.autofitColumns();

    sheet.getRangeByIndexes(0, 0, rows.length + 1, refCol + 1).format.autofitColumns();
    sheet.activate();
    sheet.getRange("A1").select();
    await context.sync();
    return { sheetName, memoName };
  });

  // Each reference cell is tagged to the item's key cell in the population file.
  const now = new Date().toISOString();
  const refLetter = columnLetter(refCol);
  await addTags(
    selection.items.map((item, i): { tag: EvidenceTag } => ({
      tag: {
        tagId: crypto.randomUUID(),
        cell: `${written.sheetName}!${refLetter}${i + 2}`,
        docId: source.docId,
        docName: source.docName,
        page: 0,
        box: { x: 0, y: 0, width: 0, height: 0 },
        sourceCell: `${pop.sheet}!${item.keyAddress}`,
        quotedText: `${item.key} (selected: ${item.reason})`,
        method: "manual",
        status: "proposed",
        createdAt: now,
      },
    })),
  );
  return written;
}
