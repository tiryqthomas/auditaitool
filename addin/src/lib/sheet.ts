export interface TargetCell {
  address: string; // Sheet!B4
  header: string; // the column's heading, used as the field to extract
}

/**
 * For a selection across one row of a testing table, pairs each selected cell
 * with its column heading (the first row of the sheet's used range).
 */
export async function selectedCellsWithHeaders(): Promise<TargetCell[]> {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const selection = context.workbook.getSelectedRange();
    const used = sheet.getUsedRange();
    selection.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    used.load(["rowIndex"]);
    sheet.load("name");
    await context.sync();
    if (selection.rowCount !== 1) throw new Error("Select the cells in one row of your testing table.");
    if (selection.rowIndex <= used.rowIndex) throw new Error("Select a row below the column headings.");
    const headers = sheet.getRangeByIndexes(used.rowIndex, selection.columnIndex, 1, selection.columnCount);
    headers.load("values");
    await context.sync();
    return headers.values[0]
      .map((h, i) => ({
        address: `${sheet.name}!${columnLetter(selection.columnIndex + i)}${selection.rowIndex + 1}`,
        header: String(h).trim(),
      }))
      .filter((t) => t.header);
  });
}

export function columnLetter(index: number): string {
  let s = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}
