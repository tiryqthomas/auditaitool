import { useEffect, useState } from "react";
import type { StoredDoc } from "../lib/docStore";
import { columnLetter } from "../lib/sheet";
import { cellContext, loadWorkbook, parseAddress, splitSourceCell, type SourceWorkbook } from "../lib/spreadsheet";
import { send } from "./Viewer";

const MAX_ROWS = 2000;

/**
 * Shows a client spreadsheet as a read-only grid.
 * mode=tag: click a cell, then "Use this value".
 * mode=trace: opens on the tag's sheet with its source cell highlighted.
 */
export function SheetViewer({ doc, mode, sourceCell }: { doc: StoredDoc; mode: "tag" | "trace"; sourceCell?: string }) {
  const [wb, setWb] = useState<SourceWorkbook | null>(null);
  const [error, setError] = useState("");
  const [target] = useState(() => (sourceCell ? splitSourceCell(sourceCell) : null));
  const [sheetName, setSheetName] = useState(target?.[0] ?? "");
  const [picked, setPicked] = useState<string | null>(mode === "trace" ? target?.[1] ?? null : null);

  useEffect(() => {
    loadWorkbook(doc.bytes, doc.name).then(
      (book) => {
        setWb(book);
        setSheetName((s) => (book.some((x) => x.name === s) ? s : book[0]?.name ?? ""));
      },
      (e: Error) => setError(`Couldn't read this spreadsheet: ${e.message}`),
    );
  }, [doc]);

  if (error) return <p className="error">{error}</p>;
  if (!wb) return <p className="muted">Reading spreadsheet…</p>;
  const sheet = wb.find((s) => s.name === sheetName);
  if (!sheet) return <p className="error">This spreadsheet has no visible sheets.</p>;

  const width = Math.max(1, ...sheet.rows.map((r) => r.length));
  const rows = sheet.rows.slice(0, MAX_ROWS);
  const pickedAt = picked ? parseAddress(picked) : null;
  const pickedCell = pickedAt ? sheet.rows[pickedAt.row]?.[pickedAt.col] ?? null : null;

  return (
    <div className="viewer">
      <div className="viewer-top">
      <header className="viewer-bar">
        <strong>{doc.name}</strong>
        {wb.length > 1 && (
          <select value={sheetName} onChange={(e) => { setSheetName(e.target.value); setPicked(null); }}>
            {wb.map((s) => (
              <option key={s.name}>{s.name}</option>
            ))}
          </select>
        )}
        {mode === "tag" && <span className="hint">Click the cell that holds the value.</span>}
      </header>
      {mode === "tag" && pickedCell && pickedAt && (
        <div className="selection">
          <code>
            {sheet.name}!{picked}: {pickedCell.text}
          </code>
          <button
            onClick={() =>
              send({
                type: "cell-selected",
                docId: doc.docId,
                sourceCell: `${sheet.name}!${picked}`,
                text: pickedCell.text,
                context: cellContext(sheet, pickedAt.row, pickedAt.col),
              })
            }
          >
            Use this value
          </button>
        </div>
      )}
      </div>
      <div className="grid-wrap">
        <table className={`grid ${mode === "tag" ? "taggable" : ""}`}>
          <thead>
            <tr>
              <th />
              {Array.from({ length: width }, (_, c) => (
                <th key={c}>{columnLetter(c)}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, r) => (
              <tr key={r}>
                <th>{r + 1}</th>
                {Array.from({ length: width }, (_, c) => {
                  const cell = row[c];
                  const address = `${columnLetter(c)}${r + 1}`;
                  const isPicked = pickedAt?.row === r && pickedAt.col === c;
                  return (
                    <td
                      key={c}
                      className={`${typeof cell?.value === "number" ? "num" : ""} ${isPicked ? `picked ${mode}` : ""}`}
                      onClick={mode === "tag" ? () => setPicked(address) : undefined}
                      // Scroll only this window, as in the PDF viewer, so Excel's page stays put.
                      ref={(el) => {
                        if (isPicked && mode === "trace" && el) {
                          const box = el.getBoundingClientRect();
                          window.scrollTo({
                            top: box.top + window.scrollY - window.innerHeight / 2,
                            left: box.left + window.scrollX - window.innerWidth / 2,
                          });
                        }
                      }}
                    >
                      {cell?.text ?? ""}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
        {sheet.rows.length > MAX_ROWS && <p className="muted">Showing the first {MAX_ROWS} rows.</p>}
      </div>
    </div>
  );
}
