import { useEffect, useState } from "react";
import { askSheet } from "../lib/ai";
import type { StoredDoc } from "../lib/docStore";
import {
  checkSheetAnswers,
  loadWorkbook,
  searchWorkbook,
  workbookAsText,
  type SourceWorkbook,
} from "../lib/spreadsheet";

export interface SheetResult {
  sheet: string;
  address: string;
  text: string;
  context: string;
  reason?: string;
  supported: boolean; // false: the AI named a cell that doesn't hold its answer
  method: "manual" | "ai";
}

/**
 * Searches a client spreadsheet: "Find" matches text or amounts instantly in
 * the browser; "Ask AI" answers a plain-language question with source cells,
 * each checked against the file before it can be used.
 */
export function SearchPanel({
  doc,
  cell,
  onShow,
  onUse,
}: {
  doc: StoredDoc;
  cell: string;
  onShow: (r: SheetResult) => void;
  onUse: (r: SheetResult) => Promise<void>;
}) {
  const [wb, setWb] = useState<SourceWorkbook | null>(null);
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<SheetResult[] | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    setWb(null);
    setResults(null);
    setError("");
    loadWorkbook(doc.bytes, doc.name).then(setWb, (e: Error) => setError(`Couldn't read ${doc.name}: ${e.message}`));
  }, [doc]);

  const find = () => {
    if (!wb) return;
    setError("");
    setResults(searchWorkbook(wb, query).map((h) => ({ ...h, supported: true, method: "manual" })));
  };

  const ask = async () => {
    if (!wb) return;
    setError("");
    setBusy(true);
    try {
      const { text, truncated } = workbookAsText(wb);
      const answers = checkSheetAnswers(wb, await askSheet(text, query));
      setResults(
        answers.map((a) => ({
          sheet: a.sheet,
          address: a.cell,
          text: a.found?.text ?? a.answer,
          context: a.context,
          reason: a.reason,
          supported: a.supported,
          method: "ai",
        })),
      );
      if (truncated) setError("This file is large, so the AI only saw its first part. Use Find for the rest.");
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const sheetCount = wb?.length ?? 0;

  return (
    <section>
      <h2>Search {doc.name}</h2>
      <p className="muted">
        {wb ? `${sheetCount} sheet${sheetCount === 1 ? "" : "s"}. ` : "Reading… "}
        Find text or an amount, or ask a question in plain language.
      </p>
      <textarea
        rows={2}
        value={query}
        onChange={(e) => setQuery(e.target.value)}
        placeholder="e.g. 10443, or: what is the amount for invoice 10443?"
      />
      <button onClick={find} disabled={!wb || !query.trim()}>
        Find
      </button>
      <button onClick={ask} disabled={!wb || !query.trim() || busy}>
        {busy ? "Asking…" : "Ask AI"}
      </button>
      {error && <p className="warn">{error}</p>}
      {results && results.length === 0 && <p className="muted">Nothing found.</p>}
      {results && results.length > 0 && (
        <ul className="hits">
          {results.map((r, i) => (
            <li key={i}>
              <div>
                <strong>{r.text}</strong> <span className="where">{r.sheet}!{r.address}</span>
              </div>
              {r.context && <div className="muted">{r.context}</div>}
              {r.reason && <div className="muted">{r.reason}</div>}
              {r.supported ? (
                <>
                  <button onClick={() => onShow(r)}>Show</button>
                  <button onClick={() => onUse(r)} disabled={!cell}>
                    Use in {cell.split("!").pop() || "cell"}
                  </button>
                </>
              ) : (
                <p className="warn">The AI pointed at a cell that doesn't hold this value, so it can't be used.</p>
              )}
            </li>
          ))}
        </ul>
      )}
    </section>
  );
}
