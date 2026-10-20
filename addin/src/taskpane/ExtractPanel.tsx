import { useState } from "react";
import { getExtractor } from "../lib/ai";
import { toBase64, type StoredDoc } from "../lib/docStore";
import { checkAgainstDocument, OUTCOME_LABEL, type CheckedField } from "../lib/extraction";
import { loadPdf, pageTextRuns } from "../lib/pdf";
import { selectedCellsWithHeaders, type TargetCell } from "../lib/sheet";
import { addTag } from "../lib/tags";

interface Row extends TargetCell {
  result?: CheckedField;
}

/**
 * Prompt-driven extraction: the selected row's column headings become the
 * fields to find, the AI answers with quotes, and each answer is checked
 * against the PDF's text before it can be written as a proposed tag.
 */
export function ExtractPanel({ doc, onWritten }: { doc?: StoredDoc; onWritten: () => void }) {
  const [instructions, setInstructions] = useState("");
  const [rows, setRows] = useState<Row[]>([]);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const extractor = getExtractor();

  const run = async () => {
    if (!doc || !extractor) return;
    setError("");
    setBusy(true);
    try {
      const targets = await selectedCellsWithHeaders();
      if (targets.length === 0) throw new Error("The selected columns have no headings to look for.");
      setRows(targets);
      const pdf = await loadPdf(doc.bytes);
      const pages = await Promise.all(
        Array.from({ length: pdf.numPages }, (_, i) => pdf.getPage(i + 1).then(pageTextRuns)),
      );
      const answers = await extractor.extract(toBase64(doc.bytes), targets.map((t) => t.header), instructions || undefined);
      const checked = checkAgainstDocument(
        targets.map((t) => answers.find((a) => a.field === t.header) ?? { field: t.header, value: null, quote: null, page: null }),
        pages,
      );
      setRows(targets.map((t, i) => ({ ...t, result: checked[i] })));
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  const supported = rows.filter((r) => r.result?.status === "supported");

  const write = async () => {
    if (!doc) return;
    for (const r of supported) {
      const res = r.result!;
      if (res.status !== "supported") continue;
      await addTag(
        {
          tagId: crypto.randomUUID(),
          cell: r.address,
          docId: doc.docId,
          docName: doc.name,
          page: res.page,
          box: res.box,
          quotedText: res.quote!,
          method: "ai",
          status: "proposed",
          createdAt: new Date().toISOString(),
        },
        res.value!,
      );
    }
    setRows([]);
    onWritten();
  };

  return (
    <section>
      <h2>Extract with AI</h2>
      <p className="muted">Select the cells to fill in one row of your testing table. The column headings say what to find.</p>
      <textarea
        placeholder="Optional instructions, e.g. use the invoice date, not the ship date"
        value={instructions}
        onChange={(e) => setInstructions(e.target.value)}
        rows={2}
      />
      {!extractor && <p className="muted">AI isn't connected yet.</p>}
      <button onClick={run} disabled={!doc || !extractor || busy}>
        {busy ? "Reading document…" : "Extract from document"}
      </button>
      {error && <p className="error">{error}</p>}
      {rows.some((r) => r.result) && (
        <>
          <ul className="results">
            {rows.map((r) => (
              <li key={r.address} className={r.result?.status === "supported" ? "ok" : "warn"}>
                <strong>{r.header}</strong> {r.result?.value ?? "—"}
                <span className="muted"> · {r.result && OUTCOME_LABEL[r.result.status]}</span>
              </li>
            ))}
          </ul>
          <button onClick={write} disabled={supported.length === 0}>
            Write {supported.length} proposed value{supported.length === 1 ? "" : "s"}
          </button>
        </>
      )}
    </section>
  );
}
