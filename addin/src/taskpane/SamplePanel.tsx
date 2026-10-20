import { useState } from "react";
import { getExtractor } from "../lib/ai";
import { toBase64, type StoredDoc } from "../lib/docStore";
import { checkAgainstDocument } from "../lib/extraction";
import { loadPdf, pageTextRuns } from "../lib/pdf";
import { planSample, type CellOutcome, type DocResult, type SamplePlan } from "../lib/sample";
import { readSampleTable } from "../lib/sheet";
import { addTags } from "../lib/tags";
import type { EvidenceTag } from "../lib/types";

const PARALLEL = 3;

/** Runs `fn` over items with at most `limit` in flight, in input order. */
async function mapLimited<T, R>(items: T[], limit: number, fn: (item: T) => Promise<R>): Promise<R[]> {
  const out: R[] = new Array(items.length);
  let next = 0;
  await Promise.all(
    Array.from({ length: Math.min(limit, items.length) }, async () => {
      while (next < items.length) {
        const i = next++;
        out[i] = await fn(items[i]);
      }
    }),
  );
  return out;
}

function guessKeyHeader(headers: string[]): string {
  return headers.find((h) => /invoice|inv\b|#|number|no\.?$|ref/i.test(h)) ?? headers[0] ?? "";
}

const summary = (outcomes: CellOutcome[], kind: CellOutcome["kind"]) => outcomes.filter((o) => o.kind === kind).length;

/**
 * Whole-sample testing: every added document is read once, matched to the
 * sample row with the same key (e.g. invoice number), and each cell is filled,
 * confirmed, or flagged as an exception. Nothing is written until the auditor
 * reviews the summary.
 */
export function SamplePanel({ docs, onWritten }: { docs: StoredDoc[]; onWritten: (message: string) => void }) {
  const [headers, setHeaders] = useState<string[]>([]);
  const [keyHeader, setKeyHeader] = useState("");
  const [progress, setProgress] = useState("");
  const [plan, setPlan] = useState<SamplePlan | null>(null);
  const [failed, setFailed] = useState<string[]>([]);
  const [error, setError] = useState("");
  const extractor = getExtractor();

  const loadHeaders = async () => {
    try {
      const table = await readSampleTable();
      setHeaders(table.headers);
      setKeyHeader((k) => (k && table.headers.includes(k) ? k : guessKeyHeader(table.headers)));
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const run = async () => {
    if (!extractor) return;
    setError("");
    setPlan(null);
    setFailed([]);
    try {
      const table = await readSampleTable();
      setHeaders(table.headers);
      const key = table.headers.includes(keyHeader) ? keyHeader : guessKeyHeader(table.headers);
      setKeyHeader(key);
      const rows = table.rows
        .map((r) => ({ ...r, key: r.cells[key]?.text.trim() ?? "" }))
        .filter((r) => r.key);
      if (rows.length === 0) throw new Error(`No sample rows have a value in the "${key}" column.`);

      let done = 0;
      const failures: string[] = [];
      setProgress(`Reading 0 of ${docs.length} documents…`);
      const results = await mapLimited(docs, PARALLEL, async (doc): Promise<DocResult | null> => {
        try {
          const pdf = await loadPdf(doc.bytes);
          const pages = await Promise.all(
            Array.from({ length: pdf.numPages }, (_, i) => pdf.getPage(i + 1).then(pageTextRuns)),
          );
          const answers = await extractor.extract(toBase64(doc.bytes), table.headers);
          const fields = checkAgainstDocument(
            table.headers.map((h) => answers.find((a) => a.field === h) ?? { field: h, value: null, quote: null, page: null }),
            pages,
          );
          return { docId: doc.docId, docName: doc.name, fields };
        } catch (e) {
          failures.push(`${doc.name}: ${(e as Error).message}`);
          return null;
        } finally {
          setProgress(`Reading ${++done} of ${docs.length} documents…`);
        }
      });
      setFailed(failures);
      setPlan(planSample(rows, key, results.filter((r): r is DocResult => r !== null)));
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setProgress("");
    }
  };

  const write = async () => {
    if (!plan) return;
    const now = new Date().toISOString();
    const writes: { tag: EvidenceTag; value?: string }[] = [];
    for (const r of plan.rows) {
      if (!r.doc) continue;
      for (const o of r.outcomes) {
        if (o.kind === "unsupported") continue;
        writes.push({
          tag: {
            tagId: crypto.randomUUID(),
            cell: o.cell.address,
            docId: r.doc.docId,
            docName: r.doc.docName,
            page: o.field.page,
            box: o.field.box,
            quotedText: o.field.quote!,
            method: "ai",
            status: o.kind === "exception" ? "exception" : "proposed",
            createdAt: now,
          },
          // Only empty cells get the document's value; existing workpaper values are never overwritten.
          value: o.kind === "fill" ? o.field.value! : undefined,
        });
      }
    }
    try {
      await addTags(writes);
      const exceptions = writes.filter((w) => w.tag.status === "exception").length;
      setPlan(null);
      onWritten(
        `Tagged ${writes.length} cells.` +
          (exceptions ? ` ${exceptions} exception${exceptions === 1 ? "" : "s"} marked red for your review.` : ""),
      );
    } catch (e) {
      setError((e as Error).message);
    }
  };

  const all = plan?.rows.flatMap((r) => r.outcomes) ?? [];

  return (
    <section>
      <h2>Fill whole sample</h2>
      <p className="muted">
        Add every invoice for the sample, then fill the table on this sheet. Each row is matched to its document by the
        column below.
      </p>
      <label className="field">
        Match rows by
        <select value={keyHeader} onFocus={loadHeaders} onChange={(e) => setKeyHeader(e.target.value)}>
          {(headers.length ? headers : keyHeader ? [keyHeader] : []).map((h) => (
            <option key={h}>{h}</option>
          ))}
        </select>
      </label>
      <button onClick={run} disabled={!extractor || docs.length === 0 || !!progress}>
        {progress || `Fill sample from ${docs.length} document${docs.length === 1 ? "" : "s"}`}
      </button>
      {error && <p className="error">{error}</p>}
      {plan && (
        <>
          <p>
            {summary(all, "fill")} to fill · {summary(all, "agree")} agree ·{" "}
            <span className={summary(all, "exception") ? "bad" : ""}>{summary(all, "exception")} exceptions</span> ·{" "}
            {summary(all, "unsupported")} need manual review
          </p>
          <ul className="results">
            {plan.rows.map((r) => (
              <li key={r.row.rowIndex} className={r.problem || r.outcomes.some((o) => o.kind !== "fill" && o.kind !== "agree") ? "warn" : "ok"}>
                <strong>{r.row.key}</strong>{" "}
                {r.problem === "no-document" && "No matching document"}
                {r.problem === "several-documents" && "More than one document matches; tag this row by hand"}
                {r.doc && (
                  <>
                    {r.doc.docName}
                    {r.outcomes
                      .filter((o) => o.kind === "exception" || o.kind === "unsupported")
                      .map((o) => (
                        <span key={o.header} className="muted">
                          {" "}
                          · {o.header}: {o.kind === "exception" ? `sheet ${o.cell.text}, document ${o.field.value}` : "not verified"}
                        </span>
                      ))}
                  </>
                )}
              </li>
            ))}
          </ul>
          {plan.unmatchedDocs.length > 0 && (
            <p className="muted">Not in the sample: {plan.unmatchedDocs.map((d) => d.docName).join(", ")}</p>
          )}
          <button onClick={write} disabled={all.every((o) => o.kind === "unsupported")}>
            Write results
          </button>
        </>
      )}
      {failed.length > 0 && <p className="error">Couldn't read: {failed.join("; ")}</p>}
    </section>
  );
}
