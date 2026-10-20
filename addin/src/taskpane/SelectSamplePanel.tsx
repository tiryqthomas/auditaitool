import { useEffect, useMemo, useState } from "react";
import type { StoredDoc } from "../lib/docStore";
import { writeSampleSheet } from "../lib/sampleSheet";
import {
  guessAmountHeader,
  listingHeaders,
  newSeed,
  readPopulation,
  selectSample,
  type SampleMethod,
  type SampleSettings,
} from "../lib/sampling";
import { asAmount, loadWorkbook, type SourceWorkbook } from "../lib/spreadsheet";

const money = (n: number) => n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function guessKey(headers: string[]): string {
  return headers.find((h) => /invoice|inv\b|#|number|no\.?$|ref|id\b/i.test(h)) ?? headers[0] ?? "";
}

/**
 * Pulls a sample from a client listing (e.g. a sales ledger) and writes it to a
 * new Sample sheet ready for Fill sample. Selection is reproducible from its seed.
 */
export function SelectSamplePanel({ doc, onWritten }: { doc: StoredDoc; onWritten: (message: string) => void }) {
  const [wb, setWb] = useState<SourceWorkbook | null>(null);
  const [sheetName, setSheetName] = useState("");
  const [keyHeader, setKeyHeader] = useState("");
  const [amountHeader, setAmountHeader] = useState("");
  const [method, setMethod] = useState<SampleMethod>("random");
  const [size, setSize] = useState("5");
  const [threshold, setThreshold] = useState("");
  const [seed, setSeed] = useState(newSeed);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    setWb(null);
    loadWorkbook(doc.bytes, doc.name).then((book) => {
      setWb(book);
      setSheetName(book[0]?.name ?? "");
    }, (e: Error) => setError(e.message));
  }, [doc]);

  const sheet = wb?.find((s) => s.name === sheetName);
  const headers = useMemo(() => (sheet ? listingHeaders(sheet).filter(Boolean) : []), [sheet]);

  useEffect(() => {
    setKeyHeader(guessKey(headers));
    setAmountHeader(guessAmountHeader(headers));
  }, [headers]);

  // Recomputed on every change so the auditor sees exactly what will be written.
  const preview = useMemo(() => {
    if (!sheet || !keyHeader) return null;
    try {
      const pop = readPopulation(sheet, keyHeader, amountHeader);
      const t = threshold.trim() ? asAmount(threshold) : undefined;
      if (t === null) throw new Error("The key item amount isn't a number.");
      if (method === "mus" && !amountHeader) throw new Error("Monetary unit sampling needs an amount column.");
      const settings: SampleSettings = { method, size: Number(size), seed, keyItemThreshold: t };
      return { pop, settings, selection: selectSample(pop, settings), error: "" };
    } catch (e) {
      return { error: (e as Error).message };
    }
  }, [sheet, keyHeader, amountHeader, method, size, threshold, seed]);

  const write = async () => {
    if (!preview || !("selection" in preview) || !preview.selection) return;
    setBusy(true);
    setError("");
    try {
      const { sheetName: written, memoName } = await writeSampleSheet(preview.pop!, preview.selection, preview.settings!, {
        docId: doc.docId,
        docName: doc.name,
      });
      onWritten(
        `Wrote ${preview.selection.items.length} items to "${written}" (how they were picked is on "${memoName}"). Add the support and use Fill whole sample to test them.`,
      );
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setBusy(false);
    }
  };

  if (!wb) return null;
  const sel = preview && "selection" in preview ? preview.selection : undefined;

  return (
    <section>
      <h2>Pull a sample</h2>
      <p className="muted">Pick items from this listing onto a new Sample sheet, then test them with Fill whole sample.</p>
      {wb.length > 1 && (
        <label className="field">
          Sheet
          <select value={sheetName} onChange={(e) => setSheetName(e.target.value)}>
            {wb.map((s) => (
              <option key={s.name}>{s.name}</option>
            ))}
          </select>
        </label>
      )}
      <label className="field">
        Item ID column
        <select value={keyHeader} onChange={(e) => setKeyHeader(e.target.value)}>
          {headers.map((h) => (
            <option key={h}>{h}</option>
          ))}
        </select>
      </label>
      <label className="field">
        Amount column
        <select value={amountHeader} onChange={(e) => setAmountHeader(e.target.value)}>
          <option value="">(none)</option>
          {headers.map((h) => (
            <option key={h}>{h}</option>
          ))}
        </select>
      </label>
      <label className="field">
        Method
        <select value={method} onChange={(e) => setMethod(e.target.value as SampleMethod)}>
          <option value="random">Random</option>
          <option value="mus">Monetary unit (bigger amounts more likely)</option>
        </select>
      </label>
      <div className="inline-fields">
        <label className="field">
          Sample size
          <input type="number" min={0} value={size} onChange={(e) => setSize(e.target.value)} />
        </label>
        <label className="field">
          Always test items over
          <input value={threshold} onChange={(e) => setThreshold(e.target.value)} placeholder="optional" />
        </label>
        <label className="field">
          Seed
          <input type="number" value={seed} onChange={(e) => setSeed(Number(e.target.value) || 0)} />
        </label>
      </div>
      <button className="link" onClick={() => setSeed(newSeed())}>
        New seed
      </button>

      {preview?.error && <p className="warn">{preview.error}</p>}
      {sel && preview && "pop" in preview && preview.pop && (
        <>
          <p className="muted">
            {preview.pop.items.length} items totalling {money(sel.populationTotal)}
            {preview.pop.skipped.length ? `, ${preview.pop.skipped.length} row${preview.pop.skipped.length === 1 ? "" : "s"} left out (${preview.pop.skipped.map((s) => s.reason).join(", ")})` : ""}.
            {sel.interval !== undefined ? ` Interval ${money(sel.interval)}.` : ""}
            {sel.notSelectable ? ` ${sel.notSelectable} zero or negative items can't be picked by monetary unit; test them separately.` : ""}
          </p>
          <p>
            {sel.items.length} selected{sel.keyItems ? `, including ${sel.keyItems} key items` : ""}:
          </p>
          <ul className="hits">
            {sel.items.map((i) => (
              <li key={i.row}>
                {i.key}
                {i.amount !== null ? ` · ${money(i.amount)}` : ""} <span className="muted">({i.reason})</span>
              </li>
            ))}
          </ul>
          <button onClick={write} disabled={busy || sel.items.length === 0}>
            {busy ? "Writing…" : "Write Sample sheet"}
          </button>
        </>
      )}
      {error && <p className="warn">{error}</p>}
    </section>
  );
}
