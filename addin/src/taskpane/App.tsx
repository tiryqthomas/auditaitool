import { useCallback, useEffect, useState, type ChangeEvent } from "react";
import { listDocs, putDoc, type StoredDoc } from "../lib/docStore";
import { sha256Hex } from "../lib/hash";
import { addTag, findTag } from "../lib/tags";
import type { EvidenceTag } from "../lib/types";
import { openViewer } from "./dialog";

async function activeCellAddress(): Promise<string> {
  return Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.load("address");
    await context.sync();
    return cell.address;
  });
}

export function App() {
  const [docs, setDocs] = useState<StoredDoc[]>([]);
  const [docId, setDocId] = useState("");
  const [cell, setCell] = useState("");
  const [tag, setTag] = useState<EvidenceTag | undefined>();
  const [message, setMessage] = useState("");

  const refreshDocs = useCallback(async () => {
    const all = await listDocs();
    setDocs(all);
    setDocId((current) => current || all[0]?.docId || "");
  }, []);

  const refreshSelection = useCallback(async () => {
    const address = await activeCellAddress();
    setCell(address);
    setTag(await findTag(address));
  }, []);

  useEffect(() => {
    refreshDocs();
    refreshSelection();
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
      refreshSelection();
    });
  }, [refreshDocs, refreshSelection]);

  const onUpload = async (e: ChangeEvent<HTMLInputElement>) => {
    for (const file of Array.from(e.target.files ?? [])) {
      const bytes = await file.arrayBuffer();
      const id = await sha256Hex(bytes);
      await putDoc({ docId: id, name: file.name, addedAt: new Date().toISOString(), bytes });
      setDocId(id);
    }
    e.target.value = "";
    refreshDocs();
  };

  const tagFromDocument = () => {
    const targetCell = cell;
    const doc = docs.find((d) => d.docId === docId);
    if (!doc || !targetCell) return;
    openViewer({ docId: doc.docId, mode: "tag" }, async (m, close) => {
      if (m.type !== "region-selected") return;
      close();
      await addTag(
        {
          tagId: crypto.randomUUID(),
          cell: targetCell,
          docId: doc.docId,
          docName: doc.name,
          page: m.page,
          box: m.box,
          quotedText: m.text,
          method: "manual",
          status: "proposed",
          createdAt: new Date().toISOString(),
        },
        m.text,
      );
      setMessage(`Tagged ${targetCell} to ${doc.name}, page ${m.page}.`);
      refreshSelection();
    });
  };

  const traceToSource = () => {
    if (!tag) return;
    openViewer({ docId: tag.docId, mode: "trace", page: tag.page, box: tag.box }, () => undefined);
  };

  return (
    <main className="pane">
      <section>
        <h2>Evidence</h2>
        <label className="upload">
          Add PDF
          <input type="file" accept="application/pdf" multiple onChange={onUpload} hidden />
        </label>
        {docs.length === 0 ? (
          <p className="muted">No documents yet.</p>
        ) : (
          <select value={docId} onChange={(e) => setDocId(e.target.value)}>
            {docs.map((d) => (
              <option key={d.docId} value={d.docId}>
                {d.name}
              </option>
            ))}
          </select>
        )}
      </section>

      <section>
        <h2>Selected cell</h2>
        <p className="cell">{cell || "…"}</p>
        {tag ? (
          <div className="tag-card">
            <p>
              <span className={`status ${tag.status}`}>{tag.status}</span> {tag.docName}, page {tag.page}
            </p>
            <blockquote>{tag.quotedText}</blockquote>
            <button onClick={traceToSource}>Open source</button>
            <button onClick={tagFromDocument} disabled={!docId}>
              Re-tag
            </button>
          </div>
        ) : (
          <button onClick={tagFromDocument} disabled={!docId || !cell}>
            Tag from document
          </button>
        )}
      </section>

      {message && <p className="muted">{message}</p>}
    </main>
  );
}
