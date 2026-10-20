import { useCallback, useEffect, useState, type ChangeEvent } from "react";
import { deleteDoc, listDocs, putDoc, type StoredDoc } from "../lib/docStore";
import { sha256Hex } from "../lib/hash";
import {
  addTag,
  historyForCell,
  latestForCell,
  nextUnreviewed,
  readTags,
  reviewTag,
  selectCell,
  summarise,
  type ReviewDecision,
  type ReviewSummary,
} from "../lib/tags";
import type { EvidenceTag } from "../lib/types";
import { EVIDENCE_ACCEPT, isSpreadsheet } from "../lib/spreadsheet";
import { openViewer } from "./dialog";
import { ExtractPanel } from "./ExtractPanel";
import { SamplePanel } from "./SamplePanel";
import { SearchPanel, type SheetResult } from "./SearchPanel";
import { SelectSamplePanel } from "./SelectSamplePanel";

async function activeCellAddress(): Promise<string> {
  return Excel.run(async (context) => {
    const cell = context.workbook.getActiveCell();
    cell.load("address");
    await context.sync();
    return cell.address;
  });
}

const REVIEWER_KEY = "audit-evidence.reviewer";

function loadReviewer(): string {
  try {
    return localStorage.getItem(REVIEWER_KEY) ?? "";
  } catch {
    return "";
  }
}

function saveReviewer(name: string) {
  try {
    localStorage.setItem(REVIEWER_KEY, name);
  } catch {
    // Storage can be blocked; the name then lasts for this session only.
  }
}

const DECISION_LABEL: Record<EvidenceTag["status"], string> = {
  proposed: "Proposed",
  verified: "Verified",
  rejected: "Rejected",
  exception: "Exception",
};

/** Where a tag's value came from: a PDF page or a spreadsheet cell. */
function sourceLabel(t: EvidenceTag): string {
  return t.sourceCell ? `${t.docName}, ${t.sourceCell}` : `${t.docName}, page ${t.page}`;
}

/** What a spreadsheet tag shows as its quote: the value and its row and column labels. */
const quoteFor = (text: string, context: string) => (context ? `${text} (${context})` : text);

const NO_BOX = { x: 0, y: 0, width: 0, height: 0 };

function formatWhen(iso: string): string {
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? iso : d.toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" });
}

export function App() {
  const [docs, setDocs] = useState<StoredDoc[]>([]);
  const [docId, setDocId] = useState("");
  const [cell, setCell] = useState("");
  const [tag, setTag] = useState<EvidenceTag | undefined>();
  const [history, setHistory] = useState<EvidenceTag[]>([]);
  const [summary, setSummary] = useState<ReviewSummary | undefined>();
  const [reviewer, setReviewer] = useState(loadReviewer);
  const [note, setNote] = useState("");
  const [message, setMessage] = useState("");

  const selectedDoc = docs.find((d) => d.docId === docId);
  const pdfs = docs.filter((d) => !isSpreadsheet(d.name));

  const refreshDocs = useCallback(async () => {
    const all = await listDocs();
    setDocs(all);
    setDocId((current) => current || all[0]?.docId || "");
  }, []);

  const refreshSelection = useCallback(async () => {
    const address = await activeCellAddress();
    const tags = await readTags();
    setCell(address);
    setTag(latestForCell(tags, address));
    setHistory(historyForCell(tags, address));
    setSummary(summarise(tags));
  }, []);

  useEffect(() => {
    refreshDocs();
    refreshSelection();
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
      refreshSelection();
    });
  }, [refreshDocs, refreshSelection]);

  // A note belongs to the cell it was typed for.
  useEffect(() => setNote(""), [cell]);

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

  const removeSelectedDoc = async () => {
    if (!docId) return;
    await deleteDoc(docId);
    setDocId("");
    refreshDocs();
  };

  const tagFromDocument = () => {
    const targetCell = cell;
    const doc = docs.find((d) => d.docId === docId);
    if (!doc || !targetCell) return;
    openViewer({ docId: doc.docId, mode: "tag" }, async (m, close) => {
      if (m.type !== "region-selected" && m.type !== "cell-selected") return;
      close();
      const base = {
        tagId: crypto.randomUUID(),
        cell: targetCell,
        docId: doc.docId,
        docName: doc.name,
        method: "manual" as const,
        status: "proposed" as const,
        createdAt: new Date().toISOString(),
      };
      const tag: EvidenceTag =
        m.type === "region-selected"
          ? { ...base, page: m.page, box: m.box, quotedText: m.text }
          : { ...base, page: 0, box: NO_BOX, sourceCell: m.sourceCell, quotedText: quoteFor(m.text, m.context) };
      await addTag(tag, m.text);
      setMessage(`Tagged ${targetCell} to ${sourceLabel(tag)}.`);
      refreshSelection();
    });
  };

  const showSheetResult = (r: SheetResult) => {
    if (selectedDoc) openViewer({ docId: selectedDoc.docId, mode: "trace", cell: `${r.sheet}!${r.address}` }, () => undefined);
  };

  const useSheetResult = async (r: SheetResult) => {
    if (!selectedDoc || !cell) return;
    const tag: EvidenceTag = {
      tagId: crypto.randomUUID(),
      cell,
      docId: selectedDoc.docId,
      docName: selectedDoc.name,
      page: 0,
      box: NO_BOX,
      sourceCell: `${r.sheet}!${r.address}`,
      quotedText: quoteFor(r.text, r.context),
      method: r.method,
      status: "proposed",
      createdAt: new Date().toISOString(),
    };
    try {
      await addTag(tag, r.text);
      setMessage(`Tagged ${cell} to ${sourceLabel(tag)}.`);
      await refreshSelection();
    } catch (e) {
      setMessage((e as Error).message);
    }
  };

  const signOff = async (decision: ReviewDecision) => {
    if (!tag) return;
    try {
      await reviewTag(tag, decision, reviewer, note);
      saveReviewer(reviewer.trim());
      setNote("");
      setMessage(`${DECISION_LABEL[decision]} ${cell}.`);
      await refreshSelection();
    } catch (e) {
      setMessage(e instanceof Error ? e.message : String(e));
    }
  };

  const goToNextUnreviewed = async () => {
    const next = nextUnreviewed(await readTags(), cell);
    if (!next) {
      setMessage("Every tagged cell has been signed off.");
      return;
    }
    await selectCell(next.cell);
    // The selection event refreshes the pane, but not every host fires it for programmatic selects.
    await refreshSelection();
  };

  const traceToSource = () => {
    if (!tag) return;
    openViewer(
      tag.sourceCell
        ? { docId: tag.docId, mode: "trace", cell: tag.sourceCell }
        : { docId: tag.docId, mode: "trace", page: tag.page, box: tag.box },
      () => undefined,
    );
  };

  return (
    <main className="pane">
      <section>
        <h2>Evidence</h2>
        <label className="upload">
          Add documents
          <input type="file" accept={EVIDENCE_ACCEPT} multiple onChange={onUpload} hidden />
        </label>
        {docs.length === 0 ? (
          <p className="muted">No documents yet.</p>
        ) : (
          <>
            <select value={docId} onChange={(e) => setDocId(e.target.value)}>
              {docs.map((d) => (
                <option key={d.docId} value={d.docId}>
                  {d.name}
                </option>
              ))}
            </select>
            <p className="muted">
              {docs.length} document{docs.length === 1 ? "" : "s"} added.{" "}
              <button className="link" onClick={removeSelectedDoc}>
                Remove selected
              </button>
            </p>
          </>
        )}
      </section>

      {selectedDoc && isSpreadsheet(selectedDoc.name) && (
        <SearchPanel doc={selectedDoc} cell={cell} onShow={showSheetResult} onUse={useSheetResult} />
      )}

      {selectedDoc && isSpreadsheet(selectedDoc.name) && (
        <SelectSamplePanel
          doc={selectedDoc}
          onWritten={(m) => {
            setMessage(m);
            refreshSelection();
          }}
        />
      )}

      <section>
        <h2>Selected cell</h2>
        <p className="cell">{cell || "…"}</p>
        {tag ? (
          <div className="tag-card">
            <p>
              <span className={`status ${tag.status}`}>{tag.status}</span> {tag.method === "ai" ? "AI · " : ""}{sourceLabel(tag)}
            </p>
            <blockquote>{tag.quotedText}</blockquote>
            {tag.reviewedBy && tag.reviewedAt && (
              <p className="muted">
                {DECISION_LABEL[tag.status]} by {tag.reviewedBy} on {formatWhen(tag.reviewedAt)}
                {tag.note ? `: ${tag.note}` : ""}
              </p>
            )}
            <button onClick={traceToSource}>Open source</button>
            <button onClick={tagFromDocument} disabled={!docId}>
              Re-tag
            </button>

            <div className="review">
              <label>
                Reviewer
                <input value={reviewer} onChange={(e) => setReviewer(e.target.value)} placeholder="Name or initials" />
              </label>
              <textarea
                rows={2}
                value={note}
                onChange={(e) => setNote(e.target.value)}
                placeholder="Note (required to reject or mark an exception)"
              />
              <button onClick={() => signOff("verified")} disabled={!reviewer.trim()}>
                Verify
              </button>
              <button onClick={() => signOff("rejected")} disabled={!reviewer.trim() || !note.trim()}>
                Reject
              </button>
              <button onClick={() => signOff("exception")} disabled={!reviewer.trim() || !note.trim()}>
                Mark exception
              </button>
            </div>

            {history.length > 1 && (
              <details>
                <summary>History ({history.length})</summary>
                <ol className="history">
                  {history.map((h, i) => (
                    <li key={i}>
                      {DECISION_LABEL[h.status]}
                      {h.reviewedBy && h.reviewedAt
                        ? ` by ${h.reviewedBy}, ${formatWhen(h.reviewedAt)}`
                        : `, ${h.method === "ai" ? "AI" : "tagged"} ${formatWhen(h.createdAt)}`}
                      {h.note ? `: ${h.note}` : ""}
                    </li>
                  ))}
                </ol>
              </details>
            )}
          </div>
        ) : (
          <button onClick={tagFromDocument} disabled={!docId || !cell}>
            Tag from document
          </button>
        )}
      </section>

      {summary && summary.total > 0 && (
        <section>
          <h2>Review</h2>
          <p>
            {summary.total - summary.awaiting} of {summary.total} tagged cells signed off
          </p>
          <p className="muted">
            {summary.verified} verified · {summary.exception} exception{summary.exception === 1 ? "" : "s"} · {summary.rejected} rejected · {summary.awaiting} to review
          </p>
          <button onClick={goToNextUnreviewed} disabled={summary.awaiting === 0}>
            Next to review
          </button>
        </section>
      )}

      <SamplePanel
        docs={pdfs}
        onWritten={(m) => {
          setMessage(m);
          refreshSelection();
        }}
      />

      <ExtractPanel
        doc={selectedDoc && !isSpreadsheet(selectedDoc.name) ? selectedDoc : undefined}
        onWritten={() => {
          setMessage("Wrote AI values as proposed. Review each one with Open source.");
          refreshSelection();
        }}
      />

      {message && <p className="muted">{message}</p>}
    </main>
  );
}
