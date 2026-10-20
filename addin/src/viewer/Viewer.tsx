import { useEffect, useRef, useState, type MouseEvent } from "react";
import { fromBase64, getDoc, type StoredDoc } from "../lib/docStore";
import { boxFromPoints, scaleBox, textInBox } from "../lib/geometry";
import { loadPdf, pageTextRuns, type PdfDoc } from "../lib/pdf";
import { isSpreadsheet } from "../lib/spreadsheet";
import type { Box, PaneMessage, ViewerMessage } from "../lib/types";
import { SheetViewer } from "./SheetViewer";

const SCALE = 1.5;

export function send(message: ViewerMessage) {
  if (Office.context?.ui?.messageParent) Office.context.ui.messageParent(JSON.stringify(message));
  else console.info("viewer message", message);
}

const MISSING = "This document isn't stored on this computer. Add it again from the task pane.";

/**
 * Loads a document from local storage, or asks the task pane for it when this
 * dialog can't see the task pane's storage (Excel on the web partitions it).
 */
async function loadDocument(docId: string): Promise<StoredDoc> {
  const local = await getDoc(docId).catch(() => undefined);
  if (local) return local;
  if (!Office.context?.ui?.addHandlerAsync) throw new Error(MISSING);
  return new Promise((resolve, reject) => {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
      const m = JSON.parse((arg as { message: string }).message) as PaneMessage;
      if (m.docId !== docId) return;
      if (m.type === "document") resolve({ docId, name: m.name, addedAt: "", bytes: fromBase64(m.base64) });
      else reject(new Error(MISSING));
    });
    send({ type: "need-document", docId });
  });
}

function boxFromParams(p: URLSearchParams): Box | null {
  const [x, y, width, height] = ["x", "y", "w", "h"].map((k) => Number(p.get(k)));
  return [x, y, width, height].every(Number.isFinite) && width > 0 ? { x, y, width, height } : null;
}

/**
 * Source document viewer, opened by the task pane as an Office dialog.
 * mode=tag: the auditor picks a value to link to a cell.
 * mode=trace: opens at a tag's source with it highlighted.
 */
export function Viewer({ params }: { params: URLSearchParams }) {
  const docId = params.get("docId") ?? "";
  const mode = params.get("mode") === "trace" ? "trace" : "tag";
  const [doc, setDoc] = useState<StoredDoc | null>(null);
  const [error, setError] = useState("");

  useEffect(() => {
    loadDocument(docId).then(setDoc, (e: Error) => setError(e.message));
  }, [docId]);

  if (error) return <p className="error">{error}</p>;
  if (!doc) return <p className="muted">Loading…</p>;
  return isSpreadsheet(doc.name) ? (
    <SheetViewer doc={doc} mode={mode} sourceCell={params.get("cell") ?? undefined} />
  ) : (
    <PdfViewer doc={doc} mode={mode} params={params} />
  );
}

/** mode=tag: drag a box around a value. mode=trace: opens at the tag's page with its region highlighted. */
function PdfViewer({ doc, mode, params }: { doc: StoredDoc; mode: "tag" | "trace"; params: URLSearchParams }) {
  const docId = doc.docId;
  const name = doc.name;
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [pdf, setPdf] = useState<PdfDoc | null>(null);
  const [pageNo, setPageNo] = useState(Number(params.get("page")) || 1);
  const [highlight, setHighlight] = useState<Box | null>(mode === "trace" ? boxFromParams(params) : null);
  const [dragStart, setDragStart] = useState<{ x: number; y: number } | null>(null);
  const [selection, setSelection] = useState<{ box: Box; text: string } | null>(null);
  const [error, setError] = useState("");

  useEffect(() => {
    loadPdf(doc.bytes).then(setPdf, (e: Error) => setError(e.message));
  }, [doc]);

  useEffect(() => {
    if (!pdf || !canvasRef.current) return;
    let cancelled = false;
    pdf.getPage(pageNo).then((page) => {
      if (cancelled) return;
      const viewport = page.getViewport({ scale: SCALE });
      const canvas = canvasRef.current!;
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      page.render({ canvasContext: canvas.getContext("2d")!, viewport });
    });
    return () => {
      cancelled = true;
    };
  }, [pdf, pageNo]);

  const toPage = (e: MouseEvent) => {
    const rect = e.currentTarget.getBoundingClientRect();
    return { x: (e.clientX - rect.left) / SCALE, y: (e.clientY - rect.top) / SCALE };
  };

  const onMouseDown = (e: MouseEvent) => {
    if (mode !== "tag") return;
    setDragStart(toPage(e));
    setSelection(null);
  };
  const onMouseMove = (e: MouseEvent) => {
    if (!dragStart) return;
    const p = toPage(e);
    setHighlight(boxFromPoints(dragStart.x, dragStart.y, p.x, p.y));
  };
  const onMouseUp = async () => {
    if (!dragStart || !highlight || !pdf) return;
    setDragStart(null);
    const runs = await pageTextRuns(await pdf.getPage(pageNo));
    setSelection({ box: highlight, text: textInBox(runs, highlight) });
  };

  if (error) return <p className="error">{error}</p>;
  const shown = highlight && scaleBox(highlight, SCALE);

  return (
    <div className="viewer">
      <header className="viewer-bar">
        <strong>{name}</strong>
        <button disabled={pageNo <= 1} onClick={() => setPageNo(pageNo - 1)}>‹</button>
        <span>
          Page {pageNo} of {pdf?.numPages ?? "…"}
        </span>
        <button disabled={!pdf || pageNo >= pdf.numPages} onClick={() => setPageNo(pageNo + 1)}>›</button>
        {mode === "tag" && <span className="hint">Drag a box around the value to tag it.</span>}
      </header>
      {selection && (
        <div className="selection">
          <code>{selection.text || "(no text found in this box; scanned pages need OCR, coming in Phase 2)"}</code>
          <button
            disabled={!selection.text}
            onClick={() => send({ type: "region-selected", docId, page: pageNo, box: selection.box, text: selection.text })}
          >
            Use this value
          </button>
        </div>
      )}
      <div
        className={`page ${mode === "tag" ? "taggable" : ""}`}
        onMouseDown={onMouseDown}
        onMouseMove={onMouseMove}
        onMouseUp={onMouseUp}
      >
        <canvas ref={canvasRef} />
        {shown && (
          <div
            className={`highlight ${mode}`}
            style={{ left: shown.x, top: shown.y, width: shown.width, height: shown.height }}
            // Scroll only this window: scrollIntoView also scrolls the Excel page
            // hosting the dialog, which pushed its close button off screen.
            ref={(el) => {
              if (mode === "trace" && el) window.scrollTo({ top: el.getBoundingClientRect().top + window.scrollY - window.innerHeight / 2 });
            }}
          />
        )}
      </div>
    </div>
  );
}
