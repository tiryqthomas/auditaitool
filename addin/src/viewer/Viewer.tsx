import { useEffect, useRef, useState, type MouseEvent } from "react";
import { getDoc } from "../lib/docStore";
import { boxFromPoints, scaleBox, textInBox } from "../lib/geometry";
import { loadPdf, pageTextRuns, type PdfDoc } from "../lib/pdf";
import type { Box, ViewerMessage } from "../lib/types";

const SCALE = 1.5;

function send(message: ViewerMessage) {
  if (Office.context?.ui?.messageParent) Office.context.ui.messageParent(JSON.stringify(message));
  else console.info("viewer message", message);
}

function boxFromParams(p: URLSearchParams): Box | null {
  const [x, y, width, height] = ["x", "y", "w", "h"].map((k) => Number(p.get(k)));
  return [x, y, width, height].every(Number.isFinite) && width > 0 ? { x, y, width, height } : null;
}

/**
 * Source document viewer, opened by the task pane as an Office dialog.
 * mode=tag: the auditor drags a box around a value to link it to a cell.
 * mode=trace: opens at a tag's page with its region highlighted.
 */
export function Viewer({ params }: { params: URLSearchParams }) {
  const docId = params.get("docId") ?? "";
  const mode = params.get("mode") === "trace" ? "trace" : "tag";
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [pdf, setPdf] = useState<PdfDoc | null>(null);
  const [name, setName] = useState("");
  const [pageNo, setPageNo] = useState(Number(params.get("page")) || 1);
  const [highlight, setHighlight] = useState<Box | null>(mode === "trace" ? boxFromParams(params) : null);
  const [dragStart, setDragStart] = useState<{ x: number; y: number } | null>(null);
  const [selection, setSelection] = useState<{ box: Box; text: string } | null>(null);
  const [error, setError] = useState("");

  useEffect(() => {
    getDoc(docId)
      .then((doc) => {
        if (!doc) throw new Error("This document isn't stored on this computer. Add it again from the task pane.");
        setName(doc.name);
        return loadPdf(doc.bytes);
      })
      .then(setPdf)
      .catch((e: Error) => setError(e.message));
  }, [docId]);

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
            ref={(el) => mode === "trace" && el?.scrollIntoView({ block: "center" })}
          />
        )}
      </div>
    </div>
  );
}
