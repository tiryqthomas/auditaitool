import { getDoc, toBase64 } from "../lib/docStore";
import type { Box, PaneMessage, ViewerMessage } from "../lib/types";

export interface ViewerOptions {
  docId: string;
  mode: "tag" | "trace";
  page?: number;
  box?: Box;
  cell?: string; // spreadsheet evidence: the source cell to highlight
}

/** Opens the source viewer in an Office dialog and relays its messages. */
export function openViewer(opts: ViewerOptions, onMessage: (m: ViewerMessage, close: () => void) => void): void {
  const q = new URLSearchParams({ docId: opts.docId, mode: opts.mode });
  if (opts.page) q.set("page", String(opts.page));
  if (opts.cell) q.set("cell", opts.cell);
  if (opts.box) {
    q.set("x", String(opts.box.x));
    q.set("y", String(opts.box.y));
    q.set("w", String(opts.box.width));
    q.set("h", String(opts.box.height));
  }
  // Relative to the task pane so it works under any hosting path.
  const url = new URL(`viewer.html?${q}`, location.href).toString();
  // displayInIframe keeps the viewer inside the Excel page on the web, so it
  // shares the task pane's storage. Desktop Excel ignores it.
  Office.context.ui.displayDialogAsync(url, { height: 85, width: 60, displayInIframe: true }, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Could not open viewer", result.error);
      return;
    }
    const dialog = result.value;
    const close = () => dialog.close();
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, async (arg) => {
      if (!("message" in arg)) return;
      const message = JSON.parse(arg.message) as ViewerMessage;
      if (message.type === "need-document") {
        const doc = await getDoc(message.docId);
        const reply: PaneMessage = doc
          ? { type: "document", docId: doc.docId, name: doc.name, base64: toBase64(doc.bytes) }
          : { type: "document-missing", docId: message.docId };
        dialog.messageChild(JSON.stringify(reply));
        return;
      }
      onMessage(message, close);
    });
    dialog.addEventHandler(Office.EventType.DialogEventReceived, () => onMessage({ type: "closed" }, close));
  });
}
