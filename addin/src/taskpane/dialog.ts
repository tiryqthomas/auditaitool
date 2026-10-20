import type { Box, ViewerMessage } from "../lib/types";

export interface ViewerOptions {
  docId: string;
  mode: "tag" | "trace";
  page?: number;
  box?: Box;
}

/** Opens the source viewer in an Office dialog and relays its messages. */
export function openViewer(opts: ViewerOptions, onMessage: (m: ViewerMessage, close: () => void) => void): void {
  const q = new URLSearchParams({ docId: opts.docId, mode: opts.mode });
  if (opts.page) q.set("page", String(opts.page));
  if (opts.box) {
    q.set("x", String(opts.box.x));
    q.set("y", String(opts.box.y));
    q.set("w", String(opts.box.width));
    q.set("h", String(opts.box.height));
  }
  // Relative to the task pane so it works under any hosting path.
  const url = new URL(`viewer.html?${q}`, location.href).toString();
  Office.context.ui.displayDialogAsync(url, { height: 85, width: 60, displayInIframe: false }, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Could not open viewer", result.error);
      return;
    }
    const dialog = result.value;
    const close = () => dialog.close();
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
      if ("message" in arg) onMessage(JSON.parse(arg.message) as ViewerMessage, close);
    });
    dialog.addEventHandler(Office.EventType.DialogEventReceived, () => onMessage({ type: "closed" }, close));
  });
}
