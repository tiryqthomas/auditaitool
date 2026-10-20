import type { Extractor } from "./extraction";

// The connection to Claude is configured in Phase 1 once hosting is chosen.
// Until then the task pane shows the extraction panel without running it.
export function getExtractor(): Extractor | null {
  return null;
}
