/** A rectangle in PDF page units at scale 1, origin at the top-left of the page. */
export interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}

export type TagStatus = "proposed" | "verified" | "rejected" | "exception";

/** Links one workbook cell to the exact place in a source document it came from. */
export interface EvidenceTag {
  tagId: string;
  cell: string; // e.g. Sheet1!B4
  docId: string; // SHA-256 of the original file
  docName: string;
  page: number; // 1-based
  box: Box;
  quotedText: string;
  method: "manual" | "ai";
  status: TagStatus;
  createdAt: string; // ISO timestamp
}

/** Messages the viewer dialog sends back to the task pane. */
export type ViewerMessage =
  | { type: "region-selected"; docId: string; page: number; box: Box; text: string }
  | { type: "closed" };
