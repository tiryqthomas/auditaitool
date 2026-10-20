// Phase 0 keeps source documents in the task pane's IndexedDB. The viewer
// dialog may not share that storage (browsers partition storage for add-ins
// framed inside Excel on the web), so it falls back to asking the task pane
// for the bytes. Phase 1 replaces this with the backend document store.

export interface StoredDoc {
  docId: string;
  name: string;
  addedAt: string;
  bytes: ArrayBuffer;
}

const DB_NAME = "audit-evidence";
const STORE = "documents";

function open(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE, { keyPath: "docId" });
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function run<T>(mode: IDBTransactionMode, fn: (s: IDBObjectStore) => IDBRequest<T>): Promise<T> {
  const db = await open();
  return new Promise((resolve, reject) => {
    const req = fn(db.transaction(STORE, mode).objectStore(STORE));
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export const putDoc = (doc: StoredDoc) => run("readwrite", (s) => s.put(doc)).then(() => undefined);
export const getDoc = (docId: string) => run<StoredDoc | undefined>("readonly", (s) => s.get(docId));
export const listDocs = () => run<StoredDoc[]>("readonly", (s) => s.getAll());

export function toBase64(bytes: ArrayBuffer): string {
  const u8 = new Uint8Array(bytes);
  let binary = "";
  for (let i = 0; i < u8.length; i += 0x8000) binary += String.fromCharCode(...u8.subarray(i, i + 0x8000));
  return btoa(binary);
}

export function fromBase64(base64: string): ArrayBuffer {
  return Uint8Array.from(atob(base64), (c) => c.charCodeAt(0)).buffer;
}

export const deleteDoc = (docId: string) => run("readwrite", (s) => s.delete(docId)).then(() => undefined);
