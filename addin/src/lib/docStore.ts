// Phase 0 keeps source documents in the browser's IndexedDB, shared by the task
// pane and the viewer dialog (same origin). Phase 1 replaces this with the
// backend document store.

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
