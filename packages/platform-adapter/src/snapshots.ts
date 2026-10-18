import type { DocumentSnapshot, DocumentSnapshotMeta } from '@mini-office/shared';

const DB_NAME = 'mini-office-store';
const DB_VERSION = 1;
const STORE = 'snapshots';
const MAX_SNAPSHOTS = 12;
const MAX_TOTAL_BYTES = 160 * 1024 * 1024;

let dbPromise: Promise<IDBDatabase | null> | null = null;

function openDb(): Promise<IDBDatabase | null> {
  if (typeof indexedDB === 'undefined') return Promise.resolve(null);
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve) => {
    try {
      const request = indexedDB.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(STORE)) {
          const store = db.createObjectStore(STORE, { keyPath: 'id' });
          store.createIndex('updatedAt', 'updatedAt');
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => resolve(null);
    } catch {
      resolve(null);
    }
  });
  return dbPromise;
}

function requestToPromise<T>(request: IDBRequest<T>): Promise<T | null> {
  return new Promise((resolve) => {
    request.onsuccess = () => resolve(request.result ?? null);
    request.onerror = () => resolve(null);
  });
}

function transactionDone(tx: IDBTransaction): Promise<void> {
  return new Promise((resolve) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => resolve();
    tx.onabort = () => resolve();
  });
}

export async function putSnapshot(snapshot: DocumentSnapshot): Promise<void> {
  const db = await openDb();
  if (!db) return;
  try {
    const tx = db.transaction(STORE, 'readwrite');
    const store = tx.objectStore(STORE);
    store.put({
      ...snapshot,
      // Copy so later mutations cannot affect the persisted record.
      bytes: snapshot.bytes.slice(),
    });
    await transactionDone(tx);
    await pruneSnapshots();
  } catch {
    // Storage failures must never block editing.
  }
}

export async function listSnapshots(): Promise<DocumentSnapshotMeta[]> {
  const db = await openDb();
  if (!db) return [];
  try {
    const tx = db.transaction(STORE, 'readonly');
    const records = await requestToPromise(tx.objectStore(STORE).getAll());
    if (!records) return [];
    return (records as DocumentSnapshot[])
      .map(({ bytes: _bytes, ...meta }) => meta)
      .sort((a, b) => b.updatedAt - a.updatedAt);
  } catch {
    return [];
  }
}

export async function getSnapshot(id: string): Promise<DocumentSnapshot | null> {
  const db = await openDb();
  if (!db) return null;
  try {
    const tx = db.transaction(STORE, 'readonly');
    const record = await requestToPromise(tx.objectStore(STORE).get(id));
    if (!record) return null;
    const snapshot = record as DocumentSnapshot;
    if (!(snapshot.bytes instanceof Uint8Array)) {
      snapshot.bytes = new Uint8Array(snapshot.bytes as unknown as ArrayBuffer);
    }
    return snapshot;
  } catch {
    return null;
  }
}

export async function removeSnapshot(id: string): Promise<void> {
  const db = await openDb();
  if (!db) return;
  try {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).delete(id);
    await transactionDone(tx);
  } catch {
    // ignore
  }
}

export async function clearSnapshots(): Promise<void> {
  const db = await openDb();
  if (!db) return;
  try {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).clear();
    await transactionDone(tx);
  } catch {
    // ignore
  }
}

async function pruneSnapshots(): Promise<void> {
  const db = await openDb();
  if (!db) return;
  try {
    const tx = db.transaction(STORE, 'readonly');
    const records = (await requestToPromise(tx.objectStore(STORE).getAll())) as DocumentSnapshot[] | null;
    if (!records || records.length <= MAX_SNAPSHOTS) {
      if (!records) return;
      let total = records.reduce((sum, record) => sum + (record.bytes?.byteLength ?? 0), 0);
      if (total <= MAX_TOTAL_BYTES) return;
    }
    const sorted = records.sort((a, b) => b.updatedAt - a.updatedAt);
    let total = 0;
    const toRemove: string[] = [];
    sorted.forEach((record, index) => {
      total += record.bytes?.byteLength ?? 0;
      if (index >= MAX_SNAPSHOTS || total > MAX_TOTAL_BYTES) toRemove.push(record.id);
    });
    if (toRemove.length === 0) return;
    const writeTx = db.transaction(STORE, 'readwrite');
    const store = writeTx.objectStore(STORE);
    for (const id of toRemove) store.delete(id);
    await transactionDone(writeTx);
  } catch {
    // ignore
  }
}
