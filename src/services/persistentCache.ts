// Minimal IndexedDB key/value store for the downloaded user directory, so a
// page view doesn't re-download the whole tenant. Every operation is
// best-effort: if IndexedDB is unavailable (private browsing, blocked site
// data, old browser) or anything fails, reads resolve undefined and writes
// are silently skipped.

const DB_NAME    = 'SmartOrgChart';
const STORE_NAME = 'userData';
const DB_VERSION = 1;
const OPEN_TIMEOUT_MS = 3000;   // never let a hung open() block the data load

let _dbPromise: Promise<IDBDatabase | null> | null = null;

function openDb(): Promise<IDBDatabase | null> {
  if (_dbPromise) return _dbPromise;
  _dbPromise = new Promise<IDBDatabase | null>(resolve => {
    let settled = false;
    const done = (db: IDBDatabase | null): void => {
      if (settled) return;
      settled = true;
      resolve(db);
    };
    setTimeout(() => done(null), OPEN_TIMEOUT_MS);
    try {
      if (typeof indexedDB === 'undefined') { done(null); return; }
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        try {
          const db = req.result;
          if (!db.objectStoreNames.contains(STORE_NAME)) db.createObjectStore(STORE_NAME);
        } catch { /* handled by onerror */ }
      };
      req.onsuccess = () => done(req.result);
      req.onerror   = () => done(null);
      req.onblocked = () => done(null);
    } catch {
      done(null);
    }
  });
  return _dbPromise;
}

function runRequest<T>(mode: IDBTransactionMode, op: (store: IDBObjectStore) => IDBRequest): Promise<T | undefined> {
  return openDb().then(db => new Promise<T | undefined>(resolve => {
    if (!db) { resolve(undefined); return; }
    try {
      const tx  = db.transaction(STORE_NAME, mode);
      const req = op(tx.objectStore(STORE_NAME));
      req.onsuccess = () => resolve(req.result as T);
      req.onerror   = () => resolve(undefined);
      tx.onabort    = () => resolve(undefined);
    } catch {
      resolve(undefined);
    }
  })).catch(() => undefined);
}

export function cacheGet<T>(key: string): Promise<T | undefined> {
  return runRequest<T>('readonly', store => store.get(key));
}

export function cacheSet(key: string, value: unknown): Promise<void> {
  return runRequest<IDBValidKey>('readwrite', store => store.put(value, key)).then(() => undefined);
}

export function cacheDelete(key: string): Promise<void> {
  return runRequest<undefined>('readwrite', store => store.delete(key)).then(() => undefined);
}
