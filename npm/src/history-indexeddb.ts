import { DocxHistoryError } from './history.js';
import type { HistoryBlobReference, HistoryHead, HistoryHeadInitializationResult, HistoryStorage } from './history.js';
import type { HistoryCheckpointJournal, HistoryCheckpointRequest } from './history-checkpoints.js';

export interface IndexedDbHistoryStore {
  readonly storage: HistoryStorage;
  journal(documentId: string): HistoryCheckpointJournal;
  /** Await active history calls first. Stored documents and pending requests are retained. */
  close(): void;
}

/** Optional browser-local persistence. Call explicitly; no storage, autosave or retention
 * policy is installed by mounting controls. Browser site-data clearing removes this store. */
export async function openIndexedDbHistoryStore(name: string): Promise<IndexedDbHistoryStore> {
  const db = await new Promise<IDBDatabase>((resolve, reject) => {
    let blocked = false;
    const request = indexedDB.open(name, 1);
    request.onupgradeneeded = () => {
      for (const store of ['blobs', 'heads', 'requests']) request.result.createObjectStore(store);
    };
    request.onsuccess = () => { if (blocked) request.result.close(); else resolve(request.result); };
    request.onerror = () => reject(request.error);
    request.onblocked = () => {
      blocked = true;
      reject(new Error('Close other tabs using this history store, then try again.'));
    };
  });
  db.onversionchange = () => db.close();

  function transaction<T>(name: string, mode: IDBTransactionMode,
    action: (store: IDBObjectStore, result: (value: T) => void) => void): Promise<T> {
    return new Promise((resolve, reject) => {
      const tx = db.transaction(name, mode);
      let value: T;
      tx.oncomplete = () => resolve(value);
      tx.onabort = () => reject(tx.error ?? new Error('History storage transaction was aborted.'));
      try { action(tx.objectStore(name), result => { value = result; }); }
      catch (error) { tx.abort(); reject(error); }
    });
  }

  const storage: HistoryStorage = {
    readBlob(reference) {
      const key = referenceKey(reference);
      return transaction<Uint8Array | null>('blobs', 'readonly', (store, result) => {
        store.get(key).onsuccess = event => result((event.target as IDBRequest<Uint8Array>).result ?? null);
      });
    },
    async putBlob(reference, bytes) {
      const key = referenceKey(reference);
      const captured = bytes.slice();
      if (captured.length !== reference.length) throw new DocxHistoryError('PayloadMismatch', 'Stored document length does not match its reference.');
      const hash = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', captured)),
        byte => byte.toString(16).padStart(2, '0')).join('');
      if (hash !== key.split(':')[0])
        throw new DocxHistoryError('PayloadMismatch', 'Stored document bytes do not match their reference.');
      await transaction<void>('blobs', 'readwrite', (store, result) => {
        store.put(captured, key); result();
      });
    },
    readHead(documentId) {
      return transaction<HistoryHead | null>('heads', 'readonly', (store, result) => {
        store.get(documentId).onsuccess = event => result((event.target as IDBRequest<HistoryHead>).result ?? null);
      });
    },
    advanceHead(documentId, expected, state) {
      referenceKey(state);
      const captured = structuredClone({ expected, state });
      if (expected) validateHead(expected);
      const revision = BigInt(expected?.revision ?? '0') + 1n;
      if (revision > 9223372036854775807n) throw new RangeError('History revision exhausted.');
      return transaction<HistoryHead | null>('heads', 'readwrite', (store, result) => {
        store.get(documentId).onsuccess = event => {
          const current = (event.target as IDBRequest<HistoryHead>).result ?? null;
          if (!equalHead(current, captured.expected)) { result(null); return; }
          const head = { revision: String(revision), state: captured.state };
          store.put(head, documentId); result(head);
        };
      });
    },
    initializeHead(documentId, head) {
      validateHead(head);
      const captured = structuredClone(head);
      // Both publication paths lock the same object store for the entire read/write transaction.
      return transaction<HistoryHeadInitializationResult>('heads', 'readwrite', (store, result) => {
        store.get(documentId).onsuccess = event => {
          const existing = (event.target as IDBRequest<HistoryHead>).result;
          if (existing) { result({ initialized: false, head: existing }); return; }
          store.put(captured, documentId); result({ initialized: true, head: captured });
        };
      });
    },
  };

  return {
    storage,
    journal(documentId) {
      return {
        read: () => transaction<HistoryCheckpointRequest | null>('requests', 'readonly', (store, result) => {
          store.get(documentId).onsuccess = event => result((event.target as IDBRequest<HistoryCheckpointRequest>).result ?? null);
        }),
        put(request) {
          if (request.documentId !== documentId) throw new Error('Checkpoint document identity does not match.');
          const captured = structuredClone(request);
          return transaction<HistoryCheckpointRequest>('requests', 'readwrite', (store, result) => {
            store.get(documentId).onsuccess = event => {
              const existing = (event.target as IDBRequest<HistoryCheckpointRequest>).result;
              if (existing) { result(existing); return; }
              store.put(captured, documentId); result(captured);
            };
          });
        },
        remove(requestId) {
          return transaction<void>('requests', 'readwrite', (store, result) => {
            store.get(documentId).onsuccess = event => {
              if ((event.target as IDBRequest<HistoryCheckpointRequest>).result?.id === requestId) store.delete(documentId);
              result();
            };
          });
        },
      };
    },
    close: () => db.close(),
  };
}

function referenceKey(reference: HistoryBlobReference): string {
  if (reference.digest.algorithm !== 'SHA-256' || !/^[a-f0-9]{64}$/.test(reference.digest.value)
    || !Number.isSafeInteger(reference.length) || reference.length < 0)
    throw new DocxHistoryError('InvalidRequest', 'Invalid history blob reference.');
  return `${reference.digest.value}:${reference.length}`;
}
function validateHead(head: HistoryHead): void {
  referenceKey(head.state);
  if (typeof head.revision !== 'string' || !/^[1-9][0-9]*$/.test(head.revision) || BigInt(head.revision) > 9223372036854775807n)
    throw new DocxHistoryError('InvalidRequest', 'Invalid history revision.');
}
function equalHead(a: HistoryHead | null, b: HistoryHead | null): boolean {
  return a === null || b === null ? a === b
    : a.revision === b.revision && a.state.length === b.state.length
      && a.state.digest.algorithm === b.state.digest.algorithm && a.state.digest.value === b.state.digest.value;
}
