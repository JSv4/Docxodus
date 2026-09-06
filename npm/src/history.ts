import type { VerificationDigest } from './types.js';

/** Nonnegative Int64 decimal string, never a JS number. Sequence orders content; time does not. */
export type HistoryPosition = string;
export interface HistoryBlobReference { digest: VerificationDigest; length: number }
export interface HistoryHead { revision: HistoryPosition; state: HistoryBlobReference }
export interface DocxSnapshotReference { blob: HistoryBlobReference; contentDigest: VerificationDigest }
export interface DocxVersionMetadata {
  author: string;
  createdAt: string;
  label?: string | null;
  message?: string | null;
  applicationMetadata?: Readonly<Record<string, string>>;
}
export interface DocxVersionRecord {
  documentId: string;
  metadata: DocxVersionMetadata;
  nonce: string;
  parent: HistoryBlobReference | null;
  restoredFrom: HistoryBlobReference | null;
  sequence: HistoryPosition;
  snapshot: DocxSnapshotReference;
}
export interface DocxStoredVersion { id: HistoryBlobReference; record: DocxVersionRecord }
export interface DocxHistoryState {
  commit: HistoryBlobReference | null;
  documentId: string;
  epoch: HistoryPosition;
  initialSnapshot: DocxSnapshotReference;
  sequence: HistoryPosition;
  snapshot: DocxSnapshotReference;
  version: HistoryBlobReference;
}
export interface DocxHistoryView { head: HistoryHead; state: DocxHistoryState; version: DocxStoredVersion }
export interface DocxVersionPage { versions: DocxStoredVersion[]; next: HistoryBlobReference | null }

/** Host-owned storage. Persist blobs before returning; advanceHead must be an atomic CAS.
 * References and returned bytes are verified in the core. Methods must settle their promises.
 * No fetching, subscription, timer, authorization or retention policy is supplied by this API. */
export interface HistoryStorage {
  readBlob(reference: HistoryBlobReference): Promise<Uint8Array | null>;
  putBlob(reference: HistoryBlobReference, bytes: Uint8Array): Promise<void>;
  readHead(documentId: string): Promise<HistoryHead | null>;
  advanceHead(documentId: string, expected: HistoryHead | null, state: HistoryBlobReference): Promise<HistoryHead | null>;
}

export interface HistoryBridge {
  Open(adapterId: number): number;
  Close(handle: number): void;
  Invoke(handle: number, requestJson: string, bytes: Uint8Array): Promise<string>;
}

export class DocxHistoryError extends Error {
  constructor(public readonly code: string, message: string) { super(message); this.name = 'DocxHistoryError'; }
}

interface Result {
  success: boolean;
  errorCode: string | null;
  message: string | null;
  view: DocxHistoryView | null;
  version: DocxStoredVersion | null;
  page: DocxVersionPage | null;
  sequence: HistoryPosition | null;
  bytes: string | null;
}

const adapters = new Map<number, HistoryStorage>();
let nextAdapter = 0;
function adapter(id: number): HistoryStorage {
  const store = adapters.get(id);
  if (!store) throw new DocxHistoryError('Closed', 'History storage adapter is closed.');
  return store;
}

function toBase64(bytes: Uint8Array): string {
  let binary = '';
  for (let offset = 0; offset < bytes.length; offset += 0x8000)
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  return btoa(binary);
}
function fromBase64(encoded: string): Uint8Array {
  const binary = atob(encoded);
  return Uint8Array.from(binary, c => c.charCodeAt(0));
}

/** For custom WASM loaders. initialize() installs these automatically. */
export function installHistoryStorageImports(setModuleImports: (name: string, imports: object) => void): void {
  setModuleImports('docxodus.history', {
    readBlob: async (id: number, json: string): Promise<string> => {
      const reference: HistoryBlobReference = JSON.parse(json);
      const bytes = await adapter(id).readBlob(reference);
      if (bytes === null) return '';
      if (bytes.length !== reference.length) throw new DocxHistoryError('PayloadMismatch', 'Stored blob length differs from its reference.');
      return bytes.length === 0 ? '=' : toBase64(bytes);
    },
    putBlob: async (id: number, json: string, bytes: Uint8Array): Promise<void> =>
      adapter(id).putBlob(JSON.parse(json), bytes.slice()),
    readHead: async (id: number, documentId: string): Promise<string> =>
      JSON.stringify(await adapter(id).readHead(documentId)),
    advanceHead: async (id: number, documentId: string, expected: string, state: string): Promise<string> =>
      JSON.stringify(await adapter(id).advanceHead(documentId, JSON.parse(expected), JSON.parse(state))),
  });
}

/** Durable package-history client. Reopening over the same adapters preserves history.
 * Await active calls before close(); it releases handles but never deletes host data. */
export class DocxHistoryClient {
  private readonly adapterId: number;
  private readonly handle: number;
  private closed = false;

  constructor(private readonly bridge: HistoryBridge, storage: HistoryStorage) {
    if (nextAdapter >= 0x7fffffff) throw new Error('History adapter identifiers exhausted.');
    this.adapterId = ++nextAdapter;
    adapters.set(this.adapterId, storage);
    try { this.handle = bridge.Open(this.adapterId); }
    catch (error) { adapters.delete(this.adapterId); throw error; }
  }

  close(): void {
    if (this.closed) return;
    this.bridge.Close(this.handle);
    adapters.delete(this.adapterId);
    this.closed = true;
  }

  async read(documentId: string): Promise<DocxHistoryView | null> {
    return (await this.call('read', documentId)).view;
  }
  async createVersion(documentId: string, expectedHead: HistoryHead | null, bytes: Uint8Array,
    metadata: DocxVersionMetadata): Promise<DocxHistoryView> {
    return (await this.call('create', documentId, { expectedHead, metadata }, bytes)).view!;
  }
  async listVersions(documentId: string, cursor: HistoryBlobReference | null = null, limit = 25): Promise<DocxVersionPage> {
    return (await this.call('list', documentId, { versionId: cursor, limit })).page!;
  }
  async getVersion(documentId: string, versionId: HistoryBlobReference): Promise<DocxStoredVersion> {
    return (await this.call('get', documentId, { versionId })).version!;
  }
  async exportVersion(documentId: string, versionId: HistoryBlobReference): Promise<Uint8Array> {
    return fromBase64((await this.call('export', documentId, { versionId })).bytes!);
  }
  async materialize(documentId: string, sequence: HistoryPosition, maxEntriesToScan = 10_000): Promise<Uint8Array> {
    return fromBase64((await this.call('materialize', documentId, { sequence, maxEntriesToScan })).bytes!);
  }
  async replay(documentId: string, sequence: HistoryPosition, maxEntriesToScan = 10_000): Promise<Uint8Array> {
    return fromBase64((await this.call('replay', documentId, { sequence, maxEntriesToScan })).bytes!);
  }
  async resolveSequenceAtTime(documentId: string, cutoff: string, maxEntriesToScan = 10_000): Promise<HistoryPosition> {
    return (await this.call('resolveTime', documentId, { cutoff, maxEntriesToScan })).sequence!;
  }
  async restoreVersion(documentId: string, expectedHead: HistoryHead, versionId: HistoryBlobReference,
    metadata: DocxVersionMetadata): Promise<DocxHistoryView> {
    return (await this.call('restore', documentId, { expectedHead, versionId, metadata })).view!;
  }

  private async call(operation: string, documentId: string, fields: object = {}, bytes: Uint8Array = new Uint8Array()): Promise<Result> {
    if (this.closed) throw new DocxHistoryError('Closed', 'History client is closed.');
    const result: Result = JSON.parse(await this.bridge.Invoke(this.handle,
      JSON.stringify({ schemaVersion: 1, operation, documentId, ...fields }), bytes));
    if (!result.success) throw new DocxHistoryError(result.errorCode!, result.message!);
    return result;
  }
}

/** Reference process-local storage. Share one instance among local clients; it is not durable
 * across a page/process restart. Every incoming blob is copied and SHA-256 verified. */
export function createMemoryHistoryStorage(maxBlobBytes = 256 * 1024 * 1024): HistoryStorage {
  if (!Number.isSafeInteger(maxBlobBytes) || maxBlobBytes <= 0) throw new RangeError('Invalid blob byte limit.');
  const blobs = new Map<string, Uint8Array>();
  const heads = new Map<string, HistoryHead>();
  const copy = <T>(value: T): T => JSON.parse(JSON.stringify(value));
  const key = (reference: HistoryBlobReference): string => {
    if (reference.digest.algorithm !== 'SHA-256' || !/^[a-f0-9]{64}$/.test(reference.digest.value)
      || !Number.isSafeInteger(reference.length) || reference.length < 0 || reference.length > maxBlobBytes)
      throw new DocxHistoryError('InvalidRequest', 'Invalid blob reference or byte limit.');
    return `${reference.digest.value}:${reference.length}`;
  };
  const equalReference = (a: HistoryBlobReference, b: HistoryBlobReference): boolean =>
    a.length === b.length && a.digest.algorithm === b.digest.algorithm && a.digest.value === b.digest.value;
  const equalHead = (a: HistoryHead | null, b: HistoryHead | null): boolean => a === null || b === null
    ? a === b : a.revision === b.revision && equalReference(a.state, b.state);
  return {
    async readBlob(reference) { return blobs.get(key(reference))?.slice() ?? null; },
    async putBlob(reference, bytes) {
      const blobKey = key(reference);
      const captured = bytes.slice();
      if (captured.length !== reference.length) throw new DocxHistoryError('PayloadMismatch', 'Blob length mismatch.');
      const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', captured)),
        n => n.toString(16).padStart(2, '0')).join('');
      if (digest !== blobKey.split(':')[0]) throw new DocxHistoryError('PayloadMismatch', 'Blob digest mismatch.');
      if (!blobs.has(blobKey)) blobs.set(blobKey, captured);
    },
    async readHead(documentId) { return copy(heads.get(documentId) ?? null); },
    async advanceHead(documentId, expected, state) {
      key(state);
      const current = heads.get(documentId) ?? null;
      if (!equalHead(current, expected)) return null;
      const revision = BigInt(current?.revision ?? '0') + 1n;
      if (revision > 9223372036854775807n) throw new RangeError('History revision exhausted.');
      const head = { revision: String(revision), state: copy(state) };
      // No awaits inside the compare/publication critical section, even under concurrent callers.
      heads.set(documentId, head);
      return copy(head);
    },
  };
}
