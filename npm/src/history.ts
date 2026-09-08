import type { VerificationDigest } from './types.js';

/** Nonnegative Int64 decimal string, never a JS number. Sequence orders content; time does not. */
export type HistoryPosition = string;
export interface HistoryBlobReference { digest: VerificationDigest; length: number }
export interface HistoryHead { revision: HistoryPosition; state: HistoryBlobReference }
export interface HistoryHeadInitializationResult { initialized: boolean; head: HistoryHead }
export interface DocxHistoryArchiveInfo { documentId: string; head: HistoryHead; blobCount: number; totalBlobBytes: HistoryPosition }
export interface DocxHistoryImportResult { archive: DocxHistoryArchiveInfo; view: DocxHistoryView; alreadyPresent: boolean }
export const MAX_HISTORY_ARCHIVE_BYTES = 64 * 1024 * 1024;
export interface HistoryRequestIdentity { id: string; fingerprint: VerificationDigest }
export interface HistoryRequestJournal {
  documentId: string;
  revision: HistoryPosition;
  index: HistoryBlobReference | null;
  current: HistoryRequestIdentity | null;
}
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
  requests?: HistoryRequestJournal;
  parentPublication?: HistoryHead;
  operation?: HistoryBlobReference;
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
export interface PackageHistoryCommit {
  after: DocxSnapshotReference;
  before: DocxSnapshotReference;
  contribution: HistoryBlobReference | null;
  documentId: string;
  epoch: HistoryPosition;
  kind: 'import' | 'restore';
  parent: HistoryBlobReference | null;
  sequence: HistoryPosition;
  version: HistoryBlobReference;
}
export interface DocxHistoryLogEntry { id: HistoryBlobReference; commit: PackageHistoryCommit; metadata: DocxVersionMetadata }
export interface DocxHistoryUpdate {
  after: HistoryHead | null;
  view: DocxHistoryView;
  entries: DocxHistoryLogEntry[];
  reset: boolean;
}
export interface DocxTextSplice { partUri: string; textNode: number; offset: number; deleteCount: number; insert: string }
export interface DocxOperationRequest {
  requestId: string; base: HistoryHead; kind: 'text' | 'package' | 'discard'; metadata: DocxVersionMetadata;
  text: DocxTextSplice | null; readParts: string[]; resolves: HistoryBlobReference | null;
}
export interface DocxOperationInput { documentId: string; request: DocxOperationRequest; candidate: HistoryBlobReference | null }
export interface DocxOperationRecord {
  documentId: string; input: HistoryBlobReference; parent: HistoryBlobReference | null; revision: HistoryPosition;
  before: HistoryHead; proposedSnapshot: DocxSnapshotReference; afterSnapshot: DocxSnapshotReference;
  version: HistoryBlobReference; contentCommit: HistoryBlobReference | null; status: 'accepted' | 'conflict';
  conflict: string | null; appliedText: DocxTextSplice | null;
}
export interface DocxStoredOperation { id: HistoryBlobReference; record: DocxOperationRecord; input: DocxOperationInput }
export interface DocxOperationUpdate { after: HistoryHead | null; view: DocxHistoryView; operations: DocxStoredOperation[] }

/** Host-owned storage. Persist blobs before returning; advanceHead must be an atomic CAS.
 * References and returned bytes are verified in the core. Methods must settle their promises.
 * No fetching, subscription, timer, authorization or retention policy is supplied by this API. */
export interface HistoryStorage {
  readBlob(reference: HistoryBlobReference): Promise<Uint8Array | null>;
  putBlob(reference: HistoryBlobReference, bytes: Uint8Array): Promise<void>;
  readHead(documentId: string): Promise<HistoryHead | null>;
  advanceHead(documentId: string, expected: HistoryHead | null, state: HistoryBlobReference): Promise<HistoryHead | null>;
  /** Optional import capability. Atomically insert the exact head only if absent; otherwise return
   * the unchanged existing head. Must share exclusion with advanceHead, never overwrite/revise it. */
  initializeHead?(documentId: string, head: HistoryHead): Promise<HistoryHeadInitializationResult>;
}

export interface HistoryBridge {
  Open(adapterId: number): number;
  OpenWithInitialization?(adapterId: number): number;
  OpenArchive?(bytes: Uint8Array): Promise<string>;
  Close(handle: number): void;
  Invoke(handle: number, requestJson: string, bytes: Uint8Array): Promise<string>;
}

export class DocxHistoryError extends Error {
  constructor(public readonly code: string, message: string) { super(message); this.name = 'DocxHistoryError'; }
}

interface Result {
  handle: number | null;
  archive: DocxHistoryArchiveInfo | null;
  import: DocxHistoryImportResult | null;
  success: boolean;
  errorCode: string | null;
  message: string | null;
  view: DocxHistoryView | null;
  version: DocxStoredVersion | null;
  page: DocxVersionPage | null;
  sequence: HistoryPosition | null;
  bytes: string | null;
  update: DocxHistoryUpdate | null;
  operationUpdate: DocxOperationUpdate | null;
  operation: DocxStoredOperation | null;
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
    initializeHead: async (id: number, documentId: string, head: string): Promise<string> => {
      const storage = adapter(id);
      if (!storage.initializeHead) throw new DocxHistoryError('InitializationUnsupported', 'Storage cannot import an exact head.');
      return JSON.stringify(await storage.initializeHead(documentId, JSON.parse(head)));
    },
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
    try { this.handle = storage.initializeHead && bridge.OpenWithInitialization
      ? bridge.OpenWithInitialization(this.adapterId) : bridge.Open(this.adapterId); }
    catch (error) { adapters.delete(this.adapterId); throw error; }
  }

  close(): void {
    if (this.closed) return;
    this.bridge.Close(this.handle);
    adapters.delete(this.adapterId);
    this.closed = true;
  }

  /** Bind an identity without reading, creating a version, or taking ownership of this client. */
  document(documentId: string): DocxHistoryDocument {
    return new DocxHistoryDocument(documentId, (operation, fields, bytes) => this.call(operation, documentId, fields, bytes));
  }

  async exportHistoryArchive(documentId: string): Promise<Uint8Array> {
    return fromBase64((await this.call('exportArchive', documentId)).bytes!);
  }
  async importHistoryArchive(bytes: Uint8Array): Promise<DocxHistoryImportResult> {
    checkArchiveSize(bytes);
    return (await this.call('importArchive', '', {}, bytes)).import!;
  }

  async read(documentId: string): Promise<DocxHistoryView | null> {
    return (await this.call('read', documentId)).view;
  }
  /** Host-triggered validated metadata tail; first join starts at the latest checkpoint. */
  async readChangesSince(documentId: string, after: HistoryHead | null, maxEntriesToScan = 10_000): Promise<DocxHistoryUpdate> {
    return (await this.call('updates', documentId, { expectedHead: after, maxEntriesToScan })).update!;
  }
  async readOperationsSince(documentId: string, after: HistoryHead | null, maxEntriesToScan = 10_000): Promise<DocxOperationUpdate> {
    return (await this.call('operations', documentId, { expectedHead: after, maxEntriesToScan })).operationUpdate!;
  }
  async getOperation(documentId: string, operationId: HistoryBlobReference): Promise<DocxStoredOperation> {
    return (await this.call('getOperation', documentId, { operationId })).operation!;
  }
  async exportOperationProposal(documentId: string, operationId: HistoryBlobReference): Promise<Uint8Array> {
    return fromBase64((await this.call('exportOperationProposal', documentId, { operationId })).bytes!);
  }
  async compareVersions(documentId: string, beforeVersionId: HistoryBlobReference, afterVersionId: HistoryBlobReference): Promise<Uint8Array> {
    return fromBase64((await this.call('compare', documentId, { beforeVersionId, afterVersionId })).bytes!);
  }
  async createVersion(documentId: string, expectedHead: HistoryHead | null, bytes: Uint8Array,
    metadata: DocxVersionMetadata, requestId?: string): Promise<DocxHistoryView> {
    return (await this.call('create', documentId, { expectedHead, metadata, requestId }, bytes)).view!;
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
    metadata: DocxVersionMetadata, requestId?: string): Promise<DocxHistoryView> {
    return (await this.call('restore', documentId, { expectedHead, versionId, metadata, requestId })).view!;
  }

  private async call(operation: string, documentId: string, fields: object = {}, bytes: Uint8Array = new Uint8Array()): Promise<Result> {
    if (this.closed) throw new DocxHistoryError('Closed', 'History client is closed.');
    return parseResult(await this.bridge.Invoke(this.handle,
      JSON.stringify({ schemaVersion: 1, operation, documentId, ...fields }), bytes));
  }
}

type HistoryCall = (operation: string, fields?: object, bytes?: Uint8Array) => Promise<Result>;

/** Document-scoped reads. No method publishes or changes an editor. */
export class DocxHistoryReader {
  readonly #documentId: string;
  get documentId(): string { return this.#documentId; }
  /** @internal Obtain through client.document() or openDocxHistoryArchive(). */
  constructor(documentId: string, protected readonly call: HistoryCall) { this.#documentId = documentId; }
  async read(): Promise<DocxHistoryView | null> { return (await this.call('read')).view; }
  async listVersions(cursor: HistoryBlobReference | null = null, limit = 25): Promise<DocxVersionPage> {
    return (await this.call('list', { versionId: cursor, limit })).page!;
  }
  async getVersion(versionId: HistoryBlobReference): Promise<DocxStoredVersion> {
    return (await this.call('get', { versionId })).version!;
  }
  /** Omit versionId for latest; pass an ID to keep a UI operation pinned to a captured version. */
  async exportDocx(versionId: HistoryBlobReference | null = null): Promise<Uint8Array> {
    return fromBase64((await this.call('exportDocx', { versionId })).bytes!);
  }
  async exportHistoryArchive(): Promise<Uint8Array> { return fromBase64((await this.call('exportArchive')).bytes!); }
  async materialize(sequence: HistoryPosition, maxEntriesToScan = 10_000): Promise<Uint8Array> {
    return fromBase64((await this.call('materialize', { sequence, maxEntriesToScan })).bytes!);
  }
  async replay(sequence: HistoryPosition, maxEntriesToScan = 10_000): Promise<Uint8Array> {
    return fromBase64((await this.call('replay', { sequence, maxEntriesToScan })).bytes!);
  }
  async resolveSequenceAtTime(cutoff: string, maxEntriesToScan = 10_000): Promise<HistoryPosition> {
    return (await this.call('resolveTime', { cutoff, maxEntriesToScan })).sequence!;
  }
  async readChangesSince(after: HistoryHead | null, maxEntriesToScan = 10_000): Promise<DocxHistoryUpdate> {
    return (await this.call('updates', { expectedHead: after, maxEntriesToScan })).update!;
  }
  async readOperationsSince(after: HistoryHead | null, maxEntriesToScan = 10_000): Promise<DocxOperationUpdate> {
    return (await this.call('operations', { expectedHead: after, maxEntriesToScan })).operationUpdate!;
  }
  async getOperation(operationId: HistoryBlobReference): Promise<DocxStoredOperation> {
    return (await this.call('getOperation', { operationId })).operation!;
  }
  async exportOperationProposal(operationId: HistoryBlobReference): Promise<Uint8Array> {
    return fromBase64((await this.call('exportOperationProposal', { operationId })).bytes!);
  }
  /** Redlined DOCX through DocxCompare's existing accepted-input revision policy. */
  async compareVersions(beforeVersionId: HistoryBlobReference, afterVersionId: HistoryBlobReference): Promise<Uint8Array> {
    return fromBase64((await this.call('compare', { beforeVersionId, afterVersionId })).bytes!);
  }
}

/** Explicit checkpoint/restore controls over host-owned storage; request IDs are required here. */
export class DocxHistoryDocument extends DocxHistoryReader {
  async createVersion(expectedHead: HistoryHead | null, bytes: Uint8Array, metadata: DocxVersionMetadata,
    requestId: string): Promise<DocxHistoryView> {
    requireRequestId(requestId);
    return (await this.call('create', { expectedHead, metadata, requestId }, bytes)).view!;
  }
  async restoreVersion(expectedHead: HistoryHead, versionId: HistoryBlobReference, metadata: DocxVersionMetadata,
    requestId: string): Promise<DocxHistoryView> {
    requireRequestId(requestId);
    return (await this.call('restore', { expectedHead, versionId, metadata, requestId })).view!;
  }
}

/** Standalone readonly file. Owns its handle and captured bytes; await calls before close(). */
export class DocxHistoryArchive extends DocxHistoryReader {
  private closed = false;
  private constructor(private readonly bridge: HistoryBridge, private readonly handle: number,
    public readonly info: DocxHistoryArchiveInfo) {
    super(info.documentId, async (operation, fields = {}, bytes = new Uint8Array()) => {
      if (this.closed) throw new DocxHistoryError('Closed', 'History archive is closed.');
      return parseResult(await bridge.Invoke(handle,
        JSON.stringify({ schemaVersion: 1, operation, documentId: this.documentId, ...fields }), bytes));
    });
  }
  /** For custom loaders; normal callers use openDocxHistoryArchive(). */
  static async open(bridge: HistoryBridge, bytes: Uint8Array): Promise<DocxHistoryArchive> {
    checkArchiveSize(bytes);
    if (!bridge.OpenArchive) throw new Error('This WASM build does not include portable history.');
    const result = parseResult(await bridge.OpenArchive(bytes));
    return new DocxHistoryArchive(bridge, result.handle!, result.archive!);
  }
  close(): void {
    if (this.closed) return;
    this.bridge.Close(this.handle); this.closed = true;
  }
}

function parseResult(json: string): Result {
  const result: Result = JSON.parse(json);
  if (!result.success) throw new DocxHistoryError(result.errorCode!, result.message!);
  return result;
}
function checkArchiveSize(bytes: Uint8Array): void {
  if (bytes.length > MAX_HISTORY_ARCHIVE_BYTES) throw new DocxHistoryError('ResourceLimit', 'History archive exceeds 64 MiB.');
}
function requireRequestId(requestId: string): void {
  if (typeof requestId !== 'string' || !requestId.trim()) throw new DocxHistoryError('InvalidRequest', 'A durable requestId is required.');
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
    async initializeHead(documentId, head) {
      key(head.state);
      if (typeof head.revision !== 'string' || !/^[1-9][0-9]*$/.test(head.revision) || BigInt(head.revision) > 9223372036854775807n)
        throw new DocxHistoryError('InvalidRequest', 'Imported head revision must be a positive Int64 string.');
      const existing = heads.get(documentId);
      if (existing) return { initialized: false, head: copy(existing) };
      // Same synchronous critical section as advanceHead; exact revision, no await/increment.
      const captured = copy(head); heads.set(documentId, captured);
      return { initialized: true, head: copy(captured) };
    },
  };
}
