import { DocxHistoryError } from './history.js';
import type { DocxHistoryDocument, DocxHistoryView, DocxVersionMetadata, HistoryBlobReference, HistoryHead } from './history.js';

/** Persist the entire request before publication. Bytes are structured-cloneable, not JSON. */
export type HistoryCheckpointRequest = {
  id: string;
  documentId: string;
  metadata: DocxVersionMetadata;
} & ({ kind: 'save'; head: HistoryHead | null; bytes: Uint8Array }
  | { kind: 'restore'; head: HistoryHead; target: HistoryBlobReference });

/** Host-owned durable outbox, scoped to one document. Methods resolve only after commit.
 * put must atomically return the existing request or insert and return the supplied one.
 * Existing inputs must never change, even for the same ID. remove must match the ID.
 * Share exclusion between tabs/clients. Never clear an uncertain publication to start another. */
export interface HistoryCheckpointJournal {
  read(): Promise<HistoryCheckpointRequest | null>;
  put(request: HistoryCheckpointRequest): Promise<HistoryCheckpointRequest>;
  remove(requestId: string): Promise<void>;
}

/** Explicit checkpoint commands. No editor ownership, timers, automatic retries or subscriptions.
 * Keep this instance for the document's editing session; refresh only when the user requests it. */
export class HistoryCheckpoints {
  private current: DocxHistoryView | null = null;
  private request: HistoryCheckpointRequest | null = null;
  private running = false;
  private stale = false;

  private constructor(readonly document: DocxHistoryDocument, private readonly journal: HistoryCheckpointJournal) {}

  /** Pass an already captured view when opening its exact version in an editor (for example, after import). */
  static async open(document: DocxHistoryDocument, journal: HistoryCheckpointJournal,
    view?: DocxHistoryView | null): Promise<HistoryCheckpoints> {
    const controls = new HistoryCheckpoints(document, journal);
    controls.request = structuredClone(await journal.read());
    if (controls.request && controls.request.documentId !== document.documentId)
      throw new DocxHistoryError('InvalidRequest', 'The pending checkpoint belongs to another document.');
    controls.current = view === undefined ? await document.read() : structuredClone(view);
    if (controls.current && controls.current.state.documentId !== document.documentId)
      throw new DocxHistoryError('InvalidRequest', 'The captured view belongs to another document.');
    return controls;
  }

  get view(): DocxHistoryView | null { return structuredClone(this.current); }
  get hasPending(): boolean { return this.request !== null; }
  get needsRefresh(): boolean { return this.stale; }

  async refresh(): Promise<DocxHistoryView | null> {
    return this.exclusive(async () => {
      this.current = await this.document.read();
      this.stale = false;
      return this.view;
    });
  }

  async save(bytes: Uint8Array, metadata: DocxVersionMetadata): Promise<DocxHistoryView> {
    return this.start({ kind: 'save', head: this.current?.head ?? null, bytes,
      id: crypto.randomUUID(), documentId: this.document.documentId, metadata });
  }

  async restore(target: HistoryBlobReference, metadata: DocxVersionMetadata): Promise<DocxHistoryView> {
    if (!this.current) throw new DocxHistoryError('NotFound', 'Save a checkpoint before restoring a version.');
    return this.start({ kind: 'restore', head: this.current.head, target,
      id: crypto.randomUUID(), documentId: this.document.documentId, metadata });
  }

  async retry(): Promise<DocxHistoryView> {
    return this.exclusive(async () => {
      if (!this.request) throw new DocxHistoryError('InvalidRequest', 'There is no pending checkpoint.');
      return this.publish();
    });
  }

  private async start(request: HistoryCheckpointRequest): Promise<DocxHistoryView> {
    return this.exclusive(async () => {
      if (this.request) throw new DocxHistoryError('PendingRequest', 'Retry the pending checkpoint first.');
      if (this.stale) throw new DocxHistoryError('StaleHead', 'Refresh history before saving again.');
      this.request = structuredClone(request);
      return this.publish();
    });
  }

  private async publish(): Promise<DocxHistoryView> {
    const intended = this.request!;
    // Retry persistence too: an outbox acknowledgement can be lost independently of publication.
    const request = await this.journal.put(structuredClone(intended));
    this.request = structuredClone(request);
    if (request.id !== intended.id)
      throw new DocxHistoryError('PendingRequest', 'Another tab has a pending checkpoint. Retry it before saving your draft.');
    let view: DocxHistoryView;
    try {
      view = request.kind === 'save'
        ? await this.document.createVersion(request.head, request.bytes, request.metadata, request.id)
        : await this.document.restoreVersion(request.head, request.target, request.metadata, request.id);
    } catch (error) {
      if (error instanceof DocxHistoryError && error.code === 'StaleHead') {
        this.stale = true;
        await this.journal.remove(request.id);
        this.request = null;
      }
      throw error;
    }
    // Idempotent retries may return an older version. Retain that exact view, never substitute latest.
    this.current = view;
    await this.journal.remove(request.id);
    this.request = null;
    return this.view!;
  }

  private async exclusive<T>(action: () => Promise<T>): Promise<T> {
    if (this.running) throw new DocxHistoryError('Busy', 'A history command is still running.');
    this.running = true;
    try { return await action(); } finally { this.running = false; }
  }
}
