import { HistoryCheckpoints } from './history-checkpoints.js';
import { mountHistoryControls, historyControlError } from './history-controls.js';
import type { HistoryControls } from './history-controls.js';
import { openIndexedDbHistoryStore } from './history-indexeddb.js';
import type { IndexedDbHistoryStore } from './history-indexeddb.js';
import { DocxHistoryError, MAX_HISTORY_ARCHIVE_BYTES } from './history.js';
import type { DocxHistoryArchive, DocxHistoryClient, HistoryStorage } from './history.js';
import type { RibbonEditor } from './ribbon.js';

export interface RibbonHistoryOptions {
  /** Browser database shared by this app's editors. No document bytes are stored until Save version. */
  storageName?: string;
  /** Stable host-owned workspace key. Reopens its last saved document across reloads. */
  workspaceId?: string;
  author?: string;
}

/** Low-level hosts supply their initialized history and preview services. createRibbonEditor wires these automatically. */
export interface RibbonHistoryBinding extends RibbonHistoryOptions {
  openHistory(storage: HistoryStorage): DocxHistoryClient;
  openArchive(bytes: Uint8Array): Promise<DocxHistoryArchive>;
  preview(container: HTMLElement, bytes: Uint8Array): Promise<{ destroy(): void }>;
}

interface DocumentIdentity { id: string; name: string }
interface SavedDocument extends DocumentIdentity { checkpoints: HistoryCheckpoints }

/** Shared ribbon lifecycle, browser persistence and separate previews. No autosave or network sync. */
export class RibbonHistory {
  private readonly dialog: HTMLDialogElement;
  private readonly previewDialog: HTMLDialogElement;
  private readonly body: HTMLElement;
  private readonly status: HTMLElement;
  private readonly source: HTMLElement;
  private readonly recent: HTMLSelectElement;
  private readonly back: HTMLButtonElement;
  private readonly resumeButton: HTMLButtonElement;
  private readonly events = new AbortController();
  private readonly prefix: string;
  private readonly lastKey: string;
  private store?: IndexedDbHistoryStore;
  private client?: DocxHistoryClient;
  private draft?: SavedDocument;
  private identity: DocumentIdentity = { id: crypto.randomUUID(), name: 'Untitled.docx' };
  private archive?: { reader: DocxHistoryArchive; bytes: Uint8Array; name: string };
  private panel?: HistoryControls;
  private preview?: { destroy(): void };
  private active?: Promise<void>;
  private destroyed = false;
  private installing = false;
  private dirty = false;
  private generation = 0;
  private captured = -1;
  private savedVersion: number | null = null;
  private observedVersion: number | null = null;
  private capturedVersion: number | null = null;
  private capturedBytes?: Uint8Array;
  private readonly panelObserver: MutationObserver;

  constructor(private readonly ribbon: RibbonEditor, private readonly options: RibbonHistoryBinding) {
    const doc = ribbon.element.ownerDocument;
    this.prefix = `docxodus:versions:${options.storageName ?? 'docxodus-editor'}:document:`;
    this.lastKey = `docxodus:versions:${options.storageName ?? 'docxodus-editor'}:workspace:${options.workspaceId ?? crypto.randomUUID()}`;
    const style = doc.createElement('style'); style.textContent = HISTORY_DRAWER_CSS;
    this.dialog = doc.createElement('dialog');
    this.dialog.className = 'dxr-history-dialog';
    this.dialog.setAttribute('aria-label', 'Version history');
    const header = doc.createElement('div'); header.className = 'dxr-history-header'; this.dialog.append(header);
    const close = this.button('Close version history', () => this.dialog.close(), header);
    close.className = 'dxr-history-close'; close.textContent = 'Close ✕';
    this.source = doc.createElement('p'); this.source.className = 'dxr-history-source';
    this.status = doc.createElement('p'); this.status.setAttribute('role', 'status');
    this.status.className = 'dxr-history-status';
    const recentLabel = doc.createElement('label'); recentLabel.textContent = 'Saved documents on this device';
    this.recent = doc.createElement('select'); this.recent.setAttribute('aria-label', recentLabel.textContent);
    recentLabel.append(this.recent);
    this.recent.addEventListener('change', () => this.command(async () => {
      const selectedId = this.recent.value;
      this.recent.value = this.identity.id;
      const selected = this.readIdentity(localStorage.getItem(this.prefix + selectedId));
      if (!selected || !this.confirmReplace()) return;
      await this.loadSaved(selected);
      await this.mountPanel();
      this.updateRecent();
    }), { signal: this.events.signal });
    this.dialog.append(this.source, this.status, recentLabel);
    this.resumeButton = this.button('Continue editing this document', () => this.command(() => this.importArchive()), this.dialog);
    this.back = this.button('Back to my document', () => this.command(async () => {
      await this.closeArchive(); await this.mountPanel();
    }), this.dialog);
    this.body = doc.createElement('div'); this.dialog.append(this.body);
    this.previewDialog = doc.createElement('dialog'); this.previewDialog.className = 'dxr-version-preview';
    this.previewDialog.setAttribute('aria-label', 'Version preview');
    this.panelObserver = new MutationObserver(() => this.syncBusy());
    ribbon.element.append(style, this.dialog, this.previewDialog);
    for (const dialog of [this.dialog, this.previewDialog]) {
      dialog.addEventListener('keydown', event => event.stopPropagation(), { signal: this.events.signal });
      dialog.addEventListener('click', event => { if (event.target === dialog) {
        const rect = dialog.getBoundingClientRect();
        if (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom) dialog.close();
      } }, { signal: this.events.signal });
    }
    this.dialog.addEventListener('close', () => ribbon.control('history')?.focus(), { signal: this.events.signal });
    ribbon.control('history')?.addEventListener('click', () => { void this.show(); }, { signal: this.events.signal });
    doc.defaultView?.addEventListener('beforeunload', event => {
      if (this.hasUnsavedChanges || this.busy) event.preventDefault();
    }, { signal: this.events.signal });
    doc.defaultView?.addEventListener('pagehide', event => {
      if (!event.persisted) void this.destroy();
    }, { signal: this.events.signal });
  }

  get busy(): boolean { return !!this.active || this.panel?.element.getAttribute('aria-busy') === 'true'; }
  private get hasUnsavedChanges(): boolean {
    return this.dirty || (this.ribbon.editor?.version ?? null) !== this.savedVersion;
  }

  /** Called by the ribbon before a programmatic replacement. */
  beforeOpen(): void {
    if (!this.installing && this.busy) throw new DocxHistoryError('Busy', 'Finish the version action before opening another document.');
  }

  /** A new DOCX is a new document identity, even if its filename matches another document. */
  documentOpened(name: string): void {
    if (this.installing) return;
    this.identity = { id: crypto.randomUUID(), name };
    this.draft = undefined; this.dirty = false; this.generation++;
    this.savedVersion = this.observedVersion = this.ribbon.editor?.version ?? null;
    this.dialog.close(); this.previewDialog.close();
    void this.closeArchive();
  }

  edited(): void {
    const version = this.ribbon.editor?.version ?? null;
    if (version !== null && version === this.observedVersion) return;
    this.observedVersion = version; this.dirty = true; this.generation++;
  }

  /** Restores a saved workspace only when requested by the host, before exposing the editor. */
  async resume(): Promise<void> {
    if (!this.options.workspaceId) return;
    try {
      const previous = this.readIdentity(localStorage.getItem(this.lastKey));
      if (previous) await this.run(() => this.loadSaved(previous));
    } catch (error) {
      // Storage denial/corruption must not prevent an ordinary document from opening.
      this.ribbon.setStatus('Your document is open. Version history is unavailable on this device.');
      this.status.textContent = this.explain(error);
    }
  }

  async show(): Promise<void> {
    if (this.destroyed) return;
    if (!this.dialog.open) this.dialog.showModal();
    if (this.busy) return;
    await this.run(async () => { await this.mountPanel(); this.updateRecent(); }).catch(() => {});
  }

  async openFile(file: File): Promise<void> {
    await this.run(async () => {
      if (/\.docxhistory$/i.test(file.name)) {
        if (file.size > MAX_HISTORY_ARCHIVE_BYTES) throw new DocxHistoryError('ResourceLimit', 'History file is too large.');
        const bytes = new Uint8Array(await file.arrayBuffer());
        const reader = await this.options.openArchive(bytes);
        try { await reader.read(); }
        catch (error) { reader.close(); throw error; }
        await this.closeArchive();
        this.archive = { reader, bytes, name: file.name };
        if (!this.dialog.open) this.dialog.showModal();
        await this.mountPanel(); this.updateRecent();
      } else {
        if (!/\.docx$/i.test(file.name)) throw new Error('Choose a Word document or a document with version history.');
        if (!this.confirmReplace()) return;
        const bytes = new Uint8Array(await file.arrayBuffer());
        this.install(bytes, { id: crypto.randomUUID(), name: file.name });
        this.draft = undefined; this.dirty = true;
        await this.closeArchive(); this.dialog.close();
      }
    }).catch(() => {});
  }

  async newDocument(): Promise<void> {
    await this.run(async () => {
      if (!this.confirmReplace()) return;
      this.installing = true;
      try { this.ribbon.openBlank('Untitled.docx'); }
      finally { this.installing = false; }
      this.identity = { id: crypto.randomUUID(), name: 'Untitled.docx' };
      this.draft = undefined; this.dirty = false; this.generation++;
      this.savedVersion = this.observedVersion = this.ribbon.editor?.version ?? null;
      // Choosing New survives a reload instead of silently reopening the previous document.
      try { localStorage.removeItem(this.lastKey); } catch { /* editing does not require storage */ }
      await this.closeArchive(); this.dialog.close();
    }).catch(() => {});
  }

  async destroy(): Promise<void> {
    if (this.destroyed) return;
    this.destroyed = true; this.events.abort(); this.dialog.remove(); this.previewDialog.remove();
    this.panelObserver.disconnect();
    await this.active?.catch(() => {}); await this.closeArchive();
    this.preview?.destroy(); this.client?.close(); this.store?.close();
  }

  private async ensureStorage(): Promise<void> {
    if (this.client) return;
    const store = await openIndexedDbHistoryStore(this.options.storageName ?? 'docxodus-editor');
    try { this.client = this.options.openHistory(store.storage); this.store = store; }
    catch (error) { store.close(); throw error; }
  }

  private async ensureDraft(): Promise<SavedDocument> {
    await this.ensureStorage();
    if (!this.draft) {
      const document = this.client!.document(this.identity.id);
      this.draft = { ...this.identity, checkpoints: await HistoryCheckpoints.open(document, this.store!.journal(this.identity.id)) };
    }
    return this.draft;
  }

  private async loadSaved(identity: DocumentIdentity): Promise<void> {
    await this.ensureStorage();
    const document = this.client!.document(identity.id);
    const checkpoints = await HistoryCheckpoints.open(document, this.store!.journal(identity.id));
    const pending = checkpoints.pendingRequest;
    const view = checkpoints.view;
    if (!view && pending?.kind !== 'save') throw new Error('This saved document is no longer on this device.');
    const bytes = pending?.kind === 'save' ? pending.bytes : await document.exportDocx(view!.version.id);
    this.install(bytes, identity); this.draft = { ...identity, checkpoints };
    this.dirty = pending?.kind === 'save';
    if (pending?.kind === 'save') this.captureState(pending.bytes);
    await this.closeArchive(); this.remember();
  }

  private install(bytes: Uint8Array, identity: DocumentIdentity): void {
    if (this.destroyed) throw new DocxHistoryError('Closed', 'The editor is closed.');
    this.installing = true;
    try { this.ribbon.open(bytes, identity.name); }
    finally { this.installing = false; }
    this.identity = identity; this.dirty = false; this.generation++;
    this.savedVersion = this.observedVersion = this.ribbon.editor?.version ?? null;
  }

  private async mountPanel(): Promise<void> {
    const current = this.archive ?? await this.ensureDraft();
    const writable = 'checkpoints' in current;
    const holder = this.body.ownerDocument.createElement('div');
    const panel = mountHistoryControls(holder, {
      reader: writable ? current.checkpoints.document : current.reader,
      checkpoints: writable ? current.checkpoints : undefined,
      author: this.options.author, documentName: current.name,
      capture: writable ? () => {
        const bytes = this.ribbon.save(); if (!bytes) throw new Error('Open a document first.');
        // Remember the identity before publication, so uncertain saves can recover after reload.
        this.remember(); this.captureState(bytes); return bytes;
      } : undefined,
      confirmRestore: writable ? title => this.confirm(`Restore ${title}?${this.hasUnsavedChanges ? '\n\nYour unsaved changes will be replaced.' : ''}\n\nAll saved versions will be kept.`) : undefined,
      onCheckpoint: async (view, action, request) => {
        if (this.destroyed) return;
        const savedCapture = action === 'save' || (action === 'retry' && request?.kind === 'save'
          && this.capturedBytes && sameBytes(this.capturedBytes, request.bytes));
        if (savedCapture && this.captured === this.generation && this.capturedVersion === (this.ribbon.editor?.version ?? null)) {
          this.dirty = false; this.savedVersion = this.capturedVersion;
        }
        this.captured = -1; this.capturedBytes = undefined;
        if ((action === 'restore' || (action === 'retry' && request?.kind === 'restore')) && writable)
          this.install(await current.checkpoints.document.exportDocx(view.version.id), this.identity);
        this.remember(); this.updateRecent();
      },
      restoreUpdatesDraft: true,
      preview: (bytes, title) => this.showPreview(bytes, title),
    });
    try { await panel.ready; }
    catch (error) { await panel.destroy(); throw error; }
    if (this.destroyed) { await panel.destroy(); return; }
    await this.panel?.destroy(); this.panel = panel; this.body.replaceChildren(holder);
    this.panelObserver.disconnect();
    this.panelObserver.observe(panel.element, { attributes: true, attributeFilter: ['aria-busy'] });
    this.source.textContent = current.name;
    this.resumeButton.hidden = this.back.hidden = writable;
    this.status.textContent = writable
      ? 'Your saved versions stay on this device.'
      : 'You’re browsing a history file. Your open document is safe.';
  }

  private async showPreview(bytes: Uint8Array, title: string): Promise<void> {
    const doc = this.previewDialog.ownerDocument;
    const holder = doc.createElement('div'); holder.className = 'dxr-version-paper';
    const preview = await this.options.preview(holder, bytes);
    if (this.destroyed) { preview.destroy(); return; }
    const heading = doc.createElement('h2'); heading.textContent = title;
    const actions = doc.createElement('div'); actions.className = 'dxr-version-actions';
    this.button('Back to version history', () => this.previewDialog.close(), actions);
    if (!this.archive) this.button('Use as draft', () => this.command(async () => {
      if (!this.confirmReplace()) return;
      this.install(bytes, this.identity); this.dirty = true; this.previewDialog.close(); this.dialog.close();
    }), actions);
    this.button('Download this version', () => download(bytes, 'version.docx', doc), actions);
    this.preview?.destroy(); this.preview = preview;
    this.previewDialog.replaceChildren(heading, actions, holder);
    if (!this.previewDialog.open) this.previewDialog.showModal();
  }

  private async importArchive(): Promise<void> {
    if (!this.archive || !this.confirmReplace()) return;
    await this.ensureStorage();
    const imported = await this.client!.importHistoryArchive(this.archive.bytes);
    const identity = { id: imported.archive.documentId, name: this.archive.name.replace(/\.docxhistory$/i, '.docx') };
    const document = this.client!.document(identity.id);
    const checkpoints = await HistoryCheckpoints.open(document, this.store!.journal(identity.id), imported.view);
    this.install(await document.exportDocx(imported.view.version.id), identity);
    this.draft = { ...identity, checkpoints }; this.remember();
    await this.closeArchive(); await this.mountPanel(); this.updateRecent();
  }

  private async closeArchive(): Promise<void> {
    this.panelObserver.disconnect();
    const panel = this.panel; this.panel = undefined;
    const archive = this.archive; this.archive = undefined;
    await panel?.destroy();
    archive?.reader.close();
  }

  private remember(): void {
    localStorage.setItem(this.prefix + this.identity.id, JSON.stringify(this.identity));
    localStorage.setItem(this.lastKey, JSON.stringify(this.identity));
  }

  private readIdentity(value: string | null): DocumentIdentity | null {
    if (!value) return null;
    const identity = JSON.parse(value);
    if (typeof identity?.id !== 'string' || !identity.id || typeof identity?.name !== 'string') return null;
    return { id: identity.id, name: identity.name };
  }

  private updateRecent(): void {
    this.recent.replaceChildren(new Option('Choose a saved document…', ''));
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i)!;
      if (!key.startsWith(this.prefix)) continue;
      try { const identity = this.readIdentity(localStorage.getItem(key));
        if (identity) this.recent.append(new Option(identity.name, identity.id));
      } catch { /* a damaged catalog entry does not hide other saved documents */ }
    }
    this.recent.value = this.identity.id;
    this.recent.parentElement!.hidden = this.recent.options.length <= 1;
  }

  private confirm(message: string): boolean { return this.dialog.ownerDocument.defaultView?.confirm(message) ?? false; }
  private captureState(bytes: Uint8Array): void {
    this.captured = this.generation; this.capturedVersion = this.ribbon.editor?.version ?? null; this.capturedBytes = bytes.slice();
  }
  private confirmReplace(): boolean { return !this.hasUnsavedChanges || this.confirm('Replace your open document? Save a version or download it first to keep your unsaved changes.'); }
  private explain(error: unknown): string { return historyControlError(error, this.draft?.checkpoints.hasPending); }
  private command(action: () => Promise<void>): void { void this.run(action).catch(() => {}); }

  private syncBusy(): void {
    const busy = this.busy;
    // Closing the drawer while publication runs must not enable edits that a restore would replace.
    this.ribbon.surface.inert = busy;
    const chrome = this.ribbon.element.querySelector<HTMLElement>('.dxr-chrome'); if (chrome) chrome.inert = busy;
    this.recent.disabled = this.resumeButton.disabled = this.back.disabled = busy;
    const file = this.ribbon.control<HTMLInputElement>('file'); if (file) file.disabled = busy;
    const create = this.ribbon.control<HTMLButtonElement>('new'); if (create) create.disabled = busy;
  }

  private async run(action: () => Promise<void>): Promise<void> {
    if (this.destroyed || this.busy) return;
    const work = Promise.resolve().then(action); this.active = work;
    this.dialog.setAttribute('aria-busy', 'true'); this.body.inert = true;
    this.syncBusy();
    this.ribbon.element.querySelector<HTMLElement>('.dxr-chrome')!.inert = true;
    this.ribbon.surface.inert = true;
    try { await work; }
    catch (error) { this.status.textContent = this.explain(error); this.ribbon.setStatus(this.explain(error)); throw error; }
    finally {
      this.active = undefined; this.body.inert = false; this.dialog.setAttribute('aria-busy', 'false');
      this.syncBusy();
      this.ribbon.element.querySelector<HTMLElement>('.dxr-chrome')?.removeAttribute('inert');
      this.ribbon.surface.inert = false;
    }
  }

  private button(label: string, action: () => void, parent: HTMLElement): HTMLButtonElement {
    const button = parent.ownerDocument.createElement('button'); button.type = 'button'; button.textContent = label;
    button.setAttribute('aria-label', label); button.addEventListener('click', action, { signal: this.events.signal });
    parent.append(button); return button;
  }
}

function download(bytes: Uint8Array, name: string, doc: Document): void {
  const url = URL.createObjectURL(new Blob([bytes.slice()], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
  const link = doc.createElement('a'); link.href = url; link.download = name; link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function sameBytes(left: Uint8Array, right: Uint8Array): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

const HISTORY_DRAWER_CSS = `
.dxr-history-dialog,.dxr-version-preview{box-sizing:border-box;color:#243b42;background:#fff;border:1px solid #dce5e5;box-shadow:0 24px 80px #163f4033;font:14px/1.5 system-ui,sans-serif;padding:24px;overscroll-behavior:contain}
.dxr-history-dialog{width:min(420px,calc(100vw - 24px));max-height:calc(100dvh - 24px);margin:12px 12px 12px auto;border-radius:18px}
.dxr-history-dialog::backdrop,.dxr-version-preview::backdrop{background:#16323333;backdrop-filter:blur(2px)}
.dxr-history-dialog [hidden],.dxr-version-preview [hidden]{display:none!important}
.dxr-history-dialog button,.dxr-version-preview button{font:inherit;color:inherit;border:1px solid #cbd9d9;background:#fff;border-radius:9px;min-height:40px;padding:8px 12px;cursor:pointer}
.dxr-history-dialog button:hover,.dxr-version-preview button:hover{background:#eff8f6}
.dxr-history-dialog :focus-visible,.dxr-version-preview :focus-visible{outline:3px solid #0f766e;outline-offset:3px}
.dxr-history-header{display:flex;justify-content:flex-end;position:sticky;top:0;z-index:2;background:#fff;box-shadow:0 0 0 6px #fff}
.dxr-history-dialog .dxr-history-close{border:0;font-size:13px;background:#fff}
.dxr-history-source{font-weight:600;overflow-wrap:anywhere;margin:12px 0 4px}
.dxr-history-status{color:#597273;margin:4px 0 18px}
.dxr-history-dialog>label{display:block;font-size:12px;color:#597273;margin:12px 0}
.dxr-history-dialog>label select{display:block;width:100%;font:inherit;padding:8px;border:1px solid #cbd9d9;border-radius:8px;background:#fff;color:#243b42}
.dxr-history-dialog .dx-history{padding:0;border:0;border-radius:0;background:transparent}
.dxr-history-dialog .dx-history h2{font-size:24px;letter-spacing:-.6px;margin:12px 0}
.dxr-history-dialog .dx-history button[data-history-action=save]{background:#0f766e;border-color:#0f766e;color:#fff;font-weight:600}
.dxr-history-dialog .dx-history select{border-radius:9px;background:#f8fbfa}
.dxr-history-dialog .dx-history option{padding:10px;font-size:13px}
.dxr-version-preview{width:min(1080px,calc(100vw - 24px));max-height:calc(100dvh - 24px);border-radius:18px}
.dxr-version-preview h2{font-size:20px;margin:0 0 12px;overflow-wrap:anywhere}
.dxr-version-actions{display:flex;flex-wrap:wrap;gap:8px;margin-bottom:16px}
.dxr-version-paper{max-height:70dvh;overflow:auto;background:#f2f6f5;padding:20px;border-radius:10px}
@media(max-width:520px){.dxr-history-dialog,.dxr-version-preview{padding:16px}.dxr-version-paper{padding:8px}}
@media(prefers-reduced-motion:no-preference){.dxr-history-dialog[open]{animation:dxr-history-in .16s ease-out}@keyframes dxr-history-in{from{opacity:0;transform:translateX(16px)}to{opacity:1;transform:translateX(0)}}}
`;
