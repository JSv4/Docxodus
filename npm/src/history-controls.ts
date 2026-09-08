import { DocxHistoryError } from './history.js';
import type { DocxHistoryReader, DocxHistoryView, DocxStoredOperation, DocxStoredVersion, HistoryBlobReference } from './history.js';
import type { HistoryCheckpoints } from './history-checkpoints.js';

export interface HistoryControlsOptions {
  reader: DocxHistoryReader;
  /** Omit for a read-only archive. Must belong to reader. */
  checkpoints?: HistoryCheckpoints;
  /** Capture the open draft on Save only. The panel never replaces editor content. */
  capture?: () => Uint8Array | Promise<Uint8Array>;
  author?: string;
  documentName?: string;
  /** Render in a separate preview; reject on failure, preserving the previous document. */
  preview: (bytes: Uint8Array, title: string) => void | Promise<void>;
  /** Defaults to a browser download. */
  download?: (bytes: Uint8Array, filename: string) => void | Promise<void>;
  /** Called after an acknowledged checkpoint, before the list refreshes. Retries may return an older view. */
  onCheckpoint?: (view: DocxHistoryView, action: 'save' | 'restore' | 'retry') => void | Promise<void>;
  /** Hosts that replace the draft on restore provide a matching confirmation and acknowledgement. */
  confirmRestore?: (title: string) => boolean | Promise<boolean>;
  restoreUpdatesDraft?: boolean;
  pageSize?: number;
}

export interface HistoryControls {
  readonly element: HTMLElement;
  /** Initial metadata load. No document export or checkpoint is implicit. */
  readonly ready: Promise<void>;
  refresh(): Promise<void>;
  /** Removes controls and awaits their active call. The host may then close its reader/client. */
  destroy(): Promise<void>;
}

/** Mount explicit document-history controls. Persistence and preview/editor ownership stay with the host. */
export function mountHistoryControls(container: HTMLElement, options: HistoryControlsOptions): HistoryControls {
  const pageSize = options.pageSize ?? 25;
  if (!Number.isInteger(pageSize) || pageSize < 1 || pageSize > 100) throw new RangeError('History page size must be 1–100.');
  if (options.checkpoints && options.checkpoints.document !== options.reader)
    throw new Error('History controls and checkpoints must use the same document.');
  return new HistoryPanel(container, options, pageSize);
}

class HistoryPanel implements HistoryControls {
  readonly element: HTMLElement;
  readonly ready: Promise<void>;
  private readonly fieldset: HTMLFieldSetElement;
  private readonly status: HTMLElement;
  private readonly versions: HTMLSelectElement;
  private readonly before: HTMLSelectElement;
  private readonly detail: HTMLElement;
  private readonly author: HTMLInputElement;
  private readonly label: HTMLInputElement;
  private readonly actions = new Map<string, HTMLButtonElement>();
  private readonly events = new AbortController();
  private records: DocxStoredVersion[] = [];
  private next: HistoryBlobReference | null = null;
  private view: DocxHistoryView | null = null;
  private active: Promise<void> | null = null;
  private destroyed = false;

  constructor(container: HTMLElement, private readonly options: HistoryControlsOptions, private readonly pageSize: number) {
    const doc = container.ownerDocument;
    this.element = doc.createElement('section');
    this.element.className = 'dx-history';
    this.element.setAttribute('aria-label', 'Version history');
    const style = doc.createElement('style'); style.textContent = HISTORY_CSS;
    const title = doc.createElement('h2'); title.textContent = 'Version history';
    this.status = doc.createElement('p'); this.status.setAttribute('role', 'status'); this.status.setAttribute('aria-live', 'polite');
    this.fieldset = doc.createElement('fieldset');
    const legend = doc.createElement('legend'); legend.textContent = 'Saved versions'; this.fieldset.append(legend);
    this.element.append(style, title, this.status, this.fieldset); container.append(this.element);
    const note = doc.createElement('p');
    note.dataset.historyExisting = '';
    note.textContent = options.checkpoints ? 'Browse saved versions without changing your draft.' : 'Read-only history. Preview or download any saved version.';
    this.fieldset.append(note);
    this.button('refresh', 'Refresh history', () => this.load(true));
    this.button('latest', 'Preview latest', async () => this.preview(await options.reader.exportDocx(), 'Latest saved version'));
    this.versions = this.select('Version'); this.versions.size = 6;
    this.detail = doc.createElement('p'); this.fieldset.append(this.detail);
    this.versions.addEventListener('change', () => this.update(), { signal: this.events.signal });
    this.button('preview', 'Preview selected', async () => this.preview(await options.reader.exportDocx(this.selected().id), versionTitle(this.selected())));
    this.button('download', 'Download selected', async () => this.download(await options.reader.exportDocx(this.selected().id), `version-${this.selected().record.sequence}.docx`));
    const comparison = this.disclosure('Compare versions');
    this.before = this.select('Compare from', comparison);
    this.before.addEventListener('change', () => this.update(), { signal: this.events.signal });
    const compareNote = doc.createElement('p'); compareNote.textContent = 'Compare from this version to the selected version above.'; comparison.append(compareNote);
    this.button('compare', 'Compare versions', async () => {
      const before = this.records[Number(this.before.value)];
      await this.preview(await options.reader.compareVersions(before.id, this.selected().id), 'Comparison with tracked changes');
    }, comparison);
    this.button('more', 'Load older versions', async () => { if (this.next) await this.page(this.next, false); });
    const saveGroup = doc.createElement('div'); saveGroup.className = 'dx-history-save';
    this.fieldset.insertBefore(saveGroup, note);
    const attribution = doc.createElement('details');
    const attributionTitle = doc.createElement('summary'); attributionTitle.textContent = 'Saved by'; attribution.append(attributionTitle);
    saveGroup.append(attribution);
    this.author = this.input('Your name', 'text', options.author ?? 'You', attribution);
    this.label = this.input('Version name (optional)', 'text', '', saveGroup);
    saveGroup.hidden = !options.checkpoints;
    this.author.parentElement!.hidden = this.label.parentElement!.hidden = !options.checkpoints;
    this.button('save', 'Save version', async () => {
      const view = await options.checkpoints!.save(await options.capture!(), this.metadata());
      await options.onCheckpoint?.(view, 'save');
      await this.load(false); this.label.value = '';
      this.status.textContent = 'Version saved. Your draft remains open.';
    }, saveGroup);
    saveGroup.append(attribution);
    this.button('restore', 'Restore selected', async () => {
      const version = this.selected();
      const confirmed = options.confirmRestore ? await options.confirmRestore(versionTitle(version))
        : doc.defaultView?.confirm(`Restore ${versionTitle(version)} as a new saved version?\n\nYour current draft and all later versions will be kept.`);
      if (!confirmed) {
        this.status.textContent = 'Restore canceled. Your draft and history are unchanged.';
        return;
      }
      const view = await options.checkpoints!.restore(version.id, this.metadata());
      await options.onCheckpoint?.(view, 'restore');
      await this.load(false);
      this.status.textContent = options.restoreUpdatesDraft ? 'Version restored. All saved versions are kept.' : 'Restored as a new saved version. Your draft and later versions are kept.';
    });
    const restoreNote = doc.createElement('p'); restoreNote.hidden = !options.checkpoints;
    if (options.checkpoints) restoreNote.dataset.historyExisting = '';
    restoreNote.textContent = options.restoreUpdatesDraft ? 'Restore returns your document to this version and keeps every saved version.' : 'Restore creates a new saved version. Preview latest to view it; your draft stays open.'; this.fieldset.append(restoreNote);
    this.button('retry', 'Retry save', async () => {
      const view = await options.checkpoints!.retry();
      await options.onCheckpoint?.(view, 'retry'); await this.load(false);
      this.status.textContent = 'Saved version recovered. Your draft remains open. Refresh history to check for newer versions.';
    });
    const sharing = this.disclosure('Download with history');
    const sharingNote = doc.createElement('p');
    sharingNote.textContent = 'Includes retained drafts and collaboration proposals. Share this file only when you want to include that history. A DOCX download keeps existing Word comments and revisions, without external history.';
    sharing.append(sharingNote);
    this.button('archive', 'Download with version history', async () => this.download(await options.reader.exportHistoryArchive(), 'docxhistory'), sharing);
    const portability = doc.createElement('p'); portability.textContent = 'Saved versions stay on this device. Download a history file to keep a portable copy; clearing browser data removes local versions.'; sharing.append(portability);
    const time = this.disclosure('Find a version by time');
    const cutoff = this.input('Saved at or before (local time)', 'datetime-local', '', time);
    this.button('time', 'Preview at time', async () => {
      if (!cutoff.value) throw new DocxHistoryError('InvalidRequest', 'Choose a date and time first.');
      const sequence = await options.reader.resolveSequenceAtTime(new Date(cutoff.value).toISOString());
      await this.preview(await options.reader.materialize(sequence), 'Version at selected time');
    }, time);
    const activity = this.disclosure('Recorded collaboration');
    const decisions = doc.createElement('ol');
    let showOlderActivity = () => {};
    this.button('activity', 'Load activity', async () => {
      const { operations } = await options.reader.readOperationsSince(null);
      const resolved = new Set(operations.filter(op => op.record.status === 'accepted').map(op => op.input.request.resolves?.digest.value));
      let shown = 0;
      showOlderActivity = () => {
        const page = operations.slice(Math.max(0, operations.length - shown - this.pageSize), operations.length - shown).reverse();
        decisions.append(...page.map(op => this.activityItem(op, resolved.has(op.id.digest.value)))); shown += page.length;
        this.actions.get('activity-more')!.hidden = shown >= operations.length;
      };
      decisions.replaceChildren(); showOlderActivity();
      this.status.textContent = operations.length ? 'Recorded activity loaded.' : 'No recorded collaboration.';
    }, activity);
    activity.append(decisions);
    this.button('activity-more', 'Load older activity', async () => showOlderActivity(), activity);
    this.actions.get('activity-more')!.hidden = true;
    this.ready = this.run('Loading history', () => this.load(false));
  }

  refresh(): Promise<void> { return this.run('Loading history', () => this.load(true)); }

  async destroy(): Promise<void> {
    this.destroyed = true; this.events.abort(); this.element.remove();
    await this.active?.catch(() => {});
  }

  private async load(refresh: boolean): Promise<void> {
    this.view = this.options.checkpoints
      ? refresh ? await this.options.checkpoints.refresh() : this.options.checkpoints.view
      : await this.options.reader.read();
    if (this.view) await this.page(this.view.version.id, true);
    else { this.records = []; this.next = null; this.versions.replaceChildren(); this.before.replaceChildren(); }
    this.status.textContent = this.options.checkpoints?.hasPending
      ? 'A save needs recovery. Retry it before saving more changes.'
      : this.view ? 'Select a version to preview, download or restore.' : 'No saved versions yet. Save your first version when you are ready.';
  }

  private async page(cursor: HistoryBlobReference, reset: boolean): Promise<void> {
    const page = await this.options.reader.listVersions(cursor, this.pageSize);
    if (reset) { this.records = []; this.versions.replaceChildren(); this.before.replaceChildren(); }
    const start = this.records.length;
    this.records.push(...page.versions); this.next = page.next;
    for (let index = start; index < this.records.length; index++) {
      for (const select of [this.versions, this.before]) {
        const option = select.ownerDocument.createElement('option'); option.value = String(index);
        option.textContent = option.title = versionTitle(this.records[index]); select.append(option);
      }
    }
    if (reset) { this.versions.value = '0'; this.before.value = this.records.length > 1 ? '1' : '0'; }
  }

  private update(): void {
    const hasVersion = this.records.length > 0;
    const pending = this.options.checkpoints?.hasPending ?? false;
    const stale = this.options.checkpoints?.needsRefresh ?? false;
    for (const key of ['latest', 'preview', 'download', 'archive', 'time', 'activity']) this.actions.get(key)!.disabled = !hasVersion;
    this.actions.get('more')!.hidden = !this.next;
    this.actions.get('compare')!.disabled = !hasVersion || this.before.value === this.versions.value;
    this.actions.get('save')!.hidden = !this.options.checkpoints || !this.options.capture;
    this.actions.get('save')!.disabled = pending || stale;
    this.actions.get('restore')!.hidden = !this.options.checkpoints;
    this.actions.get('restore')!.disabled = !hasVersion || pending || stale;
    this.actions.get('retry')!.hidden = !pending;
    this.versions.parentElement!.hidden = this.detail.hidden = !hasVersion;
    this.versions.size = Math.max(2, Math.min(6, this.records.length));
    for (const key of ['latest', 'preview', 'download']) this.actions.get(key)!.hidden = !hasVersion;
    this.actions.get('restore')!.hidden = !hasVersion || !this.options.checkpoints;
    for (const element of Array.from(this.element.querySelectorAll<HTMLElement>('[data-history-existing]'))) element.hidden = !hasVersion;
    this.detail.textContent = hasVersion ? [versionTitle(this.selected()), this.selected().record.metadata.message,
      this.selected().record.restoredFrom ? 'Restored from an earlier version.' : ''].filter(Boolean).join(' — ') : '';
  }

  private selected(): DocxStoredVersion { return this.records[Number(this.versions.value)]; }
  private activityItem(op: DocxStoredOperation, resolved: boolean): HTMLLIElement {
    const doc = this.element.ownerDocument;
    const item = doc.createElement('li');
    const state = op.record.status === 'accepted' ? 'Accepted' : resolved ? 'Conflict resolved' : 'Conflict needs review';
    const label = doc.createElement('p');
    label.textContent = `${op.input.request.metadata.author} · ${state} · ${formatTime(op.input.request.metadata.createdAt)}`;
    item.append(label);
    const description = doc.createElement('p');
    this.commandButton('View decision', async () => {
      const decision = await this.options.reader.getOperation(op.id);
      const { kind, metadata } = decision.input.request;
      description.textContent = [{ text: 'Text edit', package: 'Document edit', discard: 'Discarded proposal' }[kind],
        metadata.label, metadata.message, decision.record.conflict].filter(Boolean).join(' — ');
    }, item);
    item.append(description);
    this.commandButton('Download proposal', async () => {
      await this.download(await this.options.reader.exportOperationProposal(op.id), `proposal-${op.record.revision}.docx`);
    }, item);
    return item;
  }
  private metadata() { return { author: this.author.value.trim() || 'You', createdAt: new Date().toISOString(), label: this.label.value.trim() || undefined }; }
  private async preview(bytes: Uint8Array, title: string): Promise<void> {
    if (!this.destroyed) await this.options.preview(bytes, title);
  }
  private async download(bytes: Uint8Array, suffix: string): Promise<void> {
    if (this.destroyed) return;
    const stem = (this.options.documentName ?? 'document').replace(/\.(docx|docxhistory)$/i, '');
    const name = suffix === 'docxhistory' ? `${stem}.docxhistory` : `${stem}-${suffix}`;
    if (this.options.download) { await this.options.download(bytes, name); return; }
    const url = URL.createObjectURL(new Blob([bytes.slice()], { type: suffix === 'docxhistory' ? 'application/octet-stream' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
    const link = this.element.ownerDocument.createElement('a'); link.href = url; link.download = name;
    link.click(); setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  private async run(label: string, action: () => Promise<void>): Promise<void> {
    if (this.destroyed) throw new DocxHistoryError('Closed', 'History controls are closed.');
    if (this.active) throw new DocxHistoryError('Busy', 'A history command is still running.');
    this.fieldset.disabled = true; this.element.setAttribute('aria-busy', 'true'); this.status.textContent = `${label}…`;
    const work = Promise.resolve().then(action); this.active = work;
    try { await work; if (this.status.textContent === `${label}…`) this.status.textContent = 'Ready.'; }
    catch (error) {
      if (!this.destroyed) this.status.textContent = historyControlError(error, this.options.checkpoints?.hasPending);
      throw error;
    } finally {
      this.active = null;
      if (!this.destroyed) { this.fieldset.disabled = false; this.element.setAttribute('aria-busy', 'false'); this.update(); }
    }
  }

  private button(key: string, title: string, action: () => Promise<void>, parent: HTMLElement = this.fieldset): void {
    const button = this.commandButton(title, action, parent); button.dataset.historyAction = key;
    this.actions.set(key, button);
  }
  private commandButton(title: string, action: () => Promise<void>, parent: HTMLElement): HTMLButtonElement {
    const button = parent.ownerDocument.createElement('button'); button.type = 'button'; button.textContent = title;
    button.addEventListener('click', () => { void this.run(title, action).catch(() => {}); }, { signal: this.events.signal });
    parent.append(button); return button;
  }
  private select(title: string, parent: HTMLElement = this.fieldset): HTMLSelectElement {
    const label = this.fieldset.ownerDocument.createElement('label'); label.textContent = title;
    const select = label.ownerDocument.createElement('select'); select.setAttribute('aria-label', title);
    label.append(select); parent.append(label); return select;
  }
  private input(title: string, type: string, value = '', parent: HTMLElement = this.fieldset): HTMLInputElement {
    const label = parent.ownerDocument.createElement('label'); label.textContent = title;
    const input = label.ownerDocument.createElement('input'); input.type = type; input.value = value;
    label.append(input); parent.append(label); return input;
  }
  private disclosure(title: string): HTMLDetailsElement {
    const details = this.fieldset.ownerDocument.createElement('details');
    details.dataset.historyExisting = '';
    const summary = details.ownerDocument.createElement('summary'); summary.textContent = title;
    details.append(summary); this.fieldset.append(details); return details;
  }
}

export function historyControlError(error: unknown, pending = false): string {
  const code = error instanceof DocxHistoryError ? error.code : '';
  if (code === 'StaleHead') return 'A newer saved version exists. Your draft is safe. Refresh history, review the newer version, then save again.';
  if (code === 'ImportConflict') return 'A different local history already exists. Open this file read-only to explore it.';
  if (code === 'InitializationUnsupported') return 'This storage cannot import history. Open read-only or choose storage that supports importing.';
  if (pending) return 'The save could not be confirmed. Your draft is safe. Retry save to recover your saved version.';
  if (code === 'ResourceLimit') return 'This history file exceeds browser processing limits, which can apply even below 64 MiB. Your document is unchanged.';
  if (code === 'UnsupportedVersion') return 'This history file uses an unsupported version. Open it with a newer app. Your document is unchanged.';
  if (code === 'InvalidManifest') return 'This history file is damaged or incomplete. Choose another copy. Your document is unchanged.';
  return `History could not be loaded. Your document is unchanged. ${error instanceof Error ? error.message : 'Please try again.'}`;
}

function versionTitle(version: DocxStoredVersion): string {
  const { metadata } = version.record;
  return `${metadata.label || 'Saved version'} · ${metadata.author} · ${formatTime(metadata.createdAt)}`;
}
function formatTime(value: string): string {
  const time = new Date(value);
  return Number.isNaN(time.getTime()) ? value : new Intl.DateTimeFormat(undefined, { dateStyle: 'medium', timeStyle: 'short' }).format(time);
}

const HISTORY_CSS = `
.dx-history{font:14px/1.5 system-ui,sans-serif;color:#203047;background:#fff;border:1px solid #cbd5e1;border-radius:12px;padding:18px;min-width:0}
.dx-history *{box-sizing:border-box}.dx-history h2{font-size:20px;margin:0 0 8px}.dx-history p{margin:8px 0;overflow-wrap:anywhere}
.dx-history fieldset{border:0;padding:0;margin:0;min-width:0}.dx-history legend{position:absolute;width:1px;height:1px;overflow:hidden;clip-path:inset(50%)}
.dx-history label{display:block;font-weight:600;margin:12px 0 6px}.dx-history input,.dx-history select{display:block;width:100%;max-width:100%;margin-top:4px;font:inherit;color:inherit;border:1px solid #94a3b8;border-radius:6px;padding:8px;background:#fff}
.dx-history button{min-height:40px;margin:4px 6px 4px 0;padding:7px 12px;border:1px solid #94a3b8;border-radius:6px;background:#f8fafc;color:inherit;font:inherit;cursor:pointer}
.dx-history button:disabled{opacity:.5;cursor:wait}.dx-history :focus-visible{outline:3px solid #2563eb;outline-offset:2px}.dx-history [hidden]{display:none}
.dx-history details{margin-top:16px;border-top:1px solid #e2e8f0;padding-top:12px}.dx-history summary{cursor:pointer;font-weight:600;padding:4px 0}
.dx-history [role=status]{min-height:42px}.dx-history ol{padding-left:22px}.dx-history option{padding:6px}
`;
