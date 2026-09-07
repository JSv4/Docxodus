import {
  initialize, getWasmExports, createBlankDocx, mountRibbon, createViewer, openDocxHistory,
  openDocxHistoryArchive, openIndexedDbHistoryStore, HistoryCheckpoints, mountHistoryControls,
  MAX_HISTORY_ARCHIVE_BYTES, DocxHistoryError, historyControlError, TrackedChangeMode,
} from '../src/embed.js';
import type { DocxHistoryArchive, DocxHistoryDocument, DocxHistoryView, DocxViewer, HistoryControls, RibbonEditor, DocxEditorExports } from '../src/embed.js';

interface Draft { document: DocxHistoryDocument; checkpoints: HistoryCheckpoints; name: string }
const element = <T extends HTMLElement>(id: string): T => document.getElementById(id) as T;
const files = element<HTMLFieldSetElement>('files');
const input = element<HTMLInputElement>('file');
const editorRoot = element<HTMLElement>('editor');
const historyRoot = element<HTMLElement>('history');
const message = element<HTMLElement>('message');
const dialog = element<HTMLDialogElement>('preview-dialog');
const resume = element<HTMLButtonElement>('resume');
const back = element<HTMLButtonElement>('back');
const usePreview = element<HTMLButtonElement>('use-preview');
let ribbon: RibbonEditor | undefined;
let panel: HistoryControls | undefined;
let draft: Draft;
let archive: { reader: DocxHistoryArchive; bytes: Uint8Array; name: string } | undefined;
let preview: DocxViewer | undefined;
let previewBytes: Uint8Array | undefined;
let dirty = false;
let editGeneration = 0;
let capturedGeneration = -1;
let busy = false;
let activeHost: Promise<void> | undefined;

async function start(): Promise<void> {
  await initialize(new URL('./', location.href).href);
  const store = await openIndexedDbHistoryStore('docxodus-history-example');
  const client = openDocxHistory(store.storage);
  const stored = localStorage.getItem('docxodus-history-draft');
  const identity: { id: string; name: string } = stored ? JSON.parse(stored) : { id: crypto.randomUUID(), name: 'Untitled.docx' };

  const makeDraft = async (id: string, name: string, view?: DocxHistoryView): Promise<Draft> => {
    const document = client.document(id);
    return { document, name, checkpoints: await HistoryCheckpoints.open(document, store.journal(id), view) };
  };
  const mountPanel = (holder: HTMLElement, current: Draft | DocxHistoryArchive, name: string): HistoryControls => {
    const writable = 'checkpoints' in current;
    return mountHistoryControls(holder, {
      reader: writable ? current.document : current, documentName: name,
      checkpoints: writable ? current.checkpoints : undefined,
      capture: writable ? () => {
        const bytes = ribbon?.save(); if (!bytes) throw new Error('Open a document first.');
        capturedGeneration = editGeneration; return bytes;
      } : undefined,
      onCheckpoint: (_view, action) => {
        // A retry can recover another tab's older request. Only a fresh save confirms this capture.
        if (action === 'save' && capturedGeneration === editGeneration) dirty = false;
        capturedGeneration = -1;
      },
      preview: showPreview,
    });
  };
  const observer = new MutationObserver(() => { files.disabled = busy || panel?.element.getAttribute('aria-busy') === 'true'; });
  function watchPanel(): void {
    observer.disconnect();
    if (panel) observer.observe(panel.element, { attributes: true, attributeFilter: ['aria-busy'] });
  }
  async function installDraft(next: Draft, bytes: Uint8Array): Promise<void> {
    const editorHolder = document.createElement('div');
    const controlsHolder = document.createElement('div');
    const candidate = mountRibbon(editorHolder, {
      exports: getWasmExports() as unknown as DocxEditorExports,
      documentName: next.name, fileActions: false, loader: false, hint: false,
      onEdit: () => { dirty = true; editGeneration++; }, onCommand: () => { dirty = true; editGeneration++; },
      trackedChanges: TrackedChangeMode.RenderInline,
    });
    let controls: HistoryControls | undefined;
    try {
      candidate.open(bytes, next.name);
      controls = mountPanel(controlsHolder, next, next.name); await controls.ready;
      localStorage.setItem('docxodus-history-draft', JSON.stringify({ id: next.document.documentId, name: next.name }));
    } catch (error) { candidate.destroy(); await controls?.destroy(); throw error; }
    await panel?.destroy(); archive?.reader.close(); archive = undefined;
    ribbon?.destroy(); ribbon = candidate; panel = controls; draft = next;
    editorRoot.replaceChildren(editorHolder); historyRoot.replaceChildren(controlsHolder);
    dirty = false; editGeneration = 0; capturedGeneration = -1; resume.hidden = back.hidden = true;
    element('source').textContent = `${next.name} — checkpoints in this browser`;
    watchPanel();
  }
  async function openFile(file: File): Promise<void> {
    if (/\.docxhistory$/i.test(file.name)) {
      if (file.size > MAX_HISTORY_ARCHIVE_BYTES) throw new DocxHistoryError('ResourceLimit', 'History archives are limited to 64 MiB.');
      const bytes = new Uint8Array(await file.arrayBuffer());
      const reader = await openDocxHistoryArchive(bytes);
      const holder = document.createElement('div');
      let controls: HistoryControls | undefined;
      try { controls = mountPanel(holder, reader, file.name); await controls.ready; }
      catch (error) { await controls?.destroy(); reader.close(); throw error; }
      await panel?.destroy(); archive?.reader.close();
      archive = { reader, bytes, name: file.name }; panel = controls;
      historyRoot.replaceChildren(holder); resume.hidden = back.hidden = false;
      element('source').textContent = `${file.name} — read-only; your draft is still open`;
      watchPanel();
    } else {
      if (!/\.docx$/i.test(file.name)) throw new Error('Choose a .docx or .docxhistory file.');
      if (!replaceDraft()) return;
      await installDraft(await makeDraft(crypto.randomUUID(), file.name), new Uint8Array(await file.arrayBuffer()));
      dirty = true;
    }
  }
  async function act(action: () => Promise<void>): Promise<void> {
    if (busy || panel?.element.getAttribute('aria-busy') === 'true') return;
    busy = files.disabled = editorRoot.inert = true;
    if (panel) panel.element.inert = true;
    message.textContent = 'Opening document history…';
    try { activeHost = action(); await activeHost; message.textContent = 'Ready. Save a checkpoint to keep your changes in this browser.'; }
    catch (error) { message.textContent = historyControlError(error); }
    finally {
      activeHost = undefined;
      busy = files.disabled = editorRoot.inert = false;
      if (panel) panel.element.inert = false;
    }
  }
  input.addEventListener('change', () => {
    const file = input.files?.[0]; input.value = ''; if (file) void act(() => openFile(file));
  });
  element('new').addEventListener('click', () => void act(async () => {
    if (replaceDraft()) await installDraft(await makeDraft(crypto.randomUUID(), 'Untitled.docx'), createBlankDocx());
  }));
  resume.addEventListener('click', () => void act(async () => {
    if (!archive || !replaceDraft()) return;
    const imported = await client.importHistoryArchive(archive.bytes);
    const next = await makeDraft(imported.archive.documentId, archive.name.replace(/\.docxhistory$/i, '.docx'), imported.view);
    // Exact-head retries can return an older receipt. Never replace its version with latest.
    await installDraft(next, await next.document.exportDocx(imported.view.version.id));
  }));
  back.addEventListener('click', () => void act(async () => {
    const holder = document.createElement('div');
    const controls = mountPanel(holder, draft, draft.name);
    try { await controls.ready; } catch (error) { await controls.destroy(); throw error; }
    await panel?.destroy(); archive?.reader.close(); archive = undefined; panel = controls;
    historyRoot.replaceChildren(holder); resume.hidden = back.hidden = true;
    element('source').textContent = `${draft.name} — checkpoints in this browser`; watchPanel();
  }));
  usePreview.addEventListener('click', () => void act(async () => {
    if (!previewBytes || archive || !replaceDraft()) return;
    await installDraft(draft, previewBytes); dirty = true; dialog.close();
  }));
  element('close-preview').addEventListener('click', () => dialog.close());
  element('download-preview').addEventListener('click', () => { if (previewBytes) download(previewBytes, 'preview.docx'); });
  element('download-draft').addEventListener('click', () => { const bytes = ribbon?.save(); if (bytes) download(bytes, draft.name); });
  window.addEventListener('beforeunload', event => { if (dirty || busy || panel?.element.getAttribute('aria-busy') === 'true') event.preventDefault(); });
  // Teardown is explicit: stop accepting commands, await the panel, then release client/store handles.
  window.addEventListener('pagehide', event => {
    if (event.persisted) return;
    observer.disconnect();
    void (async () => {
      await activeHost?.catch(() => {}); await panel?.destroy();
      archive?.reader.close(); client.close(); store.close(); ribbon?.destroy(); preview?.destroy();
    })();
  });
  const initial = await makeDraft(identity.id, identity.name);
  const pending = await store.journal(identity.id).read();
  const view = initial.checkpoints.view;
  await installDraft(initial, pending?.kind === 'save' ? pending.bytes : view ? await initial.document.exportDocx(view.version.id) : createBlankDocx());
  dirty = pending?.kind === 'save';
  files.disabled = false; message.textContent = 'Ready. Save a checkpoint to keep your changes in this browser.';
}

function replaceDraft(): boolean { return !dirty || confirm('Replace your open draft? Download it or save a checkpoint first if you want to keep your changes.'); }
async function showPreview(bytes: Uint8Array, title: string): Promise<void> {
  const holder = document.createElement('div');
  const candidate = await createViewer(holder, bytes, {
    wasmBasePath: new URL('./', location.href).href, renderTrackedChanges: true,
  });
  preview?.destroy(); preview = candidate; previewBytes = bytes;
  element('preview').replaceChildren(holder); element('preview-title').textContent = title;
  usePreview.hidden = !!archive;
  if (!dialog.open) dialog.showModal();
  element('close-preview').focus();
}
function download(bytes: Uint8Array, name: string): void {
  const url = URL.createObjectURL(new Blob([bytes.slice()], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
  const link = document.createElement('a'); link.href = url; link.download = name; link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
void start().catch(error => { message.textContent = historyControlError(error); });
