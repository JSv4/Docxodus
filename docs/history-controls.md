# History controls (frontend)

`mountHistoryControls` supplies an accessible, responsive panel for an app-owned reader.
It pages through metadata, previews selected versions and comparisons, downloads DOCX or
history files, finds versions by time, and shows recorded collaboration with resolved conflicts.
Preview callbacks should render separately from the editor, so browsing never discards a draft.

```ts
import {
  initialize, openDocxHistory, openIndexedDbHistoryStore,
  HistoryCheckpoints, mountHistoryControls,
} from 'docxodus';

await initialize('/wasm/');
const store = await openIndexedDbHistoryStore('my-app-history'); // Explicit, optional local persistence.
const client = openDocxHistory(store.storage);
const document = client.document(stableDocumentId);
const checkpoints = await HistoryCheckpoints.open(document, store.journal(document.documentId));
const panel = mountHistoryControls(container, {
  reader: document, checkpoints, author: currentUserName,
  capture: () => editor.save(),
  preview: (bytes, title) => showSeparatePreview(bytes, title),
});
await panel.ready;
// On teardown: await panel.destroy(); client.close(); store.close();
```

Omit `checkpoints` and `capture` when mounting a standalone archive reader. The panel never
owns or closes its reader. It awaits preview/download callbacks and disables conflicting
controls during calls. `ready` and `refresh()` reject on errors as well as showing feedback.
The optional `download` callback lets your app choose its own download dialog.
`onCheckpoint(view, action)` reports an acknowledged save, restore or retry before refreshing
the list. Hosts can use it to mark a captured draft as saved, provided no edits arrived while
the save ran. A restore leaves the draft untouched; a retry can acknowledge an older request.

`HistoryCheckpoints` also works without the panel. Supply your own durable
`HistoryCheckpointJournal`, or use the IndexedDB store's per-document journal. Complete
requests survive reloads; Retry recovers the captured inputs, even if a later checkpoint
exists. A stale draft stays open until the user refreshes history and decides what to save.
IndexedDB data stays on this browser and can be removed by clearing site data; download a
history file for a portable copy. No automatic checkpoint, network sync or retention is installed.

The [runnable TypeScript example](../npm/examples/history.ts) pairs this panel with the ribbon
editor and a separate preview. From `npm/`, with WASM already built:

```sh
npm run build:history-example
python3 -m http.server 8088 --directory dist/wasm
```

Open `http://localhost:8088/history.html`. Open a DOCX, edit and save a checkpoint, or open one
of the [sample history files](../TestFiles/HistoryArchive/README.md). History files open read-only
first; **Resume editing this history** imports the embedded document identity. A conflicting
local history leaves the archive open read-only. **Back to my draft** returns to the existing
editor history. **Use as draft** in a saved-version preview asks before replacing unsaved work.
Malformed or unsupported files leave the current editor intact; archive size is checked before
reading upload bytes. Pending saves recover across reloads, and the example reopens their
captured draft. Ordinary unsaved edits are not autosaved.

When loading an already captured version, pass its view as the third argument to
`HistoryCheckpoints.open(document, journal, view)`. This keeps the editor's expected head
paired with its bytes, including when an import retry returns an older receipt.

Run the artifact and browser-interaction checks with:

```sh
npm run build:ts
npm run typecheck
npx playwright test history-checkpoints.spec.ts history-controls.spec.ts history-example.spec.ts --project=chromium
```

Users handle documents, not storage directories. Your app owns persistence, request IDs,
loading indicators, download dialogs and the editor; Docxodus supplies the history calls.

```ts
import { initialize, openDocxHistory, openDocxHistoryArchive } from 'docxodus';
await initialize('/wasm/');
const history = openDocxHistory(storage); // Your durable HistoryStorage adapter.
const doc = history.document(stableDocumentId);
const view = await doc.read(); // null: no history yet; no implicit checkpoint.
```

| Control | Call |
| --- | --- |
| Open latest | `await doc.exportDocx()` |
| History list | `await doc.listVersions(null, 25)` → `{ versions, next }`, newest first |
| Next page | `await doc.listVersions(page.next, 25)`; stop when `next === null` |
| Preview/download selected | `await doc.exportDocx(version.id)` |
| Compare selected pair | `await doc.compareVersions(beforeId, afterId)` → redlined DOCX |
| Save checkpoint | `await doc.createVersion(view?.head ?? null, editorBytes, metadata, requestId)` |
| Restore selected | `await doc.restoreVersion(view.head, version.id, metadata, requestId)` |
| Download with history | `await doc.exportHistoryArchive()` → `.docxhistory` bytes |

Metadata: `{ author, createdAt: new Date().toISOString(), label?, message?, applicationMetadata? }`.
Persist a request ID **and the exact head/bytes/metadata/target** before save/restore.
Retries reuse those inputs unchanged and return the original result, not necessarily latest.
Store the returned view; load its `version.id` explicitly if replacing editor content.
Restore appends history; it never modifies an open editor or deletes later versions.

Compare any pair without storing pairwise diffs:

```ts
const redline = await doc.compareVersions(beforeId, afterId); // DOCX bytes; existing revision policy.
```

For recorded collaboration: `readOperationsSince(headOrNull)` returns decisions in
`operations`; `getOperation(id)` shows one, and `exportOperationProposal(id)` returns its
exact proposed DOCX. A later accepted decision's `input.request.resolves` identifies the
conflict it resolved; historical `status: 'conflict'` alone does not mean still unresolved.

Open a history download without any storage adapter:

```ts
const reader = await openDocxHistoryArchive(new Uint8Array(await file.arrayBuffer()));
const latest = await reader.exportDocx(); // Same list/preview/time/read methods; no save/restore.
// Await pending calls before reader.close().
```

To resume editing, import into your app's store, preserving the embedded identity:

```ts
const imported = await history.importHistoryArchive(archiveBytes);
const resumed = history.document(imported.archive.documentId);
const bytes = await resumed.exportDocx(imported.view.version.id);
```

`storage.initializeHead` is optional for existing adapters, required for import. It must
atomically insert the exact head only when absent, sharing exclusion with `advanceHead`;
otherwise return the unchanged existing head. `alreadyPresent` means exact-head retry.
Never implement initialization as a read followed by an unconditional write.

UI states: disable conflicting controls while calls run; preserve unsaved drafts on
`StaleHead`. On `ImportConflict`, offer read-only opening (a different local head exists).
`InitializationUnsupported` needs a capable adapter. Malformed/unsupported/resource-limited
files must not replace the current editor. Uncertain storage errors require retrying exact
inputs. Closing a reader/client releases handles, never stored history; await active calls.

DOCX export adds **no external history**, but preserves existing Word comments/revisions.
History files carry retained drafts/proposals; treat sharing as an explicit user choice.
Byte bindings cap archives at **64 MiB** and make copies/base64 buffers; use native streams
for larger files. Conservative expanded-graph/entry/count/work limits also apply:
`ResourceLimit` can occur below 64 MiB. No transport, subscription, autosave or retention policy is installed.
See [storage contract](history-clients.md), [.NET API](history-files.md), and
[real sample archives](../TestFiles/HistoryArchive/README.md).
