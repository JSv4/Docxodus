# Session edit-cost benchmark

A standalone harness that measures what one `DocxSession` edit costs as a document grows, and how
much undo history the default memory budget keeps (issues #965 and #1022). It is not part of `Docxodus.sln`;
CI compiles it so a library change cannot rot it.

It builds a synthetic document of N paragraphs of 40 words each plus K embedded pictures of M MiB
each, then applies E `ReplaceText` edits to text paragraphs spread through the document, with the
per-op markdown patch on (the session default) or off (what the browser editor uses). It prints the
per-edit time, the bytes allocated per edit, the undo bytes each further step adds, how many undo
steps the default 128 MiB budget kept, and the managed heap afterwards.

## Running

```bash
cd benchmarks/session-edit-cost
dotnet run -c Release -- 4000 4 2 20 on   # paragraphs, images, MiB per image, edits, patch on|off
```

Wall-clock time is sensitive to machine load; allocation per edit and undo steps kept are
deterministic, so compare those first.

## Results

4,000 paragraphs, four 2 MiB pictures, 20 edits, default settings (20-step depth, 128 MiB budget):

| | Before #965 | After #965 |
|---|---|---|
| Allocated per edit (median) | 33.4 MiB | 15.8 MiB |
| Undo steps kept | 11 of 20 | 20 of 20 |
| Undo bytes counted against the budget | 124.7 MiB | 114.8 MiB |
| History trimmed for memory | yes | no |

Before, every edit copied all four pictures into its undo snapshot (8 MiB), and the budget charged
each snapshot for them, so it ran out after 11 edits. Now consecutive snapshots share each unchanged
picture's bytes, and the budget counts each shared array once.

### Block-local index refresh and shared snapshots (#1022)

Without pictures, the edit's own bookkeeping was the next cost that grew with the document: every
edit deep-copied every part into its undo snapshot, and the next lookup rebuilt the anchor index from
every part. Both now cost what the edit touched (see "What one edit costs" in
`docs/architecture/docx_mutation_api.md`). Allocated per `ReplaceText`, median of 20, no pictures,
patch **off**:

| Paragraphs | Before | After |
|---|---|---|
| 1,000 | 997 KiB | 23 KiB |
| 2,000 | 1,983 KiB | 23 KiB |
| 4,000 | 3,734 KiB | 25 KiB |
| 8,000 | 7,466 KiB | 26 KiB |

Undo retention per step went from about 1.3 KiB per paragraph (5.3 MiB at 4,000, 10.6 MiB at 8,000,
where the budget kept only 12 of 20 steps) to about 1.5 KiB per step regardless of size; the budget now
counts one shared copy of the document plus those deltas, and keeps all 20 steps at 8,000.

With the patch **on** (the default), each edit used to re-project the whole document for
`EditResult.Patch`; the patch now carries only the blocks the edit changed:

| Paragraphs | Before | After |
|---|---|---|
| 1,000 | 5,143 KiB | 27 KiB |
| 2,000 | 10,861 KiB | 28 KiB |
| 4,000 | 21,508 KiB | 29 KiB |
| 8,000 | 43,022 KiB | 31 KiB |

A document whose paragraphs are list items pays one more linear term: after each edit the session
re-verifies list numbering over every numbered paragraph (about 1.5 MiB per edit at 4,000 paragraphs,
half of them numbered).
