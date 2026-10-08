# Session edit-cost benchmark

A standalone harness that measures what one `DocxSession` edit costs as a document grows, and how
much undo history the default memory budget keeps (issue #965). It is not part of `Docxodus.sln`;
CI compiles it so a library change cannot rot it.

It builds a synthetic document of N paragraphs of 40 words each plus K embedded pictures of M MiB
each, then applies E `ReplaceText` edits to text paragraphs spread through the document. It prints
the per-edit time, the bytes allocated per edit, how many undo steps the default 128 MiB budget kept,
and the managed heap afterwards.

## Running

```bash
cd benchmarks/session-edit-cost
dotnet run -c Release -- 4000 4 2 20   # paragraphs, images, MiB per image, edits
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

What each edit still costs scales with the document: on the same 4,000 paragraphs without pictures,
the pre-edit snapshot is about 2 ms and 1.4 MiB, and the edit as a whole about 15 MiB, most of it
rebuilding the anchor index the edit invalidated. Scoping that work to the blocks an edit touched is
tracked in #1022.
