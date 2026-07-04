#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Docxodus;
using Docxodus.Tests.Ir;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// The deterministic mutation engine behind <see cref="IrDiffFuzzTests"/> (M2.3 Task 3). From a single
/// integer seed it generates a base document (paragraphs of seeded word soup + an optional 2×2 table) and
/// a small mutation list, then applies those mutations to produce the RIGHT document. Both documents are
/// emitted as DOCX bytes via <see cref="IrTestDocuments.FromBodyXml"/>.
///
/// <para><b>Determinism.</b> Every choice is drawn from a <see cref="Random"/> constructed from the seed
/// alone — no <c>DateTime.Now</c>, no unseeded RNG, no environment input. A given seed therefore yields the
/// exact same (base doc, mutation list, right doc) on every run and every machine, which is what makes a
/// failing seed a one-line repro (<see cref="IrDiffFuzzTests.ReproduceCase"/>).</para>
/// </summary>
internal static class DiffFuzzer
{
    private static readonly string[] WordBank =
    {
        "alpha", "bravo", "charlie", "delta", "echo", "foxtrot", "golf", "hotel", "india", "juliet",
        "kilo", "lima", "mike", "november", "oscar", "papa", "quebec", "romeo", "sierra", "tango",
        "uniform", "victor", "whiskey", "xray", "yankee", "zulu", "lorem", "ipsum", "dolor", "amet",
        "consectetur", "adipiscing", "elit", "tempor", "incididunt", "labore", "magna", "aliqua",
    };

    // ---------------------------------------------------------------------- the mutation vocabulary

    /// <summary>The mutation kinds the fuzzer can apply. See <see cref="Mutation"/> for the operand shape.</summary>
    public enum MutationKind
    {
        /// <summary>Replace one word in one body paragraph with a different bank word. (Comparable class.)</summary>
        EditWord,

        /// <summary>Insert a fresh word-soup paragraph at a body index. (Comparable class.)</summary>
        InsertParagraph,

        /// <summary>Delete one body paragraph. (Comparable class.)</summary>
        DeleteParagraph,

        /// <summary>Move one body paragraph to another body index. (NOT comparable — move semantics differ.)</summary>
        RelocateParagraph,

        /// <summary>Wrap one word's run in <c>w:b</c> (bold it). (NOT comparable — format reporting differs.)</summary>
        BoldWord,

        /// <summary>Replace the text of one table cell. (Comparable class.)</summary>
        EditTableCell,

        /// <summary>Append a new row to the table. (Comparable class.)</summary>
        InsertRow,

        /// <summary>Delete a table row. (Comparable class.)</summary>
        DeleteRow,

        /// <summary>Replace the text of the document's footnote. (Comparable class — WmlComparer diffs
        /// footnote scopes via GetRevisions, so a footnote-text edit is cross-engine comparable.)</summary>
        EditFootnote,

        /// <summary>Split one body paragraph at a word boundary into two paragraphs (M2.6). (NOT
        /// comparable — the engines frame a clean split differently BY CONSTRUCTION: WmlComparer
        /// reports the tail as a real Deleted+Inserted pair of IDENTICAL text (its delete-and-reinsert
        /// account), while the IR's SplitBlock keeps the content Equal and reports only the structural
        /// mark — so the new side's content bag is legitimately empty and the differential's new-empty
        /// regression gate false-alarms (500-seed artifact evidence, 2026-06-12: e.g. seed 5,
        /// old = Del+Ins "labore foxtrot", new = the SplitBlock op). The RelocateParagraph precedent.
        /// Own-oracle coverage — apply-verify, JSON round-trip, determinism — still runs on every seed.)</summary>
        SplitParagraph,

        /// <summary>Merge two adjacent body paragraphs into one (space-joined) (M2.6). (NOT comparable —
        /// the byte-mirror of <see cref="SplitParagraph"/>'s framing difference.)</summary>
        MergeParagraphs,
    }

    /// <summary>
    /// One mutation with its (already-resolved) operands. <see cref="Describe"/> renders it for the
    /// failure repro dump. Operand fields are kind-specific; irrelevant ones stay at their defaults.
    /// </summary>
    public readonly record struct Mutation(
        MutationKind Kind,
        int Index = -1,        // paragraph/row index, or source index for relocate
        int Target = -1,       // destination index (relocate / insert position), or cell/word index
        string? Payload = null) // replacement / inserted text
    {
        public string Describe() => Kind switch
        {
            MutationKind.EditWord => $"EditWord(para={Index}, word={Target}, -> \"{Payload}\")",
            MutationKind.InsertParagraph => $"InsertParagraph(at={Index}, \"{Payload}\")",
            MutationKind.DeleteParagraph => $"DeleteParagraph(para={Index})",
            MutationKind.RelocateParagraph => $"RelocateParagraph(from={Index}, to={Target})",
            MutationKind.BoldWord => $"BoldWord(para={Index}, word={Target})",
            MutationKind.EditTableCell => $"EditTableCell(row={Index}, col={Target}, \"{Payload}\")",
            MutationKind.InsertRow => $"InsertRow(at={Index}, \"{Payload}\")",
            MutationKind.DeleteRow => $"DeleteRow(row={Index})",
            MutationKind.EditFootnote => $"EditFootnote(-> \"{Payload}\")",
            MutationKind.SplitParagraph => $"SplitParagraph(para={Index}, atWord={Target})",
            MutationKind.MergeParagraphs => $"MergeParagraphs(first={Index})",
            _ => Kind.ToString(),
        };
    }

    /// <summary>
    /// A fully-resolved fuzz case: the seed, the two documents' byte wrappers, and the metadata needed to
    /// reproduce + classify it. <see cref="IsComparableClass"/> gates the differential check.
    /// </summary>
    public sealed record FuzzCase(
        int Seed,
        WmlDocument Left,
        WmlDocument Right,
        int BaseParagraphCount,
        bool HasTable,
        IReadOnlyList<Mutation> Mutations)
    {
        /// <summary>
        /// True iff EVERY mutation is in the cross-engine-comparable class. Relocate (move semantics differ
        /// between engines), BoldWord (format-change reporting differs), and SplitParagraph/MergeParagraphs
        /// (the engines frame a clean split/merge differently by construction — see the enum docs) make the
        /// whole case non-comparable, so the differential check is SKIPPED when any is present.
        /// </summary>
        public bool IsComparableClass =>
            Mutations.All(m => m.Kind is not (MutationKind.RelocateParagraph or MutationKind.BoldWord
                or MutationKind.SplitParagraph or MutationKind.MergeParagraphs));

        public string DescribeMutations() => string.Join("; ", Mutations.Select(m => m.Describe()));
    }

    /// <summary>
    /// A K-reviewer composite fuzz case (Task 5.2): one shared base plus <c>reviewerCount</c> reviewer
    /// documents, each = the base with a DISJOINT partition of comparable mutations applied (reviewer
    /// <c>i</c> edits only body paragraphs where <c>index % reviewerCount == i</c>). Built for the
    /// consolidate own-oracle — round-trip (reject ≡ base) and the composite apply-verifier — so the
    /// reviewer mutation pool is paragraph-only (<see cref="PickCompositeMutation"/>:
    /// EditWord / InsertParagraph / DeleteParagraph); EditTableCell / EditFootnote and the non-comparable
    /// Relocate / Bold / Split / Merge are all excluded for v1.
    /// </summary>
    /// <param name="Seed">The seed that deterministically generated this case.</param>
    /// <param name="Base">The shared base document bytes every reviewer revised.</param>
    /// <param name="Reviewers">Per-reviewer (author name, revised-document bytes), in reviewer order.</param>
    /// <param name="AppliedMutations">Per-reviewer list of the mutations that ACTUALLY changed that
    /// reviewer's document (parallel to <see cref="Reviewers"/>) — the repro/classification metadata.</param>
    public sealed record CompositeFuzzCase(
        int Seed,
        byte[] Base,
        (string Author, byte[] Doc)[] Reviewers,
        IReadOnlyList<IReadOnlyList<Mutation>> AppliedMutations)
    {
        public string DescribeReviewer(int i) =>
            $"{Reviewers[i].Author}: " + string.Join("; ", AppliedMutations[i].Select(m => m.Describe()));
    }

    // ---------------------------------------------------------------------- the document model

    /// <summary>A run inside a paragraph: text plus an optional bold flag (the only modeled format the fuzzer sets).</summary>
    private sealed class Run
    {
        public string Text = "";
        public bool Bold;
    }

    private sealed class Para
    {
        public List<Run> Runs = new();
        public string Text => string.Concat(Runs.Select(r => r.Text));

        public static Para Words(IEnumerable<string> words)
        {
            var p = new Para();
            // One run per word, space-joined: word, space, word, space … so a single-word edit / bold
            // touches exactly one run (mirrors how real authored docs split runs on edits).
            var list = words.ToList();
            for (int i = 0; i < list.Count; i++)
            {
                if (i > 0)
                    p.Runs.Add(new Run { Text = " " });
                p.Runs.Add(new Run { Text = list[i] });
            }
            return p;
        }

        /// <summary>The word runs (the non-space runs), in order, for word-addressed mutations.</summary>
        public List<Run> WordRuns => Runs.Where(r => r.Text.Trim().Length > 0).ToList();
    }

    /// <summary>A 2-column table modeled as a list of rows, each row a list of cell texts.</summary>
    private sealed class Table
    {
        public List<List<string>> Rows = new();
        public int Cols => Rows.Count > 0 ? Rows[0].Count : 0;
    }

    /// <summary>The mutable document model the fuzzer mutates, then serializes to body XML.</summary>
    private sealed class DocModel
    {
        public List<Para> Paragraphs = new();
        public Table? Table; // appended after the paragraphs when present
        public string? FootnoteText; // a single footnote (id=1) when present; the EditFootnote mutation edits it

        /// <summary>
        /// A deep, alias-free copy of the model. <see cref="Apply"/> mutates the model in place, so the
        /// composite fuzzer clones the shared base ONCE PER REVIEWER before applying that reviewer's
        /// partition of mutations — without this, one reviewer's edits would leak into the next reviewer's
        /// (and the base) document. Every Para/Run/Table-row container is reconstructed; no list, paragraph,
        /// or run instance is shared with the source.
        /// </summary>
        public DocModel Clone()
        {
            var copy = new DocModel { FootnoteText = FootnoteText };
            foreach (var p in Paragraphs)
                copy.Paragraphs.Add(ClibPara(p));
            if (Table is { } t)
            {
                copy.Table = new Table();
                foreach (var row in t.Rows)
                    copy.Table.Rows.Add(new List<string>(row));
            }
            return copy;
        }
    }

    // ---------------------------------------------------------------------- generation

    /// <summary>Build the fully-resolved fuzz case for <paramref name="seed"/> (base doc + mutations + right doc).</summary>
    public static FuzzCase Generate(int seed)
    {
        var rng = new Random(seed);

        var model = GenerateBase(rng, out int baseParaCount, out bool hasTable);
        var leftBytes = Serialize(model);

        var mutations = GenerateMutations(rng, model);
        var applied = new List<Mutation>();
        foreach (var m in mutations)
            if (Apply(model, m))
                applied.Add(m);

        var rightBytes = Serialize(model);

        return new FuzzCase(seed, leftBytes, rightBytes, baseParaCount, hasTable, applied);
    }

    /// <summary>
    /// Build a K-reviewer composite case for <paramref name="seed"/>: one shared base, and
    /// <paramref name="reviewerCount"/> reviewer documents each = a DEEP CLONE of the base with a disjoint
    /// partition of comparable mutations applied. Reviewer <c>i</c> edits only the base paragraphs in its
    /// residue class (<c>index % reviewerCount == i</c>); each reviewer applies 1–3 mutations restricted to
    /// the cross-reviewer-clean comparable kinds. A small deliberate-collision chance lets two reviewers
    /// touch the same paragraph — the round-trip oracle still holds.
    /// </summary>
    public static CompositeFuzzCase GenerateComposite(int seed, int reviewerCount) =>
        GenerateComposite(seed, reviewerCount, keepStructure: false);

    /// <summary>
    /// Like <see cref="GenerateComposite(int,int)"/>, but KEEPS any table / footnote the base generated
    /// instead of stripping them. The reviewer MUTATIONS are still restricted to the paragraph-only pool
    /// (EditWord / InsertParagraph / DeleteParagraph) so the merge stays on the supported v1 path — the
    /// point is to fuzz that the BASE's table/footnote STRUCTURE survives a consolidate + reject, not that
    /// table/footnote edits merge. Consumed by the composite STRUCTURAL round-trip test, which compares a
    /// table-aware projection of the body (not the paragraph-only <c>Docs.PlainText</c>).
    /// </summary>
    public static CompositeFuzzCase GenerateCompositeWithStructure(int seed, int reviewerCount) =>
        GenerateComposite(seed, reviewerCount, keepStructure: true);

    /// <summary>
    /// Like <see cref="GenerateComposite(int,int)"/>, but the reviewer mutation pool ADDITIONALLY includes
    /// the STRUCTURAL kinds (<see cref="MutationKind.RelocateParagraph"/> / <see cref="MutationKind.SplitParagraph"/>
    /// / <see cref="MutationKind.MergeParagraphs"/>) so a consolidate exercises the merger's
    /// <c>LowerStructuralOps</c> lowering + contested-relocation branch — paths the paragraph-only pool never
    /// reaches. The base stays paragraph-only (no table/footnote, like the default composite set) so the
    /// body-text apply-verifier oracle applies. Consumed by the structural-op composite fuzz test, which
    /// asserts reject ≡ base AND no content loss on accept (the apply-verifier).
    /// </summary>
    public static CompositeFuzzCase GenerateCompositeWithStructuralOps(int seed, int reviewerCount) =>
        GenerateComposite(seed, reviewerCount, keepStructure: false, keepStructuralOps: true);

    /// <summary>
    /// FOLLOW-ON B structural-fuzz generator: a base WITH a multi-row/-cell table, where DIFFERENT reviewers
    /// edit DIFFERENT table cells (and occasionally append a row), so each consolidate drives the per-cell
    /// table-composition path — disjoint cross-reviewer cell edits composing inline. Each reviewer's edits are
    /// confined to its CELL partition (cell linear index <c>row*cols+col % reviewerCount == reviewer</c>) so
    /// edits stay disjoint (the composed-table happy path); a row append by one reviewer routes by preceding
    /// base row. The base carries NO footnote (the composite path does not merge notes) but DOES keep its
    /// table; the body also has paragraphs so the surrounding structure is exercised. Consumed by the
    /// composite table fuzz test, which asserts reject ≡ base (structural, table-aware) AND the table-aware
    /// apply-verifier over the composed table.
    /// </summary>
    public static CompositeFuzzCase GenerateCompositeDisjointTableCells(int seed, int reviewerCount)
    {
        var rng = new Random(seed);
        var baseModel = GenerateBase(rng, out _, out _);
        baseModel.FootnoteText = null; // notes are not composed in the consolidate path

        // Force a table sized so every reviewer gets at least one cell in its partition. Columns are fixed at
        // 2 to match the fuzzer's InsertRow apply (which appends a 2-cell row) — a non-2 column count would
        // produce a ragged table when a reviewer appends a row, which is a separate (column-structure) path.
        // Rows are chosen so rows*2 >= reviewerCount, with multi-word cells (so cell edits are meaningful).
        const int cols = 2;
        int rows = Math.Max(2, (reviewerCount + 1) / cols + rng.Next(1, 3));
        var table = new Table();
        for (int r = 0; r < rows; r++)
        {
            var row = new List<string>();
            for (int c = 0; c < cols; c++)
                row.Add(string.Join(" ", SoupWords(rng, rng.Next(2, 5))));
            table.Rows.Add(row);
        }
        baseModel.Table = table;

        var baseBytes = Serialize(baseModel).DocumentByteArray;

        var reviewers = new (string Author, byte[] Doc)[reviewerCount];
        var appliedPerReviewer = new List<IReadOnlyList<Mutation>>(reviewerCount);
        for (int i = 0; i < reviewerCount; i++)
        {
            var model = baseModel.Clone();
            var applied = new List<Mutation>();
            int count = rng.Next(1, 4);
            for (int k = 0; k < count; k++)
            {
                var m = PickDisjointTableMutation(rng, model, i, reviewerCount);
                if (m is { } mut && Apply(model, mut))
                    applied.Add(mut);
            }
            reviewers[i] = ($"Reviewer{i + 1}", Serialize(model).DocumentByteArray);
            appliedPerReviewer.Add(applied);
        }
        return new CompositeFuzzCase(seed, baseBytes, reviewers, appliedPerReviewer);
    }

    /// <summary>
    /// Pick one DISJOINT table mutation for reviewer <paramref name="reviewer"/>: an EditTableCell confined to
    /// a cell in this reviewer's partition (cell linear index <c>% reviewerCount == reviewer</c>), or — less
    /// often — an InsertRow (routed by preceding base row, so two reviewers' inserted rows both land without
    /// conflict). Returns null when no partition cell exists. Stays disjoint so the composed-table happy path
    /// is exercised (no same-cell conflict).
    /// </summary>
    private static Mutation? PickDisjointTableMutation(Random rng, DocModel model, int reviewer, int reviewerCount)
    {
        if (model.Table is not { Rows.Count: > 0 } t)
            return null;
        int cols = t.Cols;
        int rowsN = t.Rows.Count;

        // 1-in-6: append a row (disjoint by construction — both reviewers' new rows both appear).
        if (rng.Next(6) == 0)
            return new Mutation(MutationKind.InsertRow, Index: rowsN /* append at end */, Payload: BankWord(rng));

        // Otherwise pick a cell in this reviewer's partition.
        var partitionCells = new List<(int Row, int Col)>();
        for (int r = 0; r < rowsN; r++)
            for (int c = 0; c < cols; c++)
                if ((r * cols + c) % reviewerCount == reviewer)
                    partitionCells.Add((r, c));
        if (partitionCells.Count == 0)
            return null;
        var (ri, ci) = partitionCells[rng.Next(partitionCells.Count)];
        return new Mutation(MutationKind.EditTableCell, Index: ri, Target: ci, Payload: BankWord(rng));
    }

    private static CompositeFuzzCase GenerateComposite(
        int seed, int reviewerCount, bool keepStructure, bool keepStructuralOps = false)
    {
        var rng = new Random(seed);

        var baseModel = GenerateBase(rng, out _, out _);

        if (!keepStructure)
        {
            // v1 composite set excludes FOOTNOTES entirely. The footnote-reference run tokenizes to an
            // atomic NoteRef token (MatchKey "fn") that the apply-verifier's text view counts but the test's
            // body-only Docs.PlainText (which reads w:t only) does not — a text-projection asymmetry in the
            // harness, not a consolidate defect. (The composite merger also does not yet merge note-scope
            // content — NoteOps is always null — so a reviewer footnote edit would be silently dropped; that
            // is tracked for a later note-consolidation task.) Strip the footnote from the base so neither
            // the base nor any reviewer carries one, keeping the own-oracle clean. EditFootnote is likewise
            // absent from PickCompositeMutation's pool.
            baseModel.FootnoteText = null;

            // Same asymmetry for TABLES: the apply-verifier reconstructs a non-paragraph block as its
            // content hash (BlockText → ContentHash.ToHex()), while the test's body oracle Docs.PlainText
            // reads only the body's direct-child w:p elements (table cell paragraphs are excluded). A table
            // can therefore never match in the apply-verifier's text projection, so the v1 composite set
            // carries no table and EditTableCell is omitted from the pool. Round-trip (reject ≡ base) still
            // validates the full body including any structure; only this body-text apply-verifier needs the
            // paragraph-only shape.
            baseModel.Table = null;
        }
        // When keepStructure is true the base's table/footnote (if generated) are retained. Reviewer
        // mutations stay paragraph-only regardless (PickCompositeMutation's pool), so a consolidate runs on
        // the supported v1 path while a table-aware structural oracle asserts the base structure survives.

        var baseBytes = Serialize(baseModel).DocumentByteArray;

        var reviewers = new (string Author, byte[] Doc)[reviewerCount];
        var appliedPerReviewer = new List<IReadOnlyList<Mutation>>(reviewerCount);

        for (int i = 0; i < reviewerCount; i++)
        {
            // Deep-copy the base so this reviewer's in-place mutations cannot leak into the base or any
            // sibling reviewer.
            var model = baseModel.Clone();
            int count = rng.Next(1, 4); // 1..3 mutations per reviewer
            var applied = new List<Mutation>();
            for (int k = 0; k < count; k++)
            {
                var m = PickCompositeMutation(rng, model, i, reviewerCount, keepStructuralOps);
                if (m is { } mut && Apply(model, mut))
                    applied.Add(mut);
            }

            reviewers[i] = ($"Reviewer{i + 1}", Serialize(model).DocumentByteArray);
            appliedPerReviewer.Add(applied);
        }

        return new CompositeFuzzCase(seed, baseBytes, reviewers, appliedPerReviewer);
    }

    /// <summary>
    /// Pick one comparable mutation for reviewer <paramref name="reviewer"/>, resolved INTO that reviewer's
    /// partition (<c>paraIndex % reviewerCount == reviewer</c>). A small deliberate-collision chance instead
    /// targets a foreign paragraph so the oracle is also exercised under occasional cross-reviewer overlap.
    /// Returns null when no paragraph can be targeted (an empty partition), which the caller treats as a
    /// skipped slot. The default composite pool is paragraph-only (EditWord / InsertParagraph /
    /// DeleteParagraph) — EditTableCell and EditFootnote are excluded (see <see cref="GenerateComposite"/>).
    /// <para>When <paramref name="keepStructuralOps"/> is true the pool ALSO includes the structural kinds
    /// (RelocateParagraph / SplitParagraph / MergeParagraphs), so a consolidate exercises the merger's
    /// structural-op lowering + contested-relocation branch. Their operands are resolved against the
    /// reviewer's partition paragraph; <see cref="Apply"/> clamps every index into range (and no-ops a
    /// mutation that cannot apply, e.g. a split of a &lt;4-word paragraph), so an out-of-shape pick is simply
    /// skipped.</para>
    /// </summary>
    private static Mutation? PickCompositeMutation(
        Random rng, DocModel model, int reviewer, int reviewerCount, bool keepStructuralOps = false)
    {
        var pool = new List<MutationKind>
        {
            MutationKind.EditWord, MutationKind.EditWord,
            MutationKind.InsertParagraph,
            MutationKind.DeleteParagraph,
        };
        if (keepStructuralOps)
        {
            pool.Add(MutationKind.RelocateParagraph);
            pool.Add(MutationKind.SplitParagraph);
            pool.Add(MutationKind.MergeParagraphs);
        }
        var kind = pool[rng.Next(pool.Count)];

        // 1-in-8 deliberate collision: target a foreign paragraph; otherwise this reviewer's partition.
        bool collide = rng.Next(8) == 0;
        if (PickPartitionParagraph(rng, model, reviewer, reviewerCount, collide) is not { } para)
            return null;

        return kind switch
        {
            MutationKind.EditWord =>
                new Mutation(kind, Index: para, Target: rng.Next(0, 1_000), Payload: BankWord(rng)),
            MutationKind.InsertParagraph =>
                new Mutation(kind, Index: para, Payload: BankWord(rng)),
            MutationKind.DeleteParagraph =>
                new Mutation(kind, Index: para),
            // Relocate FROM the partition paragraph to a seeded destination (Apply resolves both modulo
            // and re-targets when from == to, so a valid move always results when ≥3 paragraphs exist).
            MutationKind.RelocateParagraph =>
                new Mutation(kind, Index: para, Target: rng.Next(0, 1_000)),
            // Split the partition paragraph at a seeded word boundary (Apply no-ops if it has <4 words).
            MutationKind.SplitParagraph =>
                new Mutation(kind, Index: para, Target: rng.Next(0, 1_000)),
            // Merge the partition paragraph with its successor (Apply resolves the index modulo Count-1).
            MutationKind.MergeParagraphs =>
                new Mutation(kind, Index: para),
            _ => null,
        };
    }

    /// <summary>
    /// Resolve a body-paragraph index in reviewer <paramref name="reviewer"/>'s residue class
    /// (<c>index % reviewerCount == reviewer</c>), or — when <paramref name="collide"/> — a foreign index.
    /// For InsertParagraph the index is an insertion position (0..Count), so the partition is interpreted
    /// over positions too. Returns null when no candidate index exists.
    /// </summary>
    private static int? PickPartitionParagraph(
        Random rng, DocModel model, int reviewer, int reviewerCount, bool collide)
    {
        int n = model.Paragraphs.Count;
        if (n == 0)
            return null;
        var candidates = new List<int>();
        for (int idx = 0; idx < n; idx++)
        {
            bool mine = idx % reviewerCount == reviewer;
            if (collide ? !mine : mine)
                candidates.Add(idx);
        }
        if (candidates.Count == 0)
            candidates = collide
                ? Enumerable.Range(0, n).ToList()              // tiny doc: fall back to any index
                : null!;
        if (candidates is null || candidates.Count == 0)
            return null;
        return candidates[rng.Next(candidates.Count)];
    }

    /// <summary>Serialize the model to DOCX bytes, routing through the footnote-carrying builder when the
    /// model has a footnote (so the footnote scope is read + diffed) and the plain body builder otherwise.</summary>
    private static WmlDocument Serialize(DocModel model) =>
        model.FootnoteText is { } fn
            ? IrTestDocuments.FromBodyXmlWithFootnote(ToBodyXml(model), fn)
            : IrTestDocuments.FromBodyXml(ToBodyXml(model));

    private static DocModel GenerateBase(Random rng, out int paraCount, out bool hasTable)
    {
        var model = new DocModel();
        paraCount = rng.Next(10, 41); // 10..40 paragraphs

        // A small pool of "boilerplate" paragraphs so the corpus has occasional exact-duplicate blocks
        // (realistic; also stresses the aligner's duplicate handling). ~1 in 5 paragraphs is boilerplate.
        var boilerplate = new List<Para>();
        for (int i = 0; i < 3; i++)
            boilerplate.Add(Para.Words(SoupWords(rng, rng.Next(4, 9))));

        for (int i = 0; i < paraCount; i++)
        {
            if (boilerplate.Count > 0 && rng.Next(5) == 0)
                model.Paragraphs.Add(ClibPara(boilerplate[rng.Next(boilerplate.Count)]));
            else
                model.Paragraphs.Add(Para.Words(SoupWords(rng, rng.Next(3, 11))));
        }

        hasTable = rng.Next(5) == 0; // ~20% chance of a table
        if (hasTable)
        {
            model.Table = new Table();
            for (int r = 0; r < 2; r++)
                model.Table.Rows.Add(new List<string>
                {
                    string.Join(" ", SoupWords(rng, rng.Next(1, 4))),
                    string.Join(" ", SoupWords(rng, rng.Next(1, 4))),
                });
        }

        // ~25% chance the document carries a footnote (so the footnote-scope diff path is exercised).
        if (rng.Next(4) == 0)
            model.FootnoteText = string.Join(" ", SoupWords(rng, rng.Next(3, 8)));

        return model;
    }

    private static List<Mutation> GenerateMutations(Random rng, DocModel model)
    {
        int count = rng.Next(1, 6); // 1..5 mutations
        var list = new List<Mutation>();
        for (int i = 0; i < count; i++)
            list.Add(PickMutation(rng, model.Table is not null, model.FootnoteText is not null));
        return list;
    }

    private static Mutation PickMutation(Random rng, bool hasTable, bool hasFootnote)
    {
        // Weighted pick. Table/footnote mutations only enter the pool when that scope exists.
        var pool = new List<MutationKind>
        {
            MutationKind.EditWord, MutationKind.EditWord,
            MutationKind.InsertParagraph,
            MutationKind.DeleteParagraph,
            MutationKind.RelocateParagraph,
            MutationKind.BoldWord,
            MutationKind.SplitParagraph,
            MutationKind.MergeParagraphs,
        };
        if (hasTable)
        {
            pool.Add(MutationKind.EditTableCell);
            pool.Add(MutationKind.InsertRow);
            pool.Add(MutationKind.DeleteRow);
        }
        if (hasFootnote)
            pool.Add(MutationKind.EditFootnote);

        var kind = pool[rng.Next(pool.Count)];
        // Operands are resolved at APPLY time against the live model (indices must be in range then), so we
        // carry a freshly-drawn seed-derived payload here and let Apply clamp/resolve indices.
        return kind switch
        {
            MutationKind.EditWord or MutationKind.InsertParagraph or MutationKind.EditTableCell
                or MutationKind.InsertRow or MutationKind.EditFootnote =>
                new Mutation(kind, Payload: BankWord(rng), Index: rng.Next(0, 1_000), Target: rng.Next(0, 1_000)),
            _ => new Mutation(kind, Index: rng.Next(0, 1_000), Target: rng.Next(0, 1_000)),
        };
    }

    // ---------------------------------------------------------------------- applying a mutation

    /// <summary>
    /// Apply <paramref name="m"/> to the model IN PLACE, resolving its placeholder indices against the
    /// current model size (modulo into range). Returns false (a no-op) when the mutation cannot apply to the
    /// current model — e.g. DeleteParagraph at 1 paragraph, or any table mutation with no table. A returned
    /// mutation in <see cref="FuzzCase.Mutations"/> is one that ACTUALLY changed the document.
    /// </summary>
    private static bool Apply(DocModel model, Mutation m)
    {
        switch (m.Kind)
        {
            case MutationKind.EditWord:
            {
                if (model.Paragraphs.Count == 0)
                    return false;
                int pi = m.Index % model.Paragraphs.Count;
                var words = model.Paragraphs[pi].WordRuns;
                if (words.Count == 0)
                    return false;
                int wi = m.Target % words.Count;
                string replacement = m.Payload!;
                if (words[wi].Text == replacement)
                    replacement += "x"; // guarantee an actual change
                words[wi].Text = replacement;
                return true;
            }

            case MutationKind.InsertParagraph:
            {
                int at = m.Index % (model.Paragraphs.Count + 1);
                // Inserted paragraph text seeded off the payload word so it's distinct-ish and reproducible.
                var inserted = Para.Words(new[] { m.Payload!, m.Payload! + "-ins", "tail" });
                model.Paragraphs.Insert(at, inserted);
                return true;
            }

            case MutationKind.DeleteParagraph:
            {
                if (model.Paragraphs.Count <= 1)
                    return false; // keep at least one paragraph
                int pi = m.Index % model.Paragraphs.Count;
                model.Paragraphs.RemoveAt(pi);
                return true;
            }

            case MutationKind.RelocateParagraph:
            {
                if (model.Paragraphs.Count < 3)
                    return false;
                int from = m.Index % model.Paragraphs.Count;
                int to = m.Target % model.Paragraphs.Count;
                if (from == to)
                    to = (to + 1) % model.Paragraphs.Count;
                var p = model.Paragraphs[from];
                model.Paragraphs.RemoveAt(from);
                model.Paragraphs.Insert(to, p);
                return true;
            }

            case MutationKind.BoldWord:
            {
                if (model.Paragraphs.Count == 0)
                    return false;
                int pi = m.Index % model.Paragraphs.Count;
                var words = model.Paragraphs[pi].WordRuns;
                var unbolded = words.Where(r => !r.Bold).ToList();
                if (unbolded.Count == 0)
                    return false;
                unbolded[m.Target % unbolded.Count].Bold = true;
                return true;
            }

            case MutationKind.EditTableCell:
            {
                if (model.Table is not { Rows.Count: > 0 } t)
                    return false;
                int ri = m.Index % t.Rows.Count;
                int ci = m.Target % t.Cols;
                string replacement = m.Payload!;
                if (t.Rows[ri][ci] == replacement)
                    replacement += "x";
                t.Rows[ri][ci] = replacement;
                return true;
            }

            case MutationKind.InsertRow:
            {
                if (model.Table is not { } t)
                    return false;
                int at = m.Index % (t.Rows.Count + 1);
                t.Rows.Insert(at, new List<string> { m.Payload! + "-r0", m.Payload! + "-r1" });
                return true;
            }

            case MutationKind.DeleteRow:
            {
                if (model.Table is not { Rows.Count: > 1 } t)
                    return false; // keep at least one row
                int ri = m.Index % t.Rows.Count;
                t.Rows.RemoveAt(ri);
                return true;
            }

            case MutationKind.EditFootnote:
            {
                if (model.FootnoteText is null)
                    return false;
                string replacement = m.Payload! + " " + (m.Payload! + "-fn");
                if (model.FootnoteText == replacement)
                    replacement += "x"; // guarantee an actual change
                model.FootnoteText = replacement;
                return true;
            }

            case MutationKind.SplitParagraph:
            {
                if (model.Paragraphs.Count == 0)
                    return false;
                int pi = m.Index % model.Paragraphs.Count;
                var words = model.Paragraphs[pi].WordRuns;
                if (words.Count < 4)
                    return false; // need ≥2 words per half for a detectable split
                int at = 1 + (m.Target % (words.Count - 1)); // split AFTER word index at-1
                var first = Para.Words(words.Take(at).Select(r => r.Text));
                var second = Para.Words(words.Skip(at).Select(r => r.Text));
                model.Paragraphs[pi] = first;
                model.Paragraphs.Insert(pi + 1, second);
                return true;
            }

            case MutationKind.MergeParagraphs:
            {
                if (model.Paragraphs.Count < 2)
                    return false;
                int pi = m.Index % (model.Paragraphs.Count - 1);
                var a = model.Paragraphs[pi];
                var b = model.Paragraphs[pi + 1];
                if (a.WordRuns.Count == 0 || b.WordRuns.Count == 0)
                    return false;
                var merged = Para.Words(a.WordRuns.Select(r => r.Text).Concat(b.WordRuns.Select(r => r.Text)));
                model.Paragraphs[pi] = merged;
                model.Paragraphs.RemoveAt(pi + 1);
                return true;
            }

            default:
                return false;
        }
    }

    // ---------------------------------------------------------------------- helpers

    private static Para ClibPara(Para src)
    {
        // Deep copy of a boilerplate paragraph so later mutations don't alias two body blocks.
        var p = new Para();
        foreach (var r in src.Runs)
            p.Runs.Add(new Run { Text = r.Text, Bold = r.Bold });
        return p;
    }

    private static IEnumerable<string> SoupWords(Random rng, int n)
    {
        for (int i = 0; i < n; i++)
            yield return WordBank[rng.Next(WordBank.Length)];
    }

    private static string BankWord(Random rng) => WordBank[rng.Next(WordBank.Length)];

    // ---------------------------------------------------------------------- serialization

    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>Serialize the model to <c>w:body</c> inner XML for <see cref="IrTestDocuments.FromBodyXml"/>.</summary>
    private static string ToBodyXml(DocModel model)
    {
        var sb = new StringBuilder();
        foreach (var p in model.Paragraphs)
            AppendParagraph(sb, p);
        if (model.Table is { } t)
            AppendTable(sb, t);
        return sb.ToString();
    }

    private static void AppendParagraph(StringBuilder sb, Para p)
    {
        sb.Append("<w:p>");
        foreach (var r in p.Runs)
        {
            sb.Append("<w:r>");
            if (r.Bold)
                sb.Append("<w:rPr><w:b/></w:rPr>");
            sb.Append("<w:t xml:space=\"preserve\">").Append(Esc(r.Text)).Append("</w:t>");
            sb.Append("</w:r>");
        }
        sb.Append("</w:p>");
    }

    private static void AppendTable(StringBuilder sb, Table t)
    {
        sb.Append("<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>");
        sb.Append("<w:tblGrid>");
        for (int c = 0; c < t.Cols; c++)
            sb.Append("<w:gridCol w:w=\"2000\"/>");
        sb.Append("</w:tblGrid>");
        foreach (var row in t.Rows)
        {
            sb.Append("<w:tr>");
            foreach (var cell in row)
            {
                sb.Append("<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>");
                sb.Append("<w:p><w:r><w:t xml:space=\"preserve\">").Append(Esc(cell)).Append("</w:t></w:r></w:p>");
                sb.Append("</w:tc>");
            }
            sb.Append("</w:tr>");
        }
        sb.Append("</w:tbl>");
    }

    private static string Esc(string s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
}
