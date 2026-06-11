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
        /// between engines) and BoldWord (format-change reporting differs) make the whole case
        /// non-comparable, so the differential check is SKIPPED when either is present.
        /// </summary>
        public bool IsComparableClass =>
            Mutations.All(m => m.Kind is not (MutationKind.RelocateParagraph or MutationKind.BoldWord));

        public string DescribeMutations() => string.Join("; ", Mutations.Select(m => m.Describe()));
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
    }

    // ---------------------------------------------------------------------- generation

    /// <summary>Build the fully-resolved fuzz case for <paramref name="seed"/> (base doc + mutations + right doc).</summary>
    public static FuzzCase Generate(int seed)
    {
        var rng = new Random(seed);

        var model = GenerateBase(rng, out int baseParaCount, out bool hasTable);
        var leftBytes = IrTestDocuments.FromBodyXml(ToBodyXml(model));

        var mutations = GenerateMutations(rng, model);
        var applied = new List<Mutation>();
        foreach (var m in mutations)
            if (Apply(model, m))
                applied.Add(m);

        var rightBytes = IrTestDocuments.FromBodyXml(ToBodyXml(model));

        return new FuzzCase(seed, leftBytes, rightBytes, baseParaCount, hasTable, applied);
    }

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

        return model;
    }

    private static List<Mutation> GenerateMutations(Random rng, DocModel model)
    {
        int count = rng.Next(1, 6); // 1..5 mutations
        var list = new List<Mutation>();
        for (int i = 0; i < count; i++)
            list.Add(PickMutation(rng, model.Table is not null));
        return list;
    }

    private static Mutation PickMutation(Random rng, bool hasTable)
    {
        // Weighted pick. Table mutations only enter the pool when a table exists.
        var pool = new List<MutationKind>
        {
            MutationKind.EditWord, MutationKind.EditWord,
            MutationKind.InsertParagraph,
            MutationKind.DeleteParagraph,
            MutationKind.RelocateParagraph,
            MutationKind.BoldWord,
        };
        if (hasTable)
        {
            pool.Add(MutationKind.EditTableCell);
            pool.Add(MutationKind.InsertRow);
            pool.Add(MutationKind.DeleteRow);
        }

        var kind = pool[rng.Next(pool.Count)];
        // Operands are resolved at APPLY time against the live model (indices must be in range then), so we
        // carry a freshly-drawn seed-derived payload here and let Apply clamp/resolve indices.
        return kind switch
        {
            MutationKind.EditWord or MutationKind.InsertParagraph or MutationKind.EditTableCell
                or MutationKind.InsertRow =>
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
