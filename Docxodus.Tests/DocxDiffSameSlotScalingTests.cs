// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Block alignment's same-slot pass grows linearly with a gap's size (issue #937). Its competitor-evidence
/// guard compared each slot pair against every still-free paragraph on both sides, so a gap of K leftover
/// paragraphs cost K² word-set comparisons. The guard now visits only paragraphs that hold one of the slot's
/// rarest content words, which an exact counting argument shows reaches every paragraph that could outbid it.
/// </summary>
public class DocxDiffSameSlotScalingTests
{
    private const string Sentence = "Alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo.";

    public static TheoryData<string> Shapes => new() { "moved sentence", "rewritten clauses" };

    [Theory]
    [MemberData(nameof(Shapes))]
    public void CompetitorGuard_ChecksALinearNumberOfParagraphs(string shape)
    {
        int Checks(int paragraphs)
        {
            var (left, right) = Documents(shape, paragraphs);
            var settings = DocxCompare.ApplyFrontDoorRevisionPolicy(null).ToIrDiffSettings();
            var l = IrReader.Read(left, DocxDiff.ReadOpts);
            var r = IrReader.Read(right, DocxDiff.ReadOpts);
            int before = IrBlockAligner.CompetitorChecksOnThisThread;
            IrBlockAligner.Align(l, r, settings);
            return IrBlockAligner.CompetitorChecksOnThisThread - before;
        }

        int small = Checks(40);
        int large = Checks(160);

        // Four times the paragraphs: at most four times the checks (a quadratic guard makes sixteen).
        Assert.True(large <= 4 * small + 8, $"{small} checks at 40 paragraphs, {large} at 160");
    }

    [Fact]
    public void CompetitorGuard_StillProbesWhenASlotPairLeavesWordsUnshared()
    {
        // Each rewritten clause holds the next clause's seal, so the guard has a real candidate to rule out.
        var (left, right) = Documents("rewritten clauses", 40);
        var settings = DocxCompare.ApplyFrontDoorRevisionPolicy(null).ToIrDiffSettings();
        int before = IrBlockAligner.CompetitorChecksOnThisThread;
        IrBlockAligner.Align(IrReader.Read(left, DocxDiff.ReadOpts), IrReader.Read(right, DocxDiff.ReadOpts), settings);

        Assert.True(IrBlockAligner.CompetitorChecksOnThisThread - before > 0);
    }

    [Theory]
    [MemberData(nameof(Shapes))]
    public void SameSlotPass_PairsEveryClauseWithItsCounterpart(string shape)
    {
        var (left, right) = Documents(shape, 40);
        var settings = DocxCompare.ApplyFrontDoorRevisionPolicy(null).ToIrDiffSettings();
        var l = IrReader.Read(left, DocxDiff.ReadOpts);
        var r = IrReader.Read(right, DocxDiff.ReadOpts);
        var alignment = IrBlockAligner.Align(l, r, settings);

        int modified = 0;
        foreach (var entry in alignment.Entries)
            if (entry.Kind == IrAlignmentKind.Modified)
            {
                Assert.Equal(IndexOf(l.Body.Blocks, entry.Left!), IndexOf(r.Body.Blocks, entry.Right!));
                modified++;
            }
        Assert.Equal(40, modified);
    }

    [Fact]
    public void CompetitorIndex_AgreesWithScanningEveryParagraph()
    {
        // The index answers exactly what the old scan did, for every target, evidence level, slot partner
        // and set of already-paired paragraphs. Paragraphs draw from a skewed vocabulary that mixes function
        // words, boilerplate every paragraph holds, and rare words.
        var random = new Random(937);
        string[] vocabulary =
        {
            "the", "of", "and", "shall", "licensee", "records", "term", "notice", "party", "agreement",
            "vault", "store", "keep", "ledger", "audit", "breach", "cure", "venue", "seal", "waiver",
        };
        string Sentence(int i)
        {
            var words = new List<string> { "licensee", "the" };
            int count = random.Next(1, 9);
            for (int w = 0; w < count; w++)
                words.Add(vocabulary[Math.Min(random.Next(vocabulary.Length) * random.Next(1, 3) / 2, vocabulary.Length - 1)]);
            return string.Join(" ", words) + ".";
        }

        const int paragraphs = 30;
        var leftTexts = Enumerable.Range(0, paragraphs).Select(Sentence).ToArray();
        var rightTexts = Enumerable.Range(0, paragraphs).Select(Sentence).ToArray();
        var settings = DocxCompare.ApplyFrontDoorRevisionPolicy(null).ToIrDiffSettings();
        var left = IrReader.Read(Build(paragraphs, i => leftTexts[i]), DocxDiff.ReadOpts).Body.Blocks;
        var right = IrReader.Read(Build(paragraphs, i => rightTexts[i]), DocxDiff.ReadOpts).Body.Blocks;
        var similarity = new IrBlockSimilarity(settings);
        var candidates = Enumerable.Range(0, left.Count).Where(i => left[i] is IrParagraph).ToList();
        var index = new IrBlockAligner.ContentWordIndex(left, candidates, similarity);

        int disagreements = 0, outbids = 0;
        for (int trial = 0; trial < 2000; trial++)
        {
            var target = (IrParagraph)right[random.Next(paragraphs)];
            int partner = candidates[random.Next(candidates.Count)];
            int evidence = random.Next(0, 8);
            var match = new int[left.Count];
            for (int i = 0; i < match.Length; i++)
                match[i] = random.Next(4) == 0 ? 0 : -1;

            bool scanned = candidates.Any(c => c != partner && match[c] == -1 &&
                IrBlockAligner.SharedContentWordCount((IrParagraph)left[c], target, similarity) > evidence);
            if (scanned)
                outbids++;
            if (index.Outbids(target, partner, evidence, match, similarity) != scanned)
                disagreements++;
        }

        Assert.Equal(0, disagreements);
        Assert.InRange(outbids, 200, 1800);
    }

    private static (WmlDocument Left, WmlDocument Right) Documents(string shape, int paragraphs) => shape switch
    {
        // The issue's own shape: each paragraph moves a sentence from its head to its tail, so every slot pair
        // shares all of its content words.
        "moved sentence" =>
            (Build(paragraphs, i => $"{Sentence} Clause number {i} words follow here. {Sentence} The tail of clause {i} ends."),
             Build(paragraphs, i => $"Clause number {i} words follow here. {Sentence} The tail of clause {i} ends. {Sentence}")),
        // Shared boilerplate plus words few clauses hold: each clause shares its two "keep" words with a
        // neighbour, and its revision swaps its seal for a vault and borrows the next clause's seal. That
        // borrowed seal is the rarest word the slot pair leaves unshared, so the guard probes it and rules the
        // next clause out (it shares as many words, not more).
        "rewritten clauses" =>
            (Build(paragraphs, i => $"The licensee shall {Word("keep", i + 1)} {Word("keep", i)} {Word("seal", i)} records."),
             Build(paragraphs, i => $"The licensee shall {Word("keep", i + 1)} {Word("keep", i)} {Word("vault", i)} {Word("seal", i + 1)} records.")),
        _ => throw new ArgumentOutOfRangeException(nameof(shape)),
    };

    private static int IndexOf(IrNodeList<IrBlock> blocks, IrBlock block)
    {
        for (int i = 0; i < blocks.Count; i++)
            if (ReferenceEquals(blocks[i], block))
                return i;
        return -1;
    }

    /// <summary>A word unique to <paramref name="i"/>: the stem followed by i spelled in letters.</summary>
    private static string Word(string stem, int i)
    {
        var word = new StringBuilder(stem);
        do
        {
            word.Append((char)('a' + (i % 26)));
            i /= 26;
        }
        while (i > 0);
        return word.ToString();
    }

    private static WmlDocument Build(int paragraphs, Func<int, string> text)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var body = new Body();
            for (int i = 0; i < paragraphs; i++)
                body.Append(new Paragraph(new Run(new Text(text(i)) { Space = SpaceProcessingModeValues.Preserve })));
            main.Document = new Document(body);
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults());
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }

        return new WmlDocument("d.docx", stream.ToArray());
    }
}
