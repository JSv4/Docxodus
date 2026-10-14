// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests;

[CollectionDefinition("Cross-paragraph run timing", DisableParallelization = true)]
public sealed class CrossParagraphRunTimingCollection
{
}

/// <summary>
/// The cross-paragraph fusion decision grows linearly with a run of edited paragraphs (issue #931). When a
/// run declined to fuse, the edit-script builder retried from the next paragraph, re-segmenting nearly the
/// same run once per paragraph: a long run of edited paragraphs cost quadratic time, and an 800-paragraph
/// document took over half a minute to compare. A declined run now proves which later starts must decline
/// too, so the segmenter runs a bounded number of times however long the run is.
/// </summary>
[Collection("Cross-paragraph run timing")]
public class DocxDiffCrossParagraphRunScalingTests
{
    private const string Sentence = "Alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo.";

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DecliningRun_SegmentsABoundedNumberOfTimes(bool fused)
    {
        // Each paragraph moves a sentence from its head to its tail: every pair keeps in-slot anchors, no
        // word crosses a paragraph boundary, and the run declines to fuse. The detect-only build (fusion
        // off, recording the anchors a fused build would absorb) runs the same decision.
        int Attempts(int paragraphs)
        {
            var (left, right) = Documents(paragraphs);
            var settings = DocxCompare.ApplyFrontDoorRevisionPolicy(null).ToIrDiffSettings() with
            {
                CrossParagraphTokenDiff = fused,
            };
            var l = IrReader.Read(left, DocxDiff.ReadOpts);
            var r = IrReader.Read(right, DocxDiff.ReadOpts);
            int before = IrCrossParagraphSegmenter.AttemptsOnThisThread;
            IrEditScriptBuilder.Build(l, r, settings, fused ? null : new HashSet<string>());
            return IrCrossParagraphSegmenter.AttemptsOnThisThread - before;
        }

        int small = Attempts(20);
        int large = Attempts(80);

        Assert.True(small > 0, "the run must reach the segmenter at all");
        Assert.Equal(small, large);
    }

    [Fact]
    public void Compare_LongRunOfEditedParagraphs_FinishesWithinBudget()
    {
        var (left, right) = Documents(2000);

        var stopwatch = Stopwatch.StartNew();
        DocxCompare.Compare(left, right);
        stopwatch.Stop();

        // Over 20 minutes before the fix; a few seconds after it.
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(60), $"comparison took {stopwatch.Elapsed}");
    }

    private static (WmlDocument Left, WmlDocument Right) Documents(int paragraphs) =>
        (Build(paragraphs, i => $"{Sentence} Clause number {i} words follow here. {Sentence} The tail of clause {i} ends."),
         Build(paragraphs, i => $"Clause number {i} words follow here. {Sentence} The tail of clause {i} ends. {Sentence}"));

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
