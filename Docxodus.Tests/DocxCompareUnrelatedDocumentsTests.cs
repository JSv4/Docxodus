// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Diagnostics;
using System.IO;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests;

[CollectionDefinition("Unrelated-document alignment timing", DisableParallelization = true)]
public sealed class UnrelatedDocumentAlignmentTimingCollection
{
}

/// <summary>
/// Aligning two long, unrelated documents finishes in seconds (issue #863). Before the fix the split/merge
/// scan scored nearly every window of one document against every paragraph of the other, and the similarity
/// pairing rescanned its whole grid once per pair it formed: aligning this pair took over ten minutes. It
/// runs alone and times only the alignment, with a budget a loaded machine meets and the old path missed by
/// minutes.
/// </summary>
[Collection("Unrelated-document alignment timing")]
public class DocxCompareUnrelatedDocumentsTests
{
    private static readonly string TestFiles = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "../../../../TestFiles"));

    [Fact]
    public void Align_LongUnrelatedDocuments_FinishesWithinBudget()
    {
        var settings = DocxCompare.ApplyFrontDoorRevisionPolicy(null).ToIrDiffSettings();
        var original = IrReader.Read(RevisionProcessor.AcceptRevisions(
            new WmlDocument(Path.Combine(TestFiles, "HistoryArchive", "charter-collaboration-v5.docx"))), DocxDiff.ReadOpts);
        var revised = IrReader.Read(RevisionProcessor.AcceptRevisions(
            new WmlDocument(Path.Combine(TestFiles, "WC", "WC-BodyBookmarks-Before.docx"))), DocxDiff.ReadOpts);

        var stopwatch = Stopwatch.StartNew();
        IrBlockAligner.Align(original, revised, settings);
        stopwatch.Stop();

        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(60), $"alignment took {stopwatch.Elapsed}");
    }
}
