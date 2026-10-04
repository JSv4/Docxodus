// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Tests.Ir;
using Xunit;
using static Docxodus.Tests.DocxBackendReconciliationTests;

namespace Docxodus.Tests;

/// <summary>
/// Bookmarks keep their spans through a redline: accepting gives the revised document's bookmarks and
/// rejecting gives the original's, for a row-level bookmark in a row the diff modifies, and for a
/// reviewer's bookmark that opens in a paragraph the reviewer left unchanged (issue #864).
/// </summary>
public class DocxDiffRowAndOneSidedBookmarkTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    private static string Paragraph(string text) => $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static string Cell(string text) =>
        $"<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>{Paragraph(text)}</w:tc>";

    /// <summary>A one-row table whose row-level bookmark wraps all its cells, as Word writes a row bookmark.</summary>
    private static string RowBookmarkTable(string name, params string[] cellTexts) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid>" +
        string.Concat(cellTexts.Select(_ => "<w:gridCol w:w=\"2000\"/>")) + "</w:tblGrid>" +
        $"<w:tr><w:bookmarkStart w:id=\"7\" w:name=\"{name}\"/>" + string.Concat(cellTexts.Select(Cell)) +
        "<w:bookmarkEnd w:id=\"7\"/></w:tr></w:tbl>";

    private const string Around = "Shared opening paragraph text";

    private static XElement Body(WmlDocument document)
    {
        using var package = new ZipArchive(new MemoryStream(document.DocumentByteArray));
        using var part = package.GetEntry("word/document.xml")!.Open();
        return XElement.Load(part).Element(W + "body")!;
    }

    /// <summary>Each bookmark as "name:enclosed text", sorted; the text is every w:t between its start and its
    /// end in document order.</summary>
    private static List<string> Spans(WmlDocument document)
    {
        var open = new Dictionary<string, (string Name, System.Text.StringBuilder Text)>();
        var spans = new List<string>();
        foreach (var e in Body(document).Descendants())
        {
            if (e.Name == W + "bookmarkStart")
                open[(string)e.Attribute(W + "id")!] = ((string)e.Attribute(W + "name")!, new System.Text.StringBuilder());
            else if (e.Name == W + "bookmarkEnd" && open.Remove((string)e.Attribute(W + "id")!, out var closed))
                spans.Add($"{closed.Name}:{closed.Text}");
            else if (e.Name == W + "t")
                foreach (var openSpan in open.Values)
                    openSpan.Text.Append(e.Value);
        }
        spans.AddRange(open.Values.Select(span => $"{span.Name}:<unclosed>"));
        return spans.OrderBy(s => s, System.StringComparer.Ordinal).ToList();
    }

    private static void AssertRoundTripsBookmarks(WmlDocument original, WmlDocument revised, WmlDocument redline)
    {
        Assert.Equal(Spans(revised), Spans(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(Spans(original), Spans(RevisionProcessor.RejectRevisions(redline)));
    }

    // ------------------------------------------------------------ 1. a row-level bookmark in a modified row

    [Fact]
    public void Compare_RowBookmark_InAModifiedRow_KeepsItsSpanBothWays()
    {
        var left = IrTestDocuments.FromParts(Paragraph(Around) + RowBookmarkTable("RowMark", "Headcount", "Plan alpha"));
        var right = IrTestDocuments.FromParts(Paragraph(Around) + RowBookmarkTable("RowMark", "Headcount", "Plan beta"));

        AssertRoundTripsBookmarks(left, right, DocxCompare.Compare(left, right));
    }

    [Fact]
    public void Compare_RowBookmark_InAModifiedRow_StaysAroundTheCells()
    {
        var left = IrTestDocuments.FromParts(Paragraph(Around) + RowBookmarkTable("RowMark", "Headcount", "Plan alpha"));
        var right = IrTestDocuments.FromParts(Paragraph(Around) + RowBookmarkTable("RowMark", "Headcount", "Plan beta"));

        var row = Body(DocxCompare.Compare(left, right)).Descendants(W + "tr").Single();

        Assert.Equal(new[] { "bookmarkStart", "tc", "tc", "bookmarkEnd" },
            row.Elements().Where(e => e.Name != W + "trPr" && e.Name != W + "tblPrEx").Select(e => e.Name.LocalName));
    }

    [Fact]
    public void Compare_RowBookmark_RenamedInAModifiedRow_AcceptGivesTheRevisedAndRejectTheOriginal()
    {
        var left = IrTestDocuments.FromParts(Paragraph(Around) + RowBookmarkTable("OldRow", "Headcount", "Plan alpha"));
        var right = IrTestDocuments.FromParts(Paragraph(Around) + RowBookmarkTable("NewRow", "Headcount", "Plan beta"));

        AssertRoundTripsBookmarks(left, right, DocxCompare.Compare(left, right));
    }

    [Fact]
    public void Compare_RowBookmark_InAModifiedRow_AddsNoValidationErrors()
    {
        var left = IrTestDocuments.FromParts(Paragraph(Around) + RowBookmarkTable("OldRow", "Headcount", "Plan alpha"));
        var right = IrTestDocuments.FromParts(Paragraph(Around) + RowBookmarkTable("NewRow", "Headcount", "Plan beta"));

        NoNewValidationErrors(left.DocumentByteArray, DocxCompare.Compare(left, right).DocumentByteArray);
    }

    [Fact]
    public void Consolidate_RowBookmark_InARowAReviewerModified_KeepsItsSpanBothWays()
    {
        var original = IrTestDocuments.FromParts(Paragraph(Around) + RowBookmarkTable("RowMark", "Headcount", "Plan alpha"));
        var reviewer = IrTestDocuments.FromParts(Paragraph(Around) + RowBookmarkTable("RowMark", "Headcount", "Plan beta"));

        var consolidated = DocxDiff.Consolidate(original, new[] { new DocxDiffReviewer { Author = "Reviewer", Document = reviewer } });

        AssertRoundTripsBookmarks(original, reviewer, consolidated);
    }

    [Fact]
    public void Consolidate_RowBookmark_InARowTwoReviewersModified_StaysAroundTheCells()
    {
        var original = IrTestDocuments.FromParts(Paragraph(Around) + RowBookmarkTable("RowMark", "Headcount total", "Plan alpha"));
        var first = IrTestDocuments.FromParts(Paragraph(Around) + RowBookmarkTable("RowMark", "Headcount grand total", "Plan alpha"));
        var second = IrTestDocuments.FromParts(Paragraph(Around) + RowBookmarkTable("RowMark", "Headcount total", "Plan beta"));

        var consolidated = DocxDiff.Consolidate(original, new[]
        {
            new DocxDiffReviewer { Author = "First", Document = first },
            new DocxDiffReviewer { Author = "Second", Document = second },
        });

        var row = Body(consolidated).Descendants(W + "tr").Single();
        Assert.Equal(new[] { "bookmarkStart", "tc", "tc", "bookmarkEnd" },
            row.Elements().Where(e => e.Name != W + "trPr" && e.Name != W + "tblPrEx").Select(e => e.Name.LocalName));
        Assert.Equal(Spans(original), Spans(RevisionProcessor.RejectRevisions(consolidated)));
    }

    // ------------------------------------------- 2. a reviewer's bookmark that opens in an equal paragraph

    private static readonly WmlDocument SplitOriginal =
        IrTestDocuments.FromParts(Paragraph("Shared middle paragraph text stays") + Paragraph("Tail shared"));

    private static readonly WmlDocument SplitReviewer = IrTestDocuments.FromParts(
        "<w:p><w:bookmarkStart w:id=\"0\" w:name=\"New\"/><w:r><w:t>Shared middle paragraph text stays</w:t></w:r></w:p>" +
        "<w:p><w:r><w:t>Brand new replacement sentence</w:t></w:r><w:bookmarkEnd w:id=\"0\"/></w:p>" + Paragraph("Tail shared"));

    private static WmlDocument ConsolidateSplit() =>
        DocxDiff.Consolidate(SplitOriginal, new[] { new DocxDiffReviewer { Author = "Reviewer", Document = SplitReviewer } });

    [Fact]
    public void Consolidate_ReviewerBookmarkOpeningInAnEqualParagraph_KeepsItsSpanBothWays() =>
        AssertRoundTripsBookmarks(SplitOriginal, SplitReviewer, ConsolidateSplit());

    [Fact]
    public void Consolidate_ReviewerBookmarkInsideAnEqualParagraph_EnclosesTheSameWords()
    {
        var original = IrTestDocuments.FromParts(Paragraph("Shared middle paragraph text stays") + Paragraph("Tail shared"));
        var reviewer = IrTestDocuments.FromParts(
            "<w:p><w:r><w:t xml:space=\"preserve\">Shared </w:t></w:r><w:bookmarkStart w:id=\"3\" w:name=\"Mid\"/>" +
            "<w:r><w:t>middle</w:t></w:r><w:bookmarkEnd w:id=\"3\"/><w:r><w:t xml:space=\"preserve\"> paragraph text stays</w:t></w:r></w:p>" +
            Paragraph("Tail shared"));

        var consolidated = DocxDiff.Consolidate(original, new[] { new DocxDiffReviewer { Author = "Reviewer", Document = reviewer } });

        AssertRoundTripsBookmarks(original, reviewer, consolidated);
        Assert.Equal(new[] { "Mid:middle" }, Spans(RevisionProcessor.AcceptRevisions(consolidated)));
    }

    [Fact]
    public void Consolidate_ReviewerBookmarkOpeningInAnEqualParagraph_AddsNoValidationErrors() =>
        NoNewValidationErrors(SplitOriginal.DocumentByteArray, ConsolidateSplit().DocumentByteArray);
}
