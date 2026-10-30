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

    [Theory]
    [InlineData(false, "none")]
    [InlineData(true, "none")]
    [InlineData(false, "unchanged")]
    [InlineData(true, "unchanged")]
    [InlineData(false, "edited")]
    [InlineData(true, "edited")]
    public void Consolidate_OneSidedInlineBookmarkDeltaBesideIndependentEdits(bool addition, string math)
    {
        WmlDocument Input(bool mark, string prefix, string tail, string equation = "n") =>
            IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>" + prefix + "</w:t></w:r>"
                + (mark ? "<w:bookmarkStart w:id=\"5\" w:name=\"tab\"/>" : "")
                + "<w:r><w:tab/></w:r>" + (mark ? "<w:bookmarkEnd w:id=\"5\"/>" : "")
                + (math == "none" ? "" : "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
                    + "<m:r><m:t>" + equation + "</m:t></m:r></m:oMath>")
                + "<w:r><w:t xml:space=\"preserve\"> shared tail " + tail + "</w:t></w:r></w:p>");
        var original = Input(!addition, "III.", "original");
        var alice = Input(addition, math == "edited" ? "III." : "IIIA.", "original", math == "edited" ? "k" : "n");
        var bob = Input(!addition, "III.", "updated");
        var expected = Input(addition, math == "edited" ? "III." : "IIIA.", "updated", math == "edited" ? "k" : "n");
        foreach (bool reversed in new[] { false, true })
        {
            var reviewers = new[]
            {
                new DocxDiffReviewer { Author = "Alice", Document = alice },
                new DocxDiffReviewer { Author = "Bob", Document = bob },
            };
            var redline = DocxDiff.Consolidate(original, reversed ? reviewers.Reverse().ToArray() : reviewers);
            var accepted = RevisionProcessor.AcceptRevisions(redline);
            var rejected = RevisionProcessor.RejectRevisions(redline);
            Assert.Equal(InlineOrder(expected), InlineOrder(accepted));
            Assert.Equal(InlineOrder(original), InlineOrder(rejected));
            NoNewValidationErrors(original.DocumentByteArray, redline.DocumentByteArray);
            NoNewValidationErrors(expected.DocumentByteArray, accepted.DocumentByteArray);
            NoNewValidationErrors(original.DocumentByteArray, rejected.DocumentByteArray);
            var markers = Body(redline).Descendants().Where(e => e.Name == W + "bookmarkStart" || e.Name == W + "bookmarkEnd").ToList();
            Assert.Equal(2, markers.Count);
            Assert.All(markers, marker => Assert.Contains(marker.Ancestors(addition ? W + "ins" : W + "del"),
                revision => (string?)revision.Attribute(W + "author") == "Alice"));
            Assert.Contains(Body(redline).Descendants(W + "ins"), revision =>
                (string?)revision.Attribute(W + "author") == "Bob" && revision.Value.Contains("updated"));
        }
    }

    [Theory]
    [InlineData("add", "word")]
    [InlineData("remove", "word")]
    [InlineData("unchanged", "word")]
    [InlineData("add", "interior")]
    [InlineData("remove", "interior")]
    [InlineData("unchanged", "interior")]
    [InlineData("add", "coincident")]
    [InlineData("remove", "coincident")]
    [InlineData("unchanged", "coincident")]
    [InlineData("add", "math")]
    [InlineData("remove", "math")]
    [InlineData("unchanged", "math")]
    public void Consolidate_InlineBookmarkDeltaKeepsSourceOrder(string change, string span)
    {
        WmlDocument Input(bool marked, string prefix, string tail)
        {
            string content = span == "math"
                ? "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x</m:t></m:r></m:oMath>"
                : "<w:r><w:t>constant</w:t></w:r>";
            if (span == "coincident") content = "<w:bookmarkStart w:id=\"9\" w:name=\"shared\"/>" + content + "<w:bookmarkEnd w:id=\"9\"/>";
            if (marked) content = "<w:bookmarkStart w:id=\"7\" w:name=\"edited\"/>" + content + "<w:bookmarkEnd w:id=\"7\"/>";
            if (span == "interior") content = "<w:r><w:t>pre</w:t></w:r>" + content + "<w:r><w:t>post</w:t></w:r>";
            return IrTestDocuments.FromBodyXml("<w:p><w:r><w:t xml:space=\"preserve\">" + prefix + " </w:t></w:r>" + content
                + "<w:r><w:t xml:space=\"preserve\"> " + tail + "</w:t></w:r></w:p>");
        }
        bool before = change != "add", after = change != "remove";
        var original = Input(before, "Alder", "ledger");
        var alice = Input(after, "Birch", "ledger");
        var bob = Input(before, "Alder", "schedule");
        var expected = Input(after, "Birch", "schedule");
        foreach (bool reverse in new[] { false, true })
        {
            var reviewers = new[] { new DocxDiffReviewer { Author = "Alice", Document = alice },
                new DocxDiffReviewer { Author = "Bob", Document = bob } };
            var redline = DocxDiff.Consolidate(original, reverse ? reviewers.Reverse().ToArray() : reviewers);
            Assert.Equal(InlineOrder(expected), InlineOrder(RevisionProcessor.AcceptRevisions(redline)));
            Assert.Equal(InlineOrder(original), InlineOrder(RevisionProcessor.RejectRevisions(redline)));
            NoNewValidationErrors(original.DocumentByteArray, redline.DocumentByteArray);
            var markers = Body(redline).Descendants().Where(e => e.Name == W + "bookmarkStart" || e.Name == W + "bookmarkEnd").ToList();
            if (change == "unchanged") Assert.All(markers, marker => Assert.DoesNotContain(marker.Ancestors(), e => e.Name == W + "ins" || e.Name == W + "del"));
            else
            {
                var edited = markers.Single(e => e.Name == W + "bookmarkStart" && (string?)e.Attribute(W + "name") == "edited");
                var id = (string?)edited.Attribute(W + "id");
                var endpoints = markers.Where(e => (string?)e.Attribute(W + "id") == id).ToList();
                Assert.Equal(2, endpoints.Count);
                Assert.All(endpoints, marker =>
                    Assert.Contains(marker.Ancestors(change == "add" ? W + "ins" : W + "del"), revision => (string?)revision.Attribute(W + "author") == "Alice"));
            }
        }
    }

    [Theory]
    [InlineData(false, "replace")]
    [InlineData(true, "replace")]
    [InlineData(false, "insert")]
    [InlineData(false, "delete-middle")]
    [InlineData(true, "delete-middle")]
    public void Consolidate_OneSidedBookmarkSpanningIndependentReplacements(bool addition, string otherEdit)
    {
        WmlDocument Input(bool marked, bool prefix, bool tail, int start, int end)
        {
            string Run(string value) => "<w:r><w:t xml:space=\"preserve\">" + value + "</w:t></w:r>";
            var chunks = new[] { Run(prefix ? "IIIA." : "III."), "<w:r><w:tab/></w:r>",
                Run(" shared "), Run(tail && otherEdit == "delete-middle" ? "" : "middle "), Run("tail "),
                Run(tail && otherEdit != "delete-middle" ? (otherEdit == "insert" ? "original updated" : "updated") : "original") };
            string body = "<w:p>";
            for (int index = 0; index <= chunks.Length; index++)
            {
                if (marked && index == start) body += "<w:bookmarkStart w:id=\"5\" w:name=\"range\"/>";
                if (marked && index == end) body += "<w:bookmarkEnd w:id=\"5\"/>";
                if (index < chunks.Length) body += chunks[index];
            }
            return IrTestDocuments.FromBodyXml(body + "</w:p>");
        }
        for (int start = 0; start <= 6; start++)
        for (int end = start; end <= 6; end++)
        foreach (bool reverse in new[] { false, true })
        {
            var original = Input(!addition, false, false, start, end);
            var alice = Input(addition, true, false, start, end);
            var bob = Input(!addition, false, true, start, end);
            var expected = Input(addition, true, true, start, end);
            var reviewers = new[] { new DocxDiffReviewer { Author = "Alice", Document = alice },
                new DocxDiffReviewer { Author = "Bob", Document = bob } };
            var redline = DocxDiff.Consolidate(original, reverse ? reviewers.Reverse().ToArray() : reviewers);
            Assert.Equal(InlineOrder(expected), InlineOrder(RevisionProcessor.AcceptRevisions(redline)));
            Assert.Equal(InlineOrder(original), InlineOrder(RevisionProcessor.RejectRevisions(redline)));
            NoNewValidationErrors(original.DocumentByteArray, redline.DocumentByteArray);
            var markers = Body(redline).Descendants().Where(e => e.Name == W + "bookmarkStart" || e.Name == W + "bookmarkEnd").ToList();
            Assert.Equal(2, markers.Count);
            Assert.All(markers, marker => Assert.Contains(marker.Ancestors(addition ? W + "ins" : W + "del"),
                revision => (string?)revision.Attribute(W + "author") == "Alice"));
        }
    }

    private static string InlineOrder(WmlDocument document)
    {
        var body = Body(document);
        var names = body.Descendants(W + "bookmarkStart").ToDictionary(e => (string)e.Attribute(W + "id")!, e => (string)e.Attribute(W + "name")!);
        return string.Concat(body.Descendants().Select(e => e.Name == W + "t" ? e.Value
            : e.Name == W + "tab" ? "[tab]"
            : e.Name == W + "bookmarkStart" ? "[start:" + (string?)e.Attribute(W + "name") + "]"
            : e.Name == W + "bookmarkEnd" ? "[end:" + names.GetValueOrDefault((string)e.Attribute(W + "id")!, "orphan") + "]"
            : e.Name.LocalName == "oMath" ? "[math:" + e.Value + "]" : ""));
    }

    [Theory]
    [InlineData("Alder study", "paragraph")]
    [InlineData("Birch survey", "paragraph")]
    [InlineData("Alder ridge survey", "paragraph")]
    [InlineData("Birch study", "paragraph")]
    [InlineData("Alder", "paragraph")]
    [InlineData("Alder study", "inline")]
    [InlineData("Alder ridge survey", "inline")]
    [InlineData("Alder study", "cell")]
    public void Compare_ReplacementAtBookmarkBoundary_PreservesEnclosedText(string revisedText, string context)
    {
        WmlDocument Document(string text)
        {
            var paragraph = "<w:p><w:bookmarkStart w:id=\"5\" w:name=\"_Toc5\"/>" +
                $"<w:r><w:t>{text}</w:t></w:r><w:bookmarkEnd w:id=\"5\"/>" +
                (context == "inline" ? "<w:r><w:t xml:space=\"preserve\"> outside</w:t></w:r>" : "") + "</w:p>";
            return IrTestDocuments.FromParts(context == "cell"
                ? "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid>" +
                  "<w:tr><w:tc><w:tcPr/>" + paragraph + "</w:tc></w:tr></w:tbl>"
                : paragraph);
        }
        var original = Document("Alder survey");
        var revised = Document(revisedText);

        var redline = DocxCompare.Compare(original, revised);

        AssertRoundTripsBookmarks(original, revised, redline);
        Assert.Equal(Body(revised).Value, Body(RevisionProcessor.AcceptRevisions(redline)).Value);
        Assert.Equal(Body(original).Value, Body(RevisionProcessor.RejectRevisions(redline)).Value);
        NoNewValidationErrors(original.DocumentByteArray, redline.DocumentByteArray);
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
