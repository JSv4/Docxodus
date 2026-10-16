// Copyright (c) John Scrudato IV. All rights reserved.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A bookmark or comment range written between blocks — a direct child of the body, a table cell or a block
/// wrapper, a sibling of <c>w:p</c> — keeps its place through a comparison and through accepting or
/// rejecting revisions: accepting the redline gives the revised document's markers at the revised document's
/// block boundaries, rejecting gives the original's, and a document with no revisions keeps them (issue #1149).
/// </summary>
public class BlockLevelMarkerDiffTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    private static string P(string text) => $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static string Start(string name, int id = 7) => $"<w:bookmarkStart w:id=\"{id}\" w:name=\"{name}\"/>";

    private static string End(int id = 7) => $"<w:bookmarkEnd w:id=\"{id}\"/>";

    private static string Table(params string[] cellContents) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid>" +
        string.Concat(cellContents.Select(_ => "<w:gridCol w:w=\"3000\"/>")) + "</w:tblGrid><w:tr>" +
        string.Concat(cellContents.Select(c => $"<w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>{c}</w:tc>")) +
        "</w:tr></w:tbl>";

    private static string CustomXml(string content) =>
        $"<w:customXml w:uri=\"urn:test\" w:element=\"clause\">{content}</w:customXml>";

    private static string Sdt(string content) =>
        $"<w:sdt><w:sdtPr><w:alias w:val=\"Clause\"/><w:id w:val=\"42\"/></w:sdtPr><w:sdtContent>{content}</w:sdtContent></w:sdt>";

    private static WmlDocument Doc(string body) => IrTestDocuments.FromParts(body);

    private static WmlDocument CommentDoc(string body) =>
        IrTestDocuments.WithComment("Reviewer", "R", "2026-01-01T00:00:00Z", "Check this", body);

    private const string CommentReference =
        "<w:r><w:rPr><w:rStyle w:val=\"CommentReference\"/></w:rPr><w:commentReference w:id=\"0\"/></w:r>";

    // ------------------------------------------------------------------------------------- the layout

    private static XDocument Part(WmlDocument document, string path)
    {
        using var package = new ZipArchive(new MemoryStream(document.DocumentByteArray));
        var entry = package.GetEntry(path);
        if (entry is null)
            return new XDocument();
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static XElement Body(WmlDocument document) => Part(document, "word/document.xml").Root!.Element(W + "body")!;

    /// <summary>The block structure of the body: each paragraph as its text, each block-level bookmark or
    /// comment-range marker by kind and name (a bookmark end and a comment marker resolved to the bookmark's
    /// name or the comment's text, since ids are renumbered), and tables, cells and wrappers as nested groups.
    /// Markers inside paragraphs are left out: this is the layout at block boundaries.</summary>
    private static string Layout(WmlDocument document)
    {
        var body = Body(document);
        var bookmarkNames = body.Descendants(W + "bookmarkStart")
            .GroupBy(s => (string?)s.Attribute(W + "id") ?? "")
            .ToDictionary(g => g.Key, g => (string?)g.First().Attribute(W + "name") ?? "");
        var commentTexts = (Part(document, "word/comments.xml").Root?.Elements(W + "comment") ?? Enumerable.Empty<XElement>())
            .GroupBy(c => (string?)c.Attribute(W + "id") ?? "")
            .ToDictionary(g => g.Key, g => string.Concat(g.First().Descendants(W + "t").Select(t => t.Value)));

        string Id(XElement e) => (string?)e.Attribute(W + "id") ?? "";

        IEnumerable<string> Blocks(XElement container)
        {
            foreach (var child in container.Elements())
            {
                var n = child.Name;
                if (n == W + "p")
                    yield return "P:" + string.Concat(child.Descendants(W + "t").Select(t => t.Value));
                else if (n == W + "bookmarkStart")
                    yield return "BS:" + (string?)child.Attribute(W + "name");
                else if (n == W + "bookmarkEnd")
                    yield return "BE:" + bookmarkNames.GetValueOrDefault(Id(child), "?" + Id(child));
                else if (n == W + "commentRangeStart")
                    yield return "CS:" + commentTexts.GetValueOrDefault(Id(child), "?" + Id(child));
                else if (n == W + "commentRangeEnd")
                    yield return "CE:" + commentTexts.GetValueOrDefault(Id(child), "?" + Id(child));
                else if (n == W + "tbl")
                    yield return "TBL[" + string.Join(" | ", child.Elements(W + "tr").Select(tr =>
                        string.Join(" || ", tr.Elements(W + "tc").Select(tc => string.Join(" ", Blocks(tc)))))) + "]";
                else if (n == W + "sdt")
                    yield return "SDT{" + string.Join(" ", Blocks(child.Element(W + "sdtContent") ?? new XElement("x"))) + "}";
                else if (n == W + "customXml")
                    yield return "CX{" + string.Join(" ", Blocks(child)) + "}";
                else if (n == W + "ins" || n == W + "del")
                    yield return n.LocalName + "(" + string.Join(" ", Blocks(child)) + ")";
            }
        }

        return string.Join(" ", Blocks(body));
    }

    private static WmlDocument Accept(WmlDocument d) => RevisionProcessor.AcceptRevisions(d);

    private static WmlDocument Reject(WmlDocument d) => RevisionProcessor.RejectRevisions(d);

    public static TheoryData<string> Engines => new() { "DocxCompare", "DocxDiff" };

    private static WmlDocument Compare(string engine, WmlDocument left, WmlDocument right) =>
        engine == "DocxCompare" ? DocxCompare.Compare(left, right) : DocxDiff.Compare(left, right);

    /// <summary>Accepting the redline gives the revised layout, rejecting it the original's, and the redline
    /// adds no validation errors.</summary>
    private static WmlDocument AssertRoundTrip(string engine, WmlDocument left, WmlDocument right)
    {
        var redline = Compare(engine, left, right);
        Assert.Equal(Layout(right), Layout(Accept(redline)));
        Assert.Equal(Layout(left), Layout(Reject(redline)));
        AssertNoNewValidationErrors(redline, left, right);
        return redline;
    }

    /// <summary>Every validation error of <paramref name="redline"/> also occurs in one of the inputs, compared by
    /// description: the comparison introduced none. (A path-keyed comparison would flag an input's own error, such
    /// as the validator not knowing <c>w:customXml</c>'s attributes, again at a replaced wrapper's second copy.)</summary>
    private static void AssertNoNewValidationErrors(WmlDocument redline, params WmlDocument[] inputs)
    {
        static HashSet<string> Errors(WmlDocument document)
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
            return new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019)
                .Validate(package).Select(e => e.Description).ToHashSet(StringComparer.Ordinal);
        }
        var known = inputs.SelectMany(Errors).ToHashSet(StringComparer.Ordinal);
        Assert.All(Errors(redline), e => Assert.Contains(e, known));
    }

    // ------------------------------------------------------------- accept/reject of a revision-free document

    public static TheoryData<string> RevisionFreeBodies => new()
    {
        Start("Article") + P("Alpha") + P("Bravo") + End(),
        P("Lead") + Table(Start("InCell") + P("Alpha") + P("Bravo") + End()),
        P("Lead") + CustomXml(Start("Wrapped") + P("Alpha") + P("Bravo") + End()),
        P("Lead") + Sdt(Start("Controlled") + P("Alpha") + P("Bravo") + End()),
    };

    [Theory]
    [MemberData(nameof(RevisionFreeBodies))]
    public void AcceptAndReject_OfARevisionFreeDocument_KeepBlockLevelBookmarks(string body)
    {
        var document = Doc(body);
        Assert.Contains("BS:", Layout(document));

        Assert.Equal(Layout(document), Layout(Accept(document)));
        Assert.Equal(Layout(document), Layout(Reject(document)));
    }

    [Fact]
    public void AcceptAndReject_OfARevisionFreeDocument_KeepBlockLevelCommentRanges()
    {
        var document = CommentDoc(
            "<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo") + "<w:commentRangeEnd w:id=\"0\"/>" +
            $"<w:p>{CommentReference}<w:r><w:t>Charlie</w:t></w:r></w:p>");
        Assert.Contains("CS:Check this", Layout(document));

        Assert.Equal(Layout(document), Layout(Accept(document)));
        Assert.Equal(Layout(document), Layout(Reject(document)));
    }

    [Fact]
    public void Accept_KeepsABlockLevelBookmark_NextToARevisedParagraph()
    {
        // A revision elsewhere in the document must not take the bookmark out either.
        var document = Doc(Start("Article") + P("Alpha") +
            "<w:p><w:ins w:id=\"1\" w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\"><w:r><w:t>Bravo</w:t></w:r></w:ins></w:p>" + End());

        Assert.Equal("BS:Article P:Alpha P:Bravo BE:Article", Layout(Accept(document)));
        Assert.Equal("BS:Article P:Alpha P: BE:Article", Layout(Reject(document)));
    }

    // ------------------------------------------------------------------------------------ the body

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_AroundAnEditedParagraph_SurvivesAcceptAndReject(string engine)
    {
        // The issue's reproducer.
        var left = Doc(Start("Article") + P("Alpha") + P("Bravo monthly") + End());
        var right = Doc(Start("Article") + P("Alpha") + P("Bravo weekly") + End());

        var redline = AssertRoundTrip(engine, left, right);
        Assert.Equal("BS:Article", Layout(redline).Split(' ')[0]);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_AroundAnInsertedAndADeletedParagraph_SurvivesAcceptAndReject(string engine)
    {
        var left = Doc(P("Lead") + Start("Article") + P("Alpha") + P("Bravo") + End() + P("Tail"));
        var right = Doc(P("Lead") + Start("Article") + P("Alpha") + P("Charlie inserted") + End() + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_AddedOnlyInTheRevisedDocument_AcceptKeepsItAndRejectDropsIt(string engine)
    {
        var left = Doc(P("Lead") + P("Alpha") + P("Bravo monthly") + P("Tail"));
        var right = Doc(P("Lead") + Start("Added") + P("Alpha") + P("Bravo weekly") + End() + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_RemovedInTheRevisedDocument_AcceptDropsItAndRejectKeepsIt(string engine)
    {
        var left = Doc(P("Lead") + Start("Removed") + P("Alpha") + P("Bravo monthly") + End() + P("Tail"));
        var right = Doc(P("Lead") + P("Alpha") + P("Bravo weekly") + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_AddedWithNoOtherChange_AcceptKeepsItAndRejectDropsIt(string engine)
    {
        var left = Doc(P("Lead") + P("Alpha") + P("Bravo") + P("Tail"));
        var right = Doc(P("Lead") + Start("Added") + P("Alpha") + P("Bravo") + End() + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_WhoseEndMovesToALaterBoundary_FollowsEachSide(string engine)
    {
        var left = Doc(Start("Range") + P("Alpha") + P("Bravo monthly") + End() + P("Charlie") + P("Tail"));
        var right = Doc(Start("Range") + P("Alpha") + P("Bravo weekly") + P("Charlie") + End() + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_WhoseStartMovesToAnEarlierBoundary_FollowsEachSide(string engine)
    {
        var left = Doc(P("Lead") + P("Alpha") + Start("Range") + P("Bravo monthly") + End() + P("Tail"));
        var right = Doc(P("Lead") + Start("Range") + P("Alpha") + P("Bravo weekly") + End() + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_InsideAReplacedGap_FollowsEachSide(string engine)
    {
        // The original's bookmark opens before its second deleted paragraph and the revised document's
        // before its second inserted one: one-sided positions inside a gap of whole-block changes.
        var left = Doc(P("Lead") + P("Old one") + Start("Mid") + P("Old two") + End() + P("Tail"));
        var right = Doc(P("Lead") + P("New first") + Start("Mid", 9) + P("New second") + End(9) + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    // ------------------------------------------------------------------------------------- a table cell

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_CellBookmark_AroundAnEditedParagraph_SurvivesAcceptAndReject(string engine)
    {
        var left = Doc(P("Lead") + Table(Start("InCell") + P("Alpha") + P("Bravo monthly") + End(), P("Other")));
        var right = Doc(P("Lead") + Table(Start("InCell") + P("Alpha") + P("Bravo weekly") + End(), P("Other")));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_CellBookmark_StillBracketsItsParagraphsInTheRedline(string engine)
    {
        var left = Doc(P("Lead") + Table(P("Before") + Start("InCell") + P("Alpha") + P("Bravo monthly") + End() + P("After")));
        var right = Doc(P("Lead") + Table(P("Before") + Start("InCell") + P("Alpha") + P("Bravo weekly") + End() + P("After")));

        var cell = Body(Compare(engine, left, right)).Descendants(W + "tc").Single();
        Assert.Equal(new[] { "p", "bookmarkStart", "p", "p", "bookmarkEnd", "p" },
            cell.Elements().Where(e => e.Name != W + "tcPr").Select(e => e.Name.LocalName));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_CellBookmark_AddedOnlyInTheRevisedDocument_AcceptKeepsItAndRejectDropsIt(string engine)
    {
        var left = Doc(P("Lead") + Table(P("Alpha") + P("Bravo monthly")));
        var right = Doc(P("Lead") + Table(Start("InCell") + P("Alpha") + P("Bravo weekly") + End()));

        AssertRoundTrip(engine, left, right);
    }

    // ---------------------------------------------------------------------------------- comment ranges

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyCommentRange_AroundAnEditedParagraph_SurvivesAcceptAndReject(string engine)
    {
        var left = CommentDoc("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo monthly") +
            "<w:commentRangeEnd w:id=\"0\"/>" + $"<w:p>{CommentReference}<w:r><w:t>Charlie</w:t></w:r></w:p>");
        var right = CommentDoc("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo weekly") +
            "<w:commentRangeEnd w:id=\"0\"/>" + $"<w:p>{CommentReference}<w:r><w:t>Charlie</w:t></w:r></w:p>");

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_CellCommentRange_AroundAnEditedParagraph_SurvivesAcceptAndReject(string engine)
    {
        var left = CommentDoc(P("Lead") + Table("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo monthly") +
            "<w:commentRangeEnd w:id=\"0\"/>" + $"<w:p>{CommentReference}</w:p>"));
        var right = CommentDoc(P("Lead") + Table("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo weekly") +
            "<w:commentRangeEnd w:id=\"0\"/>" + $"<w:p>{CommentReference}</w:p>"));

        AssertRoundTrip(engine, left, right);
    }

    // ------------------------------------------------------------------------------------ block wrappers

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BookmarkBetweenAReplacedWrappersBlocks_SurvivesAcceptAndReject(string engine)
    {
        var left = Doc(P("Lead") + Sdt(P("Alpha") + Start("Inner") + P("Bravo monthly") + End()) + P("Tail"));
        var right = Doc(P("Lead") + Sdt(P("Alpha") + Start("Inner") + P("Bravo weekly") + End()) + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BookmarkBetweenAReplacedCustomXmlsBlocks_SurvivesAcceptAndReject(string engine)
    {
        var left = Doc(P("Lead") + CustomXml(P("Alpha") + Start("Inner") + P("Bravo monthly") + End()) + P("Tail"));
        var right = Doc(P("Lead") + CustomXml(P("Alpha") + Start("Inner") + P("Bravo weekly") + End()) + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_CommentRangeOpeningBeforeAReplacedWrapper_SurvivesAcceptAndReject(string engine)
    {
        var left = CommentDoc(P("Lead") + "<w:commentRangeStart w:id=\"0\"/>" + Sdt(P("Alpha") + P("Bravo monthly")) +
            $"<w:p><w:r><w:t>Charlie</w:t></w:r><w:commentRangeEnd w:id=\"0\"/>{CommentReference}</w:p>");
        var right = CommentDoc(P("Lead") + "<w:commentRangeStart w:id=\"0\"/>" + Sdt(P("Alpha") + P("Bravo weekly")) +
            $"<w:p><w:r><w:t>Charlie</w:t></w:r><w:commentRangeEnd w:id=\"0\"/>{CommentReference}</w:p>");

        AssertRoundTrip(engine, left, right);
    }

    // ---------------------------------------------------------------------------------- a real document

    /// <summary>Each body-level bookmark marker as kind, bookmark name and the number of body blocks before it.</summary>
    private static List<string> BodyMarkerPositions(WmlDocument document)
    {
        var body = Body(document);
        var names = body.Descendants(W + "bookmarkStart").GroupBy(s => (string?)s.Attribute(W + "id") ?? "")
            .ToDictionary(g => g.Key, g => (string?)g.First().Attribute(W + "name") ?? "");
        var positions = new List<string>();
        int blocks = 0;
        foreach (var child in body.Elements())
        {
            if (child.Name == W + "p" || child.Name == W + "tbl" || child.Name == W + "sdt")
                blocks++;
            else if (child.Name == W + "bookmarkStart")
                positions.Add($"BS:{(string?)child.Attribute(W + "name")}@{blocks}");
            else if (child.Name == W + "bookmarkEnd")
                positions.Add($"BE:{names.GetValueOrDefault((string?)child.Attribute(W + "id") ?? "", "?")}@{blocks}");
        }
        return positions;
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_RealDocumentWithBodyLevelBookmarks_KeepsThemAtTheirBlockBoundaries(string engine)
    {
        // A Word document whose revised version holds 916 body-level bookmark markers (the WmlComparer-era
        // fixture rule N3 was first written for).
        var dir = Path.Combine("..", "..", "..", "..", "TestFiles", "WC");
        var left = new WmlDocument(Path.Combine(dir, "WC-BodyBookmarks-Before.docx"));
        var right = new WmlDocument(Path.Combine(dir, "WC-BodyBookmarks-After.docx"));

        var redline = Compare(engine, left, right);

        Assert.Equal(916, BodyMarkerPositions(right).Count);
        Assert.Equal(BodyMarkerPositions(right), BodyMarkerPositions(Accept(redline)));
        // Rejecting puts no marker anywhere the original does not have one. (Three of the original's six are ends
        // of bookmarks that start inside a paragraph the revised document restyled; that paragraph is emitted
        // from the revised side, which lacks the start, so those bookmarks are not restored on reject. That is
        // the inline-bookmark limitation, not the block-level placement.)
        var rejected = BodyMarkerPositions(Reject(redline));
        var original = BodyMarkerPositions(left);
        Assert.NotEmpty(rejected);
        Assert.All(rejected, marker => Assert.Contains(marker, original));
    }

    // -------------------------------------------------------------------------------------- consolidate

    [Fact]
    public void Consolidate_BodyBookmark_AroundAReviewersEditedParagraph_SurvivesAcceptAndReject()
    {
        var original = Doc(Start("Article") + P("Alpha") + P("Bravo monthly") + End() + P("Tail"));
        var reviewer = Doc(Start("Article") + P("Alpha") + P("Bravo weekly") + End() + P("Tail"));

        var consolidated = DocxDiff.Consolidate(original, new[] { new DocxDiffReviewer { Author = "Reviewer", Document = reviewer } });

        Assert.Equal(Layout(reviewer), Layout(Accept(consolidated)));
        Assert.Equal(Layout(original), Layout(Reject(consolidated)));
    }

    [Fact]
    public void Consolidate_BodyBookmark_AddedByOneOfTwoReviewers_AcceptKeepsItAndRejectDropsIt()
    {
        var original = Doc(P("Lead") + P("Alpha") + P("Bravo monthly") + P("Tail"));
        var first = Doc(P("Lead") + Start("Added") + P("Alpha") + P("Bravo monthly") + End() + P("Tail"));
        var second = Doc(P("Lead") + P("Alpha") + P("Bravo weekly") + P("Tail"));

        var consolidated = DocxDiff.Consolidate(original, new[]
        {
            new DocxDiffReviewer { Author = "First", Document = first },
            new DocxDiffReviewer { Author = "Second", Document = second },
        });

        Assert.Equal("P:Lead BS:Added P:Alpha P:Bravo weekly BE:Added P:Tail", Layout(Accept(consolidated)));
        Assert.Equal(Layout(original), Layout(Reject(consolidated)));
    }
}
