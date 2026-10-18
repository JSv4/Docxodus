// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

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

    // ---------------------------------------------------------- folding a replaced gap around a marker (#1183)

    // A heading whose bookmark starts inside it and ends between blocks, as Word writes a TOC heading bookmark,
    // followed by two more paragraphs, all of it replaced. The bookmark end sits in front of the second deleted
    // paragraph, which the paragraph-mark pairing never fuses: the inserted paragraph still folds into the
    // deleted heading, as Word's compare does.
    private const string HeadingWithBookmarkEndBetweenBlocks =
        "<w:p><w:bookmarkStart w:id=\"0\" w:name=\"_Toc1\"/><w:r><w:t>Quarterly Planning Notes</w:t></w:r></w:p>" +
        "<w:bookmarkEnd w:id=\"0\"/>";

    private static readonly string ReplacedGapLeft =
        HeadingWithBookmarkEndBetweenBlocks + P("Budget review is scheduled for Monday.") + P("Hiring plan is attached.");

    private static string TwoByTwoTable()
    {
        static string Cell(string text) => $"<w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>{P(text)}</w:tc>";
        return "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"3000\"/>" +
            "<w:gridCol w:w=\"3000\"/></w:tblGrid>" + $"<w:tr>{Cell("R1C1")}{Cell("R1C2")}</w:tr><w:tr>{Cell("R2C1")}{Cell("R2C2")}</w:tr></w:tbl>";
    }

    private static int BodyParagraphs(WmlDocument document) => Body(document).Descendants(W + "p").Count();

    private static int BookmarkMarkers(WmlDocument document) =>
        Body(document).Descendants().Count(e => e.Name == W + "bookmarkStart" || e.Name == W + "bookmarkEnd");

    /// <summary>The round trip, plus the bookmark checks <see cref="Layout"/> cannot see: accepting leaves no
    /// bookmark anywhere (the revised document has none, not even inside a coalesced paragraph), and rejecting
    /// restores the heading's start inside it.</summary>
    private static WmlDocument AssertReplacedGapRoundTrip(string engine, WmlDocument left, WmlDocument right)
    {
        var redline = AssertRoundTrip(engine, left, right);
        Assert.Equal(0, BookmarkMarkers(Accept(redline)));
        var rejected = Reject(redline);
        Assert.Equal(2, BookmarkMarkers(rejected));
        Assert.Equal("_Toc1", (string?)Body(rejected).Elements(W + "p").First()
            .Elements(W + "bookmarkStart").Single().Attribute(W + "name"));
        return redline;
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_ReplacedGap_WithABookmarkEndBetweenItsDeletedBlocks_StillFoldsTheInsertedParagraph(string engine)
    {
        var left = Doc(ReplacedGapLeft);
        var right = Doc(P("Draft agenda for the offsite."));

        var redline = AssertReplacedGapRoundTrip(engine, left, right);

        // The inserted text shares the deleted heading's paragraph, and the bookmark end stays between blocks.
        Assert.Equal(3, BodyParagraphs(redline));
        var first = Body(redline).Elements(W + "p").First();
        Assert.Contains("Draft agenda for the offsite.", string.Concat(first.Descendants(W + "t").Select(t => t.Value)));
        Assert.Contains("Quarterly Planning Notes", string.Concat(first.Descendants(W + "delText").Select(t => t.Value)));
        Assert.Equal("p", Body(redline).Elements().Single(e => e.Descendants(W + "bookmarkEnd").Any() &&
            e.Name != W + "p").ElementsBeforeSelf().Last().Name.LocalName);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_ReplacedGap_WithABookmarkEndBetweenItsDeletedBlocks_StillFoldsTheParagraphAfterAnInsertedTable(string engine)
    {
        var left = Doc(ReplacedGapLeft);
        var right = Doc(TwoByTwoTable() + "<w:p/>");

        AssertReplacedGapRoundTrip(engine, left, right);

        // Four cell paragraphs and the three deleted ones: the inserted empty paragraph after the table folds
        // into the last deleted paragraph rather than standing on its own.
        Assert.Equal(7, BodyParagraphs(Compare(engine, left, right)));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_ReplacedCellGap_WithABookmarkEndBetweenItsDeletedBlocks_StillFoldsTheInsertedParagraph(string engine)
    {
        var left = Doc(P("Lead") + Table(ReplacedGapLeft));
        var right = Doc(P("Lead") + Table(P("Draft agenda for the offsite.")));

        var redline = AssertRoundTrip(engine, left, right);

        var cell = Body(redline).Descendants(W + "tc").Single();
        Assert.Equal(new[] { "p", "del", "p", "p" },
            cell.Elements().Where(e => e.Name != W + "tcPr").Select(e => e.Name.LocalName));
        Assert.Equal(0, BookmarkMarkers(Accept(redline)));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_ReplacedGap_WithABookmarkInFrontOfAnInsertedParagraph_StillFoldsTheFinalPair(string engine)
    {
        // The revised side's marker, in front of an inserted paragraph the pairing leaves alone.
        var left = Doc(P("Quarterly budget summary."));
        var right = Doc(P("Draft agenda.") + Start("Added") + P("Venue options.") + P("Travel logistics.") + End());

        var redline = AssertRoundTrip(engine, left, right);

        // The last inserted paragraph shares the deleted one's paragraph mark.
        Assert.Equal(3, BodyParagraphs(redline));
        Assert.Equal(new[] { "BS:Added@1", "BE:Added@3" }, BodyMarkerPositions(Accept(redline)));
        Assert.Equal(0, BookmarkMarkers(Reject(redline)));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_ReplacedGap_WithABookmarkEndInFrontOfTheParagraphThatKeepsItsMark_StillFolds(string engine)
    {
        // Here the marker sits in front of the deleted paragraph whose mark the pairing keeps: the pair still forms,
        // and the marker goes in front of the shared paragraph that paragraph's runs end up in.
        var left = Doc(HeadingWithBookmarkEndBetweenBlocks + P("Hiring plan is attached."));
        var right = Doc(P("Draft agenda for the offsite."));

        var redline = AssertReplacedGapRoundTrip(engine, left, right);

        Assert.Equal(2, BodyParagraphs(redline));
        var children = Body(redline).Elements().Where(e => e.Name != W + "sectPr").ToList();
        Assert.Equal(new[] { "p", "del", "p" }, children.Select(e => e.Name.LocalName));
        Assert.Equal("bookmarkEnd", children[1].Elements().Single().Name.LocalName);
        Assert.Contains("Draft agenda for the offsite.", string.Concat(children[0].Descendants(W + "t").Select(t => t.Value)));
        Assert.Contains("Quarterly Planning Notes", string.Concat(children[0].Descendants(W + "delText").Select(t => t.Value)));
        Assert.Contains("Hiring plan is attached.", string.Concat(children[2].Descendants(W + "delText").Select(t => t.Value)));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_ReplacedInteriorGap_WithABookmarkInFrontOfAnInsertedBlank_PairsAsWithoutIt(string engine)
    {
        // Between unchanged paragraphs a wordful-then-blank original faces three blanks, the first of which holds a
        // heading bookmark that ends after it. The paragraph-mark chain decides exactly as it does without the
        // bookmark: the same paragraphs come out, and the bookmark keeps its place in each view.
        string Gap(bool bookmark) =>
            (bookmark ? "<w:p><w:bookmarkStart w:id=\"21\" w:name=\"_TocR\"/></w:p><w:bookmarkEnd w:id=\"21\"/>" : "<w:p/>") +
            "<w:p/><w:p/>";
        var left = Doc(P("Lead") + P("Original wordful paragraph.") + "<w:p/>" + P("Tail"));
        var right = Doc(P("Lead") + Gap(bookmark: true) + P("Tail"));
        var unmarked = Doc(P("Lead") + Gap(bookmark: false) + P("Tail"));

        var redline = AssertRoundTrip(engine, left, right);

        Assert.Equal(BodyParagraphs(Compare(engine, left, unmarked)), BodyParagraphs(redline));
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

    // A marker-only difference inside a container the IR finds equal: the IR leaves markers out (rule N3), so
    // nothing else tells the two sides apart and the container is emitted as one side's copy.

    private static string Row(params string[] cellContents) =>
        "<w:tr>" + string.Concat(cellContents.Select(c => $"<w:tc><w:tcPr><w:tcW w:w=\"3000\" w:type=\"dxa\"/></w:tcPr>{c}</w:tc>")) + "</w:tr>";

    private static string Rows(int columns, params string[] rows) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid>" +
        string.Concat(Enumerable.Range(0, columns).Select(_ => "<w:gridCol w:w=\"3000\"/>")) + "</w:tblGrid>" +
        string.Concat(rows) + "</w:tbl>";

    private static readonly string Bracketed = P("Alpha") + Start("Old") + P("Bravo") + End();
    private static readonly string Plain = P("Alpha") + P("Bravo");

    public static TheoryData<string, string, string> MarkerOnlyContainerChanges()
    {
        var data = new TheoryData<string, string, string>();
        foreach (var engine in new[] { "DocxCompare", "DocxDiff" })
            foreach (var (withMarker, without) in new[] { (Bracketed, Plain) })
            {
                // An entirely unchanged table; a body edit elsewhere.
                data.Add(engine, P("Lead one") + Table(withMarker), P("Lead two") + Table(without));
                data.Add(engine, P("Lead one") + Table(without), P("Lead two") + Table(withMarker));
                // Another cell of the same row changes.
                data.Add(engine, Rows(2, Row(withMarker, P("Other one"))), Rows(2, Row(without, P("Other two"))));
                data.Add(engine, Rows(2, Row(without, P("Other one"))), Rows(2, Row(withMarker, P("Other two"))));
                // Another row changes.
                data.Add(engine, Rows(1, Row(withMarker), Row(P("Other one"))), Rows(1, Row(without), Row(P("Other two"))));
                data.Add(engine, Rows(1, Row(without), Row(P("Other one"))), Rows(1, Row(withMarker), Row(P("Other two"))));
                // An unchanged block wrapper and an unchanged custom-XML wrapper.
                data.Add(engine, P("Lead one") + Sdt(withMarker), P("Lead two") + Sdt(without));
                data.Add(engine, P("Lead one") + Sdt(without), P("Lead two") + Sdt(withMarker));
                data.Add(engine, P("Lead one") + CustomXml(withMarker), P("Lead two") + CustomXml(without));
                // A table nested in an unchanged table.
                data.Add(engine, P("Lead one") + Table(P("Outer") + Table(withMarker) + P("After")),
                    P("Lead two") + Table(P("Outer") + Table(without) + P("After")));
            }
        return data;
    }

    [Theory]
    [MemberData(nameof(MarkerOnlyContainerChanges))]
    public void Compare_MarkerOnlyChange_InAnEqualContainer_FollowsEachSide(string engine, string leftBody, string rightBody)
    {
        AssertRoundTrip(engine, Doc(leftBody), Doc(rightBody));
    }

    // ------------------------------------------------------------------- accepting a merged paragraph mark

    private const string DeletedMark = "<w:pPr><w:rPr><w:del w:id=\"91\" w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\"/></w:rPr></w:pPr>";

    [Fact]
    public void Accept_DeletedParagraphMark_MergesAcrossABookmarkStart_KeepsItBetweenTheTexts()
    {
        var document = Doc($"<w:p>{DeletedMark}<w:r><w:t>Alpha</w:t></w:r></w:p>" + Start("Mid") + P("Bravo") + End());

        var body = Body(Accept(document));

        Assert.Equal(new[] { "p", "bookmarkEnd" }, body.Elements().Where(e => e.Name != W + "sectPr").Select(e => e.Name.LocalName));
        Assert.Equal(new[] { "r", "bookmarkStart", "r" }, body.Element(W + "p")!.Elements().Select(e => e.Name.LocalName));
    }

    [Fact]
    public void Accept_DeletedParagraphMark_MergesAcrossAnAltChunk_KeepsTheAltChunkOutOfTheParagraph()
    {
        var source = Doc($"<w:p>{DeletedMark}<w:r><w:t>Alpha</w:t></w:r></w:p>" +
            "<w:altChunk xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdX\"/>" + P("Bravo"));
        using var stream = new MemoryStream();
        stream.Write(source.DocumentByteArray);
        using (var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, true))
        {
            var chunk = package.MainDocumentPart!
                .AddNewPart<DocumentFormat.OpenXml.Packaging.AlternativeFormatImportPart>("text/html", "rIdX");
            using var writer = new StreamWriter(chunk.GetStream(FileMode.Create));
            writer.Write("<html><body><p>Chunk</p></body></html>");
        }
        var document = new WmlDocument("altchunk.docx", stream.ToArray());

        var accepted = Accept(document);

        Assert.Empty(Body(accepted).Descendants(W + "p").Elements(W + "altChunk"));
        Assert.Single(Body(accepted).Elements(W + "altChunk"));
        AssertNoNewValidationErrors(accepted, document);
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

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyCommentRange_AddedOnlyInTheRevisedDocument_AcceptKeepsItAndRejectDropsIt(string engine)
    {
        var left = Doc(P("Alpha") + P("Bravo monthly") + P("Charlie"));
        var right = CommentDoc("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo weekly") +
            "<w:commentRangeEnd w:id=\"0\"/>" + $"<w:p>{CommentReference}<w:r><w:t>Charlie</w:t></w:r></w:p>");

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyCommentRange_RemovedInTheRevisedDocument_AcceptDropsItAndRejectKeepsIt(string engine)
    {
        var left = CommentDoc("<w:commentRangeStart w:id=\"0\"/>" + P("Alpha") + P("Bravo monthly") +
            "<w:commentRangeEnd w:id=\"0\"/>" + $"<w:p>{CommentReference}<w:r><w:t>Charlie</w:t></w:r></w:p>");
        var right = Doc(P("Alpha") + P("Bravo weekly") + P("Charlie"));

        AssertRoundTrip(engine, left, right);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_BodyBookmark_AroundAMovedParagraph_FollowsEachSide(string engine)
    {
        const string moved = "This whole paragraph moves further down the document unchanged";
        var left = Doc(P("Lead") + Start("Moved") + P(moved) + End() + P("Alpha paragraph stays") + P("Bravo paragraph stays") + P("Tail"));
        var right = Doc(P("Lead") + P("Alpha paragraph stays") + P("Bravo paragraph stays") + Start("Moved") + P(moved) + End() + P("Tail"));

        AssertRoundTrip(engine, left, right);
    }

    // ------------------------------------------------------------------------------------------ notes

    private const string Separators =
        "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
        "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>";

    private static WmlDocument FootnoteDoc(string noteContent) => IrTestDocuments.FromParts(
        "<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>",
        footnotesInnerXml: Separators + $"<w:footnote w:id=\"1\">{noteContent}</w:footnote>");

    /// <summary>The children of footnote 1: paragraphs by text, bookmark markers by kind and name.</summary>
    private static string FootnoteLayout(WmlDocument document)
    {
        var note = Part(document, "word/footnotes.xml").Root!.Elements(W + "footnote")
            .Single(n => (string?)n.Attribute(W + "id") == "1");
        var names = note.Descendants(W + "bookmarkStart").GroupBy(b => (string?)b.Attribute(W + "id") ?? "")
            .ToDictionary(g => g.Key, g => (string?)g.First().Attribute(W + "name") ?? "");
        return string.Join(" ", note.Elements().Select(e =>
            e.Name == W + "p" ? "P:" + string.Concat(e.Descendants(W + "t").Select(t => t.Value))
            : e.Name == W + "bookmarkStart" ? "BS:" + (string?)e.Attribute(W + "name")
            : e.Name == W + "bookmarkEnd" ? "BE:" + names.GetValueOrDefault((string?)e.Attribute(W + "id") ?? "", "?")
            : e.Name.LocalName));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_FootnoteBookmark_AroundAnEditedParagraph_StaysAroundItsParagraphs(string engine)
    {
        var left = FootnoteDoc(P("Note lead") + Start("InNote") + P("Alpha") + P("Bravo monthly") + End());
        var right = FootnoteDoc(P("Note lead") + Start("InNote") + P("Alpha") + P("Bravo weekly") + End());

        var redline = Compare(engine, left, right);

        Assert.Equal("P:Note lead BS:InNote P:Alpha P:Bravo weekly BE:InNote", FootnoteLayout(Accept(redline)));
        Assert.Equal("P:Note lead BS:InNote P:Alpha P:Bravo monthly BE:InNote", FootnoteLayout(Reject(redline)));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_FootnoteBookmark_MarkerOnlyChange_FollowsEachSide(string engine)
    {
        var left = FootnoteDoc(P("Note lead") + Start("InNote") + P("Alpha") + P("Bravo") + End());
        var right = FootnoteDoc(P("Note lead") + P("Alpha") + P("Bravo"));

        var redline = Compare(engine, left, right);

        Assert.Equal(FootnoteLayout(right), FootnoteLayout(Accept(redline)));
        Assert.Equal(FootnoteLayout(left), FootnoteLayout(Reject(redline)));
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

    [Fact]
    public void Consolidate_BodyBookmark_RemovedByOneReviewer_AcceptDropsItAndRejectKeepsIt()
    {
        var original = Doc(P("Lead") + Start("Removed") + P("Alpha") + P("Bravo monthly") + End() + P("Tail"));
        var first = Doc(P("Lead") + P("Alpha") + P("Bravo monthly") + P("Tail"));
        var second = Doc(P("Lead") + Start("Removed") + P("Alpha") + P("Bravo weekly") + End() + P("Tail"));

        var consolidated = DocxDiff.Consolidate(original, new[]
        {
            new DocxDiffReviewer { Author = "First", Document = first },
            new DocxDiffReviewer { Author = "Second", Document = second },
        });

        Assert.Equal("P:Lead P:Alpha P:Bravo weekly P:Tail", Layout(Accept(consolidated)));
        Assert.Equal(Layout(original), Layout(Reject(consolidated)));
        Assert.Equal("First", (string?)Body(consolidated).Elements(W + "del")
            .Single(d => d.Element(W + "bookmarkStart") is not null).Attribute(W + "author"));
    }
}
