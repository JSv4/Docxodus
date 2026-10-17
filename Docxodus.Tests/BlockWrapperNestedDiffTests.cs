// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1143: when a block wrapper (a <c>w:customXml</c> element or a block <c>w:sdt</c> content control)
/// keeps its own markup, a comparison diffs the blocks inside one copy of it, as it does a table cell's
/// blocks. Only a change to the wrapper itself (customXml attributes and <c>w:customXmlPr</c>; a control's
/// <c>w:sdtPr</c>/<c>w:sdtEndPr</c>) replaces the whole wrapper.
/// </summary>
public class BlockWrapperNestedDiffTests
{
    public enum Wrapper
    {
        CustomXml,
        ContentControl,
    }

    private const string Alpha = "Alpha clause: the parties agree to these terms.";
    private const string Bravo = "Bravo clause: backups are verified monthly by staff.";
    private const string Charlie = "Charlie clause: restoration is tested once a year.";
    private const string Tail = "Closing paragraph outside the wrapper.";

    private static string P(string text) => $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static string Wrap(Wrapper kind, string inner, string id = "a1") => kind == Wrapper.CustomXml
        ? $"<w:customXml w:uri=\"urn:example\" w:element=\"article\"><w:customXmlPr><w:attr w:name=\"id\" w:val=\"{id}\"/></w:customXmlPr>{inner}</w:customXml>"
        : $"<w:sdt><w:sdtPr><w:alias w:val=\"Article\"/><w:tag w:val=\"{id}\"/><w:id w:val=\"101\"/></w:sdtPr><w:sdtContent>{inner}</w:sdtContent></w:sdt>";

    private static XName WrapperName(Wrapper kind) => kind == Wrapper.CustomXml ? W.customXml : W.sdt;

    private static WmlDocument Doc(string body) => IrTestDocuments.FromBodyXml(body);

    private static string Three(Wrapper kind, string bravo = Bravo, string id = "a1") =>
        Wrap(kind, P(Alpha) + P(bravo) + P(Charlie), id);

    private static readonly string BravoEdited = Bravo.Replace("monthly", "weekly");

    // ---- the acceptance test -------------------------------------------------------------------

    /// <summary>The issue's case through every surface: both compare front doors and consolidate draw only the
    /// changed word, the revision lists report only it, and the semantic diff reports a text change and no
    /// change to the wrapper.</summary>
    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void OneWordEditedInsideAnUnchangedWrapper_IsReportedAsThatWordEverywhere(Wrapper kind)
    {
        var left = Doc(Three(kind) + P(Tail));
        var right = Doc(Three(kind, BravoEdited) + P(Tail));
        var reviewers = new[] { new DocxDiffReviewer { Document = right, Author = "Reviewer" } };

        var redlines = new[]
        {
            DocxCompare.Compare(left, right),
            DocxDiff.Compare(left, right, new DocxDiffSettings()),
            DocxCompare.Consolidate(left, reviewers),
        };

        foreach (var redline in redlines)
        {
            var body = Body(redline);
            AssertSingleWrapperWithoutRangeMarkers(body, kind);
            Assert.Equal("monthly", string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
            Assert.Equal("weekly", string.Concat(body.Descendants(W.ins).Descendants(W.t).Select(t => t.Value)));
            AssertUnmarked(body, Alpha, Charlie, Tail);
            AssertRoundTrip(left, right, redline);
        }
        var revisionLists = new[]
        {
            DocxDiff.GetRevisions(left, right).Select(r => (r.Type, r.Text)),
            DocxDiff.GetConsolidatedRevisions(left, reviewers).Select(r => (r.Type, r.Text)),
        };
        foreach (var revisions in revisionLists)
            Assert.Equal(
                new[] { (DocxDiffRevisionType.Deleted, "monthly"), (DocxDiffRevisionType.Inserted, "weekly") },
                revisions.Where(r => r.Type is DocxDiffRevisionType.Inserted or DocxDiffRevisionType.Deleted).OrderBy(r => r.Text));
        var changes = Docxodus.Verification.SemanticDiff.Compare(left, right).Changes;
        Assert.Contains(changes, c => c.Family == Docxodus.Verification.SemanticChangeFamily.Text);
        Assert.DoesNotContain(changes, c => c.Family is Docxodus.Verification.SemanticChangeFamily.ContentControl
            or Docxodus.Verification.SemanticChangeFamily.BlockStructure);
    }

    private static string Item(string text, int numId) =>
        $"<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"{numId}\"/></w:numPr></w:pPr>"
        + $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_ListItemsKeptInsideAnUnchangedWrapper_StayInTheirList(Wrapper kind)
    {
        // The revised document gives the same list another numId. Items the diff keeps must map back to the
        // original's list, or they start a list of their own.
        const string Definition = "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/>"
            + "<w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/></w:lvl></w:abstractNum>";
        WmlDocument List(int numId, string bravo) => IrTestDocuments.FromParts(
            Wrap(kind, Item(Alpha, numId) + Item(bravo, numId)) + P(Tail),
            numberingInnerXml: Definition + $"<w:num w:numId=\"{numId}\"><w:abstractNumId w:val=\"0\"/></w:num>");
        var left = List(1, Bravo);
        var right = List(5, BravoEdited);

        var redline = DocxCompare.Compare(left, right);

        var numIds = Body(redline).Descendants(W.numId).Select(n => (string?)n.Attribute(W.val)).Distinct().ToList();
        Assert.Equal(["1"], numIds);
        AssertSingleWrapperWithoutRangeMarkers(Body(redline), kind);
    }

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_ParagraphInsertedInsideAnUnchangedWrapper_KeepsItsNewStyle(Wrapper kind)
    {
        // Only the revised document defines the style, and only the inserted paragraph and an edited one use it.
        const string Style = "<w:style w:type=\"paragraph\" w:styleId=\"Fancy\"><w:name w:val=\"Fancy\"/>"
            + "<w:pPr><w:jc w:val=\"center\"/></w:pPr></w:style>";
        string Styled(string text) => $"<w:p><w:pPr><w:pStyle w:val=\"Fancy\"/></w:pPr><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";
        var left = Doc(Three(kind) + P(Tail));
        var right = IrTestDocuments.FromParts(
            Wrap(kind, P(Alpha) + Styled("A brand new styled paragraph appears here.") + Styled(Bravo) + P(Charlie)) + P(Tail),
            stylesInnerXml: Style);

        var redline = DocxCompare.Compare(left, right);

        AssertSingleWrapperWithoutRangeMarkers(Body(redline), kind);
        AssertRoundTrip(left, right, redline);
    }

    // ---- markers between the wrapper's blocks ----------------------------------------------------

    public enum MarkerPlace
    {
        InTheWrapper,
        InANestedWrapper,
        InATableCell,
    }

    public enum Marker
    {
        Bookmark,
        CommentRange,
    }

    public static IEnumerable<object[]> MarkersKindsAndPlaces() =>
        from marker in Enum.GetValues<Marker>()
        from kind in new[] { Wrapper.CustomXml, Wrapper.ContentControl }
        from place in Enum.GetValues<MarkerPlace>()
        select new object[] { marker, kind, place };

    // A bookmark or comment range between a wrapper's blocks is not modeled by the IR, so a wrapper holding one
    // is replaced whole (each copy keeps its markers) rather than diffed in place, which would drop them. The
    // tests read the redline itself and go through DocxDiff.Compare, because accepting revisions already loses
    // block-level markers, and so does the DocxCompare front door, which accepts its inputs first (issue #1149).

    /// <summary>A wrapper whose content brackets Alpha, Bravo and Charlie between <paramref name="start"/> and
    /// <paramref name="end"/>, block-level markers placed per <paramref name="place"/>.</summary>
    private static string Bracketed(Wrapper kind, MarkerPlace place, string bravo, string start, string end)
    {
        var inner = start + P(Alpha) + P(bravo) + P(Charlie) + end;
        return place switch
        {
            MarkerPlace.InTheWrapper => Wrap(kind, inner),
            MarkerPlace.InANestedWrapper => Wrap(kind, P(Tail) + Wrap(kind, inner, id: "inner"), id: "outer"),
            _ => Wrap(kind, TwoCellTable(P("Left cell stays."), inner)),
        };
    }

    [Theory]
    [MemberData(nameof(MarkersKindsAndPlaces))]
    public void Compare_MarkerBetweenTheWrappersBlocks_StaysInTheRedline(Marker marker, Wrapper kind, MarkerPlace place)
    {
        var (start, end) = marker == Marker.Bookmark
            ? ("<w:bookmarkStart w:id=\"7\" w:name=\"Article\"/>", "<w:bookmarkEnd w:id=\"7\"/>")
            : ("<w:commentRangeStart w:id=\"0\"/>", "<w:commentRangeEnd w:id=\"0\"/><w:p><w:r><w:commentReference w:id=\"0\"/></w:r></w:p>");
        WmlDocument Document(string bravo) => marker == Marker.Bookmark
            ? Doc(Bracketed(kind, place, bravo, start, end) + P(Tail))
            : IrTestDocuments.WithComment("Ann", "A", "2026-01-01T00:00:00Z", "Check this clause.", Bracketed(kind, place, bravo, start, end) + P(Tail));
        var left = Document(Bravo);
        var right = Document(BravoEdited);

        var redline = DocxDiff.Compare(left, right);

        var (startName, endName) = marker == Marker.Bookmark ? (W.bookmarkStart, W.bookmarkEnd) : (W.commentRangeStart, W.commentRangeEnd);
        AssertEveryRangeBracketsTheClauses(Body(redline), startName, endName);
        AssertRoundTrip(left, right, redline);
    }

    // ---- issue #1157: a comment range between the blocks of a wrapper that is replaced whole -----------

    public static IEnumerable<object[]> CommentWrapperKindsAndPlaces() =>
        from kind in new[] { Wrapper.CustomXml, Wrapper.ContentControl }
        from place in Enum.GetValues<MarkerPlace>()
        from inBodyTableCell in new[] { false, true }
        select new object[] { kind, place, inBodyTableCell };

    /// <summary>The wrapper is replaced whole (the range between its blocks is not diffed in place), so the redline
    /// holds a deleted and an inserted copy. Each copy must keep a complete comment of its own — range start, range
    /// end and reference under one id, resolving to the comment's text — so that accepting yields the revised
    /// document with the reference still in the wrapper's last paragraph, and rejecting yields the original.</summary>
    [Theory]
    [MemberData(nameof(CommentWrapperKindsAndPlaces))]
    public void Compare_CommentRangeBetweenTheWrappersBlocks_AcceptAndRejectKeepTheReferenceInTheWrapper(
        Wrapper kind, MarkerPlace place, bool inBodyTableCell)
    {
        const string Start = "<w:commentRangeStart w:id=\"0\"/>";
        const string End = "<w:commentRangeEnd w:id=\"0\"/><w:p><w:r><w:commentReference w:id=\"0\"/></w:r></w:p>";
        WmlDocument Document(string bravo)
        {
            var wrapper = Bracketed(kind, place, bravo, Start, End);
            var body = (inBodyTableCell ? TwoCellTable(P("Left cell stays."), wrapper) : wrapper) + P(Tail);
            return IrTestDocuments.WithComment("Ann", "A", "2026-01-01T00:00:00Z", CommentText, body);
        }
        var left = Document(Bravo);
        var right = Document(BravoEdited);

        var redline = DocxDiff.Compare(left, right);

        AssertRoundTrip(left, right, redline);
        AssertEachWrapperCopyHasItsOwnComment(redline.DocumentByteArray, kind, copies: 2, rangeKept: true);
        AssertEveryRangeBracketsTheClauses(Body(redline), W.commentRangeStart, W.commentRangeEnd);
        AssertNoNewValidationErrors(redline, left, right);
        AssertEachWrapperCopyHasItsOwnComment(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray), kind, copies: 1, rangeKept: false);
        AssertEachWrapperCopyHasItsOwnComment(DocxDiffOps.RejectRevisions(redline.DocumentByteArray), kind, copies: 1, rangeKept: false);
    }

    /// <summary>A comment that opens in a paragraph before the replaced wrapper and closes inside it keeps one
    /// comment: the collapse that serves its shared start is not undone by splitting the deleted copy's
    /// reference off under a fresh id, so rejecting restores the original with the reference and the range on
    /// the same comment. (Accepting still leaves the reference before the wrapper, as it did before #1157: no
    /// single set of markers can serve both copies of the wrapper.)</summary>
    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_CommentOpeningBeforeAReplacedWrapper_RejectKeepsTheReferenceOnTheSameComment(Wrapper kind)
    {
        WmlDocument Document(string bravo) => IrTestDocuments.WithComment("Ann", "A", "2026-01-01T00:00:00Z", CommentText,
            "<w:p><w:r><w:t xml:space=\"preserve\">Before the wrapper.</w:t></w:r><w:commentRangeStart w:id=\"0\"/></w:p>"
            + Wrap(kind, P(Alpha) + P(bravo) + "<w:commentRangeEnd w:id=\"0\"/><w:p><w:r><w:commentReference w:id=\"0\"/></w:r></w:p>")
            + P(Tail));
        var left = Document(Bravo);
        var right = Document(BravoEdited);

        var redline = DocxDiff.Compare(left, right);

        var rejected = Body(DocxDiffOps.RejectRevisions(redline.DocumentByteArray));
        Assert.Equal(Signature(left.DocumentByteArray), Signature(rejected));
        var start = Assert.Single(rejected.Descendants(W.commentRangeStart));
        var reference = Assert.Single(rejected.Descendants(W.commentReference));
        Assert.Equal((string?)start.Attribute(W.id), (string?)reference.Attribute(W.id));
    }

    private const string CommentText = "Check this clause.";

    /// <summary>The document holds <paramref name="copies"/> comments, each with one reference and one range start
    /// and end (at most one when not <paramref name="rangeKept"/>: accepting and rejecting revisions drop range
    /// markers between blocks, issue #1149), all inside the same innermost <paramref name="kind"/> wrapper, the
    /// reference in that wrapper's own last paragraph, and the id resolving to a definition reading
    /// <see cref="CommentText"/>. Copies of the comment carry distinct ids.</summary>
    private static void AssertEachWrapperCopyHasItsOwnComment(byte[] bytes, Wrapper kind, int copies, bool rangeKept)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var body = main.GetXDocument().Root!.Element(W.body)!;
        var definitions = main.WordprocessingCommentsPart?.GetXDocument().Root?.Elements(W.comment)
            .ToDictionary(c => (string)c.Attribute(W.id)!, c => string.Concat(c.Descendants(W.t).Select(t => t.Value)))
            ?? new Dictionary<string, string>();
        XElement? WrapperOf(XElement e) => e.Ancestors(WrapperName(kind)).FirstOrDefault();

        var references = body.Descendants(W.commentReference).ToList();
        Assert.Equal(copies, references.Count);
        Assert.Equal(copies, references.Select(r => (string?)r.Attribute(W.id)).Distinct().Count());
        foreach (var reference in references)
        {
            var id = (string)reference.Attribute(W.id)!;
            Assert.True(definitions.TryGetValue(id, out var text), $"comment {id} has no definition");
            Assert.Equal(CommentText, text);
            var wrapper = WrapperOf(reference);
            Assert.NotNull(wrapper);
            var paragraph = reference.Ancestors(W.p).First();
            Assert.Same(paragraph, wrapper!.Descendants(W.p).Last());
            foreach (var name in new[] { W.commentRangeStart, W.commentRangeEnd })
            {
                var markers = body.Descendants(name).Where(m => (string?)m.Attribute(W.id) == id).ToList();
                Assert.True(rangeKept ? markers.Count == 1 : markers.Count <= 1,
                    $"comment {id} has {markers.Count} {name.LocalName} markers");
                foreach (var marker in markers)
                    Assert.Same(wrapper, WrapperOf(marker));
            }
        }
        var referenced = references.Select(r => (string?)r.Attribute(W.id)).ToHashSet();
        Assert.All(body.Descendants().Where(e => e.Name == W.commentRangeStart || e.Name == W.commentRangeEnd),
            m => Assert.Contains((string?)m.Attribute(W.id), referenced));
    }

    /// <summary>A comment range inside an INLINE <c>w:customXml</c> whose text is edited: the fine path emits a
    /// deleted and an inserted copy of the inline element, each with its own range, and the reference after it
    /// stays shared. Each view keeps one complete range around its own copy's text.</summary>
    [Fact]
    public void Compare_CommentRangeInsideAnEditedInlineCustomXml_EachViewKeepsAWholeRange()
    {
        static WmlDocument Document(string term) => IrTestDocuments.WithComment("Ann", "A", "2026-01-01T00:00:00Z", CommentText,
            "<w:p><w:r><w:t xml:space=\"preserve\">Lead words here </w:t></w:r>"
            + "<w:customXml w:uri=\"urn:example\" w:element=\"term\"><w:commentRangeStart w:id=\"0\"/>"
            + $"<w:r><w:t xml:space=\"preserve\">{term}</w:t></w:r><w:commentRangeEnd w:id=\"0\"/></w:customXml>"
            + "<w:r><w:commentReference w:id=\"0\"/></w:r><w:r><w:t xml:space=\"preserve\"> trailing text.</w:t></w:r></w:p>"
            + P(Tail));
        var left = Document("the defined term");
        var right = Document("the amended term");

        var redline = DocxDiff.Compare(left, right);

        // Not a full AssertRoundTrip: the differ places the shared reference run before the inline element rather
        // than after it, in both views (a separate, pre-existing placement issue). This test is about the ranges.
        foreach (var (view, term) in new[]
                 {
                     (Body(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray)), "the amended term"),
                     (Body(DocxDiffOps.RejectRevisions(redline.DocumentByteArray)), "the defined term"),
                 })
        {
            var start = Assert.Single(view.Descendants(W.commentRangeStart));
            Assert.Single(view.Descendants(W.commentRangeEnd));
            var reference = Assert.Single(view.Descendants(W.commentReference));
            Assert.Equal((string?)start.Attribute(W.id), (string?)reference.Attribute(W.id));
            Assert.Equal(term, MarkedText(view, e => e == start, W.commentRangeEnd));
        }
    }

    private const string W14Ns = "http://schemas.microsoft.com/office/word/2010/wordml";
    private static readonly XNamespace W15Ns = "http://schemas.microsoft.com/office/word/2012/wordml";
    private static readonly XNamespace W16CidNs = "http://schemas.microsoft.com/office/word/2016/wordml/cid";

    /// <summary>A document whose comments part holds a parent comment and a reply to it, threaded through
    /// <c>commentsExtended</c> and given durable ids in <c>commentsIds</c>.</summary>
    private static WmlDocument ThreadedDocument(string bodyXml)
    {
        static void Write(OpenXmlPart part, string xml)
        {
            using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(xml);
        }
        static string Comment(string id, string text, string paraId) =>
            $"<w:comment w:id=\"{id}\" w:author=\"Ann\" w:initials=\"A\" w:date=\"2026-01-01T00:00:00Z\">"
            + $"<w:p w14:paraId=\"{paraId}\"><w:r><w:t>{text}</w:t></w:r></w:p></w:comment>";
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            Write(main.AddNewPart<WordprocessingCommentsPart>(),
                $"<w:comments xmlns:w=\"{W.w.NamespaceName}\" xmlns:w14=\"{W14Ns}\">"
                + Comment("0", "Parent question.", "00000001") + Comment("1", "Child reply.", "00000002") + "</w:comments>");
            Write(main.AddNewPart<WordprocessingCommentsExPart>(),
                $"<w15:commentsEx xmlns:w15=\"{W15Ns.NamespaceName}\">"
                + "<w15:commentEx w15:paraId=\"00000001\" w15:done=\"0\"/>"
                + "<w15:commentEx w15:paraId=\"00000002\" w15:paraIdParent=\"00000001\" w15:done=\"0\"/></w15:commentsEx>");
            Write(main.AddNewPart<WordprocessingCommentsIdsPart>(),
                $"<w16cid:commentsIds xmlns:w16cid=\"{W16CidNs.NamespaceName}\">"
                + "<w16cid:commentId w16cid:paraId=\"00000001\" w16cid:durableId=\"11111111\"/>"
                + "<w16cid:commentId w16cid:paraId=\"00000002\" w16cid:durableId=\"22222222\"/></w16cid:commentsIds>");
            Write(main, $"<w:document xmlns:w=\"{W.w.NamespaceName}\"><w:body>{bodyXml}</w:body></w:document>");
        }
        return new WmlDocument("threaded.docx", ms.ToArray());
    }

    /// <summary>A parent comment and its reply both range over the blocks of a wrapper that is replaced whole. Each
    /// copy of the wrapper gets its own thread: the deleted copy's reply answers the deleted copy's parent, never
    /// the inserted copy's, and the clones take fresh paraIds and durable ids. So each view shows one intact
    /// thread whose reply's parent is the other comment referenced in that view.</summary>
    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_ThreadBetweenTheWrappersBlocks_EachViewKeepsItsOwnThread(Wrapper kind)
    {
        WmlDocument Document(string bravo) => ThreadedDocument(Wrap(kind,
            "<w:commentRangeStart w:id=\"0\"/><w:commentRangeStart w:id=\"1\"/>" + P(Alpha) + P(bravo)
            + "<w:commentRangeEnd w:id=\"0\"/><w:commentRangeEnd w:id=\"1\"/>"
            + "<w:p><w:r><w:commentReference w:id=\"0\"/></w:r><w:r><w:commentReference w:id=\"1\"/></w:r></w:p>") + P(Tail));
        var left = Document(Bravo);
        var right = Document(BravoEdited);

        var redline = DocxDiff.Compare(left, right);

        AssertRoundTrip(left, right, redline);
        AssertThreadingIdsUnique(redline.DocumentByteArray);
        AssertEachReferencedReplyAnswersAReferencedParent(redline.DocumentByteArray, threads: 2);
        AssertEachReferencedReplyAnswersAReferencedParent(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray), threads: 1);
        AssertEachReferencedReplyAnswersAReferencedParent(DocxDiffOps.RejectRevisions(redline.DocumentByteArray), threads: 1);
    }

    /// <summary>Issue #1170: a <c>w14:paraId</c> is unique across the whole package, not just the comments part.
    /// The body's closing paragraph already uses the paraIds one and two past the comments part's largest, so the
    /// deleted copies' cloned definitions must take ids no body paragraph or threading entry holds.</summary>
    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_CommentClonesForAReplacedWrapper_TakeParaIdsUnusedAnywhereInThePackage(Wrapper kind)
    {
        const string TakenParaIds = $"<w:p xmlns:w14=\"{W14Ns}\" w14:paraId=\"00000003\"><w:r><w:t>{Tail}</w:t></w:r></w:p>"
            + $"<w:p xmlns:w14=\"{W14Ns}\" w14:paraId=\"00000004\"><w:r><w:t>{Charlie}</w:t></w:r></w:p>";
        WmlDocument Document(string bravo) => ThreadedDocument(Wrap(kind,
            "<w:commentRangeStart w:id=\"0\"/><w:commentRangeStart w:id=\"1\"/>" + P(Alpha) + P(bravo)
            + "<w:commentRangeEnd w:id=\"0\"/><w:commentRangeEnd w:id=\"1\"/>"
            + "<w:p><w:r><w:commentReference w:id=\"0\"/></w:r><w:r><w:commentReference w:id=\"1\"/></w:r></w:p>") + TakenParaIds);
        var left = Document(Bravo);
        var right = Document(BravoEdited);

        var redline = DocxDiff.Compare(left, right);

        AssertRoundTrip(left, right, redline);
        using var stream = new MemoryStream(redline.DocumentByteArray);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var definitions = main.WordprocessingCommentsPart!.GetXDocument().Root!;
        Assert.Equal(4, definitions.Elements(W.comment).Count()); // the two comments and their deleted-view clones
        var paraIds = main.GetXDocument().Root!.Descendants().Concat(definitions.Descendants())
            .Attributes(XName.Get("paraId", W14Ns)).Select(a => a.Value).ToList();
        Assert.Equal(6, paraIds.Count);
        Assert.True(paraIds.Count == paraIds.Distinct().Count(), $"duplicate w14:paraId: {string.Join(", ", paraIds)}");
        AssertThreadingIdsUnique(redline.DocumentByteArray);
    }

    /// <summary>Comment ids, definition paraIds and durable ids are each unique in the package.</summary>
    private static void AssertThreadingIdsUnique(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        static void Unique(IEnumerable<string?> values, string what)
        {
            var list = values.ToList();
            Assert.True(list.Count == list.Distinct().Count(), $"duplicate {what}: {string.Join(", ", list)}");
        }
        Unique(main.WordprocessingCommentsPart!.GetXDocument().Root!.Elements(W.comment).Select(c => (string?)c.Attribute(W.id)), "comment ids");
        Unique(main.WordprocessingCommentsExPart!.GetXDocument().Root!.Elements(W15Ns + "commentEx")
            .Select(e => (string?)e.Attribute(W15Ns + "paraId")), "commentsExtended paraIds");
        var ids = main.WordprocessingCommentsIdsPart!.GetXDocument().Root!.Elements(W16CidNs + "commentId").ToList();
        Unique(ids.Select(e => (string?)e.Attribute(W16CidNs + "paraId")), "commentsIds paraIds");
        Unique(ids.Select(e => (string?)e.Attribute(W16CidNs + "durableId")), "durable ids");
    }

    /// <summary>The body references <paramref name="threads"/> parent comments and as many replies; each referenced
    /// reply's <c>paraIdParent</c> names a parent whose reference sits in the same paragraph.</summary>
    private static void AssertEachReferencedReplyAnswersAReferencedParent(byte[] bytes, int threads)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var paraIdOf = main.WordprocessingCommentsPart!.GetXDocument().Root!.Elements(W.comment).ToDictionary(
            c => (string)c.Attribute(W.id)!, c => (string?)c.Elements(W.p).Last().Attribute(XName.Get("paraId", W14Ns)));
        var parentOf = main.WordprocessingCommentsExPart!.GetXDocument().Root!.Elements(W15Ns + "commentEx")
            .ToDictionary(e => (string)e.Attribute(W15Ns + "paraId")!, e => (string?)e.Attribute(W15Ns + "paraIdParent"));
        var references = main.GetXDocument().Root!.Element(W.body)!.Descendants(W.commentReference).ToList();
        Assert.Equal(2 * threads, references.Count);
        var replies = references.Where(r => parentOf[paraIdOf[(string)r.Attribute(W.id)!]!] != null).ToList();
        Assert.Equal(threads, replies.Count);
        foreach (var reply in replies)
        {
            var parentPara = parentOf[paraIdOf[(string)reply.Attribute(W.id)!]!];
            var parent = Assert.Single(references, r => paraIdOf[(string)r.Attribute(W.id)!] == parentPara);
            Assert.Same(reply.Ancestors(W.p).First(), parent.Ancestors(W.p).First());
        }
    }

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_BookmarkInsideAParagraphOfAnUnchangedWrapper_StillDiffsInPlace(Wrapper kind)
    {
        // Only markers between blocks force a whole replacement; one inside a paragraph is carried by it.
        string Marked(string bravo) => Wrap(kind, P(Alpha)
            + $"<w:p><w:bookmarkStart w:id=\"7\" w:name=\"Article\"/><w:r><w:t xml:space=\"preserve\">{bravo}</w:t></w:r>"
            + "<w:bookmarkEnd w:id=\"7\"/></w:p>" + P(Charlie));
        var left = Doc(Marked(Bravo) + P(Tail));
        var right = Doc(Marked(BravoEdited) + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        AssertSingleWrapperWithoutRangeMarkers(body, kind);
        Assert.Equal("monthly", string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
        AssertRoundTrip(left, right, redline);
    }

    /// <summary>The redline holds at least one <paramref name="startName"/> range, and each one closes after
    /// all three clauses of its copy of the wrapper.</summary>
    private static void AssertEveryRangeBracketsTheClauses(XElement body, XName startName, XName endName)
    {
        var starts = body.Descendants(startName).ToList();
        Assert.NotEmpty(starts);
        foreach (var start in starts)
        {
            var text = MarkedText(body, e => e == start, endName);
            Assert.StartsWith(Alpha, text, StringComparison.Ordinal);
            Assert.EndsWith(Charlie, text, StringComparison.Ordinal);
        }
    }

    /// <summary>The text, deleted or not, between the single start marker matching <paramref name="isStart"/>
    /// and the <paramref name="endName"/> marker with the same id.</summary>
    private static string MarkedText(XElement body, Func<XElement, bool> isStart, XName endName)
    {
        var all = body.Descendants().ToList();
        var start = Assert.Single(all, e => isStart(e));
        var id = (string?)start.Attribute(W.id);
        int from = all.IndexOf(start);
        int to = all.FindIndex(from, e => e.Name == endName && (string?)e.Attribute(W.id) == id);
        Assert.True(to > from, "the range has no end after its start");
        return string.Concat(all.Skip(from).Take(to - from).Where(e => e.Name == W.t || e.Name == W.delText).Select(e => e.Value));
    }

    // ---- the wrapper itself changed ------------------------------------------------------------

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_WrapperPropertiesChanged_StillReplacesTheWholeWrapper(Wrapper kind)
    {
        var left = Doc(Three(kind) + P(Tail));
        var right = Doc(Three(kind, BravoEdited, id: "a2") + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        Assert.Equal(2, body.Descendants(WrapperName(kind)).Count());
        Assert.NotEmpty(body.Descendants(W.customXmlDelRangeStart));
        Assert.NotEmpty(body.Descendants(W.customXmlInsRangeStart));
        AssertRoundTrip(left, right, redline);

        var revisions = DocxDiff.GetRevisions(left, right);
        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Deleted && r.Text.Contains(Alpha));
        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Inserted && r.Text.Contains(Alpha));
        Assert.Contains(Docxodus.Verification.SemanticDiff.Compare(left, right).Changes,
            c => c.Path.EndsWith(".envelope", StringComparison.Ordinal));
    }

    // ---- nesting -------------------------------------------------------------------------------

    [Theory]
    [InlineData(Wrapper.CustomXml, Wrapper.ContentControl)]
    [InlineData(Wrapper.ContentControl, Wrapper.CustomXml)]
    [InlineData(Wrapper.CustomXml, Wrapper.CustomXml)]
    public void Compare_EditInsideNestedUnchangedWrappers_MarksOnlyTheWord(Wrapper outer, Wrapper inner)
    {
        string Nested(string bravo) => Wrap(outer, P(Alpha) + Wrap(inner, P(bravo) + P(Charlie), id: "inner"), id: "outer");
        var left = Doc(Nested(Bravo) + P(Tail));
        var right = Doc(Nested(BravoEdited) + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        Assert.DoesNotContain(body.Descendants(), IsRangeMarker);
        Assert.Equal(outer == inner ? 2 : 1, body.Descendants(WrapperName(outer)).Count());
        Assert.Equal(outer == inner ? 2 : 1, body.Descendants(WrapperName(inner)).Count());
        Assert.Equal("monthly", string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
        AssertUnmarked(body, Alpha, Charlie, Tail);
        AssertRoundTrip(left, right, redline);
    }

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_InnerWrapperChangedInsideAnUnchangedOuter_ReplacesOnlyTheInnerWrapper(Wrapper kind)
    {
        string Nested(string innerId) => Wrap(kind, P(Alpha) + Wrap(kind, P(Bravo), id: innerId), id: "outer");
        var left = Doc(Nested("inner-1") + P(Tail));
        var right = Doc(Nested("inner-2") + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        var outerWrappers = body.Elements(WrapperName(kind)).ToList();
        Assert.Single(outerWrappers);
        Assert.NotEmpty(body.Descendants(W.customXmlDelRangeStart));
        AssertUnmarked(body, Alpha, Tail);
        AssertRoundTrip(left, right, redline);
    }

    // ---- block insertions and deletions inside the wrapper ---------------------------------------

    public static IEnumerable<object[]> InsertedAndDeletedChildren()
    {
        const string NewClause = "Delta clause: a brand new obligation appears here.";
        foreach (var kind in new[] { Wrapper.CustomXml, Wrapper.ContentControl })
        {
            // name, left children, right children, paragraphs that must stay unmarked
            yield return new object[] { kind, "insert middle", new[] { Alpha, Charlie }, new[] { Alpha, NewClause, Charlie }, new[] { Alpha, Charlie } };
            yield return new object[] { kind, "insert first", new[] { Alpha, Charlie }, new[] { NewClause, Alpha, Charlie }, new[] { Alpha, Charlie } };
            yield return new object[] { kind, "insert last", new[] { Alpha, Charlie }, new[] { Alpha, Charlie, NewClause }, new[] { Alpha, Charlie } };
            yield return new object[] { kind, "delete middle", new[] { Alpha, Bravo, Charlie }, new[] { Alpha, Charlie }, new[] { Alpha, Charlie } };
            yield return new object[] { kind, "delete first", new[] { Alpha, Bravo, Charlie }, new[] { Bravo, Charlie }, new[] { Charlie } };
            yield return new object[] { kind, "delete last", new[] { Alpha, Bravo, Charlie }, new[] { Alpha, Bravo }, new[] { Alpha } };
        }
    }

    [Theory]
    [MemberData(nameof(InsertedAndDeletedChildren))]
    public void Compare_ParagraphInsertedOrDeletedInsideAnUnchangedWrapper_KeepsOneWrapper(
        Wrapper kind, string name, string[] leftChildren, string[] rightChildren, string[] unmarked)
    {
        var left = Doc(Wrap(kind, string.Concat(leftChildren.Select(P))) + P(Tail));
        var right = Doc(Wrap(kind, string.Concat(rightChildren.Select(P))) + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        AssertSingleWrapperWithoutRangeMarkers(body, kind);
        AssertUnmarked(body, unmarked.Append(Tail).ToArray());
        AssertRoundTrip(left, right, redline, name);
    }

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_TableCellEditedInsideAnUnchangedWrapper_DiffsTheCell(Wrapper kind)
    {
        string Table(string cell) => TwoCellTable(P("Left cell stays."), P(cell));
        var left = Doc(Wrap(kind, P(Alpha) + Table(Bravo)) + P(Tail));
        var right = Doc(Wrap(kind, P(Alpha) + Table(BravoEdited)) + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        AssertSingleWrapperWithoutRangeMarkers(body, kind);
        Assert.Single(body.Descendants(W.tbl));
        Assert.Equal("monthly", string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
        AssertUnmarked(body, Alpha, "Left cell stays.", Tail);
        AssertRoundTrip(left, right, redline);
    }

    // ---- the wrapper pair is found even when it is not the lone residue of its gap ----------------

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_ParagraphInsertedBeforeAnEditedWrapper_StillDiffsInsideTheWrapper(Wrapper kind)
    {
        const string Intro = "A new introductory paragraph was added.";
        var left = Doc(Three(kind) + P(Tail));
        var right = Doc(P(Intro) + Three(kind, BravoEdited) + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        AssertSingleWrapperWithoutRangeMarkers(body, kind);
        Assert.Equal("monthly", string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
        AssertUnmarked(body, Alpha, Charlie, Tail);
        AssertRoundTrip(left, right, redline);
    }

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_TwoAdjacentEditedWrappers_DiffInsideEach(Wrapper kind)
    {
        string Two(string first, string second) => TwoWrappers(kind, P(Alpha) + P(first), P(second) + P(Charlie));
        var left = Doc(Two(Bravo, Bravo) + P(Tail));
        var right = Doc(Two(BravoEdited, Bravo.Replace("staff", "auditors")) + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        Assert.Equal(2, body.Descendants(WrapperName(kind)).Count());
        Assert.DoesNotContain(body.Descendants(), IsRangeMarker);
        Assert.Equal("monthlystaff", string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
        AssertUnmarked(body, Alpha, Charlie, Tail);
        AssertRoundTrip(left, right, redline);
    }

    // ---- consolidate ---------------------------------------------------------------------------

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Consolidate_TwoReviewersEditingDifferentWrappers_DiffInsideEach(Wrapper kind)
    {
        string Two(string first, string second) => TwoWrappers(kind, P(Alpha) + P(first), P(second) + P(Charlie));
        var baseline = Doc(Two(Bravo, Bravo) + P(Tail));
        var first = Doc(Two(BravoEdited, Bravo) + P(Tail));
        var second = Doc(Two(Bravo, Bravo.Replace("staff", "auditors")) + P(Tail));
        var merged = Doc(Two(BravoEdited, Bravo.Replace("staff", "auditors")) + P(Tail));

        var redline = DocxCompare.Consolidate(baseline,
            [new DocxDiffReviewer { Document = first, Author = "One" }, new DocxDiffReviewer { Document = second, Author = "Two" }]);

        var body = Body(redline);
        Assert.Equal(2, body.Descendants(WrapperName(kind)).Count());
        Assert.DoesNotContain(body.Descendants(), IsRangeMarker);
        AssertUnmarked(body, Alpha, Charlie, Tail);
        AssertRoundTrip(baseline, merged, redline);
    }

    // ---- moves ---------------------------------------------------------------------------------

    private static string TwoWrappers(Wrapper kind, string first, string second) =>
        Wrap(kind, first, id: "w1") + Wrap(kind, second, id: "w2");

    private static string TwoCellTable(string first, string second) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"4000\"/><w:gridCol w:w=\"4000\"/></w:tblGrid>"
        + "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/></w:tcPr>" + first + "</w:tc>"
        + "<w:tc><w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/></w:tcPr>" + second + "</w:tc></w:tr></w:tbl>";

    private static readonly string Abc = P(Alpha) + P(Bravo) + P(Charlie);
    private static readonly string MovedAbc = P(Bravo) + P(Charlie) + P(Alpha);

    [Theory]
    [InlineData(Wrapper.CustomXml, false)]
    [InlineData(Wrapper.ContentControl, false)]
    [InlineData(Wrapper.CustomXml, true)]
    [InlineData(Wrapper.ContentControl, true)]
    public void Compare_ParagraphMovedWithinOrIntoAnUnchangedWrapper_IsANativeMove(Wrapper kind, bool fromBody)
    {
        var left = Doc((fromBody ? P(Bravo) + Wrap(kind, P(Alpha) + P(Charlie)) : Wrap(kind, Abc)) + P(Tail));
        var right = Doc(Wrap(kind, fromBody ? Abc : MovedAbc) + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline);
        AssertSingleWrapperWithoutRangeMarkers(body, kind);
        Assert.NotEmpty(body.Elements(WrapperName(kind)).Descendants(W.moveTo));
        AssertMoveMarkersUnique(redline.DocumentByteArray, moves: 1);
        AssertUnmarked(body, fromBody ? Alpha : Bravo, Charlie, Tail);
        AssertRoundTrip(left, right, redline);
        Assert.Equal(2, DocxDiff.GetRevisions(left, right).Count(r => r.Type == DocxDiffRevisionType.Moved));
    }

    public enum MoveScope
    {
        Wrappers,
        WrappersInTableCells,
        TableCells,
        WrappersInFootnotes,
    }

    /// <summary>A document with two scopes holding <paramref name="first"/> and <paramref name="second"/>: two
    /// wrappers, two wrappers in the cells of one table, the two cells themselves, or two wrappers in different
    /// footnotes.</summary>
    private static WmlDocument TwoScopes(MoveScope scope, Wrapper kind, string first, string second) => scope switch
    {
        MoveScope.Wrappers => Doc(TwoWrappers(kind, first, second) + P(Tail)),
        MoveScope.WrappersInTableCells => Doc(TwoCellTable(Wrap(kind, first, id: "w1"), Wrap(kind, second, id: "w2")) + P(Tail)),
        MoveScope.TableCells => Doc(TwoCellTable(first, second) + P(Tail)),
        _ => IrTestDocuments.FromParts(
            P(Tail) + "<w:p><w:r><w:t>One</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r>"
                + "<w:r><w:t>Two</w:t></w:r><w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p>",
            footnotesInnerXml:
                "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
                + "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
                + $"<w:footnote w:id=\"1\">{Wrap(kind, first, id: "n1")}</w:footnote>"
                + $"<w:footnote w:id=\"2\">{Wrap(kind, second, id: "n2")}</w:footnote>"),
    };

    public static IEnumerable<object[]> WrapperKindsAndMoveScopes()
    {
        foreach (var scope in Enum.GetValues<MoveScope>())
        {
            yield return new object[] { Wrapper.CustomXml, scope };
            if (scope != MoveScope.TableCells)
                yield return new object[] { Wrapper.ContentControl, scope };
        }
    }

    /// <summary>Each reviewer's comparison numbers its moves from 1, so when two reviewers move paragraphs in
    /// different scopes (wrappers, table cells with or without a wrapper, footnotes), consolidate must give each
    /// move its own name; a shared name reads as one move with two sources.</summary>
    [Theory]
    [MemberData(nameof(WrapperKindsAndMoveScopes))]
    public void Consolidate_TwoReviewersMovingParagraphsInDifferentScopes_KeepsEachMovePaired(Wrapper kind, MoveScope scope)
    {
        var baseline = TwoScopes(scope, kind, Abc, Abc);
        var reviewers = new[]
        {
            new DocxDiffReviewer { Document = TwoScopes(scope, kind, MovedAbc, Abc), Author = "One" },
            new DocxDiffReviewer { Document = TwoScopes(scope, kind, Abc, MovedAbc), Author = "Two" },
        };

        var redline = DocxCompare.Consolidate(baseline, reviewers);

        Assert.Empty(DocxDiff.GetConflicts(baseline, reviewers));
        AssertMoveMarkersUnique(redline.DocumentByteArray, moves: 2);
        AssertRoundTrip(baseline, TwoScopes(scope, kind, MovedAbc, MovedAbc), redline);
    }

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void Compare_MovesInTheBodyAndInsideAWrapperInACell_HaveDistinctNames(Wrapper kind)
    {
        string Layout(string cell, string body) => TwoCellTable(P("Left cell stays."), Wrap(kind, cell)) + body;
        string BodyAbc = P("Delta body line one here.") + P("Echo body line two here.") + P("Foxtrot body line three here.");
        string BodyMoved = P("Echo body line two here.") + P("Foxtrot body line three here.") + P("Delta body line one here.");
        var left = Doc(Layout(Abc, BodyAbc) + P(Tail));
        var right = Doc(Layout(MovedAbc, BodyMoved) + P(Tail));

        var redline = DocxCompare.Compare(left, right);

        AssertMoveMarkersUnique(redline.DocumentByteArray, moves: 2);
        AssertRoundTrip(left, right, redline);
    }

    /// <summary>Across the main document and its notes, the redline holds <paramref name="moves"/> moves: every
    /// move name labels exactly one source range and one destination range, and every move range id is used by
    /// one start/end pair only.</summary>
    private static void AssertMoveMarkersUnique(byte[] bytes, int moves)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var roots = new[] { main.GetXDocument().Root!, main.FootnotesPart?.GetXDocument().Root, main.EndnotesPart?.GetXDocument().Root }
            .Where(r => r is not null).Select(r => r!).ToList();
        var markers = roots.SelectMany(r => r.Descendants()).ToList();
        var from = markers.Where(e => e.Name == W.moveFromRangeStart).Select(e => (string?)e.Attribute(W.name)).ToList();
        var to = markers.Where(e => e.Name == W.moveToRangeStart).Select(e => (string?)e.Attribute(W.name)).ToList();
        Assert.Equal(moves, from.Count);
        Assert.Equal(from.Count, from.Distinct().Count());
        Assert.Equal(to.Count, to.Distinct().Count());
        Assert.Equal(from.OrderBy(n => n), to.OrderBy(n => n));
        foreach (var start in new[] { W.moveFromRangeStart, W.moveToRangeStart })
        {
            var ids = markers.Where(e => e.Name == start).Select(e => (string?)e.Attribute(W.id)).ToList();
            Assert.Equal(ids.Count, ids.Distinct().Count());
        }
    }

    // ---- conflicts -----------------------------------------------------------------------------

    public static IEnumerable<object[]> WrapperKindsAndPolicies() =>
        from kind in new[] { Wrapper.CustomXml, Wrapper.ContentControl }
        from policy in Enum.GetValues<ConflictResolution>()
        select new object[] { kind, policy };

    [Theory]
    [MemberData(nameof(WrapperKindsAndPolicies))]
    public void Consolidate_TwoReviewersEditingInsideTheSameWrapper_IsAConflictThatStaysReversible(
        Wrapper kind, ConflictResolution policy)
    {
        var baseline = Doc(Three(kind) + P(Tail));
        var first = Doc(Wrap(kind, P(Alpha.Replace("terms", "rules")) + P(Bravo) + P(Charlie)) + P(Tail));
        var second = Doc(Wrap(kind, P(Alpha) + P(Bravo) + P(Charlie.Replace("once", "twice"))) + P(Tail));
        var reviewers = new[]
        {
            new DocxDiffReviewer { Document = first, Author = "One" },
            new DocxDiffReviewer { Document = second, Author = "Two" },
        };
        var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };

        var redline = DocxCompare.Consolidate(baseline, reviewers, settings);

        var conflict = Assert.Single(DocxDiff.GetConflicts(baseline, reviewers, settings));
        Assert.Collection(conflict.Competitors,
            one => Assert.Contains("rules", one.ResultText),
            two => Assert.Contains("twice", two.ResultText));
        Assert.Equal(Signature(baseline.DocumentByteArray), Signature(DocxDiffOps.RejectRevisions(redline.DocumentByteArray)));
        var accepted = Signature(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray));
        if (policy == ConflictResolution.BaseWins)
            Assert.Equal(Signature(baseline.DocumentByteArray), accepted);
        else if (policy == ConflictResolution.FirstReviewerWins)
            Assert.Equal(Signature(first.DocumentByteArray), accepted);
    }

    // ---- edit script ---------------------------------------------------------------------------

    [Theory]
    [InlineData(Wrapper.CustomXml)]
    [InlineData(Wrapper.ContentControl)]
    public void EditScript_CarriesTheWrapperChildOps_AndRoundTripsThroughJson(Wrapper kind)
    {
        var left = Doc(Three(kind) + P(Tail));
        var right = Doc(Three(kind, BravoEdited) + P(Tail));
        var options = new IrReaderOptions { RetainSources = false, RevisionView = RevisionView.Accept };
        var script = IrEditScriptBuilder.Build(IrReader.Read(left, options), IrReader.Read(right, options), new IrDiffSettings());

        var wrapperOp = Assert.Single(script.Operations, op => op.WrapperOps is not null);
        Assert.Equal(IrEditOpKind.ModifyBlock, wrapperOp.Kind);
        Assert.Equal(
            new[] { IrEditOpKind.EqualBlock, IrEditOpKind.ModifyBlock, IrEditOpKind.EqualBlock },
            wrapperOp.WrapperOps!.Select(op => op.Kind));

        var json = IrEditScriptJson.Write(script);
        Assert.Contains("\"wrapperOps\"", json);
        Assert.Equal(script, IrEditScriptJson.Read(json));
    }

    // ---- a field spanning the wrapper's paragraphs (a Word table of contents) --------------------

    public enum FieldEnd
    {
        /// <summary>Word's own shape: the field's end character sits alone in a trailing paragraph.</summary>
        OwnParagraph,

        /// <summary>The end character closes the last entry paragraph.</summary>
        LastEntry,
    }

    private const string TocShell =
        "<w:sdtPr><w:id w:val=\"-1511212882\"/><w:docPartObj><w:docPartGallery w:val=\"Table of Contents\"/><w:docPartUnique/></w:docPartObj></w:sdtPr>";

    /// <summary>One entry as Word writes it: a hyperlink around the heading text, a tab and a PAGEREF field.
    /// The first entry opens the TOC field itself before its hyperlink. The hyperlink's anchor comes from the
    /// entry's first word, so editing the rest of the text leaves the anchor alone, as Word does.</summary>
    private static string TocEntry(string text, bool opensToc, bool closesToc)
    {
        var anchor = "_Toc" + text.Split(' ')[0];
        var sb = new StringBuilder("<w:p>");
        if (opensToc)
            sb.Append("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>")
                .Append("<w:r><w:instrText xml:space=\"preserve\"> TOC \\o \"1-3\" \\h \\z \\u </w:instrText></w:r>")
                .Append("<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>");
        sb.Append($"<w:hyperlink w:anchor=\"{anchor}\" w:history=\"1\">")
            .Append($"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r><w:r><w:tab/></w:r>")
            .Append("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>")
            .Append($"<w:r><w:instrText xml:space=\"preserve\"> PAGEREF {anchor} \\h </w:instrText></w:r>")
            .Append("<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>1</w:t></w:r>")
            .Append("<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:hyperlink>");
        if (closesToc)
            sb.Append("<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>");
        return sb.Append("</w:p>").ToString();
    }

    private static string TableOfContents(FieldEnd end, string[] entries)
    {
        var sb = new StringBuilder("<w:sdt>").Append(TocShell).Append("<w:sdtContent>").Append(P("Contents"));
        for (var i = 0; i < entries.Length; i++)
            sb.Append(TocEntry(entries[i], opensToc: i == 0, closesToc: end == FieldEnd.LastEntry && i == entries.Length - 1));
        if (end == FieldEnd.OwnParagraph)
            sb.Append("<w:p><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>");
        return sb.Append("</w:sdtContent></w:sdt>").ToString();
    }

    public static IEnumerable<object[]> TableOfContentsEdits()
    {
        string[] three = { "Scope", "Term", "Fees" };
        foreach (var end in new[] { FieldEnd.OwnParagraph, FieldEnd.LastEntry })
        {
            yield return new object[] { end, "entry inserted first", new[] { "Term", "Fees" }, three };
            yield return new object[] { end, "entry inserted last", new[] { "Scope", "Term" }, three };
            yield return new object[] { end, "first entry deleted", three, new[] { "Term", "Fees" } };
            yield return new object[] { end, "middle entry deleted", three, new[] { "Scope", "Fees" } };
            yield return new object[] { end, "last entry deleted", three, new[] { "Scope", "Term" } };
            yield return new object[] { end, "entry text edited", three, new[] { "Scope", "Term and renewal", "Fees" } };
        }
    }

    /// <summary>A table of contents is a content control holding one field that opens in its first entry
    /// paragraph and closes paragraphs later. Whatever the comparison does inside it, accepting and rejecting
    /// the redline must each leave every field with its begin, separate and end characters.</summary>
    [Theory]
    [MemberData(nameof(TableOfContentsEdits))]
    public void Compare_TableOfContentsEntriesChanged_KeepsEveryFieldBalanced(
        FieldEnd end, string name, string[] leftEntries, string[] rightEntries)
    {
        var left = Doc(TableOfContents(end, leftEntries) + P(Tail));
        var right = Doc(TableOfContents(end, rightEntries) + P(Tail));

        AssertFieldsSurviveRoundTrip(left, right, DocxCompare.Compare(left, right), name);
    }

    [Fact]
    public void Compare_RealTableOfContentsWithAnEntryRemoved_KeepsEveryFieldBalanced()
    {
        var original = new WmlDocument(Path.Combine(FixtureDir.FullName, "HC022-Table-Of-Contents.docx"));
        var revised = EditBody(original, body =>
        {
            var toc = body.Descendants(W.sdt).First(s => s.Descendants(W.docPartGallery).Any());
            var entries = toc.Element(W.sdtContent)!.Elements(W.p).Where(p => p.Elements(W.hyperlink).Any()).ToList();
            entries[^1].Remove();
            var heading = entries[0].Descendants(W.t).First();
            heading.Value = heading.Value.Replace("Video", "Audio");
        });

        AssertFieldsSurviveRoundTrip(original, revised, DocxCompare.Compare(original, revised), "HC022");
    }

    /// <summary>A Word-authored document whose block content control holds a heading, a prose paragraph, a
    /// seven-row table and a prose paragraph (the pair under <c>TestFiles/Issue1143</c>, see its README): one
    /// word edited inside the control is drawn as that word inside one copy of the control, through the
    /// <see cref="DocxCompare"/> front door.</summary>
    [Fact]
    public void Compare_RealClauseControlWithOneWordEdited_DiffsInsideTheControl()
    {
        var original = new WmlDocument(Path.Combine(FixtureDir.FullName, "Issue1143/clause-control-original.docx"));
        var revised = new WmlDocument(Path.Combine(FixtureDir.FullName, "Issue1143/clause-control-revised.docx"));

        var redline = DocxCompare.Compare(original, revised);

        var body = Body(redline);
        AssertSingleWrapperWithoutRangeMarkers(body, Wrapper.ContentControl);
        Assert.Equal("only", string.Concat(body.Descendants(W.delText).Select(t => t.Value)).Trim());
        Assert.Equal("just", string.Concat(body.Descendants(W.ins).Descendants(W.t).Select(t => t.Value)).Trim());
        Assert.Equal(
            new[] { (DocxDiffRevisionType.Inserted, "just"), (DocxDiffRevisionType.Deleted, "only") },
            DocxDiff.GetRevisions(original, revised)
                .Where(r => r.Type is DocxDiffRevisionType.Inserted or DocxDiffRevisionType.Deleted)
                .Select(r => (r.Type, r.Text.Trim())).OrderBy(r => r.Item2));
        AssertNoNewValidationErrors(redline, original, revised);
        AssertRoundTrip(original, revised, redline, "Issue1143");
    }

    /// <summary>The redline adds no validation error, accepts to <paramref name="revised"/> and rejects to
    /// <paramref name="original"/> block for block, and each view holds the field characters of its document.</summary>
    private static void AssertFieldsSurviveRoundTrip(WmlDocument original, WmlDocument revised, WmlDocument redline, string name)
    {
        AssertNoNewValidationErrors(redline, original, revised);
        AssertRoundTrip(original, revised, redline, name);
        AssertSameFieldCharacters(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray), revised.DocumentByteArray, "accept, " + name);
        AssertSameFieldCharacters(DocxDiffOps.RejectRevisions(redline.DocumentByteArray), original.DocumentByteArray, "reject, " + name);
    }

    private static readonly DirectoryInfo FixtureDir = new("../../../../TestFiles/");

    private static WmlDocument EditBody(WmlDocument document, Action<XElement> edit)
    {
        using var stream = new MemoryStream();
        stream.Write(document.DocumentByteArray, 0, document.DocumentByteArray.Length);
        using (var package = WordprocessingDocument.Open(stream, true))
        {
            var main = package.MainDocumentPart!;
            edit(main.GetXDocument().Root!.Element(W.body)!);
            main.PutXDocument();
        }
        return new WmlDocument("revised.docx", stream.ToArray());
    }

    /// <summary>The body of <paramref name="actual"/> holds as many field begin, separate and end characters
    /// as the body of <paramref name="expected"/>.</summary>
    private static void AssertSameFieldCharacters(byte[] actual, byte[] expected, string because)
    {
        static string Counts(byte[] bytes) => string.Join(", ", Body(bytes).Descendants(W.fldChar)
            .GroupBy(f => (string?)f.Attribute(W.fldCharType))
            .OrderBy(g => g.Key, StringComparer.Ordinal)
            .Select(g => g.Key + "=" + g.Count()));
        Assert.True(Counts(expected) == Counts(actual), $"field characters ({because}): expected {Counts(expected)}, got {Counts(actual)}");
    }

    /// <summary>Every validation error of <paramref name="redline"/> also occurs in one of the
    /// <paramref name="inputs"/>: the comparison introduced none.</summary>
    private static void AssertNoNewValidationErrors(WmlDocument redline, params WmlDocument[] inputs)
    {
        static HashSet<string> Errors(WmlDocument document)
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            return new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package)
                .Select(e => e.Description).ToHashSet(StringComparer.Ordinal);
        }
        var known = inputs.SelectMany(Errors).ToHashSet(StringComparer.Ordinal);
        Assert.All(Errors(redline), e => Assert.Contains(e, known));
    }

    // ---- markers inside a nested wrapper -------------------------------------------------------

    private const string InnerBookmarkStart = "<w:bookmarkStart w:id=\"7\" w:name=\"Article\"/>";
    private const string InnerBookmarkEnd = "<w:bookmarkEnd w:id=\"7\"/>";

    /// <summary>An outer wrapper holding <paramref name="alpha"/> and an inner wrapper whose blocks
    /// (<paramref name="bravo"/> and Charlie) are bracketed by a bookmark between them.</summary>
    private static string OuterAroundBookmarkedInner(Wrapper outer, Wrapper inner, string alpha, string bravo) =>
        Wrap(outer, P(alpha) + Wrap(inner, InnerBookmarkStart + P(bravo) + P(Charlie) + InnerBookmarkEnd, id: "inner"), id: "outer");

    public static IEnumerable<object[]> NestedWrapperKindsAndEditPlaces() =>
        from pair in new[] { (Wrapper.CustomXml, Wrapper.ContentControl), (Wrapper.ContentControl, Wrapper.CustomXml), (Wrapper.CustomXml, Wrapper.CustomXml) }
        from editInside in new[] { false, true }
        select new object[] { pair.Item1, pair.Item2, editInside };

    /// <summary>The bookmark sits between the INNER wrapper's blocks, so only that wrapper needs the whole copy
    /// that keeps it: the outer wrapper is diffed in place whether the edit is beside the inner wrapper (which
    /// is then kept as it is) or inside it (which is then replaced whole, each copy bracketed).</summary>
    [Theory]
    [MemberData(nameof(NestedWrapperKindsAndEditPlaces))]
    public void Compare_MarkersInANestedWrapper_KeepOnlyThatWrapperWhole(Wrapper outer, Wrapper inner, bool editInside)
    {
        var left = Doc(OuterAroundBookmarkedInner(outer, inner, Alpha, Bravo) + P(Tail));
        var right = Doc(OuterAroundBookmarkedInner(outer, inner,
            editInside ? Alpha : Alpha.Replace("terms", "rules"), editInside ? BravoEdited : Bravo) + P(Tail));

        var redline = DocxDiff.Compare(left, right);

        var body = Body(redline);
        var outerWrapper = Assert.Single(body.Elements(WrapperName(outer)));
        Assert.Equal(editInside ? 2 : 1, outerWrapper.Descendants(WrapperName(inner)).Count());
        Assert.Equal(editInside ? Bravo + Charlie : "terms", string.Concat(body.Descendants(W.delText).Select(t => t.Value)));
        AssertUnmarked(body, Tail, editInside ? Alpha : Charlie);
        var starts = outerWrapper.Descendants(W.bookmarkStart).ToList();
        Assert.Equal(editInside ? 2 : 1, starts.Count);
        foreach (var start in starts)
        {
            var text = MarkedText(body, e => e == start, W.bookmarkEnd);
            Assert.StartsWith("Bravo clause", text, StringComparison.Ordinal);
            Assert.EndsWith(Charlie, text, StringComparison.Ordinal);
        }
        AssertRoundTrip(left, right, redline);
    }

    // ---- helpers -------------------------------------------------------------------------------

    private static bool IsRangeMarker(XElement e) =>
        e.Name.LocalName.StartsWith("customXmlIns", StringComparison.Ordinal)
        || e.Name.LocalName.StartsWith("customXmlDel", StringComparison.Ordinal);

    private static void AssertSingleWrapperWithoutRangeMarkers(XElement body, Wrapper kind)
    {
        Assert.Single(body.Descendants(WrapperName(kind)));
        Assert.DoesNotContain(body.Descendants(), IsRangeMarker);
    }

    /// <summary>Each paragraph whose text is one of <paramref name="texts"/> carries no revision markup at all:
    /// no run-level <c>w:ins</c>/<c>w:del</c>, no paragraph-mark revision and no property change.</summary>
    private static void AssertUnmarked(XElement body, params string[] texts)
    {
        foreach (var text in texts)
        {
            var paragraph = Assert.Single(body.Descendants(W.p), p => string.Concat(p.Descendants(W.t).Select(t => t.Value)) == text
                && !p.Descendants(W.delText).Any());
            Assert.DoesNotContain(paragraph.Descendants(), e => e.Name == W.ins || e.Name == W.del
                || e.Name == W.moveFrom || e.Name == W.moveTo || e.Name.LocalName.EndsWith("Change", StringComparison.Ordinal));
        }
    }

    /// <summary>accept(redline) equals <paramref name="revised"/> and reject(redline) equals
    /// <paramref name="original"/>, block for block: every wrapper with its exact properties, every table, and
    /// every paragraph's properties and formatted text (adjacent runs with the same properties merged, since
    /// the redline may split a run around a change). Bookmarks and comment ranges between blocks are left out:
    /// accepting revisions drops them (issue #1149).</summary>
    private static void AssertRoundTrip(WmlDocument original, WmlDocument revised, WmlDocument redline, string? because = null)
    {
        var accepted = Signature(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray));
        var rejected = Signature(DocxDiffOps.RejectRevisions(redline.DocumentByteArray));
        Assert.True(Signature(revised.DocumentByteArray) == accepted,
            $"accept ≠ revised ({because}):\n{Signature(revised.DocumentByteArray)}\n---\n{accepted}");
        Assert.True(Signature(original.DocumentByteArray) == rejected,
            $"reject ≠ original ({because}):\n{Signature(original.DocumentByteArray)}\n---\n{rejected}");
    }

    /// <summary>The body's signature followed by each footnote's.</summary>
    private static string Signature(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var sb = new StringBuilder(Signature(main.GetXDocument().Root!.Element(W.body)!));
        foreach (var note in main.FootnotesPart?.GetXDocument().Root!.Elements(W.footnote) ?? Enumerable.Empty<XElement>())
            if (note.Attribute(W.type) is null)
                sb.Append("FOOTNOTE ").Append((string?)note.Attribute(W.id)).Append('\n').Append(Signature(note));
        return sb.ToString();
    }

    private static string Signature(XElement container)
    {
        var sb = new StringBuilder();
        AppendBlocks(container, sb, 0);
        return sb.ToString();
    }

    private static void AppendBlocks(XElement container, StringBuilder sb, int depth)
    {
        foreach (var child in container.Elements())
        {
            var indent = new string(' ', depth * 2);
            if (child.Name == W.p)
            {
                sb.Append(indent).Append("P ").Append(Canonical(child.Element(W.pPr))).Append(" |");
                string? lastProps = null;
                foreach (var run in child.Descendants(W.r))
                {
                    var props = Canonical(run.Element(W.rPr));
                    if (props != lastProps)
                        sb.Append(" {").Append(props).Append("} ");
                    lastProps = props;
                    sb.Append(string.Concat(run.Elements().Where(e => e.Name != W.rPr)
                        .Select(e => e.Name == W.t ? e.Value : "<" + e.Name.LocalName + ">")));
                }
                sb.Append('\n');
            }
            else if (child.Name == W.tbl)
            {
                sb.Append(indent).Append("TBL ").Append(Canonical(child.Element(W.tblPr)))
                    .Append(Canonical(child.Element(W.tblGrid))).Append('\n');
                foreach (var row in child.Elements(W.tr))
                {
                    sb.Append(indent).Append(" TR ").Append(Canonical(row.Element(W.trPr))).Append('\n');
                    foreach (var cell in row.Elements(W.tc))
                    {
                        sb.Append(indent).Append("  TC ").Append(Canonical(cell.Element(W.tcPr))).Append('\n');
                        AppendBlocks(cell, sb, depth + 2);
                    }
                }
            }
            else if (child.Name == W.customXml)
            {
                sb.Append(indent).Append("CUSTOMXML ")
                    .Append(Canonical(new XElement(W.customXml, child.Attributes(), child.Element(W.customXmlPr)))).Append('\n');
                AppendBlocks(child, sb, depth + 1);
            }
            else if (child.Name == W.sdt)
            {
                sb.Append(indent).Append("SDT ").Append(Canonical(child.Element(W.sdtPr)))
                    .Append(Canonical(child.Element(W.sdtEndPr))).Append('\n');
                AppendBlocks(child.Element(W.sdtContent)!, sb, depth + 1);
            }
            else if (child.Name != W.sectPr && child.Name != W.customXmlPr && child.Name != W.pPr
                && child.Name != W.bookmarkStart && child.Name != W.bookmarkEnd
                && child.Name != W.commentRangeStart && child.Name != W.commentRangeEnd)
            {
                sb.Append(indent).Append("OTHER ").Append(child.Name.LocalName).Append('\n');
            }
        }
    }

    /// <summary>Canonical XML of a property element: rsid and PowerTools attributes dropped, attributes sorted, and
    /// an empty element read as an absent one.</summary>
    private static string Canonical(XElement? element)
    {
        if (element is null || (!element.HasElements && !element.HasAttributes && element.Value.Length == 0))
            return "-";
        static XElement Clean(XElement e) => new(e.Name,
            e.Attributes()
                .Where(a => !a.IsNamespaceDeclaration
                    && !a.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)
                    && a.Name.NamespaceName != "http://powertools.codeplex.com/2011")
                .OrderBy(a => a.Name.ToString(), StringComparer.Ordinal)
                .Select(a => new XAttribute(a.Name, a.Value)),
            e.Elements().Where(c => !c.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)).Select(Clean),
            e.HasElements ? null : e.Value);
        return Clean(element).ToString(SaveOptions.DisableFormatting);
    }

    private static XElement Body(WmlDocument document) => Body(document.DocumentByteArray);

    private static XElement Body(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return new XElement(document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!);
    }
}
