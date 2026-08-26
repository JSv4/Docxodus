// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionLinkBookmarkTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static string[] Paragraphs(DocxSession session, string scope = "body") =>
        session.Project().AnchorIndex.Values
            .Where(a => a.Anchor.Scope == scope && a.Anchor.Kind is "p" or "h" or "li")
            .Select(a => a.Anchor.Id).Distinct().ToArray();

    [Fact]
    public void LB001_ExternalCrud_ReusesOwnerRelationship_AndCleansOnlyAfterLastReference()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchors = Paragraphs(session);

        var first = session.AddHyperlink(anchors[0], new CharSpan(0, 5),
            HyperlinkTarget.External("https://example.test/shared"));
        var second = session.AddHyperlink(anchors[1], new CharSpan(0, 6),
            HyperlinkTarget.External("https://example.test/shared"));

        Assert.True(first.Success, first.Error?.Message);
        Assert.True(second.Success, second.Error?.Message);
        var links = session.ListHyperlinks();
        Assert.Equal(2, links.Count);
        Assert.Single(links.Select(l => l.RelationshipId).Distinct());
        Assert.Single(HyperlinkRelationships(session.Save(true), m => m));

        Assert.True(session.RemoveHyperlink(first.HyperlinkId!).Success);
        Assert.Single(HyperlinkRelationships(session.Save(true), m => m));
        Assert.True(session.RemoveHyperlink(second.HyperlinkId!).Success);
        Assert.Empty(HyperlinkRelationships(session.Save(true), m => m));
    }

    [Fact]
    public void LB002_InternalLink_IsRelationshipFree_RenameRetargetsCrossPart_AndRemoveCannotDangle()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var body = Paragraphs(session);
        Assert.True(session.AddBookmark("TargetOne", DocumentRange.In(body[0], new CharSpan(0, 5))).Success);
        Assert.True(session.SetHeaderText(body[0], HeaderFooterKind.Default, "jump").Success);
        var header = Assert.Single(Paragraphs(session, "hdr1"));

        var add = session.AddHyperlink(header, new CharSpan(0, 4), HyperlinkTarget.Internal("TargetOne"));
        Assert.True(add.Success, add.Error?.Message);
        var link = Assert.Single(session.ListHyperlinks(ProjectionScopes.Headers));
        Assert.Equal(HyperlinkKind.Internal, link.Kind);
        Assert.Null(link.RelationshipId);
        Assert.Empty(HyperlinkRelationships(session.Save(true), m => m.HeaderParts.Single()));

        Assert.True(session.RenameBookmark("TargetOne", "TargetTwo").Success);
        Assert.Equal("TargetTwo", Assert.Single(session.ListHyperlinks(ProjectionScopes.Headers)).Target);
        var blocked = session.RemoveBookmark("TargetTwo");
        Assert.False(blocked.Success);
        Assert.Equal(EditErrorCode.BookmarkInUse, blocked.Error!.Code);

        Assert.True(session.UpdateHyperlink(add.HyperlinkId!,
            HyperlinkTarget.External("https://example.test/out")).Success);
        Assert.True(session.RemoveBookmark("TargetTwo").Success);
        Assert.Single(HyperlinkRelationships(session.Save(true), m => m.HeaderParts.Single()));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(5)]
    [InlineData(16)]
    public void LB003_CollapsedBookmark_StartAlwaysPrecedesEnd(int offset)
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = Paragraphs(session)[0];
        var result = session.AddBookmark("Point" + offset,
            new DocumentRange(anchor, offset, anchor, offset));
        Assert.True(result.Success, result.Error?.Message);

        var saved = session.Save();
        using var doc = WordprocessingDocument.Open(new MemoryStream(saved), false);
        var paragraph = doc.MainDocumentPart!.GetXDocument().Descendants(W + "p").First();
        var nodes = paragraph.DescendantsAndSelf().ToList();
        var start = nodes.Single(e => e.Name == W + "bookmarkStart");
        var end = nodes.Single(e => e.Name == W + "bookmarkEnd");
        Assert.True(XNode.DocumentOrderComparer.Compare(start, end) < 0);
    }

    [Fact]
    public void LB004_MultiParagraphBookmark_EnumeratesPreciseSegments_AndMoveKeepsPairId()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchors = Paragraphs(session);
        Assert.True(session.AddBookmark("AcrossParas",
            new DocumentRange(anchors[0], 6, anchors[1], 6)).Success);
        var before = Assert.Single(session.ListBookmarks());
        Assert.True(before.IsValid);
        Assert.Equal(2, before.Segments.Count);
        Assert.Equal("paragraph.\nSecond", before.Text);

        Assert.True(session.MoveBookmark("AcrossParas", DocumentRange.In(anchors[1], new CharSpan(7, 9))).Success);
        var after = Assert.Single(session.ListBookmarks());
        Assert.Equal(before.BookmarkId, after.BookmarkId);
        Assert.Equal("paragraph", after.Text);
    }

    [Fact]
    public void LB005_CrossPartBookmarkMutation_IsStructuredUnsupported()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var body = Paragraphs(session);
        Assert.True(session.SetHeaderText(body[0], HeaderFooterKind.Default, "header").Success);
        var header = Assert.Single(Paragraphs(session, "hdr1"));

        var result = session.AddBookmark("CrossPart",
            new DocumentRange(body[0], 0, header, 1));
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnsupportedInlineBoundary, result.Error!.Code);
        Assert.Empty(session.ListBookmarks());
    }

    [Fact]
    public void LB006_MarkdownInternalLink_WritesAnchorNotRelationship_AndMissingTargetIsStructured()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchors = Paragraphs(session);
        Assert.True(session.AddBookmark("Clause", DocumentRange.In(anchors[0], new CharSpan(0, 5))).Success);

        var ok = session.ReplaceText(anchors[1], "[go](#Clause)");
        Assert.True(ok.Success, ok.Error?.Message);
        var link = Assert.Single(session.ListHyperlinks());
        Assert.Equal(HyperlinkKind.Internal, link.Kind);
        Assert.Equal("Clause", link.Target);
        Assert.Empty(HyperlinkRelationships(session.Save(true), m => m));

        var missing = session.ReplaceText(anchors[1], "[bad](#Missing)");
        Assert.False(missing.Success);
        Assert.Equal(EditErrorCode.MissingBookmarkTarget, missing.Error!.Code);
    }

    [Fact]
    public void LB007_UndoRestoresRelationshipTopology_AndPersistedIdsRoundTrip()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = Paragraphs(session)[0];
        var add = session.AddHyperlink(anchor, new CharSpan(0, 5),
            HyperlinkTarget.External("https://example.test/a"));
        Assert.True(add.Success, add.Error?.Message);
        var persisted = session.Save(true);
        using (var reopened = new DocxSession(persisted))
            Assert.Equal(add.HyperlinkId, Assert.Single(reopened.ListHyperlinks()).Id);

        Assert.True(session.RemoveHyperlink(add.HyperlinkId!).Success);
        Assert.Empty(HyperlinkRelationships(session.Save(true), m => m));
        Assert.True(session.Undo());
        Assert.Single(HyperlinkRelationships(session.Save(true), m => m));
        Assert.Equal(add.HyperlinkId, Assert.Single(session.ListHyperlinks()).Id);
    }

    [Fact]
    public void LB008_HighOrphanEndId_IsNotReused_AndSavedPackageValidates()
    {
        var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        bytes = MutatePackage(bytes, doc =>
        {
            doc.MainDocumentPart!.GetXDocument().Descendants(W + "p").First()
                .Add(new XElement(W + "bookmarkEnd", new XAttribute(W + "id", "99")));
            doc.MainDocumentPart.PutXDocument();
        });

        using var session = new DocxSession(bytes);
        var anchor = Paragraphs(session)[0];
        Assert.True(session.AddBookmark("Fresh", DocumentRange.In(anchor, new CharSpan(0, 1))).Success);
        Assert.Equal("100", Assert.Single(session.ListBookmarks()).BookmarkId);
        var saved = session.Save();
        using var reopenedStream = new MemoryStream(saved);
        using var reopened = WordprocessingDocument.Open(reopenedStream, false);
        var realErrors = new OpenXmlValidator().Validate(reopened)
            .Where(e => !(e.Description ?? string.Empty).Contains("powertools.codeplex.com", StringComparison.Ordinal))
            .ToList();
        Assert.Empty(realErrors);
    }

    [Fact]
    public void LB010_DuplicateNumericIdsInDifferentParts_DoNotConfuseMoveOrRemove()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var bodyAnchors = Paragraphs(seed);
        Assert.True(seed.AddBookmark("BodyMark", DocumentRange.In(bodyAnchors[0], new CharSpan(0, 1))).Success);
        Assert.True(seed.SetHeaderText(bodyAnchors[0], HeaderFooterKind.Default, "header").Success);
        var headerAnchor = Assert.Single(Paragraphs(seed, "hdr1"));
        Assert.True(seed.AddBookmark("HeaderMark", DocumentRange.In(headerAnchor, new CharSpan(0, 1))).Success);
        var bytes = seed.Save();

        bytes = MutatePackage(bytes, doc =>
        {
            var main = doc.MainDocumentPart!;
            var bodyId = (string)main.GetXDocument()
                .Descendants(W + "bookmarkStart").Single().Attribute(W + "id")!;
            var header = main.HeaderParts.Single();
            foreach (var marker in header.GetXDocument().Descendants()
                .Where(e => e.Name == W + "bookmarkStart" || e.Name == W + "bookmarkEnd"))
                marker.SetAttributeValue(W + "id", bodyId);
            header.PutXDocument();
        });

        using var session = new DocxSession(bytes);
        var anchors = Paragraphs(session);
        Assert.Equal(2, session.ListBookmarks().Count);
        Assert.True(session.MoveBookmark("BodyMark", DocumentRange.In(anchors[1], new CharSpan(0, 1))).Success);
        Assert.True(session.RemoveBookmark("BodyMark").Success);
        var survivor = Assert.Single(session.ListBookmarks());
        Assert.Equal("HeaderMark", survivor.Name);
        Assert.True(survivor.IsPaired);
    }

    [Fact]
    public void LB009_DeleteLinkedBlock_CleansItsOrphan_ButKeepsSharedLiveRelationship()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchors = Paragraphs(session);
        Assert.True(session.AddHyperlink(anchors[0], new CharSpan(0, 5),
            HyperlinkTarget.External("https://example.test/shared")).Success);
        Assert.True(session.AddHyperlink(anchors[1], new CharSpan(0, 6),
            HyperlinkTarget.External("https://example.test/shared")).Success);

        Assert.True(session.DeleteBlock(anchors[0]).Success);
        Assert.Single(HyperlinkRelationships(session.Save(true), m => m));
        Assert.True(session.DeleteBlock(anchors[1]).Success);
        Assert.Empty(HyperlinkRelationships(session.Save(true), m => m));
    }

    [Fact]
    public void LB011_ReplaceCellContent_RejectsDeletingTargetedBookmark()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var body = Paragraphs(session);
        var inserted = session.InsertTable(body[0], Position.After, 1, 1,
            new TableInsertOptions { CellContents = new[] { "Cell text" } });
        Assert.True(inserted.Success, inserted.Error?.Message);
        // #450 reports canonical table identities through TableAnchors; locate the new
        // mutation-ready paragraph independently of that structural result envelope.
        var cellParagraph = Assert.Single(Paragraphs(session).Except(body));
        Assert.True(session.AddBookmark("CellTarget",
            DocumentRange.In(cellParagraph, new CharSpan(0, 4))).Success);
        Assert.True(session.AddHyperlink(body[1], new CharSpan(0, 6),
            HyperlinkTarget.Internal("CellTarget")).Success);

        var cell = session.Project().AnchorIndex.Keys.Single(id => id.StartsWith("tc:", StringComparison.Ordinal));
        var blocked = session.ReplaceCellContent(cell, "Replacement");
        Assert.False(blocked.Success);
        Assert.Equal(EditErrorCode.BookmarkInUse, blocked.Error!.Code);
        Assert.Equal("CellTarget", Assert.Single(session.ListBookmarks()).Name);
    }

    [Fact]
    public void LB012_SetHeaderText_RejectsDeletingTargetedBookmark()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var body = Paragraphs(session);
        Assert.True(session.SetHeaderText(body[0], HeaderFooterKind.Default, "Header target").Success);
        var header = Assert.Single(Paragraphs(session, "hdr1"));
        Assert.True(session.AddBookmark("HeaderTarget",
            DocumentRange.In(header, new CharSpan(0, 6))).Success);
        Assert.True(session.AddHyperlink(body[1], new CharSpan(0, 6),
            HyperlinkTarget.Internal("HeaderTarget")).Success);

        var blocked = session.SetHeaderText(body[0], HeaderFooterKind.Default, "Replacement");
        Assert.False(blocked.Success);
        Assert.Equal(EditErrorCode.BookmarkInUse, blocked.Error!.Code);
        Assert.Equal("HeaderTarget", Assert.Single(session.ListBookmarks()).Name);
    }

    [Fact]
    public void LB013_WholeParagraphReplacement_RetainsBookmarkCoordinatesAndClampsEnd()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = Paragraphs(session)[0];
        Assert.True(session.AddBookmark("StableRange",
            DocumentRange.In(anchor, new CharSpan(3, 5))).Success);

        Assert.True(session.ReplaceText(anchor, "abcdef").Success);
        var bookmark = Assert.Single(session.ListBookmarks());
        Assert.Equal(new CharSpan(3, 3), Assert.Single(bookmark.Segments).Span);
        Assert.Equal("def", bookmark.Text);
    }

    [Fact]
    public void LB014_MarkdownLinks_OwnRelationshipsInFooterFootnoteAndEndnoteParts()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var body = Paragraphs(session);
        Assert.True(session.SetFooterText(body[0], HeaderFooterKind.Default,
            "[footer](https://example.test/footer)").Success);
        Assert.True(session.InsertFootnote(body[0], 1,
            "[footnote](https://example.test/footnote)").Success);
        Assert.True(session.InsertEndnote(body[1], 1,
            "[endnote](https://example.test/endnote)").Success);

        var links = session.ListHyperlinks();
        Assert.Contains(links, link => link.Scope.StartsWith("ftr", StringComparison.Ordinal));
        Assert.Contains(links, link => link.Scope == "fn");
        Assert.Contains(links, link => link.Scope == "en");
        var saved = session.Save(true);
        Assert.Empty(HyperlinkRelationships(saved, m => m));
        Assert.Single(HyperlinkRelationships(saved, m => m.FooterParts.Single()));
        Assert.Single(HyperlinkRelationships(saved, m => m.FootnotesPart!));
        Assert.Single(HyperlinkRelationships(saved, m => m.EndnotesPart!));
    }

    [Fact]
    public void LB015_PartialFormattedSpan_PreservesRunProperties_AndTrackedMetadataOpsRejectCleanly()
    {
        using (var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs()))
        {
            var anchor = Paragraphs(session)[0];
            Assert.True(session.ReplaceText(anchor, "**Hello** world").Success);
            Assert.True(session.AddHyperlink(anchor, new CharSpan(1, 3),
                HyperlinkTarget.External("https://example.test/formatted")).Success);
            using var doc = WordprocessingDocument.Open(new MemoryStream(session.Save()), false);
            var paragraph = doc.MainDocumentPart!.GetXDocument().Descendants(W + "p").First();
            var linkRun = paragraph.Descendants(W + "hyperlink").Single().Element(W + "r")!;
            Assert.NotNull(linkRun.Element(W + "rPr")?.Element(W + "b"));
            Assert.Equal("ell", linkRun.Value);
            Assert.Equal("H", paragraph.Elements(W + "r").First().Value);
        }

        var settings = new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline };
        using var tracked = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), settings);
        var trackedAnchor = Paragraphs(tracked)[0];
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported,
            tracked.AddHyperlink(trackedAnchor, new CharSpan(0, 1),
                HyperlinkTarget.External("https://example.test")).Error!.Code);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported,
            tracked.AddBookmark("TrackedBookmark",
                DocumentRange.In(trackedAnchor, new CharSpan(0, 1))).Error!.Code);
        Assert.Empty(tracked.ListHyperlinks());
        Assert.Empty(tracked.ListBookmarks());
    }

    [Fact]
    public void LB016_TrackedWholeReplacementWithBookmark_RejectsBeforeSnapshot_ThenAcceptModeWorks()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = Paragraphs(session)[0];
        Assert.True(session.AddBookmark("TrackedBoundary",
            DocumentRange.In(anchor, new CharSpan(3, 5))).Success);
        var before = Assert.Single(session.ListBookmarks());
        int undoCount = session.UndoCount;

        session.SetTrackedChanges(TrackedChangeMode.RenderInline);
        var blocked = session.ReplaceText(anchor, "abcdef");
        Assert.False(blocked.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, blocked.Error!.Code);
        Assert.Equal(undoCount, session.UndoCount);
        Assert.Empty(session.ListRevisions());
        var unchanged = Assert.Single(session.ListBookmarks());
        Assert.Equal(before.Range, unchanged.Range);
        Assert.Equal(before.Text, unchanged.Text);

        session.SetTrackedChanges(TrackedChangeMode.Accept);
        Assert.True(session.ReplaceText(anchor, "abcdef").Success);
        Assert.Equal(new CharSpan(3, 3), Assert.Single(session.ListBookmarks()).Segments.Single().Span);
    }

    [Fact]
    public void LB017_DuplicateNamesAcrossStories_AreDiagnosticsAndAmbiguousTargetsFail()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var body = Paragraphs(seed);
        Assert.True(seed.AddBookmark("Duplicate", DocumentRange.In(body[0], new CharSpan(0, 1))).Success);
        Assert.True(seed.SetHeaderText(body[0], HeaderFooterKind.Default, "header").Success);
        var header = Assert.Single(Paragraphs(seed, "hdr1"));
        Assert.True(seed.AddBookmark("HeaderName", DocumentRange.In(header, new CharSpan(0, 1))).Success);
        var bytes = MutatePackage(seed.Save(), doc =>
        {
            doc.MainDocumentPart!.HeaderParts.Single().GetXDocument()
                .Descendants(W + "bookmarkStart").Single()
                .SetAttributeValue(W + "name", "Duplicate");
            doc.MainDocumentPart.HeaderParts.Single().PutXDocument();
        });

        using var session = new DocxSession(bytes);
        var diagnostics = session.ListBookmarks();
        Assert.Equal(2, diagnostics.Count);
        Assert.All(diagnostics, bookmark =>
        {
            Assert.False(bookmark.IsValid);
            Assert.Contains("duplicated", bookmark.ValidationError);
        });
        var blocked = session.AddHyperlink(Paragraphs(session)[1], new CharSpan(0, 1),
            HyperlinkTarget.Internal("Duplicate"));
        Assert.False(blocked.Success);
        Assert.Equal(EditErrorCode.DuplicateBookmarkName, blocked.Error!.Code);
    }

    [Fact]
    public void LB018_SplitAndMerge_PreserveBoundaryPointAndSpanningBookmarkPrecisely()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var first = Paragraphs(session)[0];
        Assert.True(session.AddBookmark("SpansSplit",
            DocumentRange.In(first, new CharSpan(2, 10))).Success);
        Assert.True(session.AddBookmark("PointAtSplit",
            new DocumentRange(first, 6, first, 6)).Success);
        var originalIds = session.ListBookmarks().ToDictionary(b => b.Name, b => b.BookmarkId);

        var split = session.SplitParagraph(first, 6);
        Assert.True(split.Success, split.Error?.Message);
        var second = Assert.Single(split.Created).Id;
        var afterSplit = session.ListBookmarks().ToDictionary(b => b.Name);

        var point = afterSplit["PointAtSplit"];
        Assert.True(point.IsValid, point.ValidationError);
        Assert.Equal(originalIds[point.Name], point.BookmarkId);
        Assert.Equal(new DocumentRange(second, 0, second, 0), point.Range);
        var pointSegment = Assert.Single(point.Segments);
        Assert.Equal(new CharSpan(0, 0), pointSegment.Span);
        Assert.Equal(string.Empty, point.Text);

        var spanning = afterSplit["SpansSplit"];
        Assert.True(spanning.IsValid, spanning.ValidationError);
        Assert.Equal(originalIds[spanning.Name], spanning.BookmarkId);
        Assert.Equal(new DocumentRange(first, 2, second, 6), spanning.Range);
        Assert.Collection(spanning.Segments,
            segment =>
            {
                Assert.Equal(first, segment.AnchorId);
                Assert.Equal(new CharSpan(2, 4), segment.Span);
                Assert.Equal("rst ", segment.Text);
            },
            segment =>
            {
                Assert.Equal(second, segment.AnchorId);
                Assert.Equal(new CharSpan(0, 6), segment.Span);
                Assert.Equal("paragr", segment.Text);
            });
        Assert.Equal("rst \nparagr", spanning.Text);
        AssertBookmarkPairsAndPackageValidity(session.Save(), originalIds);

        var merge = session.MergeParagraphs(first, second);
        Assert.True(merge.Success, merge.Error?.Message);
        var afterMerge = session.ListBookmarks().ToDictionary(b => b.Name);
        point = afterMerge["PointAtSplit"];
        Assert.Equal(new DocumentRange(first, 6, first, 6), point.Range);
        Assert.Equal(originalIds[point.Name], point.BookmarkId);
        spanning = afterMerge["SpansSplit"];
        Assert.Equal(new DocumentRange(first, 2, first, 12), spanning.Range);
        var mergedSegment = Assert.Single(spanning.Segments);
        Assert.Equal(new CharSpan(2, 10), mergedSegment.Span);
        Assert.Equal("rst paragr", spanning.Text);
        Assert.Equal(originalIds[spanning.Name], spanning.BookmarkId);
        AssertBookmarkPairsAndPackageValidity(session.Save(), originalIds);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LB019_MalformedBookmarkPair_CannotBeRenamedOrTargeted(bool duplicateEnd)
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var seedAnchors = Paragraphs(seed);
        Assert.True(seed.AddBookmark("Malformed",
            DocumentRange.In(seedAnchors[0], new CharSpan(0, 5))).Success);
        Assert.True(seed.AddHyperlink(seedAnchors[1], new CharSpan(0, 1),
            HyperlinkTarget.Internal("Malformed")).Success);
        var bytes = MutatePackage(seed.Save(), document =>
        {
            var end = document.MainDocumentPart!.GetXDocument().Descendants(W + "bookmarkEnd").Single();
            if (duplicateEnd) end.AddAfterSelf(new XElement(end));
            else end.Remove();
            document.MainDocumentPart.PutXDocument();
        });

        using var session = new DocxSession(bytes);
        var anchors = Paragraphs(session);
        Assert.False(Assert.Single(session.ListBookmarks()).IsValid);
        Assert.True(Assert.Single(session.ListHyperlinks()).IsBroken);
        int undoCount = session.UndoCount;

        var rename = session.RenameBookmark("Malformed", "Renamed");
        Assert.False(rename.Success);
        Assert.Equal(EditErrorCode.BookmarkNotFound, rename.Error!.Code);
        var link = session.AddHyperlink(anchors[1], new CharSpan(2, 1),
            HyperlinkTarget.Internal("Malformed"));
        Assert.False(link.Success);
        Assert.Equal(EditErrorCode.MissingBookmarkTarget, link.Error!.Code);
        Assert.Equal(undoCount, session.UndoCount);
        Assert.Equal("Malformed", Assert.Single(session.ListBookmarks()).Name);
    }

    [Fact]
    public void LB020_UnknownWireAndEnumHyperlinkKinds_AreStructuredErrorsWithoutMutation()
    {
        var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var probe = new DocxSession(bytes);
        var anchor = Paragraphs(probe)[0];
        var invalidEnum = probe.AddHyperlink(anchor, new CharSpan(0, 1),
            new HyperlinkTarget((HyperlinkKind)99, "https://example.test"));
        Assert.False(invalidEnum.Success);
        Assert.Equal(EditErrorCode.InvalidHyperlinkTarget, invalidEnum.Error!.Code);
        Assert.Empty(probe.ListHyperlinks());

        int handle = Docxodus.Internal.DocxSessionOps.OpenSession(bytes, null);
        try
        {
            using var add = JsonDocument.Parse(Docxodus.Internal.DocxSessionOps.AddHyperlink(
                handle, anchor, 0, 1, "externl", "https://example.test"));
            Assert.False(add.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal("invalid_hyperlink_target",
                add.RootElement.GetProperty("error").GetProperty("code").GetString());
            using (var emptyLinks = JsonDocument.Parse(
                Docxodus.Internal.DocxSessionOps.ListHyperlinks(handle)))
                Assert.Empty(emptyLinks.RootElement.EnumerateArray());

            using var validAdd = JsonDocument.Parse(Docxodus.Internal.DocxSessionOps.AddHyperlink(
                handle, anchor, 0, 1, "external", "https://example.test/original"));
            var hyperlinkId = validAdd.RootElement.GetProperty("hyperlinkId").GetString()!;
            using var update = JsonDocument.Parse(Docxodus.Internal.DocxSessionOps.UpdateHyperlink(
                handle, hyperlinkId, "internla", "Replacement"));
            Assert.False(update.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal("invalid_hyperlink_target",
                update.RootElement.GetProperty("error").GetProperty("code").GetString());
            using var links = JsonDocument.Parse(Docxodus.Internal.DocxSessionOps.ListHyperlinks(handle));
            Assert.Equal("https://example.test/original",
                Assert.Single(links.RootElement.EnumerateArray()).GetProperty("target").GetString());
        }
        finally
        {
            Docxodus.Internal.DocxSessionOps.CloseSession(handle);
        }
    }

    // A bookmark's OTHER consumer family: REF/PAGEREF/NOTEREF/HYPERLINK \l cross-reference fields.
    // Word writes every TOC entry that way, so a rename that retargets only w:hyperlink/@w:anchor
    // leaves "Error! Bookmark not defined." behind and a removal that only counts anchors reports
    // success while dangling the field. Both instruction carriers are covered, and the fldChar one
    // deliberately SPLITS its instruction across two w:instrText runs (Word does this constantly).
    [Fact]
    public void LB021_CrossReferenceFields_AreRetargetedByRename_AndBlockRemoval()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var seedAnchors = Paragraphs(seed);
        Assert.True(seed.AddBookmark("ClauseOne",
            DocumentRange.In(seedAnchors[0], new CharSpan(0, 5))).Success);
        var bytes = AppendBodyParagraphs(seed.Save(),
            SplitInstructionField(" PAGEREF Clause", "One \\h ", "3"),
            new XElement(W + "p",
                new XElement(W + "fldSimple", new XAttribute(W + "instr", " REF ClauseOne \\h "),
                    new XElement(W + "r", new XElement(W + "t", "First")))),
            SplitInstructionField(" HYPERLINK \\l \"ClauseOne", "\" ", "jump"));

        using var session = new DocxSession(bytes);
        var blocked = session.RemoveBookmark("ClauseOne");
        Assert.False(blocked.Success);
        Assert.Equal(EditErrorCode.BookmarkInUse, blocked.Error!.Code);

        Assert.True(session.RenameBookmark("ClauseOne", "ClauseTwo").Success);
        Assert.Equal(new[] { " PAGEREF ClauseTwo \\h ", " REF ClauseTwo \\h ", " HYPERLINK \\l \"ClauseTwo\" " },
            Instructions(session.Save()));
        // Retargeting is what keeps removal blocked: the fields now point at the NEW name.
        Assert.Equal(EditErrorCode.BookmarkInUse, session.RemoveBookmark("ClauseTwo").Error!.Code);
        Assert.Equal("ClauseTwo", Assert.Single(session.ListBookmarks()).Name);
        AssertBookmarkPairsAndPackageValidity(session.Save(),
            new System.Collections.Generic.Dictionary<string, string>
            {
                ["ClauseTwo"] = Assert.Single(session.ListBookmarks()).BookmarkId,
            });
    }

    // Word owns the _GoBack/_Toc*/_Ref*/_Hlt*/_Hlk* namespace and reallocates it whenever a TOC or
    // cross-reference is refreshed, so creating a name inside it is refused. Names Word already put
    // there stay fully mutable — with their cross-reference fields retargeted like any other.
    [Theory]
    [InlineData("_GoBack")]
    [InlineData("_Toc12345")]
    [InlineData("_Ref99")]
    [InlineData("_Hlt7")]
    [InlineData("_Hlk7")]
    public void LB022_ReservedWordBookmarkNames_CannotBeCreated(string reserved)
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = Paragraphs(session)[0];

        var created = session.AddBookmark(reserved, DocumentRange.In(anchor, new CharSpan(0, 5)));
        Assert.False(created.Success);
        Assert.Equal(EditErrorCode.InvalidBookmarkName, created.Error!.Code);
        Assert.Empty(session.ListBookmarks());

        Assert.True(session.AddBookmark("Ordinary", DocumentRange.In(anchor, new CharSpan(0, 5))).Success);
        var renamed = session.RenameBookmark("Ordinary", reserved);
        Assert.False(renamed.Success);
        Assert.Equal(EditErrorCode.InvalidBookmarkName, renamed.Error!.Code);
        Assert.Equal("Ordinary", Assert.Single(session.ListBookmarks()).Name);
    }

    [Fact]
    public void LB023_ExistingReservedTocBookmark_StaysRenamableWithItsPageRefField()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var seedAnchors = Paragraphs(seed);
        Assert.True(seed.AddBookmark("Heading", DocumentRange.In(seedAnchors[0], new CharSpan(0, 5))).Success);
        var bytes = AppendBodyParagraphs(seed.Save(), SplitInstructionField(" PAGEREF _Toc", "1 \\h ", "1"));
        bytes = MutatePackage(bytes, document =>
        {
            document.MainDocumentPart!.GetXDocument().Descendants(W + "bookmarkStart").Single()
                .SetAttributeValue(W + "name", "_Toc1");
            document.MainDocumentPart.PutXDocument();
        });

        using var session = new DocxSession(bytes);
        Assert.Equal(EditErrorCode.BookmarkInUse, session.RemoveBookmark("_Toc1").Error!.Code);
        var renamed = session.RenameBookmark("_Toc1", "Intro");
        Assert.True(renamed.Success, renamed.Error?.Message);
        Assert.Equal(new[] { " PAGEREF Intro \\h " }, Instructions(session.Save()));
        // The retargeted field keeps protecting the bookmark under its NEW name; deleting the field
        // is what releases it.
        Assert.Equal(EditErrorCode.BookmarkInUse, session.RemoveBookmark("Intro").Error!.Code);
        Assert.True(session.DeleteBlock(Paragraphs(session).Last()).Success);
        Assert.True(session.RemoveBookmark("Intro").Success);
    }

    // AddHyperlink used to move only the selected w:r elements, stranding every zero-width marker
    // that sat BETWEEN them after the finished w:hyperlink. For a bookmark whose start is inside the
    // span that puts the start after its own end: the pair stops resolving and the bookmark becomes
    // permanently unmutatable and untargetable. Comment ranges break by the same mechanism.
    [Fact]
    public void LB024_HyperlinkOverARangeMarker_RelocatesItInsteadOfStrandingIt()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = Paragraphs(session)[0];
        Assert.True(session.AddBookmark("Inner", DocumentRange.In(anchor, new CharSpan(5, 3))).Success);
        Assert.True(session.AddComment(anchor, new CharSpan(9, 4), "Reviewer", "note").Success);

        var wrap = session.AddHyperlink(anchor, new CharSpan(0, 15),
            HyperlinkTarget.External("https://example.test/wrap"));
        Assert.True(wrap.Success, wrap.Error?.Message);

        var bookmark = Assert.Single(session.ListBookmarks());
        Assert.True(bookmark.IsValid, bookmark.ValidationError);
        Assert.Equal(new CharSpan(5, 3), Assert.Single(bookmark.Segments).Span);
        Assert.Equal(" pa", bookmark.Text);
        var saved = session.Save();
        using (var document = WordprocessingDocument.Open(new MemoryStream(saved), false))
        {
            var link = document.MainDocumentPart!.GetXDocument().Descendants(W + "hyperlink").Single();
            // Relocated INTO the link, in their original order, so document order still has each
            // start ahead of its end.
            Assert.Single(link.Descendants(W + "bookmarkStart"));
            Assert.Single(link.Descendants(W + "bookmarkEnd"));
            Assert.Single(link.Descendants(W + "commentRangeStart"));
            Assert.Single(link.Descendants(W + "commentRangeEnd"));
        }
        AssertBookmarkPairsAndPackageValidity(saved,
            new System.Collections.Generic.Dictionary<string, string> { ["Inner"] = bookmark.BookmarkId });

        // Still fully mutable: the pair resolves, so rename/move/remove all reach it.
        Assert.True(session.RenameBookmark("Inner", "Renamed").Success);
        Assert.True(session.RemoveBookmark("Renamed").Success);
    }

    // Bookmark w:id is story-part scoped and Word reuses the same decimal in several parts, so a
    // cross-part move that carries its id can collide with a bookmark already living there. Because
    // pairing demands exactly one start and one end per id, the collision breaks BOTH bookmarks.
    [Fact]
    public void LB025_CrossPartMove_TakesAFreshId_InsteadOfCollidingWithTheDestinationPart()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var bodyAnchors = Paragraphs(seed);
        Assert.True(seed.AddBookmark("BodyMark", DocumentRange.In(bodyAnchors[0], new CharSpan(0, 1))).Success);
        Assert.True(seed.SetHeaderText(bodyAnchors[0], HeaderFooterKind.Default, "header text").Success);
        var headerAnchor = Assert.Single(Paragraphs(seed, "hdr1"));
        Assert.True(seed.AddBookmark("HeaderMark", DocumentRange.In(headerAnchor, new CharSpan(0, 1))).Success);

        var bytes = MutatePackage(seed.Save(), document =>
        {
            var main = document.MainDocumentPart!;
            var bodyId = (string)main.GetXDocument()
                .Descendants(W + "bookmarkStart").Single().Attribute(W + "id")!;
            var header = main.HeaderParts.Single();
            foreach (var marker in header.GetXDocument().Descendants()
                .Where(e => e.Name == W + "bookmarkStart" || e.Name == W + "bookmarkEnd"))
                marker.SetAttributeValue(W + "id", bodyId);
            header.PutXDocument();
        });

        using var session = new DocxSession(bytes);
        var header2 = Assert.Single(Paragraphs(session, "hdr1"));
        var before = session.ListBookmarks().Single(b => b.Name == "BodyMark").BookmarkId;

        Assert.True(session.MoveBookmark("BodyMark",
            DocumentRange.In(header2, new CharSpan(2, 3))).Success);

        var after = session.ListBookmarks();
        Assert.Equal(2, after.Count);
        Assert.All(after, bookmark => Assert.True(bookmark.IsValid, bookmark.ValidationError));
        Assert.Equal(2, after.Select(bookmark => bookmark.BookmarkId).Distinct().Count());
        Assert.NotEqual(before, after.Single(b => b.Name == "BodyMark").BookmarkId);
        // Both survivors stay individually addressable, which the id collision used to prevent.
        Assert.True(session.RenameBookmark("BodyMark", "MovedMark").Success);
        Assert.True(session.RemoveBookmark("HeaderMark").Success);
        AssertPackageValidity(session.Save());
    }

    // Splitting a paragraph inside a hyperlink cuts the w:hyperlink in two. The halves used to copy
    // the SAME PtOpenXml.Unid, so both projected to one hl:body:<unid> id and only the first was
    // ever found — Update/RemoveHyperlink then silently affected half the link.
    [Fact]
    public void LB026_SplitInsideAHyperlink_GivesEachHalfItsOwnAddressableId()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = Paragraphs(session)[0];
        var add = session.AddHyperlink(anchor, new CharSpan(0, 15),
            HyperlinkTarget.External("https://example.test/whole"));
        Assert.True(add.Success, add.Error?.Message);

        var split = session.SplitParagraph(anchor, 6);
        Assert.True(split.Success, split.Error?.Message);

        var halves = session.ListHyperlinks();
        Assert.Equal(2, halves.Count);
        Assert.Equal(2, halves.Select(link => link.Id).Distinct().Count());
        Assert.Equal(add.HyperlinkId, halves[0].Id);

        // The SECOND half is independently addressable — distinct ids alone would not prove that.
        Assert.True(session.UpdateHyperlink(halves[1].Id,
            HyperlinkTarget.External("https://example.test/second")).Success);
        var retargeted = session.ListHyperlinks();
        Assert.Equal("https://example.test/whole", retargeted[0].Target);
        Assert.Equal("https://example.test/second", retargeted[1].Target);

        Assert.True(session.RemoveHyperlink(halves[1].Id).Success);
        Assert.Equal(add.HyperlinkId, Assert.Single(session.ListHyperlinks()).Id);
        AssertPackageValidity(session.Save());
    }

    // Widening BookmarkInUse to cross-reference fields also widens the structural-deletion guard,
    // and the population that reaches is large: Word puts a _Toc bookmark on every heading and the
    // matching PAGEREF lives in a DIFFERENT paragraph, so deleting a heading of a TOC'd document is
    // refused in DEFAULT (untracked) mode with no force/opt-out. That is the PR's existing
    // no-dangling-reference policy (LB011/LB012 pin the same rule for w:anchor links), but it is
    // broad enough that it must be an explicit, pinned decision rather than an emergent one.
    // Deleting the citing field WITH the heading is still allowed — the guard is reference-scoped,
    // not marker-scoped.
    [Fact]
    public void LB028_DeletingABookmarkedHeading_IsRefusedWhileACrossReferenceFieldSurvives()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var seedAnchors = Paragraphs(seed);
        Assert.True(seed.AddBookmark("Heading", DocumentRange.In(seedAnchors[0], new CharSpan(0, 5))).Success);
        var bytes = AppendBodyParagraphs(seed.Save(), SplitInstructionField(" PAGEREF Head", "ing \\h ", "1"));

        using var session = new DocxSession(bytes);
        var anchors = Paragraphs(session);
        var blocked = session.DeleteBlock(anchors[0]);
        Assert.False(blocked.Success);
        Assert.Equal(EditErrorCode.BookmarkInUse, blocked.Error!.Code);
        Assert.Equal("Heading", Assert.Single(session.ListBookmarks()).Name);

        // Removing the citing field first releases the heading; there is no force flag.
        Assert.True(session.DeleteBlock(anchors.Last()).Success);
        Assert.True(session.DeleteBlock(anchors[0]).Success);
        Assert.Empty(session.ListBookmarks());
    }

    // The comments part is a story like any other: its paragraphs are anchor-addressable
    // (p:cmt:<unid>) and editable, so it owns its own hyperlink relationships and can hold bookmark
    // markers. Leaving it out of the story-part list did not narrow the answer, it broke it —
    // ProjectionScopes.Comments returned silently empty and FindOwner returned null, which is what
    // made ReplaceText on a comment paragraph carrying a link throw (DS364).
    [Fact]
    public void LB027_CommentStory_IsAFirstClassScopeForLinksAndBookmarks()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var body = Paragraphs(session)[0];
        var comment = session.AddComment(body, new CharSpan(0, 5), "Alice", "Review this clause.");
        Assert.True(comment.Success, comment.Error?.Message);
        var commentParagraph = comment.Created.First(a => a.Kind == "p" && a.Scope == "cmt").Id;

        var link = session.AddHyperlink(commentParagraph, new CharSpan(0, 6),
            HyperlinkTarget.External("https://example.test/cmt"));
        Assert.True(link.Success, link.Error?.Message);
        var listed = Assert.Single(session.ListHyperlinks(ProjectionScopes.Comments));
        Assert.Equal("cmt", listed.Scope);
        Assert.Equal(link.HyperlinkId, listed.Id);
        Assert.Equal("Review", listed.Text);
        Assert.Empty(session.ListHyperlinks(ProjectionScopes.Body));
        // The relationship belongs to the comments part, not the main document part.
        Assert.Empty(HyperlinkRelationships(session.Save(true), m => m));
        Assert.Single(HyperlinkRelationships(session.Save(true), m => m.WordprocessingCommentsPart!));

        Assert.True(session.AddBookmark("InComment",
            DocumentRange.In(commentParagraph, new CharSpan(0, 6))).Success);
        var bookmark = Assert.Single(session.ListBookmarks(ProjectionScopes.Comments));
        Assert.Equal("cmt", bookmark.StartScope);
        Assert.True(bookmark.IsValid, bookmark.ValidationError);
        Assert.Equal("Review", bookmark.Text);
        Assert.Empty(session.ListBookmarks(ProjectionScopes.Body));

        Assert.True(session.RemoveHyperlink(link.HyperlinkId!).Success);
        Assert.Empty(HyperlinkRelationships(session.Save(true), m => m.WordprocessingCommentsPart!));
        AssertPackageValidity(session.Save());
    }

    /// <summary>A fldChar field whose instruction is deliberately split across two w:instrText runs.</summary>
    private static XElement SplitInstructionField(string head, string tail, string cachedResult) =>
        new(W + "p",
            new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"))),
            new XElement(W + "r", new XElement(W + "instrText",
                new XAttribute(XNamespace.Xml + "space", "preserve"), head)),
            new XElement(W + "r", new XElement(W + "instrText",
                new XAttribute(XNamespace.Xml + "space", "preserve"), tail)),
            new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))),
            new XElement(W + "r", new XElement(W + "t", cachedResult)),
            new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))));

    /// <summary>Every field instruction in the saved body, fldSimple and fldChar alike, in document order.</summary>
    private static string[] Instructions(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var body = document.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!;
        return body.Elements(W + "p")
            .Select(paragraph => paragraph.Descendants(W + "fldSimple").FirstOrDefault() is { } simple
                ? (string?)simple.Attribute(W + "instr")
                : paragraph.Descendants(W + "instrText").Any()
                    ? string.Concat(paragraph.Descendants(W + "instrText").Select(i => i.Value))
                    : null)
            .Where(instruction => instruction is not null)
            .ToArray()!;
    }

    private static byte[] AppendBodyParagraphs(byte[] bytes, params XElement[] paragraphs) =>
        MutatePackage(bytes, document =>
        {
            var body = document.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!;
            if (body.Element(W + "sectPr") is { } sectPr) sectPr.AddBeforeSelf(paragraphs);
            else body.Add(paragraphs);
            document.MainDocumentPart.PutXDocument();
        });

    private static void AssertPackageValidity(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var realErrors = new OpenXmlValidator().Validate(document)
            .Where(error => !(error.Description ?? string.Empty)
                .Contains("powertools.codeplex.com", StringComparison.Ordinal))
            .ToList();
        Assert.Empty(realErrors);
    }

    private static void AssertBookmarkPairsAndPackageValidity(byte[] bytes,
        System.Collections.Generic.IReadOnlyDictionary<string, string> expectedIds)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var root = document.MainDocumentPart!.GetXDocument();
        foreach (var (name, id) in expectedIds)
        {
            var start = Assert.Single(root.Descendants(W + "bookmarkStart"),
                marker => (string?)marker.Attribute(W + "name") == name);
            var end = Assert.Single(root.Descendants(W + "bookmarkEnd"),
                marker => (string?)marker.Attribute(W + "id") == id);
            Assert.Equal(id, (string?)start.Attribute(W + "id"));
            Assert.True(XNode.DocumentOrderComparer.Compare(start, end) < 0);
        }
        var realErrors = new OpenXmlValidator().Validate(document)
            .Where(error => !(error.Description ?? string.Empty)
                .Contains("powertools.codeplex.com", StringComparison.Ordinal))
            .ToList();
        Assert.Empty(realErrors);
    }

    private static HyperlinkRelationship[] HyperlinkRelationships(
        byte[] bytes, Func<MainDocumentPart, OpenXmlPart> owner)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return owner(document.MainDocumentPart!).HyperlinkRelationships.ToArray();
    }

    private static MemoryStream Expandable(byte[] bytes)
    {
        var stream = new MemoryStream(bytes.Length + 4096);
        stream.Write(bytes);
        stream.Position = 0;
        return stream;
    }

    private static byte[] MutatePackage(byte[] bytes, Action<WordprocessingDocument> mutate)
    {
        using var stream = Expandable(bytes);
        using (var document = WordprocessingDocument.Open(stream, true)) mutate(document);
        return stream.ToArray();
    }

    // ─── InsertCrossReference (issue #545) ────────────────────────────────

    [Fact]
    public void LB029_InsertCrossReference_WritesRefFieldWithCachedTargetText()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body")
            .Select(t => t.Anchor.Id).ToArray();
        var first = paragraphs[0];
        var second = paragraphs[1];
        var firstText = session.GetAnchorInfo(first)!.VisibleText;
        var word = firstText.Split(' ')[0];
        Assert.True(session.AddBookmark("defs",
            DocumentRange.In(first, new CharSpan(0, word.Length))).Success);

        var result = session.InsertCrossReference(second, 0, "defs");

        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(second, Assert.Single(result.Modified).Id);
        var xml = XElement.Parse(session.Raw.GetXml(second));
        var field = Assert.Single(xml.Descendants(
            XName.Get("fldSimple", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")));
        var instr = (string?)field.Attribute(
            XName.Get("instr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"));
        Assert.Contains("REF defs", instr);
        Assert.DoesNotContain("\\r", instr);
        Assert.DoesNotContain("\\h", instr);
        // The cached result run holds the bookmarked text, so a renderer that does not
        // recompute fields shows the referenced content.
        Assert.Equal(word, field.Value);
        Assert.StartsWith(word, session.GetAnchorInfo(second)!.VisibleText);
    }

    [Fact]
    public void LB030_InsertCrossReference_SwitchesShapeInstructionAndCachedResult()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body")
            .Select(t => t.Anchor.Id).ToArray();
        Assert.True(session.AddBookmark("target",
            DocumentRange.In(paragraphs[0], new CharSpan(0, 4))).Success);

        // \r on an unnumbered target caches Word's own value for a numberless paragraph: 0.
        var number = session.InsertCrossReference(paragraphs[1], 0, "target",
            new CrossReferenceOptions { ReferenceNumber = true, Hyperlink = true });
        Assert.True(number.Success, number.Error?.Message);
        var xml = XElement.Parse(session.Raw.GetXml(paragraphs[1]));
        var w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var field = Assert.Single(xml.Descendants(XName.Get("fldSimple", w)));
        var instr = (string?)field.Attribute(XName.Get("instr", w));
        Assert.Contains("\\r", instr);
        Assert.Contains("\\h", instr);
        Assert.Equal("0", field.Value);

        // \p alone caches only the position word; the bookmark is ABOVE this insertion point.
        var position = session.InsertCrossReference(paragraphs[1], 0, "target",
            new CrossReferenceOptions { IncludePosition = true });
        Assert.True(position.Success, position.Error?.Message);
        var fields = XElement.Parse(session.Raw.GetXml(paragraphs[1]))
            .Descendants(XName.Get("fldSimple", w)).ToArray();
        Assert.Equal(2, fields.Length);
        var positional = fields.Single(f =>
            ((string?)f.Attribute(XName.Get("instr", w)))!.Contains("\\p"));
        Assert.Equal("above", positional.Value);

        // Referencing from BEFORE the bookmark caches "below".
        var below = session.InsertCrossReference(paragraphs[0], 0, "target",
            new CrossReferenceOptions { IncludePosition = true });
        Assert.True(below.Success, below.Error?.Message);
        var belowField = XElement.Parse(session.Raw.GetXml(paragraphs[0]))
            .Descendants(XName.Get("fldSimple", w))
            .Single(f => ((string?)f.Attribute(XName.Get("instr", w)))!.Contains("\\p"));
        Assert.Equal("below", belowField.Value);
    }

    [Fact]
    public void LB031_InsertCrossReference_StructuredFailures()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body").Anchor.Id;

        var missing = session.InsertCrossReference(anchor, 0, "no_such_bookmark");
        Assert.False(missing.Success);
        Assert.Equal(EditErrorCode.MissingBookmarkTarget, missing.Error!.Code);

        Assert.True(session.AddBookmark("real",
            DocumentRange.In(anchor, new CharSpan(0, 3))).Success);
        var outOfRange = session.InsertCrossReference(anchor, 10_000, "real");
        Assert.False(outOfRange.Success);
        Assert.Equal(EditErrorCode.OffsetOutOfRange, outOfRange.Error!.Code);

        var badName = session.InsertCrossReference(anchor, 0, "has space");
        Assert.False(badName.Success);
        Assert.Equal(EditErrorCode.InvalidBookmarkName, badName.Error!.Code);

        // Undo restores the pre-field paragraph in one step.
        var before = session.GetAnchorInfo(anchor)!.VisibleText;
        Assert.True(session.InsertCrossReference(anchor, 0, "real").Success);
        Assert.True(session.Undo());
        Assert.Equal(before, session.GetAnchorInfo(anchor)!.VisibleText);
    }
}
