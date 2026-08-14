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
}
