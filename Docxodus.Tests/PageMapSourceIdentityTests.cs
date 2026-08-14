#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class PageMapSourceIdentityTests
{
    [Fact]
    public void PM100_FinalConverterTreesCarryCanonicalIdentityAcrossEveryRenderedStoryAndTableLevel()
    {
        using var session = new DocxSession(
            DocxSessionTests.BuildDS003_TableWithCells(),
            new DocxSessionSettings
            {
                TrackedChanges = TrackedChangeMode.RenderInline,
                RevisionAuthor = "PageMap test",
            });
        var body = session.Project().AnchorIndex.Values
            .First(target => target.Anchor.Scope == "body" && target.TextPreview == "After table.");

        Assert.True(session.SetHeaderText(body.Anchor.Id, HeaderFooterKind.Default, "After table.").Success);
        Assert.True(session.SetFooterText(body.Anchor.Id, HeaderFooterKind.Default, "Footer source").Success);
        Assert.True(session.InsertFootnote(body.Anchor.Id, 5, "Footnote source").Success);
        Assert.True(session.InsertEndnote(body.Anchor.Id, 6, "Endnote source").Success);
        Assert.True(session.AddComment(
            body.Anchor.Id,
            new CharSpan(0, 5),
            "Reviewer",
            "Comment source").Success);

        Assert.True(session.ReplaceText(body.Anchor.Id, "Tracked replacement").Success);

        var projection = session.Project();
        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(session, new HtmlConversionOptions
        {
            StampAnchors = true,
            FabricateCssClasses = false,
            RenderHeadersAndFooters = true,
            RenderFootnotesAndEndnotes = true,
            CommentRenderMode = (int)CommentRenderMode.EndnoteStyle,
            RenderTrackedChanges = true,
        }));
        var sourceIds = html.DescendantsAndSelf()
            .Attributes("data-source-anchor-id")
            .Select(attribute => attribute.Value)
            .ToHashSet(StringComparer.Ordinal);

        void AssertRendered(string kind, Func<string, bool> scope)
        {
            var candidates = projection.AnchorIndex.Values
                .Where(target => target.Anchor.Kind == kind && scope(target.Anchor.Scope))
                .Select(target => target.Anchor.Id)
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            Assert.NotEmpty(candidates);
            Assert.Contains(candidates, sourceIds.Contains);
        }

        AssertRendered("p", scope => scope == "body");
        AssertRendered("tbl", scope => scope == "body");
        AssertRendered("tr", scope => scope == "body");
        AssertRendered("tc", scope => scope == "body");
        AssertRendered("p", scope => scope.StartsWith("hdr", StringComparison.Ordinal));
        AssertRendered("p", scope => scope.StartsWith("ftr", StringComparison.Ordinal));
        AssertRendered("fn", scope => scope == "fn");
        AssertRendered("p", scope => scope == "fn");
        AssertRendered("en", scope => scope == "en");
        AssertRendered("p", scope => scope == "en");
        AssertRendered("cmt", scope => scope == "cmt");
        AssertRendered("p", scope => scope == "cmt");

        var tracked = projection.AnchorIndex.Values.Single(target =>
            target.Anchor.Scope == "body" && target.TextPreview == "Tracked replacement");
        Assert.Contains(tracked.Anchor.Id, sourceIds);
        Assert.Contains("rev-", html.ToString(SaveOptions.DisableFormatting));

        using var inlineSession = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var inlineBody = inlineSession.Project().AnchorIndex.Values
            .First(target => target.Anchor.Scope == "body" && target.TextPreview == "First paragraph.");
        Assert.True(inlineSession.AddComment(
            inlineBody.Anchor.Id, new CharSpan(0, 5), "Reviewer", "Inline first.\n\nInline second.").Success);
        var inlineProjection = inlineSession.Project();
        var inlineHtml = XElement.Parse(HtmlConversionOps.ConvertToHtml(inlineSession, new HtmlConversionOptions
        {
            StampAnchors = true,
            FabricateCssClasses = false,
            CommentRenderMode = (int)CommentRenderMode.Inline,
        }));
        var commentDefinition = inlineProjection.AnchorIndex.Values.Single(target => target.Anchor.Kind == "cmt");
        Assert.Contains(inlineHtml.DescendantsAndSelf(), element =>
            (string?)element.Attribute("data-source-anchor-id") == commentDefinition.Anchor.Id);
        var commentParagraphs = inlineProjection.AnchorIndex.Values
            .Where(target => target.Anchor.Scope == "cmt" && target.Anchor.Kind == "p")
            .Select(target => target.Anchor.Id)
            .ToArray();
        Assert.Equal(2, commentParagraphs.Length);
        Assert.All(commentParagraphs, anchor => Assert.Contains(
            inlineHtml.DescendantsAndSelf(), element =>
                (string?)element.Attribute("data-source-anchor-id") == anchor));
    }

    [Fact]
    public void PM101_BareUnidCollisionAcrossStoriesKeepsDistinctCanonicalSourceIdentity()
    {
        byte[] collisionBytes;
        using (var seedSession = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs()))
        {
            var body = seedSession.Project().AnchorIndex.Values.First(target => target.TextPreview == "First paragraph.");
            Assert.True(seedSession.SetHeaderText(
                body.Anchor.Id,
                HeaderFooterKind.Default,
                "First paragraph.").Success);

            using var packageStream = new MemoryStream();
            packageStream.Write(seedSession.Save(persistAnchorIds: true));
            packageStream.Position = 0;
            using (var document = WordprocessingDocument.Open(packageStream, true))
            {
                var mainPart = document.MainDocumentPart!;
                var bodyParagraph = mainPart.GetXDocument().Descendants(W.p)
                    .First(paragraph => paragraph.Value == "First paragraph.");
                var headerPart = mainPart.HeaderParts.Single();
                var headerParagraph = headerPart.GetXDocument().Descendants(W.p)
                    .First(paragraph => paragraph.Value == "First paragraph.");
                var sharedUnid = (string)bodyParagraph.Attribute(PtOpenXml.Unid)!;
                headerParagraph.SetAttributeValue(PtOpenXml.Unid, sharedUnid);
                headerPart.PutXDocument();
            }

            collisionBytes = packageStream.ToArray();
        }

        using var session = new DocxSession(collisionBytes);

        var projection = session.Project();
        var bodySame = projection.AnchorIndex.Values.Single(target =>
            target.Anchor.Scope == "body" && target.TextPreview == "First paragraph.");
        var headerSame = projection.AnchorIndex.Values.Single(target =>
            target.Anchor.Scope.StartsWith("hdr", StringComparison.Ordinal)
            && target.TextPreview == "First paragraph.");
        Assert.Equal(bodySame.Unid, headerSame.Unid);
        Assert.NotEqual(bodySame.Anchor.Id, headerSame.Anchor.Id);

        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(session, new HtmlConversionOptions
        {
            StampAnchors = true,
            FabricateCssClasses = false,
            RenderHeadersAndFooters = true,
        }));
        var identities = html.DescendantsAndSelf()
            .Attributes("data-source-anchor-id")
            .Select(attribute => attribute.Value)
            .ToArray();
        Assert.Contains(bodySame.Anchor.Id, identities);
        Assert.Contains(headerSame.Anchor.Id, identities);
    }

    [Fact]
    public void PM102_PaginatedStatelessHtmlAlwaysCarriesCanonicalIdentityAndStagesComments()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var body = session.Project().AnchorIndex.Values.First(target =>
            target.Anchor.Scope == "body" && target.Anchor.Kind == "p");
        Assert.True(session.AddComment(
            body.Anchor.Id, new CharSpan(0, 5), "Reviewer", "First.\n\nSecond.").Success);

        XElement Convert(CommentRenderMode mode) => XElement.Parse(HtmlConversionOps.ConvertToHtml(
            session,
            new HtmlConversionOptions
            {
                StampAnchors = false,
                FabricateCssClasses = false,
                PaginationMode = (int)PaginationMode.Paginated,
                CommentRenderMode = (int)mode,
            }));

        var endnoteHtml = Convert(CommentRenderMode.EndnoteStyle);
        Assert.DoesNotContain(endnoteHtml.DescendantsAndSelf().Attributes("data-anchor"), _ => true);
        Assert.Contains(endnoteHtml.DescendantsAndSelf().Attributes("data-source-anchor-id"), _ => true);
        var staging = endnoteHtml.Descendants().Single(element =>
            (string?)element.Attribute("id") == "pagination-staging");
        var finalSection = staging.Descendants().Last(element =>
            element.Attribute("data-section-index") is not null);
        Assert.Contains(finalSection.Descendants(), element =>
            ((string?)element.Attribute("class"))?.Contains("comments-section", StringComparison.Ordinal) == true);

        var marginHtml = Convert(CommentRenderMode.Margin);
        var marginStaging = marginHtml.Descendants().Single(element =>
            (string?)element.Attribute("id") == "pagination-staging");
        var registry = marginStaging.Descendants().Single(element =>
            (string?)element.Attribute("id") == "pagination-comment-margin-registry");
        Assert.Contains(registry.DescendantsAndSelf().Attributes("data-source-anchor-id"), _ => true);
        Assert.DoesNotContain(marginStaging.Descendants()
            .Where(element => element.Attribute("data-section-index") is not null)
            .SelectMany(element => element.Descendants()), element =>
                (string?)element.Attribute("id") == "pagination-comment-margin-registry");
    }

    [Fact]
    public void PM103_InlineTableCommentIdentitiesStayInsideTheTablePresentation()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS003_TableWithCells());
        var cell = session.Project().AnchorIndex.Values.First(target =>
            target.Anchor.Kind == "p" && target.TextPreview == "R0C0");
        Assert.True(session.AddComment(
            cell.Anchor.Id, new CharSpan(0, 2), "Reviewer", "First.\n\nSecond.").Success);
        var commentIds = session.Project().AnchorIndex.Values
            .Where(target => target.Anchor.Scope == "cmt" && target.Anchor.Kind is "cmt" or "p")
            .Select(target => target.Anchor.Id)
            .ToHashSet(StringComparer.Ordinal);

        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(session, new HtmlConversionOptions
        {
            StampAnchors = false,
            FabricateCssClasses = false,
            PaginationMode = (int)PaginationMode.Paginated,
            CommentRenderMode = (int)CommentRenderMode.Inline,
        }));
        var tableCell = html.Descendants().First(element => element.Name.LocalName == "td");
        var presentedCommentIds = tableCell.DescendantsAndSelf()
            .Attributes("data-source-anchor-id")
            .Select(attribute => attribute.Value)
            .Where(commentIds.Contains)
            .ToHashSet(StringComparer.Ordinal);
        Assert.Equal(commentIds, presentedCommentIds);
    }
}
