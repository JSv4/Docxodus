#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class PageMapTests
{
    private const string Fingerprint = "chromium-140|docxodus-pagination-v1";

    private static PageMapPage Page(
        int pageNumber = 1,
        int pageInSection = 1,
        double width = 612,
        double height = 792,
        string pageName = "docxodus-section-0",
        int? sectionIndex = 0) => new()
        {
            PageNumber = pageNumber,
            PageInSection = pageInSection,
            Width = width,
            Height = height,
            SectionIndex = sectionIndex,
            PageName = pageName,
        };

    private static PageMapFragment Fragment(
        string anchorId,
        int fragmentIndex = 0,
        int pageNumber = 1,
        PageMapStory story = PageMapStory.Body,
        bool inTableCell = false,
        PageMapRect? geometry = null) => new()
        {
            FragmentId = $"p{pageNumber}-f{fragmentIndex}-{anchorId}",
            AnchorId = anchorId,
            FragmentIndex = fragmentIndex,
            PageNumber = pageNumber,
            Geometry = geometry ?? new PageMapRect(72, 90, 300, 18),
            Story = story,
            InTableCell = inTableCell,
        };

    private static PageMap AvailableMap(
        DocxSession session,
        IReadOnlyList<PageMapFragment> fragments,
        IReadOnlyList<PageMapPage>? pages = null,
        string fingerprint = Fingerprint) => new()
        {
            Mode = PageMapMode.Paginated,
            Availability = PageMapAvailability.Available,
            DocumentVersion = session.Version,
            RendererFingerprint = fingerprint,
            Pages = pages ?? new[] { Page() },
            Fragments = fragments,
        };

    [Fact]
    public void PM001_NoMapAndContinuousModeAreExplicitlyUnavailable()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var request = new PageCitationRequest(session.Version, Fingerprint);
        var anchor = session.Project().AnchorIndex.Values.First().Anchor.Id;

        Assert.Equal(PageCitationUnavailableReason.NoPageMap,
            session.GetPageMapStatus(request).UnavailableReason);
        Assert.Equal(PageCitationUnavailableReason.NoPageMap,
            session.GetPageCitation(anchor, request).UnavailableReason);

        var continuous = new PageMap
        {
            Mode = PageMapMode.Continuous,
            Availability = PageMapAvailability.Unavailable,
            DocumentVersion = session.Version,
            RendererFingerprint = Fingerprint,
        };
        Assert.True(session.RegisterPageMap(continuous).Success);
        Assert.Equal(PageCitationUnavailableReason.ContinuousMode,
            session.GetPageMapStatus(request).UnavailableReason);
        Assert.Equal(PageCitationUnavailableReason.ContinuousMode,
            session.GetPageCitation(anchor, request).UnavailableReason);
    }

    [Fact]
    public void PM002_ValidMapFeedsCitationsIntoSearchAndScopedProjection()
    {
        var annotated = AnnotationManager.AddAnnotation(
            new WmlDocument("PM002.docx", DocxSessionTests.BuildDS001_SimpleTwoParagraphs()),
            new DocumentAnnotation("page-map-annotation", "PAGE_MAP", "Page map", "#FFFF00"),
            AnnotationRange.FromSearch("First paragraph."));
        using var session = new DocxSession(annotated.DocumentByteArray);
        var anchors = session.Project().AnchorIndex.Values
            .Where(target => target.Anchor.Kind == "p")
            .Select(target => target.Anchor.Id)
            .ToArray();
        var map = AvailableMap(session, new[]
        {
            Fragment(anchors[0], fragmentIndex: 0, pageNumber: 1),
            Fragment(anchors[0], fragmentIndex: 1, pageNumber: 2),
            Fragment(anchors[1], fragmentIndex: 0, pageNumber: 2),
        }, new[] { Page(), Page(2, 2) });

        Assert.True(session.RegisterPageMap(map, Fingerprint).Success);
        var request = new PageCitationRequest(session.Version, Fingerprint);
        var citation = session.GetPageCitation(anchors[0], request);
        Assert.Equal(PageMapAvailability.Available, citation.Availability);
        Assert.Equal(new[] { 1, 2 }, citation.Fragments.Select(fragment => fragment.PageNumber));

        var match = Assert.Single(session.Grep("First", citationRequest: request));
        Assert.Equal(anchors[0], match.Citation?.AnchorId);
        Assert.Equal(PageMapAvailability.Available, match.Citation?.Availability);

        var projection = session.ProjectAnchor(
            anchors[0], ProjectionDepth.SelfOnly, citationRequest: request);
        Assert.NotNull(projection.PageCitations);
        Assert.Equal(PageMapAvailability.Available,
            Assert.Single(projection.PageCitations!).Value.Availability);

        var labeled = Assert.Single(session.FindByLabel("PAGE_MAP", request));
        Assert.All(labeled.Value, target =>
            Assert.Equal(PageMapAvailability.Available, target.Citation?.Availability));
    }

    [Fact]
    public void PM003_MutationAndRendererMismatchCannotConsumeARegisteredMap()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values.First(target => target.Anchor.Kind == "p").Anchor.Id;
        Assert.True(session.RegisterPageMap(AvailableMap(session, new[] { Fragment(anchor) })).Success);

        var wrongRenderer = new PageCitationRequest(session.Version, "firefox-other-layout");
        Assert.Equal(PageCitationUnavailableReason.RendererFingerprintMismatch,
            session.GetPageCitation(anchor, wrongRenderer).UnavailableReason);

        var versionBeforeEdit = session.Version;
        Assert.True(session.ReplaceText(anchor, "Changed paragraph.").Success);
        Assert.True(session.Version > versionBeforeEdit);
        var stale = session.GetPageMapStatus(new PageCitationRequest(versionBeforeEdit, Fingerprint));
        Assert.Equal(PageCitationUnavailableReason.StaleDocumentVersion, stale.UnavailableReason);
        Assert.Equal(PageMapAvailability.Unavailable, stale.Availability);
    }

    [Fact]
    public void PM004_RegistrationRejectsWrongVersionAndExpectedRenderer()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values.First(target => target.Anchor.Kind == "p").Anchor.Id;
        var map = AvailableMap(session, new[] { Fragment(anchor) });

        Assert.Equal(PageMapRegistrationError.RendererFingerprintMismatch,
            session.RegisterPageMap(map, "another-renderer").Error);
        Assert.Equal(PageMapRegistrationError.StaleDocumentVersion,
            session.RegisterPageMap(map with { DocumentVersion = session.Version + 1 }).Error);
    }

    [Fact]
    public void PM005_ValidatorEnforcesPageFragmentAndOwnershipInvariants()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS003_TableWithCells());
        var projection = session.Project();
        var cellParagraph = projection.AnchorIndex.Values.First(target =>
            target.Anchor.Kind == "p" && target.TextPreview == "R0C0");

        PageMapRegistrationResult Register(PageMapFragment fragment,
            IReadOnlyList<PageMapPage>? pages = null) =>
            session.RegisterPageMap(AvailableMap(session, new[] { fragment }, pages));

        Assert.Equal(PageMapRegistrationError.InvalidMap,
            Register(Fragment(cellParagraph.Anchor.Id, inTableCell: false)).Error);
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            Register(Fragment(cellParagraph.Anchor.Id, story: PageMapStory.Header, inTableCell: true)).Error);
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            Register(Fragment(cellParagraph.Anchor.Id, inTableCell: true,
                geometry: new PageMapRect(600, 780, 20, 20))).Error);
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            Register(Fragment(cellParagraph.Anchor.Id, fragmentIndex: 1, inTableCell: true)).Error);
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            Register(Fragment(cellParagraph.Anchor.Id, inTableCell: true),
                new[] { Page(2, 1) }).Error);
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            Register(Fragment(cellParagraph.Anchor.Id, inTableCell: true),
                new[] { Page(pageName: "") }).Error);
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            Register(Fragment(cellParagraph.Anchor.Id, inTableCell: true),
                new[] { Page(sectionIndex: -1) }).Error);
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            Register(Fragment(cellParagraph.Anchor.Id, inTableCell: true),
                new[] { Page(2), Page(1, 2) }).Error);
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            Register(Fragment(cellParagraph.Anchor.Id, inTableCell: true),
                new[] { Page(), Page(2, 3) }).Error);
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            session.RegisterPageMap(AvailableMap(session, new[]
            {
                Fragment(cellParagraph.Anchor.Id, fragmentIndex: 1, pageNumber: 1, inTableCell: true),
                Fragment(cellParagraph.Anchor.Id, fragmentIndex: 0, pageNumber: 2, inTableCell: true),
            }, new[] { Page(), Page(2, 2) })).Error);
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            session.RegisterPageMap(AvailableMap(session, new[]
            {
                Fragment(cellParagraph.Anchor.Id, fragmentIndex: 0, pageNumber: 2, inTableCell: true),
                Fragment(cellParagraph.Anchor.Id, fragmentIndex: 1, pageNumber: 1, inTableCell: true),
            }, new[] { Page(), Page(2, 2) })).Error);

        var secondCellParagraph = projection.AnchorIndex.Values.First(target =>
            target.Anchor.Kind == "p" && target.TextPreview == "R0C1");
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            session.RegisterPageMap(AvailableMap(session, new[]
            {
                Fragment(cellParagraph.Anchor.Id, pageNumber: 2, inTableCell: true),
                Fragment(secondCellParagraph.Anchor.Id, pageNumber: 1, inTableCell: true),
            }, new[] { Page(), Page(2, 2) })).Error);
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            session.RegisterPageMap(AvailableMap(session, new PageMapFragment[] { null! })).Error);
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            session.RegisterPageMap(AvailableMap(session, new[]
            {
                Fragment(cellParagraph.Anchor.Id, inTableCell: true) with { Geometry = null! },
            })).Error);
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            session.RegisterPageMap(AvailableMap(
                session, new[] { Fragment(cellParagraph.Anchor.Id, inTableCell: true) },
                new PageMapPage[] { null! })).Error);

        Assert.True(Register(Fragment(cellParagraph.Anchor.Id, inTableCell: true)).Success);
    }

    [Fact]
    public void PM006_UnknownContractDiscriminatorsAreNeverSilentlyCoerced()
    {
        const string typoStory = """
            {
              "schemaVersion":1,
              "mode":"paginated",
              "availability":"available",
              "documentVersion":0,
              "rendererFingerprint":"renderer",
              "pages":[],
              "fragments":[{
                "fragmentId":"f", "anchorId":"p:body:u", "fragmentIndex":0,
                "pageNumber":1, "geometry":{"x":0,"y":0,"width":1,"height":1},
                "story":"boddy", "inTableCell":false
              }]
            }
            """;
        Assert.Throws<FormatException>(() => DocxSessionJson.ParsePageMap(typoStory));

        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values.First(target => target.Anchor.Kind == "p").Anchor.Id;
        var invalid = AvailableMap(session, new[]
        {
            Fragment(anchor) with { Story = (PageMapStory)999 },
        });
        Assert.Equal(PageMapRegistrationError.InvalidMap, session.RegisterPageMap(invalid).Error);
        Assert.Equal(PageMapRegistrationError.InvalidMap,
            session.RegisterPageMap(invalid with { Mode = (PageMapMode)999 }).Error);
    }

    [Fact]
    public void PM007_PlaceholderSearchCanAttachACitation()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values
            .First(target => target.Anchor.Kind == "p").Anchor.Id;
        Assert.True(session.ReplaceText(anchor, "Complete this: [___]").Success);
        Assert.True(session.RegisterPageMap(AvailableMap(session, new[] { Fragment(anchor) })).Success);

        var match = Assert.Single(session.FindPlaceholders(
            citationRequest: new PageCitationRequest(session.Version, Fingerprint)));
        Assert.Equal(anchor, match.Match.Citation?.AnchorId);
        Assert.Equal(PageMapAvailability.Available, match.Match.Citation?.Availability);
    }
}
