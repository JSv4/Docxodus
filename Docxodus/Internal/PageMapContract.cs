// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.Internal;

/// <summary>
/// Portable validation and citation projection for the public PageMap contract. Validation that
/// requires a live Open XML anchor remains in <see cref="DocxSession.RegisterPageMap"/>.
/// </summary>
internal static class PageMapContract
{
    public static PageMapRegistrationResult ValidatePortable(
        PageMap pageMap,
        long? expectedDocumentVersion = null,
        string? expectedRendererFingerprint = null)
    {
        ArgumentNullException.ThrowIfNull(pageMap);

        static PageMapRegistrationResult Fail(PageMapRegistrationError error, string message) =>
            new() { Success = false, Error = error, Message = message };

        if (pageMap.SchemaVersion != PageMap.CurrentSchemaVersion)
            return Fail(PageMapRegistrationError.UnsupportedSchemaVersion,
                $"unsupported PageMap schemaVersion {pageMap.SchemaVersion}; expected {PageMap.CurrentSchemaVersion}");
        if (expectedDocumentVersion is { } documentVersion
            && pageMap.DocumentVersion != documentVersion)
            return Fail(PageMapRegistrationError.StaleDocumentVersion,
                $"PageMap documentVersion {pageMap.DocumentVersion} does not match document version {documentVersion}");
        if (!Enum.IsDefined(pageMap.Mode) || !Enum.IsDefined(pageMap.Availability))
            return Fail(PageMapRegistrationError.InvalidMap, "PageMap mode or availability discriminator is invalid");
        if (string.IsNullOrWhiteSpace(pageMap.RendererFingerprint))
            return Fail(PageMapRegistrationError.InvalidMap, "rendererFingerprint must be non-empty");
        if (pageMap.Pages is null || pageMap.Fragments is null)
            return Fail(PageMapRegistrationError.InvalidMap, "PageMap pages and fragments arrays are required");
        if (expectedRendererFingerprint is not null
            && !string.Equals(pageMap.RendererFingerprint, expectedRendererFingerprint,
                StringComparison.Ordinal))
            return Fail(PageMapRegistrationError.RendererFingerprintMismatch,
                "PageMap rendererFingerprint does not match the expected renderer");

        if (pageMap.Mode == PageMapMode.Continuous)
        {
            if (pageMap.Availability != PageMapAvailability.Unavailable
                || pageMap.Pages.Count != 0 || pageMap.Fragments.Count != 0)
                return Fail(PageMapRegistrationError.InvalidMap,
                    "continuous PageMaps must be unavailable and contain no pages or fragments");
        }
        else if (pageMap.Availability != PageMapAvailability.Available)
            return Fail(PageMapRegistrationError.InvalidMap,
                "paginated PageMaps must be explicitly available");
        else if (pageMap.Pages.Count == 0 || pageMap.Fragments.Count == 0)
            return Fail(PageMapRegistrationError.InvalidMap,
                "an available paginated PageMap must contain at least one page and fragment");

        var pagesByNumber = new Dictionary<int, PageMapPage>();
        var seenSectionIndices = new HashSet<int>();
        PageMapPage? previousPage = null;
        for (var pageIndex = 0; pageIndex < pageMap.Pages.Count; pageIndex++)
        {
            var page = pageMap.Pages[pageIndex];
            if (page is null)
                return Fail(PageMapRegistrationError.InvalidMap,
                    "PageMap pages cannot contain null entries");
            if (page.PageNumber < 1 || page.PageInSection < 1
                || !double.IsFinite(page.Width) || page.Width <= 0
                || !double.IsFinite(page.Height) || page.Height <= 0
                || string.IsNullOrWhiteSpace(page.PageName)
                || page.SectionIndex is < 0)
                return Fail(PageMapRegistrationError.InvalidMap,
                    "pages require non-negative sectionIndex, positive numbering, a pageName, and finite positive geometry");
            if (page.PageNumber != pageIndex + 1)
                return Fail(PageMapRegistrationError.InvalidMap,
                    "pages must appear in contiguous document order starting at 1");

            if (pageIndex == 0)
            {
                if (page.PageInSection != 1)
                    return Fail(PageMapRegistrationError.InvalidMap,
                        "the first page must start at pageInSection 1");
                if (page.SectionIndex is int firstSection) seenSectionIndices.Add(firstSection);
            }
            else if (page.PageInSection == 1)
            {
                if (page.SectionIndex is int newSection && !seenSectionIndices.Add(newSection))
                    return Fail(PageMapRegistrationError.InvalidMap,
                        $"sectionIndex {newSection} appears in multiple discontiguous page runs");
                if (page.SectionIndex == previousPage!.SectionIndex && page.SectionIndex is not null)
                    return Fail(PageMapRegistrationError.InvalidMap,
                        $"pageInSection resets within sectionIndex {page.SectionIndex}");
            }
            else if (page.PageInSection != previousPage!.PageInSection + 1
                || page.SectionIndex != previousPage.SectionIndex)
                return Fail(PageMapRegistrationError.InvalidMap,
                    "pageInSection must be contiguous and reset to 1 when the section changes");

            pagesByNumber[page.PageNumber] = page;
            previousPage = page;
        }

        var fragmentIds = new HashSet<string>(StringComparer.Ordinal);
        var fragmentSequence = new Dictionary<string, (int NextIndex, int LastPage)>(StringComparer.Ordinal);
        var lastFragmentPage = 0;
        foreach (var fragment in pageMap.Fragments)
        {
            if (fragment is null || fragment.Geometry is null)
                return Fail(PageMapRegistrationError.InvalidMap,
                    "PageMap fragments and fragment geometry cannot be null");
            if (string.IsNullOrWhiteSpace(fragment.FragmentId)
                || !fragmentIds.Add(fragment.FragmentId)
                || string.IsNullOrWhiteSpace(fragment.AnchorId)
                || !Enum.IsDefined(fragment.Story)
                || fragment.FragmentIndex < 0
                || !pagesByNumber.ContainsKey(fragment.PageNumber)
                || !ValidRect(fragment.Geometry))
                return Fail(PageMapRegistrationError.InvalidMap,
                    "fragments require unique ids, canonical anchors, mapped pages, and finite non-negative geometry");
            if (fragment.PageNumber < lastFragmentPage)
                return Fail(PageMapRegistrationError.InvalidMap,
                    "PageMap fragments must appear in nondecreasing page order");
            lastFragmentPage = fragment.PageNumber;

            var page = pagesByNumber[fragment.PageNumber];
            const double geometryTolerance = 0.25;
            if (fragment.Geometry.X + fragment.Geometry.Width > page.Width + geometryTolerance
                || fragment.Geometry.Y + fragment.Geometry.Height > page.Height + geometryTolerance)
                return Fail(PageMapRegistrationError.InvalidMap,
                    $"PageMap fragment geometry exceeds page {fragment.PageNumber}");

            if (!fragmentSequence.TryGetValue(fragment.AnchorId, out var sequence))
                sequence = (NextIndex: 0, LastPage: fragment.PageNumber);
            if (fragment.FragmentIndex != sequence.NextIndex)
                return Fail(PageMapRegistrationError.InvalidMap,
                    $"PageMap fragmentIndex values must appear contiguously from 0: {fragment.AnchorId}");
            if (fragment.PageNumber < sequence.LastPage)
                return Fail(PageMapRegistrationError.InvalidMap,
                    $"PageMap fragment pages run backward for anchor: {fragment.AnchorId}");
            fragmentSequence[fragment.AnchorId] = (sequence.NextIndex + 1, fragment.PageNumber);
        }

        return new PageMapRegistrationResult { Success = true };
    }

    public static PageCitation ProjectCitation(PageMap pageMap, string anchorId, PageCitationRequest request)
    {
        ArgumentNullException.ThrowIfNull(pageMap);
        ArgumentNullException.ThrowIfNull(anchorId);
        ArgumentNullException.ThrowIfNull(request);
        var fragments = pageMap.Fragments
            .Where(fragment => string.Equals(fragment.AnchorId, anchorId, StringComparison.Ordinal))
            .OrderBy(fragment => fragment.PageNumber)
            .ThenBy(fragment => fragment.FragmentIndex)
            .ToArray();
        if (fragments.Length == 0)
            return new PageCitation
            {
                AnchorId = anchorId,
                Availability = PageMapAvailability.Unavailable,
                UnavailableReason = PageCitationUnavailableReason.AnchorNotMapped,
                DocumentVersion = request.DocumentVersion,
                RendererFingerprint = request.RendererFingerprint,
            };
        var citedPageNumbers = fragments.Select(fragment => fragment.PageNumber).ToHashSet();
        return new PageCitation
        {
            AnchorId = anchorId,
            Availability = PageMapAvailability.Available,
            DocumentVersion = request.DocumentVersion,
            RendererFingerprint = request.RendererFingerprint,
            Pages = pageMap.Pages.Where(page => citedPageNumbers.Contains(page.PageNumber))
                .OrderBy(page => page.PageNumber).ToArray(),
            Fragments = fragments,
        };
    }

    public static bool StoryMatchesScope(PageMapStory story, string scope) => story switch
    {
        PageMapStory.Header => scope.StartsWith("hdr", StringComparison.Ordinal),
        PageMapStory.Footer => scope.StartsWith("ftr", StringComparison.Ordinal),
        PageMapStory.Footnote => scope == "fn",
        PageMapStory.Endnote => scope == "en",
        PageMapStory.Comment => scope == "cmt",
        _ => scope == "body",
    };

    private static bool ValidRect(PageMapRect rect) =>
        rect is not null
        && double.IsFinite(rect.X) && rect.X >= 0
        && double.IsFinite(rect.Y) && rect.Y >= 0
        && double.IsFinite(rect.Width) && rect.Width > 0
        && double.IsFinite(rect.Height) && rect.Height > 0;
}
