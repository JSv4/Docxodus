// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Xml.Linq;

namespace Docxodus.Verification;

/// <summary>
/// Bounded package reachability plus a role-aware Wordprocessing topology rooted at the main part.
/// Generic reachability supports relationship/media diagnostics; semantic definition and story
/// parts are admitted only through the OPC relationship type and owner allowed for that role.
/// </summary>
internal sealed record WordprocessingInspectionGraph
{
    private const string TransitionalWord =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string StrictWord = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    private const string CustomXmlPropertiesContentType =
        "application/vnd.openxmlformats-officedocument.customXmlProperties+xml";
    private const string TransitionalCustomXml =
        "http://schemas.openxmlformats.org/officeDocument/2006/customXml";
    private const string StrictCustomXml =
        "http://purl.oclc.org/ooxml/officeDocument/customXml";

    required internal IReadOnlyList<PackageManifestInspectionEntry> ReachableEntries { get; init; }
    required internal IReadOnlyList<PackageManifestInspectionEntry> StoryParts { get; init; }
    required internal IReadOnlyList<PackageManifestInspectionEntry> CommentParts { get; init; }
    required internal IReadOnlyList<PackageManifestInspectionEntry> FootnoteParts { get; init; }
    required internal IReadOnlyList<PackageManifestInspectionEntry> EndnoteParts { get; init; }
    required internal IReadOnlyList<PackageManifestInspectionEntry> NumberingParts { get; init; }
    required internal IReadOnlyList<PackageManifestInspectionEntry> StyleParts { get; init; }
    required internal IReadOnlyList<PackageManifestInspectionEntry> SettingsParts { get; init; }
    required internal IReadOnlyList<PackageManifestInspectionEntry> CustomXmlPropertyParts { get; init; }
    required internal IReadOnlyList<SemanticRelationshipEdge> SemanticRelationshipEdges { get; init; }
    required internal IReadOnlyList<PackageRelationship> ReachableRelationships { get; init; }
    required internal bool Complete { get; init; }

    internal static WordprocessingInspectionGraph Build(
        PackageManifestInspection inspection,
        DeliverableInspectionBudget budget)
    {
        var uniqueEntries = inspection.Entries
            .GroupBy(entry => entry.Uri, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() == 1)
            .ToDictionary(group => group.Key, group => group.Single(), StringComparer.OrdinalIgnoreCase);
        var mainUri = inspection.Manifest.Facts.MainDocumentUri;
        if (mainUri is null || !uniqueEntries.TryGetValue(mainUri, out var mainEntry)
            || !MatchesRole(mainEntry, SemanticRole.MainDocument))
            return Empty(complete: true);

        var byOwner = inspection.Manifest.Relationships
            .GroupBy(relationship => relationship.OwnerUri, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group
                .OrderBy(item => item.Id, StringComparer.Ordinal)
                .ThenBy(item => item.Type, StringComparer.Ordinal)
                .ThenBy(item => item.Target, StringComparer.Ordinal)
                .ToArray(), StringComparer.OrdinalIgnoreCase);

        // Keep generic reachability for media and relationship diagnostics. It never confers a
        // semantic role: an arbitrary edge may make bytes inspectable without making them a valid
        // Word story or definition part.
        var pending = new Queue<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var entries = new List<PackageManifestInspectionEntry>();
        var relationships = new List<PackageRelationship>();
        pending.Enqueue(mainUri);
        while (pending.Count > 0 && !budget.Exhausted)
        {
            var uri = pending.Dequeue();
            if (!seen.Add(uri) || !uniqueEntries.TryGetValue(uri, out var entry)) continue;
            entries.Add(entry);
            if (!byOwner.TryGetValue(uri, out var owned)) continue;
            foreach (var relationship in owned)
            {
                if (!budget.Relationship() || !budget.Step()) break;
                relationships.Add(relationship);
                if (InternalTarget(relationship, uniqueEntries) is { } target)
                    pending.Enqueue(target);
            }
        }

        // The generic pass charged each retained raw relationship exactly once. Semantic role
        // classification operates only over that bounded edge set and charges steps, not the raw
        // relationship budget a second time.
        var retainedByOwner = relationships
            .GroupBy(relationship => relationship.OwnerUri, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToArray(),
                StringComparer.OrdinalIgnoreCase);
        var roleQueue = new Queue<(string Uri, SemanticRole Role)>();
        var semanticRoles = new Dictionary<string, HashSet<SemanticRole>>(
            StringComparer.OrdinalIgnoreCase);
        var semanticRelationshipEdges = new List<SemanticRelationshipEdge>();
        // glossaryDocument remains generic-only: its definitions cannot satisfy active stories.
        // Inspecting glossary content later would require a separate semantic branch.
        roleQueue.Enqueue((mainUri, SemanticRole.MainDocument));
        while (roleQueue.Count > 0 && !budget.Exhausted)
        {
            var (uri, role) = roleQueue.Dequeue();
            if (!uniqueEntries.TryGetValue(uri, out var entry) || !MatchesRole(entry, role)) continue;
            if (!semanticRoles.TryGetValue(uri, out var roles))
                semanticRoles.Add(uri, roles = new HashSet<SemanticRole>());
            if (!roles.Add(role) || !retainedByOwner.TryGetValue(uri, out var owned)) continue;
            foreach (var relationship in owned)
            {
                if (!budget.Step()) break;
                var targetRole = NextRole(role, relationship.Type);
                if (targetRole is null) continue;
                // Retain the exact owner/type fact before target admission so relationship
                // cardinality remains visible for missing, external, or mistyped target parts.
                semanticRelationshipEdges.Add(new SemanticRelationshipEdge(
                    role, targetRole.Value, relationship));
                if (InternalTarget(relationship, uniqueEntries) is not { } target) continue;
                if (!uniqueEntries.TryGetValue(target, out var targetEntry)
                    || !MatchesRole(targetEntry, targetRole.Value))
                    continue;
                roleQueue.Enqueue((target, targetRole.Value));
            }
        }

        var storyParts = SemanticEntries(uniqueEntries, semanticRoles,
            SemanticRole.MainDocument, SemanticRole.Header, SemanticRole.Footer,
            SemanticRole.Footnotes, SemanticRole.Endnotes, SemanticRole.Comments);
        var numberingParts = SemanticEntries(uniqueEntries, semanticRoles, SemanticRole.Numbering);
        var commentParts = SemanticEntries(uniqueEntries, semanticRoles, SemanticRole.Comments);
        var footnoteParts = SemanticEntries(uniqueEntries, semanticRoles, SemanticRole.Footnotes);
        var endnoteParts = SemanticEntries(uniqueEntries, semanticRoles, SemanticRole.Endnotes);
        var styleParts = SemanticEntries(uniqueEntries, semanticRoles, SemanticRole.Styles);
        var settingsParts = SemanticEntries(uniqueEntries, semanticRoles, SemanticRole.Settings);
        var customXmlProperties = SemanticEntries(
            uniqueEntries, semanticRoles, SemanticRole.CustomXmlProperties);
        return new WordprocessingInspectionGraph
        {
            ReachableEntries = entries,
            StoryParts = storyParts,
            CommentParts = commentParts,
            FootnoteParts = footnoteParts,
            EndnoteParts = endnoteParts,
            NumberingParts = numberingParts,
            StyleParts = styleParts,
            SettingsParts = settingsParts,
            CustomXmlPropertyParts = customXmlProperties,
            SemanticRelationshipEdges = semanticRelationshipEdges,
            ReachableRelationships = relationships,
            Complete = !budget.Exhausted,
        };
    }

    private static string? InternalTarget(
        PackageRelationship relationship,
        IReadOnlyDictionary<string, PackageManifestInspectionEntry> uniqueEntries) =>
        relationship.TargetMode == "Internal"
        && relationship.IsTargetPresent == true
        && relationship.ResolvedTargetUri is { } target
        && uniqueEntries.ContainsKey(target)
            ? target
            : null;

    private static SemanticRole? NextRole(SemanticRole ownerRole, string relationshipType)
    {
        return ownerRole switch
        {
            SemanticRole.MainDocument when IsRelationshipType(relationshipType, "header") =>
                SemanticRole.Header,
            SemanticRole.MainDocument when IsRelationshipType(relationshipType, "footer") =>
                SemanticRole.Footer,
            SemanticRole.MainDocument when IsRelationshipType(relationshipType, "footnotes") =>
                SemanticRole.Footnotes,
            SemanticRole.MainDocument when IsRelationshipType(relationshipType, "endnotes") =>
                SemanticRole.Endnotes,
            SemanticRole.MainDocument when IsRelationshipType(relationshipType, "comments") =>
                SemanticRole.Comments,
            SemanticRole.MainDocument when IsRelationshipType(relationshipType, "numbering") =>
                SemanticRole.Numbering,
            SemanticRole.MainDocument when IsRelationshipType(relationshipType, "styles") =>
                SemanticRole.Styles,
            SemanticRole.MainDocument when IsRelationshipType(relationshipType, "settings") =>
                SemanticRole.Settings,
            SemanticRole.MainDocument when IsRelationshipType(relationshipType, "customXml") =>
                SemanticRole.CustomXml,
            SemanticRole.CustomXml when IsRelationshipType(relationshipType, "customXmlProps") =>
                SemanticRole.CustomXmlProperties,
            _ => null,
        };
    }

    private static bool IsRelationshipType(string value, string localType) =>
        OpenXmlRelationshipVocabulary.IsOfficeType(value, localType);

    private static bool MatchesRole(PackageManifestInspectionEntry entry, SemanticRole role)
    {
        if (role == SemanticRole.CustomXml)
            return entry.Xml?.Root is not null;
        if (role == SemanticRole.CustomXmlProperties)
            return string.Equals(entry.ManifestEntry.ContentType, CustomXmlPropertiesContentType,
                StringComparison.OrdinalIgnoreCase)
                && entry.Xml?.Root is { Name.LocalName: "datastoreItem" } propertiesRoot
                && propertiesRoot.Name.NamespaceName is TransitionalCustomXml or StrictCustomXml;
        if (entry.Xml?.Root is not { } root || !IsWord(root)) return false;
        var expectedContentType = role switch
        {
            SemanticRole.Header =>
                "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
            SemanticRole.Footer =>
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
            SemanticRole.Footnotes =>
                "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
            SemanticRole.Endnotes =>
                "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
            SemanticRole.Comments =>
                "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
            SemanticRole.Numbering =>
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
            SemanticRole.Styles =>
                "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
            SemanticRole.Settings =>
                "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
            _ => null,
        };
        if (expectedContentType is not null
            && !string.Equals(entry.ManifestEntry.ContentType, expectedContentType,
                StringComparison.OrdinalIgnoreCase))
            return false;
        return role switch
        {
            SemanticRole.MainDocument => root.Name.LocalName == "document",
            SemanticRole.Header => root.Name.LocalName == "hdr",
            SemanticRole.Footer => root.Name.LocalName == "ftr",
            SemanticRole.Footnotes => root.Name.LocalName == "footnotes",
            SemanticRole.Endnotes => root.Name.LocalName == "endnotes",
            SemanticRole.Comments => root.Name.LocalName == "comments",
            SemanticRole.Numbering => root.Name.LocalName == "numbering",
            SemanticRole.Styles => root.Name.LocalName == "styles",
            SemanticRole.Settings => root.Name.LocalName == "settings",
            _ => false,
        };
    }

    private static PackageManifestInspectionEntry[] SemanticEntries(
        IReadOnlyDictionary<string, PackageManifestInspectionEntry> uniqueEntries,
        IReadOnlyDictionary<string, HashSet<SemanticRole>> rolesByUri,
        params SemanticRole[] roles)
    {
        var allowed = roles.ToHashSet();
        return rolesByUri
            .Where(pair => pair.Value.Overlaps(allowed))
            .OrderBy(pair => pair.Value.Where(allowed.Contains).Select(role => (int)role).Min())
            .ThenBy(pair => pair.Key, StringComparer.Ordinal)
            .Select(pair => uniqueEntries[pair.Key])
            .ToArray();
    }

    private static bool IsWord(XElement element) =>
        element.Name.NamespaceName is TransitionalWord or StrictWord;

    private static WordprocessingInspectionGraph Empty(bool complete) => new()
    {
        ReachableEntries = Array.Empty<PackageManifestInspectionEntry>(),
        StoryParts = Array.Empty<PackageManifestInspectionEntry>(),
        CommentParts = Array.Empty<PackageManifestInspectionEntry>(),
        FootnoteParts = Array.Empty<PackageManifestInspectionEntry>(),
        EndnoteParts = Array.Empty<PackageManifestInspectionEntry>(),
        NumberingParts = Array.Empty<PackageManifestInspectionEntry>(),
        StyleParts = Array.Empty<PackageManifestInspectionEntry>(),
        SettingsParts = Array.Empty<PackageManifestInspectionEntry>(),
        CustomXmlPropertyParts = Array.Empty<PackageManifestInspectionEntry>(),
        SemanticRelationshipEdges = Array.Empty<SemanticRelationshipEdge>(),
        ReachableRelationships = Array.Empty<PackageRelationship>(),
        Complete = complete,
    };

    internal sealed record SemanticRelationshipEdge(
        SemanticRole OwnerRole,
        SemanticRole TargetRole,
        PackageRelationship Relationship);

    internal enum SemanticRole
    {
        MainDocument,
        Header,
        Footer,
        Footnotes,
        Endnotes,
        Comments,
        Numbering,
        Styles,
        Settings,
        CustomXml,
        CustomXmlProperties,
    }
}
