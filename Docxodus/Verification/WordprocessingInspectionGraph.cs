// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Xml.Linq;

namespace Docxodus.Verification;

/// <summary>Relationship-reachable bounded package view rooted at the Word main document.</summary>
internal sealed record WordprocessingInspectionGraph
{
    private const string TransitionalWord =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string StrictWord = "http://purl.oclc.org/ooxml/wordprocessingml/main";

    required internal IReadOnlyList<PackageManifestInspectionEntry> ReachableEntries { get; init; }
    required internal IReadOnlyList<PackageManifestInspectionEntry> WordParts { get; init; }
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
        if (mainUri is null || !uniqueEntries.ContainsKey(mainUri))
            return Empty(complete: true);

        var byOwner = inspection.Manifest.Relationships
            .GroupBy(relationship => relationship.OwnerUri, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group
                .OrderBy(item => item.Id, StringComparer.Ordinal)
                .ThenBy(item => item.Type, StringComparer.Ordinal)
                .ThenBy(item => item.Target, StringComparer.Ordinal)
                .ToArray(), StringComparer.OrdinalIgnoreCase);
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
                if (relationship.TargetMode == "Internal"
                    && relationship.IsTargetPresent == true
                    && relationship.ResolvedTargetUri is { } target
                    && uniqueEntries.ContainsKey(target))
                    pending.Enqueue(target);
            }
        }

        var wordParts = entries.Where(entry => entry.Xml?.Root is { } root
            && root.Name.NamespaceName is TransitionalWord or StrictWord).ToArray();
        return new WordprocessingInspectionGraph
        {
            ReachableEntries = entries,
            WordParts = wordParts,
            ReachableRelationships = relationships,
            Complete = !budget.Exhausted,
        };
    }

    private static WordprocessingInspectionGraph Empty(bool complete) => new()
    {
        ReachableEntries = Array.Empty<PackageManifestInspectionEntry>(),
        WordParts = Array.Empty<PackageManifestInspectionEntry>(),
        ReachableRelationships = Array.Empty<PackageRelationship>(),
        Complete = complete,
    };
}
