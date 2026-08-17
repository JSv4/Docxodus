// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>
/// A live, part-aware registry of every tracked revision the session can see. The
/// registry is intentionally rebuilt after each resolution: rejecting a property
/// change can expose older revision markup from its archived property shell, while
/// resolving a structural parent can detach nested revisions.
/// </summary>
internal sealed class RevisionRegistry
{
    internal sealed record Part(string PartUri, string Scope, XElement Root);

    private readonly IReadOnlyList<Part> _parts;

    private RevisionRegistry(IReadOnlyList<Part> parts, List<RevisionOps.RevisionGroup> entries)
    {
        _parts = parts;
        Entries = entries;
    }

    internal IReadOnlyList<RevisionOps.RevisionGroup> Entries { get; }

    internal static RevisionRegistry Build(IReadOnlyList<Part> parts) =>
        new(parts, RevisionOps.Enumerate(parts
            .Select(p => (p.PartUri, p.Scope, p.Root)).ToList()));

    internal RevisionOps.RevisionGroup? Find(string id)
    {
        var exact = Entries.FirstOrDefault(entry => entry.Id == id);
        if (exact is not null) return exact;

        // Backward-compatible input only: legacy revNNN ids are accepted when they
        // identify exactly one current group. Listings always return the stable rev2 id.
        var legacy = Entries.Where(entry => RevisionOps.LegacyId(entry) == id).ToList();
        return legacy.Count == 1 ? legacy[0] : null;
    }

    internal static RevisionDiagnostic? ResolutionDiagnostic(RevisionOps.RevisionGroup group) =>
        group.ResolutionStatus == RevisionResolutionStatus.Supported
            ? null
            : group.Diagnostic ?? new RevisionDiagnostic(
                "unresolved_revision",
                "The revision cannot be resolved safely.");

    internal List<XElement> Resolve(
        RevisionOps.RevisionGroup group,
        bool accept,
        bool preserveUnrelatedMarkup = false,
        IReadOnlyCollection<string>? protectedEmptyContainerKeys = null) =>
        RevisionOps.Apply(
            group, accept, preserveUnrelatedMarkup, protectedEmptyContainerKeys);

    /// <summary>
    /// Resolve every currently live revision through the same selective resolver used
    /// by individual operations. Rebuild after every group so newly exposed archived
    /// revisions are handled and detached nested groups disappear naturally.
    /// </summary>
    internal List<XElement> ResolveAll(
        bool accept,
        bool preserveUnrelatedMarkup = false,
        IReadOnlyCollection<string>? protectedEmptyContainerKeys = null)
    {
        var removed = new List<XElement>();
        var registry = this;
        var attemptedElements = new HashSet<XElement>();
        for (int guard = 0; guard < 100_000; guard++)
        {
            if (registry.Entries.Count == 0) return removed;

            var blocked = registry.Entries.FirstOrDefault(entry =>
                entry.ResolutionStatus != RevisionResolutionStatus.Supported);
            if (blocked is not null)
                throw new RevisionResolutionException(blocked);

            var next = registry.Entries[0];
            var progressElement = next.Units.FirstOrDefault()?.Element
                ?? next.RangeMarkers.FirstOrDefault();
            if (progressElement is null || !attemptedElements.Add(progressElement))
                throw new InvalidOperationException(
                    "Bulk revision resolution made no progress; the document was left unchanged.");

            removed.AddRange(registry.Resolve(
                next, accept, preserveUnrelatedMarkup, protectedEmptyContainerKeys));
            registry = Build(_parts);
        }

        throw new InvalidOperationException(
            "Bulk revision resolution exceeded its safety limit; the document was left unchanged.");
    }
}

internal sealed class RevisionResolutionException : InvalidOperationException
{
    internal RevisionResolutionException(RevisionOps.RevisionGroup group)
        : base(group.Diagnostic?.Message ?? "revision cannot be resolved safely")
    {
        Group = group;
    }

    internal RevisionOps.RevisionGroup Group { get; }
}
