// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Relationship ids for parts a redline creates, chosen so that generating them twice gives the same
/// answer.
/// <para><c>AddNewPart&lt;T&gt;()</c> without an explicit id asks the Open XML SDK to mint one, and the
/// SDK's generator is <c>"R"</c> + sixteen random hex characters. Every part the comparison engines add to
/// their output therefore carried a different id on each run, so two <c>DocxDiff.Compare</c> calls over the
/// same inputs produced different bytes — the opposite of what
/// <see cref="Ir.Diff.IrDiffSettings.Deterministic"/> promises. Naming the id explicitly is the whole fix;
/// this type exists so the rule has one owner instead of a literal repeated at every call site.</para>
/// </summary>
internal static class DeterministicPartIds
{
    /// <summary>
    /// Add a part under a reproducible relationship id: <paramref name="stem"/>, or the lowest free
    /// <paramref name="stem"/> + ordinal when the stem is taken. Ordinals cover the part kinds a document
    /// can hold several of (headers, footers); for the single-instance kinds they are a safety net against
    /// a source document that happens to use the stem, not the normal path.
    /// </summary>
    internal static T AddDeterministicPart<T>(this OpenXmlPartContainer owner, string stem)
        where T : OpenXmlPart, IFixedContentTypePart =>
        owner.AddNewPart<T>(FreeRelationshipId(owner, stem));

    /// <summary>
    /// Every relationship id in use on <paramref name="owner"/>, across all four kinds. An id free among
    /// part relationships can still be taken by a hyperlink, an external link or a data-part reference,
    /// and reusing it makes the packaging layer throw.
    /// </summary>
    internal static HashSet<string> UsedRelationshipIds(OpenXmlPartContainer owner)
    {
        var used = new HashSet<string>(StringComparer.Ordinal);
        foreach (var pair in owner.Parts)
            used.Add(pair.RelationshipId);
        foreach (var rel in owner.HyperlinkRelationships)
            used.Add(rel.Id);
        foreach (var rel in owner.ExternalRelationships)
            used.Add(rel.Id);
        foreach (var rel in owner.DataPartReferenceRelationships)
            used.Add(rel.Id);
        return used;
    }

    /// <summary>The lowest free <paramref name="stem"/> [+ ordinal] on <paramref name="owner"/>.</summary>
    private static string FreeRelationshipId(OpenXmlPartContainer owner, string stem)
    {
        var used = UsedRelationshipIds(owner);
        if (!used.Contains(stem))
            return stem;
        var n = 2;
        while (used.Contains(stem + n.ToString(CultureInfo.InvariantCulture)))
            n++;
        return stem + n.ToString(CultureInfo.InvariantCulture);
    }
}
