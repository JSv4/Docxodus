// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>
/// Native identity for Word structured-document tags. A valid, unique <c>w:sdtPr/w:id</c>
/// is the durable identity Word itself persists; the projector's <c>pt:Unid</c> is only a
/// cache. This helper makes the cache a deterministic function of the native id so an
/// <c>sdt:</c> anchor survives the default <c>Save(false)</c> / reopen path.
/// </summary>
internal static class ContentControlIdentity
{
    internal sealed record Entry(
        XElement Element,
        string Unid,
        string? NativeId,
        bool HasValidNativeId,
        bool IsDuplicateNativeId,
        int DocumentOrdinal,
        int DuplicateOrdinal)
    {
        internal bool HasMutableIdentity => HasValidNativeId && !IsDuplicateNativeId;
    }

    /// <summary>
    /// Assign native-derived Unids to every SDT below one story root and return the same
    /// outer-before-inner registry order used by the public content-control listing.
    /// Malformed controls still receive deterministic inspection identities, but callers
    /// must gate mutation on <see cref="Entry.HasMutableIdentity"/>.
    /// </summary>
    internal static IReadOnlyList<Entry> AssignStableUnids(XElement storyRoot) =>
        AssignStableUnids(storyRoot, out _);

    internal static IReadOnlyList<Entry> AssignStableUnids(XElement storyRoot, out bool changed)
    {
        ArgumentNullException.ThrowIfNull(storyRoot);
        changed = false;
        var controls = storyRoot.DescendantsAndSelf(W.sdt).ToList();
        if (controls.Count == 0) return Array.Empty<Entry>();

        var parsed = controls.Select((element, ordinal) =>
        {
            var raw = (string?)element.Element(W.sdtPr)?.Element(W.id)?.Attribute(W.val);
            var valid = TryCanonicalizeNativeId(raw, out var canonical);
            return (element, ordinal, raw, valid, canonical);
        }).ToList();

        var counts = parsed.Where(value => value.valid)
            .GroupBy(value => value.canonical!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        var duplicateOrdinals = new Dictionary<string, int>(StringComparer.Ordinal);
        var result = new List<Entry>(parsed.Count);

        foreach (var value in parsed)
        {
            int duplicateOrdinal = 0;
            bool duplicate = value.valid && counts[value.canonical!] > 1;
            if (duplicate)
            {
                duplicateOrdinals.TryGetValue(value.canonical!, out duplicateOrdinal);
                duplicateOrdinals[value.canonical!] = duplicateOrdinal + 1;
            }

            // Unique, valid native ids are location- and content-independent. The fallback
            // discriminator is intentionally only for non-writable malformed documents.
            var seed = value.valid
                ? duplicate
                    ? $"duplicate\0{value.canonical}\0{duplicateOrdinal}"
                    : $"native\0{value.canonical}"
                : $"malformed\0{value.ordinal}\0{value.raw ?? "<missing>"}";
            var unid = HashToUnid(seed);
            if (!string.Equals((string?)value.element.Attribute(PtOpenXml.Unid), unid,
                StringComparison.Ordinal))
            {
                value.element.SetAttributeValue(PtOpenXml.Unid, unid);
                changed = true;
            }
            result.Add(new Entry(value.element, unid,
                value.valid ? value.canonical : value.raw,
                value.valid, duplicate, value.ordinal, duplicateOrdinal));
        }
        return result;
    }

    internal static bool TryCanonicalizeNativeId(string? raw, out string? canonical)
    {
        canonical = null;
        if (string.IsNullOrWhiteSpace(raw)
            || !int.TryParse(raw, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out var id))
            return false;
        canonical = id.ToString(CultureInfo.InvariantCulture);
        return true;
    }

    internal static string HashToUnid(string seed)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes("docxodus-content-control\0" + seed));
        return Convert.ToHexString(bytes.AsSpan(0, 16)).ToLowerInvariant();
    }
}
