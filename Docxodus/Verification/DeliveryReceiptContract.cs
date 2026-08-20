// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Globalization;

namespace Docxodus.Verification;

/// <summary>
/// Constants and pure functions of the receipt wire contract that the producer
/// (<see cref="DeliveryChangeReceiptBuilder"/>) and the adversarial verifier
/// (<see cref="DeliveryChangeReceiptVerifier"/>) must agree on. The two sides
/// deliberately share no serialization or parsing code — a verifier that trusts
/// producer code cannot catch producer bugs — so everything they must not drift
/// on lives here instead.
/// </summary>
internal static class DeliveryReceiptContract
{
    /// <summary>
    /// Owner schema for a generic evidence kind, or null when the kind has no
    /// generic owner schema (typed semantic evidence uses the #457 binding).
    /// </summary>
    public static string? EvidenceSchemaFor(DeliveryEvidenceKind kind) => kind switch
    {
        DeliveryEvidenceKind.ValidationResult => DeliverableVerificationResult.SchemaId,
        DeliveryEvidenceKind.RedlineReversibility => RedlineReversibilityProof.SchemaId,
        _ => null,
    };

    /// <summary>
    /// Property count of one serialized EditResult anchor. The producer is
    /// <c>DeliveryChangeReceiptBuilder.WriteAnchor</c>; the verifier re-reads the
    /// object with exactly this arity.
    /// </summary>
    public const int EditResultAnchorPropertyCount = 4;

    /// <summary>The scope segment of a canonical <c>kind:scope:unid</c> anchor.</summary>
    public static string ScopeFromAnchor(string anchorId)
    {
        var first = anchorId.IndexOf(':');
        var second = first < 0 ? -1 : anchorId.IndexOf(':', first + 1);
        if (first <= 0 || second <= first + 1 || second == anchorId.Length - 1)
        {
            throw new DeliveryReceiptValidationException(
                "invalid_anchor_id", $"'{anchorId}' is not a canonical kind:scope:unid anchor.");
        }
        return anchorId[(first + 1)..second];
    }

    /// <summary>
    /// The digest-token grammar — <c>sha256:</c> followed by 64 lower-case hex
    /// characters — shared by transaction entry ids, request fingerprints, and
    /// redaction tokens.
    /// </summary>
    public static bool IsSha256DigestToken(string? value)
    {
        if (value is null
            || value.Length != 71
            || !value.StartsWith("sha256:", StringComparison.Ordinal))
        {
            return false;
        }
        foreach (var character in value.AsSpan(7))
        {
            if (character is not (>= '0' and <= '9') and not (>= 'a' and <= 'f'))
                return false;
        }
        return true;
    }

    /// <summary>
    /// The redacted free-text shape: <c>{label}; {count} characters; {digestToken}</c>.
    /// The producer formats with it and the verifier re-parses the identical shape.
    /// </summary>
    public static string RedactedFreeText(string label, int characterCount, string digestToken) =>
        $"{label}; {characterCount.ToString(CultureInfo.InvariantCulture)}"
        + $" characters; {digestToken}";

    /// <summary>True when <paramref name="value"/> matches the
    /// <see cref="RedactedFreeText"/> shape for <paramref name="label"/>.</summary>
    public static bool IsRedactedFreeText(string value, string label)
    {
        var prefix = $"{label}; ";
        const string separator = " characters; ";
        if (!value.StartsWith(prefix, StringComparison.Ordinal))
            return false;
        int separatorIndex = value.IndexOf(separator, prefix.Length, StringComparison.Ordinal);
        return separatorIndex > prefix.Length
            && int.TryParse(
                value.AsSpan(prefix.Length, separatorIndex - prefix.Length),
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out var count)
            && count >= 0
            && IsSha256DigestToken(value[(separatorIndex + separator.Length)..]);
    }

    /// <summary>
    /// Coordinate equality between a projected citation and claimed coordinates.
    /// Availability metadata is compared by the producer only; the wire citation
    /// carries no availability, so this is the largest shape both sides check.
    /// </summary>
    public static bool CitationCoordinatesEqual(
        PageCitation projected,
        string? anchorId,
        long documentVersion,
        string? rendererFingerprint,
        IReadOnlyList<PageMapPage> pages,
        IReadOnlyList<PageMapFragment> fragments) =>
        string.Equals(anchorId, projected.AnchorId, StringComparison.Ordinal)
        && documentVersion == projected.DocumentVersion
        && string.Equals(rendererFingerprint, projected.RendererFingerprint,
            StringComparison.Ordinal)
        && pages.SequenceEqual(projected.Pages)
        && fragments.SequenceEqual(projected.Fragments);
}
