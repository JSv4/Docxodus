// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json;
using System.Text.Json.Serialization;

namespace Docxodus.Verification;

/// <summary>How a revision in a redline relates to the selected baseline.</summary>
public enum RedlineRevisionDisposition
{
    PreExisting,
    Generated,
    Conflicted,
}

/// <summary>The proof path that was evaluated.</summary>
public enum RedlineProofDirection
{
    AcceptToFinal,
    RejectToBaseline,
}

/// <summary>How a package entry differs from the path's expected document.</summary>
public enum RedlinePackageDivergenceKind
{
    Added,
    Removed,
    Modified,
}

/// <summary>
/// Settings for <see cref="RedlineReversibilityVerifier.Prove"/>. Package limits are shared with
/// package-manifest generation; proof verification never opens a package that fails that preflight.
/// </summary>
public sealed record RedlineReversibilityProofOptions
{
    public PackageManifestOptions PackageManifestOptions { get; init; } = new();

    /// <summary>
    /// Require identical ZIP bytes in addition to modeled and normalized whole-package equality.
    /// This is normally false because harmless ZIP timestamps, order, and compression may differ.
    /// </summary>
    public bool RequireExactPackageBytes { get; init; }
}

/// <summary>A stable, part-qualified identity for one native Word revision.</summary>
public sealed record RedlineRevisionIdentity
{
    required public string Id { get; init; }
    required public string PartUri { get; init; }
    required public string Scope { get; init; }
    required public string Type { get; init; }
    required public RevisionFamily Family { get; init; }
    required public IReadOnlyList<string> ConstituentIds { get; init; }
    required public string Author { get; init; }
    public string? Date { get; init; }
    required public string Text { get; init; }
    public string? AnchorId { get; init; }
    required public IReadOnlyList<string> AffectedAnchorIds { get; init; }
    required public RevisionResolutionStatus ResolutionStatus { get; init; }
    public RevisionDiagnostic? Diagnostic { get; init; }
}

/// <summary>Classification of a baseline/redline revision identity pair.</summary>
public sealed record RedlineRevisionClassification
{
    required public RedlineRevisionDisposition Disposition { get; init; }
    public RedlineRevisionIdentity? Baseline { get; init; }
    public RedlineRevisionIdentity? Redline { get; init; }
    required public string Reason { get; init; }
}

/// <summary>Input or output package identity recorded by the proof.</summary>
public sealed record RedlineProofPackageIdentity
{
    required public VerificationDigest RawPackageBytesDigest { get; init; }
    public VerificationDigest? OrderedOpcContentDigest { get; init; }
    public VerificationDigest? NormalizedWholePackageDigest { get; init; }
}

/// <summary>
/// Modeled semantic comparison result. The semantic schema and change count are explicit so a
/// caller never mistakes an empty modeled change set for complete package equality.
/// </summary>
public sealed record RedlineModeledSemanticComparison
{
    required public bool Available { get; init; }
    public bool? Equivalent { get; init; }
    public string? Schema { get; init; }
    public int? ChangeCount { get; init; }
    public string? Diagnostic { get; init; }
}

/// <summary>One added, removed, or modified package entry.</summary>
public sealed record RedlinePackageDivergence
{
    required public RedlinePackageDivergenceKind Kind { get; init; }
    required public string PartUri { get; init; }
    required public int Occurrence { get; init; }
    public string? AnchorId { get; init; }
    required public IReadOnlyList<string> ApplicableRevisionIds { get; init; }
    public VerificationDigest? ExpectedRawDigest { get; init; }
    public VerificationDigest? ActualRawDigest { get; init; }
    public VerificationDigest? ExpectedNormalizedDigest { get; init; }
    public VerificationDigest? ActualNormalizedDigest { get; init; }
    /// <summary>Whether the semantic change set reports a modeled change for this part.</summary>
    required public bool HasModeledSemanticChange { get; init; }
    /// <summary>
    /// Whether the normalized entry difference may contain content outside the modeled semantic
    /// projection. This is deliberately conservative because a modeled change in one part does
    /// not prove that every other change in that part was modeled.
    /// </summary>
    required public bool UnknownOrUnmodeled { get; init; }
}

/// <summary>A structured, actionable proof finding.</summary>
public sealed record RedlineProofFinding
{
    required public string Code { get; init; }
    required public VerificationFindingSeverity Severity { get; init; }
    required public string Message { get; init; }
    public RedlineProofDirection? Direction { get; init; }
    public ChangeLocation? Location { get; init; }
    public string? AnchorId { get; init; }
    required public IReadOnlyList<string> RevisionIds { get; init; }
    public string? Remediation { get; init; }
}

/// <summary>Result of accepting or rejecting only the generated revision set.</summary>
public sealed record RedlineProofPathResult
{
    required public RedlineProofDirection Direction { get; init; }
    required public bool Completed { get; init; }
    required public bool Equivalent { get; init; }
    required public IReadOnlyList<string> RequestedRevisionIds { get; init; }
    required public IReadOnlyList<string> ResolvedRevisionIds { get; init; }
    required public IReadOnlyList<string> ImplicitlyResolvedRevisionIds { get; init; }
    required public IReadOnlyList<RedlineRevisionIdentity> SurvivingPreExistingRevisions { get; init; }
    required public bool PreExistingRevisionsPreserved { get; init; }
    required public RedlineModeledSemanticComparison ModeledSemantic { get; init; }
    required public bool NormalizedWholePackageEquivalent { get; init; }
    required public bool OrderedOpcContentEquivalent { get; init; }
    required public bool ExactPackageBytesEquivalent { get; init; }
    required public RedlineProofPackageIdentity ExpectedPackage { get; init; }
    public RedlineProofPackageIdentity? ActualPackage { get; init; }
    public RedlinePackageDivergence? FirstDivergence { get; init; }
    required public IReadOnlyList<RedlinePackageDivergence> Divergences { get; init; }
    required public IReadOnlyList<RedlineProofFinding> Findings { get; init; }
}

/// <summary>
/// Versioned, receipt-embeddable proof that generated changes accept to the intended final and
/// reject to the selected baseline without consuming pre-existing review state.
/// </summary>
public sealed record RedlineReversibilityProof
{
    public const string SchemaId =
        "https://docxodus.dev/schemas/verification/redline-reversibility-proof/v1";

    public string Schema { get; init; } = SchemaId;
    public int SchemaVersion { get; init; } = 1;
    required public bool Success { get; init; }
    required public bool RequireExactPackageBytes { get; init; }
    required public RedlineProofPackageIdentity BaselinePackage { get; init; }
    required public RedlineProofPackageIdentity IntendedFinalPackage { get; init; }
    required public RedlineProofPackageIdentity RedlinePackage { get; init; }
    required public IReadOnlyList<RedlineRevisionClassification> RevisionClassifications { get; init; }
    public RedlineProofPathResult? AcceptToFinal { get; init; }
    public RedlineProofPathResult? RejectToBaseline { get; init; }
    required public IReadOnlyList<RedlineProofFinding> Findings { get; init; }

    /// <summary>Serialize the canonical, compact proof JSON used by delivery receipts.</summary>
    public string ToCanonicalJson() => JsonSerializer.Serialize(this, JsonOptions.Canonical);

    /// <summary>Serialize proof JSON, optionally indented for a human-viewable artifact.</summary>
    public string ToJson(bool indented = true) => JsonSerializer.Serialize(
        this, indented ? JsonOptions.Indented : JsonOptions.Canonical);

    private static class JsonOptions
    {
        internal static readonly JsonSerializerOptions Canonical = Create(indented: false);
        internal static readonly JsonSerializerOptions Indented = Create(indented: true);

        private static JsonSerializerOptions Create(bool indented)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = indented,
            };
            options.Converters.Add(new JsonStringEnumConverter(JsonNamingPolicy.CamelCase));
            return options;
        }
    }
}

/// <summary>
/// Proof metadata plus the two concrete output packages. Package bytes are deliberately outside
/// the JSON proof so a delivery receipt embeds hashes and structured evidence rather than base64.
/// </summary>
public sealed record RedlineReversibilityProofRun
{
    required public RedlineReversibilityProof Proof { get; init; }
    public byte[]? AcceptedPackageBytes { get; init; }
    public byte[]? RejectedPackageBytes { get; init; }
}
