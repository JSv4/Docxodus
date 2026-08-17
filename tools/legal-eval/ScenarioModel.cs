// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json.Nodes;
using Docxodus.Verification;

namespace LegalEval;

public enum EvalTier
{
    Fast,
    Full,
}

public enum ScoreKind
{
    EngineBaseline,
    ModelPlanning,
}

public enum ArtifactRenderMode
{
    Disabled,
    TrustedDocuments,
}

public sealed record RenderedDocumentEvidence(
    string? PdfPath,
    IReadOnlyList<string> PagePaths,
    string? Error);

public sealed record RenderedVisualDiffEvidence(
    IReadOnlyList<string> Paths,
    string? Error,
    long? DifferentPixels = null);

/// <summary>
/// Injectable boundary around process-based rendering. Tests can exercise both sides of the
/// availability contract without invoking an office suite, while production uses the isolated
/// disposable-profile adapter in <see cref="ArtifactWriter"/>.
/// </summary>
public interface IEvaluationArtifactRenderer
{
    RenderedDocumentEvidence RenderDocument(
        string artifactDirectory,
        string sourceDocxPath,
        string outputPrefix);

    RenderedVisualDiffEvidence ComparePages(
        string artifactDirectory,
        IReadOnlyList<string> targetPages,
        IReadOnlyList<string> candidatePages);
}

public sealed record FixtureReference(
    string Path,
    string ProvenanceId,
    string SourceSha256);

public sealed record ExpectedDocumentReference(
    string Path,
    string ProvenanceId,
    string SourceSha256);

public sealed record ChangeBudget(
    IReadOnlySet<string> AllowedChangedParts,
    IReadOnlySet<string> AllowedRelationshipOwners,
    int MaximumChangedAnchors);

public sealed record ExpectedArtifact(
    string Id,
    string MediaType,
    string Role,
    bool Required);

public sealed record DeterministicInvariant(
    string Id,
    string Metric,
    JsonObject Probe,
    string Operator,
    JsonNode? Expected);

public enum RedlineReversibilityApplicability
{
    Required,
    NotApplicable,
}

public sealed record RedlineReversibilityPolicy(
    RedlineReversibilityApplicability Applicability,
    string? Reason);

public sealed record LegalScenario(
    string FilePath,
    string Id,
    string Title,
    EvalTier Tier,
    FixtureReference Fixture,
    ExpectedDocumentReference ExpectedDocument,
    string Instruction,
    IReadOnlyList<string> Constraints,
    RedlineReversibilityPolicy RedlineReversibility,
    IReadOnlyList<ExpectedArtifact> ExpectedOutputs,
    IReadOnlyList<JsonObject> BaselineOperations,
    IReadOnlyList<DeterministicInvariant> Invariants,
    ChangeBudget ChangeBudget);

public sealed record FixtureProvenance(
    string Id,
    string Title,
    string Origin,
    string Author,
    string Created,
    string License,
    string RedistributionPermission,
    string SourcePath,
    string SourceSha256,
    string? RecipePath,
    string? RecipeSha256);

public sealed record ExpectedDocumentProvenance(
    string Id,
    string ScenarioId,
    string Origin,
    string GeneratedBy,
    string Created,
    string ReviewStatus,
    string ReviewNotes,
    string License,
    string RedistributionPermission,
    string SourcePath,
    string SourceSha256);

public sealed record LegalCorpus(
    string RootDirectory,
    IReadOnlyList<LegalScenario> Scenarios,
    IReadOnlyDictionary<string, FixtureProvenance> Provenance,
    IReadOnlyDictionary<string, ExpectedDocumentProvenance> ExpectedDocumentProvenance);

public sealed record MetricResult(
    string Id,
    string Category,
    string Status,
    string Detail,
    double? Score = null);

public sealed record ArtifactRecord(
    string Id,
    string Status,
    string? Path,
    string? Sha256,
    string? UnavailableReason,
    string MediaType,
    long? SizeBytes,
    string Role = "evidence");

internal sealed record EvaluationEvidence
{
    public string? ScenarioContractJson { get; init; }
    public byte[]? Input { get; init; }
    public byte[]? Candidate { get; init; }
    public byte[]? Expected { get; init; }
    public PackageManifest? InputManifest { get; init; }
    public PackageManifest? CandidateManifest { get; init; }
    public PackageManifest? ExpectedManifest { get; init; }
    public SemanticChangeSet? CandidateChanges { get; init; }
    public SemanticChangeSet? TargetChanges { get; init; }
    public SemanticChangeSet? CandidateTargetChanges { get; init; }
    public DeliverableVerificationResult? DeliverableVerification { get; init; }
    public byte[]? RedlineBytes { get; init; }
    public RedlineReversibilityProofRun? RedlineProofRun { get; init; }
    public DeliveryChangeReceipt? DeliveryReceipt { get; init; }
    public IReadOnlyDictionary<string, byte[]> DeliveryReceiptArtifacts { get; init; } =
        new Dictionary<string, byte[]>(StringComparer.Ordinal);
    public string? DeliveryReceiptUnavailableReason { get; init; }
    public string? CandidateHtml { get; init; }
    public string? FailureReason { get; init; }
    public string? InputSafetyError { get; init; }
    public string? CandidateSafetyError { get; init; }
    public string? ExpectedSafetyError { get; init; }
}

public sealed record EvaluationScore(
    string ScenarioId,
    ScoreKind Kind,
    string Status,
    IReadOnlyList<MetricResult> Metrics,
    IReadOnlyList<ArtifactRecord> Artifacts,
    string? ArtifactDirectory);

public sealed record ScenarioRunResult(
    string ScenarioId,
    EvaluationScore EngineBaseline,
    EvaluationScore? ModelPlanning);

public sealed class ScenarioValidationException : Exception
{
    public ScenarioValidationException(string message) : base(message) { }
}
