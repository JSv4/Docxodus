// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json.Nodes;

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
    string? Error);

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

public sealed record LegalScenario(
    string FilePath,
    string Id,
    string Title,
    EvalTier Tier,
    FixtureReference Fixture,
    ExpectedDocumentReference ExpectedDocument,
    string Instruction,
    IReadOnlyList<string> Constraints,
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
    long? SizeBytes);

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
