// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Docxodus.Verification;

namespace LegalEval;

public sealed record EvaluationRunOptions(
    string CorpusPath,
    string Subset,
    string? ScenarioId,
    string? CandidateDirectory,
    string ArtifactRoot,
    string? ReportPath,
    ArtifactRenderMode RenderMode = ArtifactRenderMode.Disabled);

public sealed record EvaluationRunOutcome(
    int ExitCode,
    string ArtifactRoot,
    bool ArtifactsPublished,
    string? SummaryPath,
    IReadOnlyList<ScenarioRunResult> Results,
    string? FatalError,
    string? ReportError);

/// <summary>
/// Single artifact-producing orchestration path used by both the CLI and xUnit. Scenario-level
/// failures are converted to evidence bundles and never prevent later scenarios or the portable
/// run summary from being published.
/// </summary>
public sealed class LegalEvaluationRunner
{
    private const long MaximumCandidateBytes = 64L * 1024 * 1024;
    private const long MaximumOwnedReportBytes = 64L * 1024 * 1024;
    private const string ArtifactRootMarkerName = ".docxodus-legal-eval-root";
    private const string ArtifactRootMarkerContent = "docxodus.legal-evaluation-artifacts/1.0\n";
    private const string RunSummaryDocumentKind = "docxodus.legal-evaluation-run-summary";
    private readonly Func<LegalScenario, LegalScenario>? _scenarioTransform;
    private readonly IEvaluationArtifactRenderer? _artifactRenderer;
    private readonly Action<string, string> _externalReportCommitter;

    public LegalEvaluationRunner(
        Func<LegalScenario, LegalScenario>? scenarioTransform = null,
        IEvaluationArtifactRenderer? artifactRenderer = null,
        Action<string, string>? externalReportCommitter = null)
    {
        _scenarioTransform = scenarioTransform;
        _artifactRenderer = artifactRenderer;
        _externalReportCommitter = externalReportCommitter ?? CommitStagedExternalReport;
    }

    public EvaluationRunOutcome Run(
        EvaluationRunOptions options,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        output ??= TextWriter.Null;
        error ??= TextWriter.Null;
        var artifactRoot = NormalizeDirectoryPath(options.ArtifactRoot);
        var corpusPath = ResolveCanonicalPath(options.CorpusPath);
        var corpusRoot = Path.GetDirectoryName(corpusPath)
            ?? throw new ScenarioValidationException("corpus path has no parent directory");
        var candidateDirectory = options.CandidateDirectory is null
            ? null
            : NormalizeDirectoryPath(options.CandidateDirectory);
        ValidateArtifactRootScope(artifactRoot, corpusRoot, candidateDirectory);
        ValidateArtifactRootOwnership(artifactRoot);
        var requestedSummaryPath = Path.Combine(artifactRoot, "run-summary.json");
        var externalReportPath = ResolveExternalReportPath(
            options.ReportPath, artifactRoot, requestedSummaryPath, corpusRoot, candidateDirectory);
        var stagingRoot = CreateStagingRoot(artifactRoot);
        output.WriteLine($"Artifacts: {artifactRoot}");
        var results = new List<ScenarioRunResult>();
        string? fatalError = null;
        string? reportError = null;
        var exitCode = 0;
        string? summaryPath = requestedSummaryPath;
        var published = false;
        StagedExternalReport? stagedExternalReport = null;
        var publicationPhase = "assemble the artifact root";

        try
        {
            var corpus = ScenarioLoader.LoadCorpus(corpusPath);
            var scenarios = corpus.Scenarios
                .Where(value => options.ScenarioId is null || value.Id == options.ScenarioId)
                .Where(value => options.Subset == "full" || value.Tier == EvalTier.Fast)
                .Select(value => _scenarioTransform?.Invoke(value) ?? value)
                .ToList();
            if (scenarios.Count == 0)
                throw new ScenarioValidationException("No scenarios matched the requested subset/filter.");

            var executor = new ScriptedBaselineExecutor();
            var scorer = new EvaluationScorer(artifactRenderer: _artifactRenderer);
            foreach (var scenario in scenarios)
            {
                BaselineExecution? baseline = null;
                EvaluationScore engineScore;
                try
                {
                    baseline = executor.ExecuteCheckpointed(scenario);
                    if (!baseline.Succeeded)
                    {
                        engineScore = FailureScore(scenario, baseline, ScoreKind.EngineBaseline,
                            stagingRoot, baseline.Error ?? "scripted baseline failed", options.RenderMode,
                            artifactRenderer: _artifactRenderer);
                    }
                    else
                    {
                        engineScore = scorer.Score(scenario, baseline, baseline.Output,
                            ScoreKind.EngineBaseline, stagingRoot, options.RenderMode);
                    }
                }
                catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
                {
                    baseline ??= EmergencyBaseline(scenario, exception);
                    engineScore = FailureScore(scenario, baseline, ScoreKind.EngineBaseline,
                        stagingRoot, ExceptionDetail(exception,
                            Path.GetDirectoryName(scenario.FilePath), stagingRoot, artifactRoot),
                        options.RenderMode,
                        artifactRenderer: _artifactRenderer);
                }

                EvaluationScore? planningScore = null;
                if (engineScore.Status == "passed" && candidateDirectory is not null)
                {
                    var candidatePath = Path.Combine(
                        candidateDirectory, scenario.Id + ".docx");
                    if (!File.Exists(candidatePath))
                    {
                        var reason = $"candidate file is absent: {scenario.Id}.docx in the requested candidate directory";
                        planningScore = FailureScore(scenario, baseline!, ScoreKind.ModelPlanning,
                            stagingRoot, reason, ArtifactRenderMode.Disabled, status: "incomplete",
                            candidate: null, artifactRenderer: _artifactRenderer);
                    }
                    else
                    {
                        try
                        {
                            var candidate = ReadCandidateBounded(candidatePath);
                            planningScore = scorer.Score(scenario, baseline!, candidate,
                                ScoreKind.ModelPlanning, stagingRoot, options.RenderMode);
                        }
                        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
                        {
                            byte[]? candidate = null;
                            try { candidate = ReadCandidateBounded(candidatePath); }
                            catch (Exception readException)
                                when (DeliverableExceptionBoundary.IsRecoverable(readException)) { }
                            planningScore = FailureScore(scenario, baseline!, ScoreKind.ModelPlanning,
                                stagingRoot, ExceptionDetail(exception,
                                    candidateDirectory, stagingRoot, artifactRoot),
                                ArtifactRenderMode.Disabled,
                                candidate: candidate, artifactRenderer: _artifactRenderer);
                        }
                    }
                }

                results.Add(new ScenarioRunResult(scenario.Id, engineScore, planningScore));
                output.WriteLine($"{engineScore.Status.ToUpperInvariant()} {scenario.Id} engine-baseline"
                    + (planningScore is null ? string.Empty : $"; {planningScore.Status} model-planning")
                    + $"; artifacts={RemapPath(stagingRoot, artifactRoot, engineScore.ArtifactDirectory!)}");
                if (engineScore.Status != "passed"
                    || planningScore?.Status is "failed" or "incomplete")
                    exitCode = 1;
            }
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            fatalError = ExceptionDetail(exception,
                corpusRoot, stagingRoot, artifactRoot);
            error.WriteLine(exception.Message);
            exitCode = 2;
        }
        finally
        {
            try
            {
                var summaryJson = BuildPortableSummary(
                    options, stagingRoot, results, fatalError);
                AtomicWriteText(Path.Combine(stagingRoot, "run-summary.json"), summaryJson);
                AtomicWriteText(Path.Combine(stagingRoot, "index.html"),
                    BuildRunIndex(results, fatalError));
                publicationPhase = "stage the external report";
                if (externalReportPath is not null)
                    stagedExternalReport = StageExternalReport(externalReportPath, summaryJson);

                var publishedResults = RemapResults(results, stagingRoot, artifactRoot);
                publicationPhase = "publish the artifact root";
                PublishStagedRoot(stagingRoot, artifactRoot);
                published = true;
                results = publishedResults;
                if (stagedExternalReport is not null)
                {
                    try
                    {
                        ValidateExternalReportDestinationOwnership(
                            stagedExternalReport.DestinationPath);
                        _externalReportCommitter(
                            stagedExternalReport.TemporaryPath,
                            stagedExternalReport.DestinationPath);
                        stagedExternalReport = null;
                    }
                    catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
                    {
                        reportError = "external report publication failed after the artifact root "
                            + "was published; the artifact root remains valid: "
                            + ExceptionDetail(exception, artifactRoot,
                                Path.GetDirectoryName(externalReportPath));
                        error.WriteLine(reportError);
                        exitCode = 2;
                    }
                }
                output.WriteLine($"Summary: {summaryPath}");
            }
            catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
            {
                var publicationError = published
                    ? "artifact root publication completed, but final reporting failed; "
                        + "the published artifact root remains valid: "
                        + ExceptionDetail(exception, stagingRoot, artifactRoot,
                            externalReportPath is null
                                ? null
                                : Path.GetDirectoryName(externalReportPath))
                    : $"failed to {publicationPhase} before artifact root publication: "
                        + ExceptionDetail(exception, stagingRoot, artifactRoot,
                            externalReportPath is null
                                ? null
                                : Path.GetDirectoryName(externalReportPath));
                fatalError = fatalError is null
                    ? publicationError
                    : fatalError + Environment.NewLine + publicationError;
                error.WriteLine(exception.Message);
                exitCode = 2;
                if (!published)
                {
                    results.Clear();
                    summaryPath = null;
                }
            }
            finally
            {
                if (stagedExternalReport is not null)
                    DeleteFileBestEffort(stagedExternalReport.TemporaryPath);
                if (!published) DeleteDirectoryBestEffort(stagingRoot);
            }
        }

        return new EvaluationRunOutcome(exitCode, artifactRoot, published, summaryPath,
            results, fatalError, reportError);
    }

    private static string CreateStagingRoot(string artifactRoot)
    {
        var parent = Path.GetDirectoryName(artifactRoot)
            ?? throw new InvalidOperationException("artifact root has no parent directory");
        var name = Path.GetFileName(artifactRoot);
        if (name.Length == 0)
            throw new InvalidOperationException("artifact root must name a directory");
        Directory.CreateDirectory(parent);
        var stagingRoot = Path.Combine(parent,
            "." + name + ".stage-" + Guid.NewGuid().ToString("N"));
        try
        {
            Directory.CreateDirectory(stagingRoot);
            File.WriteAllText(Path.Combine(stagingRoot, ArtifactRootMarkerName),
                ArtifactRootMarkerContent,
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            return stagingRoot;
        }
        catch
        {
            DeleteDirectoryBestEffort(stagingRoot);
            throw;
        }
    }

    private static void PublishStagedRoot(string stagingRoot, string artifactRoot)
    {
        if (!HasArtifactRootMarker(stagingRoot))
            throw new IOException("staged artifact root is missing its ownership marker");
        ValidateArtifactRootOwnership(artifactRoot);
        if (File.Exists(artifactRoot))
            throw new IOException("artifact root is an existing file");
        var parent = Path.GetDirectoryName(artifactRoot)
            ?? throw new InvalidOperationException("artifact root has no parent directory");
        var backup = Path.Combine(parent, "." + Path.GetFileName(artifactRoot)
            + ".backup-" + Guid.NewGuid().ToString("N"));
        var hadExisting = Directory.Exists(artifactRoot);
        if (hadExisting) Directory.Move(artifactRoot, backup);
        try
        {
            Directory.Move(stagingRoot, artifactRoot);
        }
        catch
        {
            if (hadExisting && Directory.Exists(backup) && !Directory.Exists(artifactRoot))
                Directory.Move(backup, artifactRoot);
            throw;
        }
        DeleteDirectoryBestEffort(backup);
    }

    private static void ValidateArtifactRootScope(
        string artifactRoot, string corpusRoot, string? candidateDirectory)
    {
        var currentDirectory = NormalizeDirectoryPath(Directory.GetCurrentDirectory());
        if (IsUnderRoot(artifactRoot, currentDirectory))
            throw new ScenarioValidationException(
                "artifact root must not equal or contain the current working directory");
        if (PathsOverlap(artifactRoot, corpusRoot))
            throw new ScenarioValidationException(
                "artifact root must not overlap the legal evaluation corpus or its source files");
        if (candidateDirectory is not null && PathsOverlap(artifactRoot, candidateDirectory))
            throw new ScenarioValidationException(
                "artifact root must not overlap the model candidate directory");
    }

    private static void ValidateArtifactRootOwnership(string artifactRoot)
    {
        if (File.Exists(artifactRoot))
            throw new ScenarioValidationException("artifact root is an existing file");
        if (!Directory.Exists(artifactRoot)
            || !Directory.EnumerateFileSystemEntries(artifactRoot).Any())
            return;
        var markerPath = Path.Combine(artifactRoot, ArtifactRootMarkerName);
        if (File.Exists(markerPath))
        {
            if (HasArtifactRootMarker(artifactRoot)) return;
            throw new ScenarioValidationException(
                "existing artifact root has an invalid legal-eval ownership marker");
        }
        throw new ScenarioValidationException(
            "existing artifact root is not owned by legal-eval; choose an empty directory or a prior legal-eval root");
    }

    private static bool HasArtifactRootMarker(string artifactRoot)
    {
        var marker = Path.Combine(artifactRoot, ArtifactRootMarkerName);
        try
        {
            var info = new FileInfo(marker);
            return info.Exists
                && info.LinkTarget is null
                && info.Length == Encoding.UTF8.GetByteCount(ArtifactRootMarkerContent)
                && string.Equals(File.ReadAllText(marker), ArtifactRootMarkerContent,
                    StringComparison.Ordinal);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    private static string? ResolveExternalReportPath(
        string? reportPath,
        string artifactRoot,
        string summaryPath,
        string corpusRoot,
        string? candidateDirectory)
    {
        if (reportPath is null) return null;
        var fullReportPath = ResolveCanonicalPath(reportPath);
        if (IsUnderRoot(artifactRoot, fullReportPath))
        {
            if (string.Equals(fullReportPath, summaryPath, PathComparison)) return null;
            throw new ScenarioValidationException(
                "--report inside the artifact root must be the canonical run-summary.json path");
        }
        if (IsUnderRoot(corpusRoot, fullReportPath))
            throw new ScenarioValidationException(
                "external report path must not overwrite the corpus or its source files");
        if (candidateDirectory is not null && IsUnderRoot(candidateDirectory, fullReportPath))
            throw new ScenarioValidationException(
                "external report path must not overwrite a model candidate");
        ValidateExternalReportDestinationOwnership(fullReportPath);
        return fullReportPath;
    }

    private static StagedExternalReport StageExternalReport(string destinationPath, string value)
    {
        if (Directory.Exists(destinationPath))
            throw new IOException("external report path is an existing directory");
        ValidateExternalReportDestinationOwnership(destinationPath);
        var directory = Path.GetDirectoryName(destinationPath)
            ?? throw new InvalidOperationException("external report path has no parent directory");
        Directory.CreateDirectory(directory);
        var temporaryPath = destinationPath + ".stage-" + Guid.NewGuid().ToString("N");
        try
        {
            File.WriteAllText(temporaryPath, value,
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            return new StagedExternalReport(destinationPath, temporaryPath);
        }
        catch
        {
            DeleteFileBestEffort(temporaryPath);
            throw;
        }
    }

    private static void CommitStagedExternalReport(string temporaryPath, string destinationPath)
    {
        if (!File.Exists(destinationPath))
        {
            // Do not overwrite a file created after the final ownership check.
            File.Move(temporaryPath, destinationPath);
            return;
        }
        ValidateExternalReportDestinationOwnership(destinationPath);
        File.Move(temporaryPath, destinationPath, overwrite: true);
    }

    private static void ValidateExternalReportDestinationOwnership(string destinationPath)
    {
        if (!File.Exists(destinationPath)) return;
        try
        {
            using var stream = File.OpenRead(destinationPath);
            if (stream.Length > MaximumOwnedReportBytes)
                throw new ScenarioValidationException(
                    "existing external report is not owned by legal-eval");
            using var document = JsonDocument.Parse(stream);
            var root = document.RootElement;
            if (root.ValueKind == JsonValueKind.Object
                && root.TryGetProperty("documentKind", out var documentKind)
                && documentKind.ValueKind == JsonValueKind.String
                && documentKind.GetString() == RunSummaryDocumentKind
                && root.TryGetProperty("schemaVersion", out var version)
                && version.ValueKind == JsonValueKind.String
                && version.GetString() == "1.0"
                && root.TryGetProperty("artifactRoot", out var reportedRoot)
                && reportedRoot.ValueKind == JsonValueKind.String
                && reportedRoot.GetString() == "."
                && root.TryGetProperty("corpus", out var corpus)
                && corpus.ValueKind == JsonValueKind.String
                && root.TryGetProperty("subset", out var subset)
                && subset.ValueKind == JsonValueKind.String
                && root.TryGetProperty("results", out var results)
                && results.ValueKind == JsonValueKind.Array)
                return;
        }
        catch (Exception exception) when (exception is JsonException or IOException
            or UnauthorizedAccessException or InvalidOperationException)
        {
            // The uniform refusal below avoids treating malformed or unreadable files as owned.
        }
        throw new ScenarioValidationException(
            "existing external report is not owned by legal-eval");
    }

    private static void DeleteFileBestEffort(string path)
    {
        try
        {
            if (File.Exists(path)) File.Delete(path);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            // A failed best-effort cleanup must not invalidate a published artifact root.
        }
    }

    private static void DeleteDirectoryBestEffort(string path)
    {
        try
        {
            if (Directory.Exists(path)) Directory.Delete(path, recursive: true);
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            // Publication is complete (or the prior root has already been restored). A leftover
            // hidden stage/backup is safer than invalidating the visible root during cleanup.
        }
    }

    private static List<ScenarioRunResult> RemapResults(
        IReadOnlyList<ScenarioRunResult> results, string stagingRoot, string artifactRoot) =>
        results.Select(result => result with
        {
            EngineBaseline = RemapScore(result.EngineBaseline, stagingRoot, artifactRoot),
            ModelPlanning = result.ModelPlanning is null
                ? null
                : RemapScore(result.ModelPlanning, stagingRoot, artifactRoot),
        }).ToList();

    private static EvaluationScore RemapScore(
        EvaluationScore score, string stagingRoot, string artifactRoot) =>
        score with
        {
            ArtifactDirectory = score.ArtifactDirectory is null
                ? null
                : RemapPath(stagingRoot, artifactRoot, score.ArtifactDirectory),
            Artifacts = score.Artifacts.Select(artifact => artifact.Path is null
                    ? artifact
                    : artifact with
                    {
                        Path = RemapPath(stagingRoot, artifactRoot, artifact.Path),
                    })
                .ToList(),
        };

    private static string RemapPath(string sourceRoot, string destinationRoot, string path) =>
        Path.GetFullPath(Path.Combine(destinationRoot, RelativeUnderRoot(sourceRoot, path)));

    private static string NormalizeDirectoryPath(string path) =>
        Path.TrimEndingDirectorySeparator(ResolveCanonicalPath(path));

    private static bool PathsOverlap(string left, string right) =>
        IsUnderRoot(left, right) || IsUnderRoot(right, left);

    private static bool IsUnderRoot(string root, string path)
    {
        var fullRoot = ResolveCanonicalPath(root);
        var fullPath = ResolveCanonicalPath(path);
        var prefix = fullRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        return fullPath.StartsWith(prefix, PathComparison)
            || string.Equals(fullPath, fullRoot, PathComparison);
    }

    /// <summary>
    /// Resolves every existing path component (including directory symlinks/junctions), then
    /// appends any not-yet-existing suffix. This makes scope checks reflect the location that a
    /// later create or replace operation will actually address.
    /// </summary>
    private static string ResolveCanonicalPath(string path)
    {
        var remainingComponents = 1024;
        var remainingLinks = 64;
        return ResolveCanonicalPathCore(
            Path.GetFullPath(path), ref remainingComponents, ref remainingLinks);
    }

    private static string ResolveCanonicalPathCore(
        string fullPath, ref int remainingComponents, ref int remainingLinks)
    {
        var root = Path.GetPathRoot(fullPath)
            ?? throw new ScenarioValidationException("path has no filesystem root");
        var remainder = fullPath[root.Length..];
        var segments = remainder.Split(
            new[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar },
            StringSplitOptions.RemoveEmptyEntries);

        var current = root;
        var existingPrefix = true;
        foreach (var segment in segments)
        {
            if (--remainingComponents < 0)
                throw new ScenarioValidationException("path resolution exceeded its component limit");
            var candidate = Path.Combine(current, segment);
            if (existingPrefix && (Directory.Exists(candidate) || File.Exists(candidate)))
            {
                FileSystemInfo info = Directory.Exists(candidate)
                    ? new DirectoryInfo(candidate)
                    : new FileInfo(candidate);
                var resolved = info.ResolveLinkTarget(returnFinalTarget: true);
                if (resolved is null)
                {
                    current = candidate;
                }
                else
                {
                    if (--remainingLinks < 0)
                        throw new ScenarioValidationException(
                            "path resolution exceeded its symbolic-link limit");
                    // A link target can itself contain symlinked ancestors even when its final
                    // component is not a link, so canonicalize the returned target from its root.
                    current = ResolveCanonicalPathCore(
                        Path.GetFullPath(resolved.FullName),
                        ref remainingComponents,
                        ref remainingLinks);
                }
            }
            else
            {
                existingPrefix = false;
                current = candidate;
            }
        }
        return Path.GetFullPath(current);
    }

    private static EvaluationScore FailureScore(
        LegalScenario scenario,
        BaselineExecution baseline,
        ScoreKind kind,
        string artifactRoot,
        string reason,
        ArtifactRenderMode renderMode,
        string status = "failed",
        byte[]? candidate = null,
        IEvaluationArtifactRenderer? artifactRenderer = null)
    {
        candidate ??= kind == ScoreKind.EngineBaseline ? baseline.Output : null;
        var metrics = new[]
        {
            new MetricResult($"{KindName(kind)}.execution", "task_completion", status,
                reason, status == "failed" ? 0 : null),
        };
        var directory = ArtifactWriter.ResolveScoreDirectory(artifactRoot, scenario.Id, kind);
        var inputSafetyError = SafetyError(baseline.Input, "evaluation input");
        var candidateSafetyError = candidate is null
            ? null
            : SafetyError(candidate, "candidate document");
        var expectedSafetyError = SafetyError(baseline.Expected, "pinned expected document");
        var rendererSafe = inputSafetyError is null
            && candidateSafetyError is null
            && expectedSafetyError is null;
        var scorer = new EvaluationScorer(artifactRenderer: artifactRenderer);
        var evidence = scorer.AnalyzeAvailableEvidence(
            baseline, candidate, kind, reason,
            inputSafetyError, candidateSafetyError, expectedSafetyError);
        var publication = ArtifactWriter.WriteIncomplete(directory, scenario.Id, kind, status,
            metrics, scenario.ExpectedOutputs, reason, baseline.OperationLog, evidence,
            rendererSafe ? renderMode : ArtifactRenderMode.Disabled,
            allowExternalRenderer: rendererSafe && kind == ScoreKind.EngineBaseline,
            artifactRenderer: artifactRenderer,
            inputSafetyError: inputSafetyError,
            candidateSafetyError: candidateSafetyError,
            expectedSafetyError: expectedSafetyError);
        return new EvaluationScore(scenario.Id, kind, publication.Status, publication.Metrics,
            publication.Artifacts, directory);
    }

    private static BaselineExecution EmergencyBaseline(
        LegalScenario scenario, Exception exception)
    {
        byte[] input = Array.Empty<byte>();
        byte[] expected = Array.Empty<byte>();
        try
        {
            input = ReadFileBounded(scenario.Fixture.Path, MaximumCandidateBytes,
                "evaluation input");
        }
        catch (Exception readException) when (DeliverableExceptionBoundary.IsRecoverable(readException)) { }
        try
        {
            expected = ReadFileBounded(scenario.ExpectedDocument.Path, MaximumCandidateBytes,
                "pinned expected document");
        }
        catch (Exception readException) when (DeliverableExceptionBoundary.IsRecoverable(readException)) { }
        var detail = ExceptionDetail(exception, Path.GetDirectoryName(scenario.FilePath));
        return new BaselineExecution(input, input, expected,
            EvaluationScorer.ErrorEnvelope("semantic-diff", "failed", detail),
            SemanticDiffSucceeded: false,
            new[] { $"baseline-initialization-failed:{exception.GetType().Name}:{exception.Message}" },
            Succeeded: false,
            Error: exception.ToString());
    }

    private static byte[] ReadCandidateBounded(string path) =>
        ReadFileBounded(path, MaximumCandidateBytes, "candidate");

    private static byte[] ReadFileBounded(string path, long maximumBytes, string label)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
            bufferSize: 81920, FileOptions.SequentialScan);
        if (stream.Length > maximumBytes)
            throw new ScenarioValidationException(
                $"{label} exceeds the {maximumBytes}-byte read limit");
        var bytes = new byte[checked((int)stream.Length)];
        stream.ReadExactly(bytes);
        if (stream.ReadByte() != -1)
            throw new ScenarioValidationException($"{label} changed while it was being read");
        return bytes;
    }

    private static string? SafetyError(byte[] bytes, string label)
    {
        try
        {
            _ = new EvaluationPackageValidator().Inspect(bytes, label);
            return null;
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            return ExceptionDetail(exception);
        }
    }

    private static string BuildPortableSummary(
        EvaluationRunOptions options,
        string artifactRoot,
        IReadOnlyList<ScenarioRunResult> results,
        string? fatalError)
    {
        object Artifact(ArtifactRecord artifact) => new
        {
            artifact.Id,
            artifact.Status,
            path = artifact.Path is null ? null : RelativeUnderRoot(artifactRoot, artifact.Path),
            artifact.MediaType,
            artifact.SizeBytes,
            artifact.Sha256,
            artifact.UnavailableReason,
        };
        object Score(EvaluationScore score) => new
        {
            score.ScenarioId,
            score.Kind,
            score.Status,
            score.Metrics,
            artifacts = score.Artifacts.Select(Artifact),
            artifactDirectory = score.ArtifactDirectory is null
                ? null : RelativeUnderRoot(artifactRoot, score.ArtifactDirectory),
        };
        return JsonSerializer.Serialize(new
        {
            documentKind = RunSummaryDocumentKind,
            schemaVersion = "1.0",
            corpus = Path.GetFileName(options.CorpusPath),
            subset = options.Subset,
            artifactRoot = ".",
            fatalError,
            engineBaselinePassed = results.Count(value => value.EngineBaseline.Status == "passed"),
            engineBaselineFailed = results.Count(value => value.EngineBaseline.Status == "failed"),
            modelPlanningScored = results.Count(value => value.ModelPlanning?.Status is "passed" or "failed"),
            modelPlanningIncomplete = results.Count(value => value.ModelPlanning?.Status == "incomplete"),
            results = results.Select(result => new
            {
                result.ScenarioId,
                engineBaseline = Score(result.EngineBaseline),
                modelPlanning = result.ModelPlanning is null ? null : Score(result.ModelPlanning),
            }),
        }, JsonOptions);
    }

    private static string BuildRunIndex(
        IReadOnlyList<ScenarioRunResult> results, string? fatalError)
    {
        var builder = new StringBuilder("<!doctype html><html><head><meta charset=\"utf-8\"><meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; style-src 'unsafe-inline'; base-uri 'none'\"><title>Legal evaluation run</title><style>body{font-family:sans-serif;max-width:70rem;margin:2rem auto}li{margin:.5rem}</style></head><body><h1>Legal evaluation run</h1>");
        if (fatalError is not null)
            builder.Append("<h2>Fatal error</h2><pre>")
                .Append(WebUtility.HtmlEncode(fatalError)).Append("</pre>");
        builder.Append("<ul>");
        foreach (var result in results)
        {
            var scenario = Uri.EscapeDataString(result.ScenarioId);
            builder.Append("<li><a href=\"").Append(scenario)
                .Append("/engine-baseline/index.html\">")
                .Append(WebUtility.HtmlEncode(result.ScenarioId)).Append(" engine baseline</a> — ")
                .Append(WebUtility.HtmlEncode(result.EngineBaseline.Status));
            if (result.ModelPlanning is not null)
                builder.Append("; <a href=\"").Append(scenario)
                    .Append("/model-planning/index.html\">model planning</a> — ")
                    .Append(WebUtility.HtmlEncode(result.ModelPlanning.Status));
            builder.Append("</li>");
        }
        return builder.Append("</ul><p><a href=\"run-summary.json\">run-summary.json</a></p></body></html>")
            .ToString();
    }

    private static string RelativeUnderRoot(string root, string path)
    {
        var fullRoot = Path.GetFullPath(root);
        var fullPath = Path.GetFullPath(path);
        var prefix = fullRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        if (!fullPath.StartsWith(prefix, PathComparison) && !string.Equals(fullPath, fullRoot, PathComparison))
            throw new ScenarioValidationException("artifact report path escapes the artifact root");
        return Path.GetRelativePath(fullRoot, fullPath).Replace(Path.DirectorySeparatorChar, '/');
    }

    private static void AtomicWriteText(string path, string value)
    {
        var temporary = path + ".stage-" + Guid.NewGuid().ToString("N");
        try
        {
            File.WriteAllText(temporary, value,
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            File.Move(temporary, path, overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
    }

    private static string KindName(ScoreKind kind) =>
        kind == ScoreKind.EngineBaseline ? "engine-baseline" : "model-planning";

    private static string ExceptionDetail(Exception exception, params string?[] roots)
    {
        var detail = $"{exception.GetType().Name}: {exception.Message}";
        foreach (var root in roots.Append(Directory.GetCurrentDirectory())
            .Where(value => !string.IsNullOrWhiteSpace(value)))
        {
            var fullRoot = Path.GetFullPath(root!).TrimEnd(
                Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            detail = detail.Replace(fullRoot + Path.DirectorySeparatorChar,
                string.Empty, PathComparison);
            detail = detail.Replace(fullRoot + Path.AltDirectorySeparatorChar,
                string.Empty, PathComparison);
            detail = detail.Replace(fullRoot, ".", PathComparison);
        }
        return detail;
    }

    private static StringComparison PathComparison =>
        OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

    private sealed record StagedExternalReport(string DestinationPath, string TemporaryPath);

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },
    };
}
