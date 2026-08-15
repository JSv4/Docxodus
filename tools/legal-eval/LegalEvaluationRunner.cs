// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

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
    string SummaryPath,
    IReadOnlyList<ScenarioRunResult> Results,
    string? FatalError);

/// <summary>
/// Single artifact-producing orchestration path used by both the CLI and xUnit. Scenario-level
/// failures are converted to evidence bundles and never prevent later scenarios or the portable
/// run summary from being published.
/// </summary>
public sealed class LegalEvaluationRunner
{
    private const long MaximumCandidateBytes = 64L * 1024 * 1024;
    private readonly Func<LegalScenario, LegalScenario>? _scenarioTransform;
    private readonly IEvaluationArtifactRenderer? _artifactRenderer;

    public LegalEvaluationRunner(
        Func<LegalScenario, LegalScenario>? scenarioTransform = null,
        IEvaluationArtifactRenderer? artifactRenderer = null)
    {
        _scenarioTransform = scenarioTransform;
        _artifactRenderer = artifactRenderer;
    }

    public EvaluationRunOutcome Run(
        EvaluationRunOptions options,
        TextWriter? output = null,
        TextWriter? error = null)
    {
        output ??= TextWriter.Null;
        error ??= TextWriter.Null;
        var artifactRoot = Path.GetFullPath(options.ArtifactRoot);
        Directory.CreateDirectory(artifactRoot);
        output.WriteLine($"Artifacts: {artifactRoot}");
        var results = new List<ScenarioRunResult>();
        string? fatalError = null;
        var exitCode = 0;
        var summaryPath = Path.Combine(artifactRoot, "run-summary.json");

        try
        {
            var corpus = ScenarioLoader.LoadCorpus(options.CorpusPath);
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
                            artifactRoot, baseline.Error ?? "scripted baseline failed", options.RenderMode,
                            artifactRenderer: _artifactRenderer);
                    }
                    else
                    {
                        engineScore = scorer.Score(scenario, baseline, baseline.Output,
                            ScoreKind.EngineBaseline, artifactRoot, options.RenderMode);
                    }
                }
                catch (Exception exception)
                {
                    baseline ??= EmergencyBaseline(scenario, exception);
                    engineScore = FailureScore(scenario, baseline, ScoreKind.EngineBaseline,
                        artifactRoot, ExceptionDetail(exception,
                            Path.GetDirectoryName(scenario.FilePath), artifactRoot), options.RenderMode,
                        artifactRenderer: _artifactRenderer);
                }

                EvaluationScore? planningScore = null;
                if (engineScore.Status == "passed" && options.CandidateDirectory is not null)
                {
                    var candidatePath = Path.Combine(
                        Path.GetFullPath(options.CandidateDirectory), scenario.Id + ".docx");
                    if (!File.Exists(candidatePath))
                    {
                        var reason = $"candidate file is absent: {scenario.Id}.docx in the requested candidate directory";
                        planningScore = FailureScore(scenario, baseline!, ScoreKind.ModelPlanning,
                            artifactRoot, reason, ArtifactRenderMode.Disabled, status: "incomplete",
                            candidate: null, artifactRenderer: _artifactRenderer);
                    }
                    else
                    {
                        try
                        {
                            var candidate = ReadCandidateBounded(candidatePath);
                            planningScore = scorer.Score(scenario, baseline!, candidate,
                                ScoreKind.ModelPlanning, artifactRoot, options.RenderMode);
                        }
                        catch (Exception exception)
                        {
                            byte[]? candidate = null;
                            try { candidate = ReadCandidateBounded(candidatePath); }
                            catch { }
                            planningScore = FailureScore(scenario, baseline!, ScoreKind.ModelPlanning,
                                artifactRoot, ExceptionDetail(exception,
                                    options.CandidateDirectory, artifactRoot), ArtifactRenderMode.Disabled,
                                candidate: candidate, artifactRenderer: _artifactRenderer);
                        }
                    }
                }

                results.Add(new ScenarioRunResult(scenario.Id, engineScore, planningScore));
                output.WriteLine($"{engineScore.Status.ToUpperInvariant()} {scenario.Id} engine-baseline"
                    + (planningScore is null ? string.Empty : $"; {planningScore.Status} model-planning")
                    + $"; artifacts={engineScore.ArtifactDirectory}");
                if (engineScore.Status != "passed"
                    || planningScore?.Status is "failed" or "incomplete")
                    exitCode = 1;
            }
        }
        catch (Exception exception)
        {
            fatalError = ExceptionDetail(exception,
                Path.GetDirectoryName(Path.GetFullPath(options.CorpusPath)), artifactRoot);
            error.WriteLine(exception.Message);
            exitCode = 2;
        }
        finally
        {
            try
            {
                var summaryJson = BuildPortableSummary(
                    options, artifactRoot, results, fatalError);
                AtomicWriteText(summaryPath, summaryJson);
                AtomicWriteText(Path.Combine(artifactRoot, "index.html"),
                    BuildRunIndex(results, fatalError));
                if (options.ReportPath is not null)
                {
                    var reportPath = Path.GetFullPath(options.ReportPath);
                    var reportDirectory = Path.GetDirectoryName(reportPath);
                    if (reportDirectory is not null) Directory.CreateDirectory(reportDirectory);
                    if (!string.Equals(reportPath, summaryPath, PathComparison))
                        AtomicWriteText(reportPath, summaryJson);
                }
                output.WriteLine($"Summary: {summaryPath}");
            }
            catch (Exception exception)
            {
                var publicationError = "run summary publication failed: "
                    + ExceptionDetail(exception, artifactRoot,
                        options.ReportPath is null
                            ? null
                            : Path.GetDirectoryName(Path.GetFullPath(options.ReportPath)));
                fatalError = fatalError is null
                    ? publicationError
                    : fatalError + Environment.NewLine + publicationError;
                error.WriteLine(exception.Message);
                exitCode = 2;
            }
        }

        return new EvaluationRunOutcome(exitCode, artifactRoot, summaryPath, results, fatalError);
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
        var targetSemanticDiff = inputSafetyError is null && expectedSafetyError is null
            ? EvaluationScorer.SemanticDiffArtifact(
                baseline.Input, baseline.Expected, "target-semantic-diff")
            : (Json: EvaluationScorer.ErrorEnvelope("target-semantic-diff", "failed",
                inputSafetyError ?? expectedSafetyError ?? "package safety validation failed"),
                Succeeded: false);
        var rendererSafe = inputSafetyError is null
            && candidateSafetyError is null
            && expectedSafetyError is null;
        var artifacts = ArtifactWriter.WriteIncomplete(directory, scenario.Id, kind, status,
            metrics, reason, baseline.OperationLog, baseline.Input, candidate, baseline.Expected,
            baseline.SemanticDiffJson,
            baseline.SemanticDiffSucceeded,
            targetSemanticDiff.Json,
            targetSemanticDiff.Succeeded,
            rendererSafe ? renderMode : ArtifactRenderMode.Disabled,
            allowExternalRenderer: rendererSafe && kind == ScoreKind.EngineBaseline,
            artifactRenderer: artifactRenderer,
            inputSafetyError: inputSafetyError,
            candidateSafetyError: candidateSafetyError,
            expectedSafetyError: expectedSafetyError);
        return new EvaluationScore(scenario.Id, kind, status, metrics, artifacts, directory);
    }

    private static BaselineExecution EmergencyBaseline(
        LegalScenario scenario, Exception exception)
    {
        byte[] input = Array.Empty<byte>();
        byte[] expected = Array.Empty<byte>();
        try { input = File.ReadAllBytes(scenario.Fixture.Path); } catch { }
        try { expected = File.ReadAllBytes(scenario.ExpectedDocument.Path); } catch { }
        var detail = ExceptionDetail(exception, Path.GetDirectoryName(scenario.FilePath));
        return new BaselineExecution(input, input, expected,
            EvaluationScorer.ErrorEnvelope("semantic-diff", "failed", detail),
            SemanticDiffSucceeded: false,
            new[] { $"baseline-initialization-failed:{exception.GetType().Name}:{exception.Message}" },
            Succeeded: false,
            Error: exception.ToString());
    }

    private static byte[] ReadCandidateBounded(string path)
    {
        var info = new FileInfo(path);
        if (info.Length > MaximumCandidateBytes)
            throw new ScenarioValidationException(
                $"candidate exceeds the {MaximumCandidateBytes}-byte read limit");
        return File.ReadAllBytes(path);
    }

    private static string? SafetyError(byte[] bytes, string label)
    {
        try
        {
            new InterimEvaluationPackageValidator().Validate(bytes, label);
            return null;
        }
        catch (Exception exception)
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

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },
    };
}
