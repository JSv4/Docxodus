// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Diagnostics;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Docxodus.Verification;
using SkiaSharp;

namespace LegalEval;

internal sealed record ArtifactPublication(
    string Status,
    IReadOnlyList<MetricResult> Metrics,
    IReadOnlyList<ArtifactRecord> Artifacts);

internal static class ArtifactWriter
{
    private const int MaximumRendererDiagnosticCharacters = 64 * 1024;
    private const int MaximumRenderedPages = 100;
    private const long MaximumRenderedFileBytes = 64L * 1024 * 1024;
    private const long MaximumRenderedAggregateBytes = 512L * 1024 * 1024;
    private const int MaximumRasterDimension = 16_384;
    private const long MaximumRasterPixels = 50_000_000;
    private const string DocxMediaType =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    public static string ResolveScoreDirectory(string artifactRoot, string scenarioId, ScoreKind kind)
    {
        if (scenarioId.Length == 0 || scenarioId.Any(character =>
                !(character is >= 'a' and <= 'z' or >= '0' and <= '9' or '-')))
            throw new ScenarioValidationException($"unsafe scenario artifact slug '{scenarioId}'");
        var root = Path.GetFullPath(artifactRoot);
        var directory = Path.GetFullPath(Path.Combine(root, scenarioId,
            kind == ScoreKind.EngineBaseline ? "engine-baseline" : "model-planning"));
        var rootPrefix = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        if (!directory.StartsWith(rootPrefix, PathComparison))
            throw new ScenarioValidationException("scenario artifact directory escapes the artifact root");
        return directory;
    }

    public static ArtifactPublication WriteIncomplete(
        string directory,
        string scenarioId,
        ScoreKind scoreKind,
        string status,
        IReadOnlyList<MetricResult> metrics,
        IReadOnlyList<ExpectedArtifact> expectedOutputs,
        string reason,
        IReadOnlyList<string>? operationLog,
        EvaluationEvidence evidence,
        ArtifactRenderMode renderMode = ArtifactRenderMode.Disabled,
        bool allowExternalRenderer = false,
        IEvaluationArtifactRenderer? artifactRenderer = null,
        string? inputSafetyError = null,
        string? candidateSafetyError = null,
        string? expectedSafetyError = null) =>
        Write(directory, scenarioId, scoreKind, status, metrics, expectedOutputs,
            operationLog ?? Array.Empty<string>(), evidence with
            {
                FailureReason = reason,
                InputSafetyError = inputSafetyError ?? evidence.InputSafetyError,
                CandidateSafetyError = candidateSafetyError ?? evidence.CandidateSafetyError,
                ExpectedSafetyError = expectedSafetyError ?? evidence.ExpectedSafetyError,
            }, renderMode, allowExternalRenderer, artifactRenderer);

    public static ArtifactPublication Write(
        string directory,
        string scenarioId,
        ScoreKind scoreKind,
        string status,
        IReadOnlyList<MetricResult> metrics,
        IReadOnlyList<ExpectedArtifact> expectedOutputs,
        IReadOnlyList<string> operationLog,
        EvaluationEvidence evidence,
        ArtifactRenderMode renderMode,
        bool allowExternalRenderer,
        IEvaluationArtifactRenderer? artifactRenderer = null)
    {
        return PublishFresh(directory, stage =>
        {
            var records = new List<ArtifactRecord>();
            var reason = evidence.FailureReason ?? "evidence was unavailable";
            records.Add(WriteOptionalDocx(stage, "input.docx", "input-docx",
                evidence.Input, reason, "source"));
            records.Add(WriteOptionalDocx(stage, "candidate.docx", "candidate-docx",
                evidence.Candidate, reason, "candidate"));
            records.Add(WriteOptionalDocx(stage, "expected.docx", "expected-docx",
                evidence.Expected, reason, "oracle"));
            AddSemanticRecord(stage, records, "semantic-change-set-v1.json",
                "semantic-change-set-v1", evidence.CandidateChanges, reason);
            AddSemanticRecord(stage, records, "target-semantic-change-set-v1.json",
                "target-semantic-change-set-v1", evidence.TargetChanges, reason);
            AddSemanticRecord(stage, records, "candidate-target-semantic-change-set-v1.json",
                "candidate-target-semantic-change-set-v1", evidence.CandidateTargetChanges, reason);
            records.Add(WritePreview(stage, "before.html", "before-html", evidence.Input,
                evidence.Input is null ? reason : evidence.InputSafetyError));
            records.Add(WritePreview(stage, "after.html", "after-html", evidence.Candidate,
                evidence.Candidate is null ? reason : evidence.CandidateSafetyError,
                evidence.CandidateHtml));
            records.Add(WritePreview(stage, "target.html", "target-html", evidence.Expected,
                evidence.Expected is null ? reason : evidence.ExpectedSafetyError));
            records.Add(evidence.RedlineBytes is null
                ? Unavailable("redline-docx", reason, DocxMediaType, "review")
                : WriteBytes(stage, "redline.docx", "redline-docx",
                    evidence.RedlineBytes, DocxMediaType, "review"));
            AddFoundationRecords(stage, records, evidence, reason);
            var renderMetric = AddRenderArtifacts(stage, records, renderMode, allowExternalRenderer,
                evidence.Input, evidence.Candidate, evidence.Expected, evidence.RedlineBytes,
                artifactRenderer);
            var finalMetrics = renderMetric is null ? metrics : metrics.Append(renderMetric).ToList();
            return FinalizeScore(stage, scenarioId, scoreKind, status, finalMetrics,
                expectedOutputs, operationLog, records);
        });
    }

    private static ArtifactPublication FinalizeScore(
        string directory,
        string scenarioId,
        ScoreKind scoreKind,
        string status,
        IReadOnlyList<MetricResult> metrics,
        IReadOnlyList<ExpectedArtifact> expectedOutputs,
        IReadOnlyList<string> operationLog,
        List<ArtifactRecord> records)
    {
        var finalMetrics = metrics.Append(RequiredOutputMetric(expectedOutputs, records)).ToList();
        var finalStatus = status == "incomplete"
            ? "incomplete"
            : finalMetrics.Any(value => value.Status == "failed") ? "failed" : "passed";
        records.Insert(0, WriteMetrics(directory, scenarioId, scoreKind, finalStatus, finalMetrics));
        records.Insert(1, WriteText(directory, "operation-log.json", "operation-log",
            JsonSerializer.Serialize(new
            {
                schemaVersion = "1.0",
                operations = operationLog,
            }, JsonOptions), "application/json"));
        records.Add(WriteText(directory, "summary.md", "scenario-summary",
            BuildSummary(scenarioId, scoreKind, finalStatus, finalMetrics), "text/markdown"));
        FinalizeIndexes(directory, scenarioId, scoreKind, finalStatus, finalMetrics,
            operationLog, records);
        return new ArtifactPublication(finalStatus, finalMetrics, records);
    }

    private static MetricResult RequiredOutputMetric(
        IReadOnlyList<ExpectedArtifact> expectedOutputs,
        IReadOnlyList<ArtifactRecord> artifacts)
    {
        var required = expectedOutputs.Where(value => value.Required).ToList();
        var missing = required.Where(output => !artifacts.Any(artifact =>
                artifact.Id == output.Id
                && artifact.Status == "available"
                && artifact.MediaType == output.MediaType
                && artifact.Role == output.Role))
            .Select(value => value.Id).Order(StringComparer.Ordinal).ToList();
        return new MetricResult("task-completion.required-artifacts", "task_completion",
            missing.Count == 0 ? "passed" : "failed",
            missing.Count == 0
                ? $"all required artifacts are available: {string.Join(", ", required.Select(value => value.Id))}"
                : $"required artifacts are absent or unavailable: {string.Join(", ", missing)}",
            missing.Count == 0 ? 1 : 0);
    }

    private static ArtifactRecord WriteMetrics(
        string directory,
        string scenarioId,
        ScoreKind scoreKind,
        string status,
        IReadOnlyList<MetricResult> metrics) =>
        WriteText(directory, "metrics.json", "metrics-json", JsonSerializer.Serialize(new
        {
            schemaVersion = "1.0",
            scenarioId,
            scoreKind = scoreKind.ToString(),
            status,
            metrics,
        }, JsonOptions), "application/json");

    private static void AddFoundationRecords(
        string directory,
        List<ArtifactRecord> records,
        EvaluationEvidence evidence,
        string reason)
    {
        AddManifestRecord(directory, records, "input-package-manifest-v1",
            "input-package-manifest-v1.json", evidence.InputManifest, reason);
        AddManifestRecord(directory, records, "candidate-package-manifest-v1",
            "candidate-package-manifest-v1.json", evidence.CandidateManifest, reason);
        AddManifestRecord(directory, records, "target-package-manifest-v1",
            "target-package-manifest-v1.json", evidence.ExpectedManifest, reason);

        records.Add(evidence.DeliverableVerification is null
            ? Unavailable("deliverable-verification-v1", reason, "application/json", "verification")
            : WriteBytes(directory, "deliverable-verification-v1.json",
                "deliverable-verification-v1",
                evidence.DeliverableVerification.ToCanonicalUtf8Bytes(),
                "application/json", "verification"));

        if (evidence.RedlineProofRun is null)
        {
            records.Add(Unavailable("redline-reversibility-proof-v1",
                reason, "application/json", "verification"));
            records.Add(Unavailable("redline-accepted-path-docx", reason, DocxMediaType, "proof-output"));
            records.Add(Unavailable("redline-rejected-path-docx", reason, DocxMediaType, "proof-output"));
        }
        else
        {
            records.Add(WriteText(directory, "redline-reversibility-proof-v1.json",
                "redline-reversibility-proof-v1",
                evidence.RedlineProofRun.Proof.ToCanonicalJson(),
                "application/json", "verification"));
            records.Add(evidence.RedlineProofRun.AcceptedPackageBytes is null
                ? Unavailable("redline-accepted-path-docx",
                    "#464 did not produce an accepted-path package", DocxMediaType, "proof-output")
                : WriteBytes(directory, "redline-accepted-path.docx",
                    "redline-accepted-path-docx",
                    evidence.RedlineProofRun.AcceptedPackageBytes, DocxMediaType, "proof-output"));
            records.Add(evidence.RedlineProofRun.RejectedPackageBytes is null
                ? Unavailable("redline-rejected-path-docx",
                    "#464 did not produce a rejected-path package", DocxMediaType, "proof-output")
                : WriteBytes(directory, "redline-rejected-path.docx",
                    "redline-rejected-path-docx",
                    evidence.RedlineProofRun.RejectedPackageBytes, DocxMediaType, "proof-output"));
        }

        if (evidence.DeliveryReceipt is null)
        {
            records.Add(Unavailable("delivery-change-receipt-v1",
                evidence.DeliveryReceiptUnavailableReason ?? reason,
                "application/json", "verification"));
        }
        else
        {
            records.Add(WriteBytes(directory, "delivery-change-receipt-v1.json",
                "delivery-change-receipt-v1", evidence.DeliveryReceipt.ToJsonBytes(),
                "application/json", "verification"));
            foreach (var artifact in evidence.DeliveryReceiptArtifacts
                .Where(value => value.Key.StartsWith("transaction-semantic-change-set-",
                    StringComparison.Ordinal)).OrderBy(value => value.Key, StringComparer.Ordinal))
            {
                records.Add(WriteBytes(directory, artifact.Key + ".json", artifact.Key,
                    artifact.Value, "application/json", "receipt-evidence"));
            }
        }
    }

    private static void AddManifestRecord(
        string directory,
        ICollection<ArtifactRecord> records,
        string id,
        string fileName,
        PackageManifest? manifest,
        string reason)
    {
        if (manifest is null)
        {
            records.Add(Unavailable(id, reason, "application/json", "verification"));
            return;
        }
        records.Add(WriteBytes(directory, fileName, id, manifest.ToJsonBytes(),
            "application/json", "verification"));
    }

    private static void AddSemanticRecord(
        string directory,
        ICollection<ArtifactRecord> records,
        string fileName,
        string id,
        SemanticChangeSet? changeSet,
        string reason)
    {
        if (changeSet is null)
        {
            var failed = WriteText(directory, fileName, id,
                EvaluationScorer.ErrorEnvelope(id, "failed", reason),
                "application/json", "verification");
            records.Add(failed with { Status = "failed", UnavailableReason = reason });
            return;
        }
        records.Add(WriteBytes(directory, fileName, id,
            changeSet.ToCanonicalUtf8Bytes(), "application/json", "verification"));
    }

    private static MetricResult? AddRenderArtifacts(
        string directory,
        List<ArtifactRecord> records,
        ArtifactRenderMode renderMode,
        bool allowExternalRenderer,
        byte[]? input,
        byte[]? candidate,
        byte[]? expected,
        byte[]? redline,
        IEvaluationArtifactRenderer? artifactRenderer)
    {
        if (renderMode == ArtifactRenderMode.Disabled)
        {
            AddUnavailableRenderSet(records, "external document rendering was not requested");
            return null;
        }
        if (!allowExternalRenderer)
        {
            AddUnavailableRenderSet(records,
                "external rendering is disabled for untrusted model candidates; sanitized HTML remains available");
            return new MetricResult("rendering-regression.visual-layout",
                "rendering_regression", "failed",
                "visual-layout scoring was requested but external rendering is disabled for untrusted candidates",
                0);
        }

        artifactRenderer ??= new ExternalArtifactRenderer();
        long renderedBytes = 0;

        var documents = new[]
        {
            (Prefix: "before", PdfId: "input-pdf", SourceName: "input.docx", Bytes: input),
            (Prefix: "candidate", PdfId: "candidate-pdf", SourceName: "candidate.docx", Bytes: candidate),
            (Prefix: "target", PdfId: "target-pdf", SourceName: "expected.docx", Bytes: expected),
            (Prefix: "redline", PdfId: "redline-pdf", SourceName: "redline.docx", Bytes: redline),
        };
        var pages = new Dictionary<string, IReadOnlyList<string>>(StringComparer.Ordinal);
        foreach (var document in documents)
        {
            if (document.Bytes is null)
            {
                records.Add(Unavailable(document.PdfId, $"{document.Prefix} DOCX is unavailable",
                    "application/pdf", "review"));
                records.Add(Unavailable($"{document.Prefix}-visual",
                    $"{document.Prefix} DOCX is unavailable", "image/png", "review"));
                continue;
            }
            var sourcePath = Path.Combine(directory, document.SourceName);
            var rendered = artifactRenderer.RenderDocument(
                directory, sourcePath, document.Prefix);
            var pdfPath = rendered.PdfPath is null
                ? null
                : RequireRenderedPath(directory, rendered.PdfPath);
            if (pdfPath is null || !File.Exists(pdfPath))
            {
                records.Add(Unavailable(document.PdfId,
                    rendered.Error ?? "renderer produced no PDF", "application/pdf", "review"));
                records.Add(Unavailable($"{document.Prefix}-visual",
                    "PDF was unavailable, so no raster pages were generated", "image/png", "review"));
                continue;
            }
            if (!TryIncludeRenderedFiles(
                    new[] { pdfPath }, ref renderedBytes, out var pdfLimitError))
            {
                records.Add(Unavailable(document.PdfId, pdfLimitError,
                    "application/pdf", "review"));
                records.Add(Unavailable($"{document.Prefix}-visual",
                    pdfLimitError, "image/png", "review"));
                continue;
            }
            records.Add(Record(document.PdfId, pdfPath, "application/pdf", "review"));
            var pagePaths = rendered.PagePaths
                .Select(value => RequireRenderedPath(directory, value)).ToList();
            if (pagePaths.Count > MaximumRenderedPages)
            {
                records.Add(Unavailable($"{document.Prefix}-visual",
                    $"renderer exceeded the {MaximumRenderedPages}-page evidence limit",
                    "image/png", "review"));
                continue;
            }
            if (rendered.Error is not null || pagePaths.Count == 0)
            {
                records.Add(Unavailable($"{document.Prefix}-visual",
                    rendered.Error ?? "renderer produced no pages", "image/png", "review"));
                continue;
            }
            if (!TryIncludeRenderedFiles(
                    pagePaths, ref renderedBytes, out var pageLimitError))
            {
                records.Add(Unavailable($"{document.Prefix}-visual",
                    pageLimitError, "image/png", "review"));
                continue;
            }
            pages[document.Prefix] = pagePaths;
            for (var index = 0; index < pagePaths.Count; index++)
            {
                var id = $"{document.Prefix}-visual-page-{index + 1:D3}";
                records.Add(Record(id, pagePaths[index], "image/png", "review"));
            }
            records.Add(Record($"{document.Prefix}-visual", pagePaths[0], "image/png", "review"));
        }

        if (!pages.TryGetValue("candidate", out var candidatePages)
            || !pages.TryGetValue("target", out var targetPages))
        {
            records.Add(Unavailable("candidate-target-visual-diff",
                "candidate and target raster pages are both required for visual diffs",
                "image/png", "review"));
            return new MetricResult("rendering-regression.visual-layout",
                "rendering_regression", "failed",
                "candidate and target raster pages were not both available", 0);
        }
        var diff = artifactRenderer.ComparePages(directory, targetPages, candidatePages);
        var diffPaths = diff.Paths
            .Select(value => RequireRenderedPath(directory, value)).ToList();
        if (diffPaths.Count > MaximumRenderedPages)
        {
            records.Add(Unavailable("candidate-target-visual-diff",
                $"visual comparator exceeded the {MaximumRenderedPages}-page evidence limit",
                "image/png", "review"));
            return new MetricResult("rendering-regression.visual-layout",
                "rendering_regression", "failed",
                $"visual comparator exceeded the {MaximumRenderedPages}-page evidence limit", 0);
        }
        if (diff.Error is not null || diffPaths.Count == 0)
        {
            records.Add(Unavailable("candidate-target-visual-diff",
                diff.Error ?? "visual comparator produced no output", "image/png", "review"));
            return new MetricResult("rendering-regression.visual-layout",
                "rendering_regression", "failed",
                diff.Error ?? "visual comparator produced no output", 0);
        }
        if (!TryIncludeRenderedFiles(diffPaths, ref renderedBytes, out var diffLimitError))
        {
            records.Add(Unavailable("candidate-target-visual-diff",
                diffLimitError, "image/png", "review"));
            return new MetricResult("rendering-regression.visual-layout",
                "rendering_regression", "failed", diffLimitError, 0);
        }
        for (var index = 0; index < diffPaths.Count; index++)
            records.Add(Record($"candidate-target-visual-diff-page-{index + 1:D3}",
                diffPaths[index], "image/png", "review"));
        records.Add(Record("candidate-target-visual-diff", diffPaths[0], "image/png", "review"));
        var equivalent = diff.DifferentPixels == 0;
        return new MetricResult("rendering-regression.visual-layout",
            "rendering_regression", equivalent ? "passed" : "failed",
            diff.DifferentPixels is null
                ? "visual comparator did not report an exact pixel-difference count"
                : $"candidate/target raster difference pixels={diff.DifferentPixels}",
            equivalent ? 1 : 0);
    }

    private static void AddUnavailableRenderSet(List<ArtifactRecord> records, string reason)
    {
        foreach (var prefix in new[] { "input", "candidate", "target", "redline" })
        {
            records.Add(Unavailable($"{prefix}-pdf", reason, "application/pdf", "review"));
            records.Add(Unavailable($"{(prefix == "input" ? "before" : prefix)}-visual",
                reason, "image/png", "review"));
        }
        records.Add(Unavailable("candidate-target-visual-diff", reason, "image/png", "review"));
    }

    private static bool TryIncludeRenderedFiles(
        IReadOnlyList<string> paths,
        ref long aggregateBytes,
        out string error)
    {
        foreach (var path in paths)
        {
            var info = new FileInfo(path);
            if (!info.Exists)
            {
                error = $"renderer output is absent: {Path.GetFileName(path)}";
                return false;
            }
            if (info.Length > MaximumRenderedFileBytes)
            {
                error = $"renderer output '{Path.GetFileName(path)}' exceeds the "
                    + $"{MaximumRenderedFileBytes}-byte per-file limit";
                return false;
            }
            if (info.Length > MaximumRenderedAggregateBytes - aggregateBytes)
            {
                error = $"renderer output exceeds the "
                    + $"{MaximumRenderedAggregateBytes}-byte aggregate limit";
                return false;
            }
            aggregateBytes += info.Length;
        }
        error = string.Empty;
        return true;
    }

    private static (string? Path, string? Diagnostics) TryLibreOfficePdf(
        string directory, string sourceDocxPath, string outputPrefix)
    {
        var executable = FindExecutable("libreoffice") ?? FindExecutable("soffice");
        if (executable is null)
            return (null, "LibreOffice/soffice is not installed");
        string? profileRoot = null;
        try
        {
            profileRoot = Path.Combine(Path.GetTempPath(), $"docxodus-legal-eval-lo-{Guid.NewGuid():N}");
            var loHome = Path.Combine(profileRoot, "home");
            var runtime = Path.Combine(profileRoot, "runtime");
            Directory.CreateDirectory(loHome);
            Directory.CreateDirectory(runtime);
            if (!OperatingSystem.IsWindows())
            {
                File.SetUnixFileMode(loHome,
                    UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
                File.SetUnixFileMode(runtime,
                    UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
            }
            var generatedPdf = Path.Combine(directory,
                Path.GetFileNameWithoutExtension(sourceDocxPath) + ".pdf");
            var pdf = Path.Combine(directory, outputPrefix + ".pdf");
            var startInfo = new ProcessStartInfo
            {
                FileName = executable,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
            };
            startInfo.ArgumentList.Add($"-env:UserInstallation={new Uri(loHome + Path.DirectorySeparatorChar).AbsoluteUri}");
            startInfo.ArgumentList.Add("--headless");
            startInfo.ArgumentList.Add("--convert-to");
            startInfo.ArgumentList.Add("pdf:writer_pdf_Export");
            startInfo.ArgumentList.Add("--outdir");
            startInfo.ArgumentList.Add(directory);
            startInfo.ArgumentList.Add(sourceDocxPath);
            startInfo.Environment["HOME"] = loHome;
            startInfo.Environment["XDG_RUNTIME_DIR"] = runtime;
            startInfo.Environment["XDG_CONFIG_HOME"] = Path.Combine(loHome, ".config");
            startInfo.Environment["XDG_CACHE_HOME"] = Path.Combine(loHome, ".cache");
            startInfo.Environment["SAL_USE_VCLPLUGIN"] = "svp";
            var result = RunProcess(startInfo, TimeSpan.FromSeconds(45));
            if (result.ExitCode != 0 || !File.Exists(generatedPdf))
                return (null, result.Diagnostics);
            if (!string.Equals(generatedPdf, pdf, PathComparison))
                File.Move(generatedPdf, pdf, overwrite: true);
            return (pdf, null);
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            return (null, exception.Message);
        }
        finally
        {
            if (profileRoot is not null && Directory.Exists(profileRoot))
            {
                try { Directory.Delete(profileRoot, recursive: true); }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }
            }
        }
    }

    private static (IReadOnlyList<string> Paths, string? Error) TryRasterAllPages(
        string directory, string prefix)
    {
        var executable = FindExecutable("pdftoppm");
        if (executable is null) return (Array.Empty<string>(), "pdftoppm is not installed");
        try
        {
            var outputPrefix = Path.Combine(directory, prefix + "-page");
            var startInfo = new ProcessStartInfo
            {
                FileName = executable,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
            };
            foreach (var argument in new[]
            {
                "-png", "-r", "144", Path.Combine(directory, prefix + ".pdf"), outputPrefix,
            })
                startInfo.ArgumentList.Add(argument);
            var result = RunProcess(startInfo, TimeSpan.FromSeconds(45));
            var paths = Directory.EnumerateFiles(directory, prefix + "-page-*.png")
                .Order(StringComparer.Ordinal).Take(MaximumRenderedPages + 1).ToList();
            if (paths.Count > MaximumRenderedPages)
                return (Array.Empty<string>(),
                    $"renderer exceeded the {MaximumRenderedPages}-page evidence limit");
            return result.ExitCode == 0 && paths.Count != 0
                ? (paths, null)
                : (Array.Empty<string>(), result.Diagnostics);
        }
        catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
        {
            return (Array.Empty<string>(), exception.Message);
        }
    }

    private static (IReadOnlyList<string> Paths, string? Error, long? DifferentPixels) TryVisualDiffs(
        string directory,
        IReadOnlyList<string> targetPages,
        IReadOnlyList<string> candidatePages)
    {
        var executable = FindExecutable("compare");
        if (executable is null)
            return TryVisualDiffsInProcess(directory, targetPages, candidatePages);
        if (targetPages.Count != candidatePages.Count)
            return (Array.Empty<string>(),
                $"page count differs: target={targetPages.Count}, candidate={candidatePages.Count}", null);
        var paths = new List<string>(targetPages.Count);
        long differentPixels = 0;
        for (var index = 0; index < targetPages.Count; index++)
        {
            var output = Path.Combine(directory, $"candidate-target-diff-page-{index + 1:D3}.png");
            var startInfo = new ProcessStartInfo
            {
                FileName = executable,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                UseShellExecute = false,
            };
            foreach (var argument in new[]
            {
                "-metric", "AE", targetPages[index], candidatePages[index], output,
            })
                startInfo.ArgumentList.Add(argument);
            var result = RunProcess(startInfo, TimeSpan.FromSeconds(30));
            // ImageMagick compare returns 1 when pixels differ and still writes the diff.
            if (result.ExitCode is not (0 or 1) || !File.Exists(output))
                return (Array.Empty<string>(), result.Diagnostics, null);
            var token = result.Diagnostics.Split(
                (char[]?)null, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
            if (!long.TryParse(token, System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture, out var pageDifference)
                || pageDifference < 0)
                return (Array.Empty<string>(),
                    $"ImageMagick compare returned an invalid AE metric: {result.Diagnostics}", null);
            differentPixels = checked(differentPixels + pageDifference);
            paths.Add(output);
        }
        return (paths, null, differentPixels);
    }

    private static (IReadOnlyList<string> Paths, string? Error, long? DifferentPixels)
        TryVisualDiffsInProcess(
            string directory,
            IReadOnlyList<string> targetPages,
            IReadOnlyList<string> candidatePages)
    {
        if (targetPages.Count != candidatePages.Count)
            return (Array.Empty<string>(),
                $"page count differs: target={targetPages.Count}, candidate={candidatePages.Count}", null);

        var paths = new List<string>(targetPages.Count);
        long differentPixels = 0;
        for (var index = 0; index < targetPages.Count; index++)
        {
            try
            {
                var targetDimensions = ReadPngDimensions(targetPages[index]);
                var candidateDimensions = ReadPngDimensions(candidatePages[index]);
                if (targetDimensions != candidateDimensions)
                    return (Array.Empty<string>(),
                        $"page {index + 1} dimensions differ: target={targetDimensions.Width}x{targetDimensions.Height}, "
                        + $"candidate={candidateDimensions.Width}x{candidateDimensions.Height}", null);

                using var target = SKBitmap.Decode(targetPages[index]);
                using var candidate = SKBitmap.Decode(candidatePages[index]);
                if (target is null || candidate is null)
                    return (Array.Empty<string>(),
                        $"page {index + 1} could not be decoded as PNG", null);
                using var visual = new SKBitmap(target.Width, target.Height,
                    SKColorType.Rgba8888, SKAlphaType.Premul);
                for (var y = 0; y < target.Height; y++)
                {
                    for (var x = 0; x < target.Width; x++)
                    {
                        var changed = target.GetPixel(x, y) != candidate.GetPixel(x, y);
                        if (changed) differentPixels = checked(differentPixels + 1);
                        visual.SetPixel(x, y, changed ? SKColors.Red : SKColors.Transparent);
                    }
                }

                var output = Path.Combine(directory,
                    $"candidate-target-diff-page-{index + 1:D3}.png");
                using var image = SKImage.FromBitmap(visual);
                using var encoded = image.Encode(SKEncodedImageFormat.Png, 100)
                    ?? throw new InvalidOperationException("SkiaSharp did not encode the visual diff");
                using var stream = new FileStream(output, FileMode.Create, FileAccess.Write,
                    FileShare.None, bufferSize: 81920, FileOptions.SequentialScan);
                encoded.SaveTo(stream);
                paths.Add(output);
            }
            catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
            {
                return (Array.Empty<string>(),
                    $"in-process visual comparison failed for page {index + 1}: {exception.Message}",
                    null);
            }
        }
        return (paths, null, differentPixels);
    }

    private static (int Width, int Height) ReadPngDimensions(string path)
    {
        Span<byte> header = stackalloc byte[24];
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
            bufferSize: 24, FileOptions.SequentialScan);
        stream.ReadExactly(header);
        ReadOnlySpan<byte> signature = stackalloc byte[]
            { 137, 80, 78, 71, 13, 10, 26, 10 };
        if (!header[..8].SequenceEqual(signature)
            || !header.Slice(12, 4).SequenceEqual("IHDR"u8))
            throw new InvalidDataException($"renderer output is not a PNG: {Path.GetFileName(path)}");
        var width = System.Buffers.Binary.BinaryPrimitives.ReadInt32BigEndian(header.Slice(16, 4));
        var height = System.Buffers.Binary.BinaryPrimitives.ReadInt32BigEndian(header.Slice(20, 4));
        if (width <= 0 || height <= 0
            || width > MaximumRasterDimension || height > MaximumRasterDimension
            || (long)width * height > MaximumRasterPixels)
            throw new InvalidDataException(
                $"renderer PNG dimensions exceed the {MaximumRasterDimension}x{MaximumRasterDimension} "
                + $"and {MaximumRasterPixels}-pixel limits: {width}x{height}");
        return (width, height);
    }

    private sealed class ExternalArtifactRenderer : IEvaluationArtifactRenderer
    {
        public RenderedDocumentEvidence RenderDocument(
            string artifactDirectory,
            string sourceDocxPath,
            string outputPrefix)
        {
            var pdf = TryLibreOfficePdf(artifactDirectory, sourceDocxPath, outputPrefix);
            if (pdf.Path is null)
                return new RenderedDocumentEvidence(null, Array.Empty<string>(), pdf.Diagnostics);
            var pages = TryRasterAllPages(artifactDirectory, outputPrefix);
            return new RenderedDocumentEvidence(pdf.Path, pages.Paths, pages.Error);
        }

        public RenderedVisualDiffEvidence ComparePages(
            string artifactDirectory,
            IReadOnlyList<string> targetPages,
            IReadOnlyList<string> candidatePages)
        {
            var result = TryVisualDiffs(artifactDirectory, targetPages, candidatePages);
            return new RenderedVisualDiffEvidence(
                result.Paths, result.Error, result.DifferentPixels);
        }
    }

    private static (int ExitCode, string Diagnostics) RunProcess(
        ProcessStartInfo startInfo, TimeSpan timeout)
    {
        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException($"{startInfo.FileName} did not start");
        var outputTask = ReadBoundedOutputAsync(process.StandardOutput);
        var errorTask = ReadBoundedOutputAsync(process.StandardError);
        using var cancellation = new CancellationTokenSource(timeout);
        try
        {
            process.WaitForExitAsync(cancellation.Token).GetAwaiter().GetResult();
        }
        catch (OperationCanceledException)
        {
            process.Kill(entireProcessTree: true);
            return (-1, $"{Path.GetFileName(startInfo.FileName)} timed out after {timeout.TotalSeconds:0} seconds");
        }
        var diagnostics = string.Join(" ", new[]
        {
            errorTask.GetAwaiter().GetResult().Trim(),
            outputTask.GetAwaiter().GetResult().Trim(),
        }.Where(value => value.Length != 0));
        return (process.ExitCode,
            diagnostics.Length == 0
                ? $"{Path.GetFileName(startInfo.FileName)} exit {process.ExitCode}"
                : diagnostics);
    }

    private static async Task<string> ReadBoundedOutputAsync(StreamReader reader)
    {
        var buffer = new char[4096];
        var value = new StringBuilder();
        var truncated = false;
        while (await reader.ReadAsync(buffer.AsMemory()).ConfigureAwait(false) is var count
               && count != 0)
        {
            var retained = Math.Min(count,
                Math.Max(0, MaximumRendererDiagnosticCharacters - value.Length));
            if (retained > 0) value.Append(buffer, 0, retained);
            if (retained != count) truncated = true;
        }
        if (truncated) value.Append(" [diagnostics truncated]");
        return value.ToString();
    }

    private static void FinalizeIndexes(
        string directory,
        string scenarioId,
        ScoreKind scoreKind,
        string status,
        IReadOnlyList<MetricResult> metrics,
        IReadOnlyList<string> operationLog,
        List<ArtifactRecord> records)
    {
        records.RemoveAll(value => value.Id is "artifact-index-markdown" or "artifact-index-html"
            or "evaluation-bundle-manifest-v2" or "artifact-status");
        // The bundle manifest covers content evidence, status covers content plus the manifest, and the
        // indexes link both. Excluding the indexes from the hashed documents avoids a digest cycle.
        records.Add(WriteEvaluationBundleManifest(directory, scenarioId, scoreKind, status,
            metrics, operationLog, records));
        records.Add(WriteArtifactStatus(directory, records));
        var indexedRecords = records.ToList();
        records.Add(WriteText(directory, "index.md", "artifact-index-markdown",
            BuildArtifactIndexMarkdown(scenarioId, status, indexedRecords), "text/markdown"));
        records.Add(WriteText(directory, "index.html", "artifact-index-html",
            BuildArtifactIndexHtml(scenarioId, status, indexedRecords), "text/html"));
    }

    private static ArtifactRecord WriteEvaluationBundleManifest(
        string directory,
        string scenarioId,
        ScoreKind scoreKind,
        string status,
        IReadOnlyList<MetricResult> metrics,
        IReadOnlyList<string> operationLog,
        IReadOnlyList<ArtifactRecord> records)
    {
        var fingerprintComponents = new List<string>
        {
            scenarioId,
            scoreKind.ToString(),
            status,
        };
        fingerprintComponents.AddRange(operationLog);
        fingerprintComponents.AddRange(records.OrderBy(value => value.Id, StringComparer.Ordinal)
            .Select(value => $"{value.Id}:{value.Status}:{value.Sha256}:{value.SizeBytes}"));
        var fingerprintInput = string.Join("\n", fingerprintComponents);
        var runId = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(fingerprintInput)))
            .ToLowerInvariant();
        return WriteText(directory, "evaluation-bundle-manifest-v2.json",
            "evaluation-bundle-manifest-v2",
            JsonSerializer.Serialize(new
            {
                schemaVersion = "docxodus.evaluation-bundle-manifest/2.0",
                bundleKind = "legal-workflow-evaluation",
                runId,
                contentScope = "content artifacts only; excludes this bundle manifest, artifact status, and indexes to avoid digest cycles",
                rendererOutputSemantics = records.Any(value => value.Status == "available" && IsRenderedArtifact(value.Id))
                    ? "renderer outputs are content-addressed for this run; the runId is reproducible only when the renderer and its output are reproducible"
                    : "no external renderer output is included in this receipt",
                scenarioId,
                scoreKind,
                status,
                operations = operationLog,
                metricStatus = metrics.Select(value => new { value.Id, value.Status }),
                artifacts = records.OrderBy(value => value.Id, StringComparer.Ordinal).Select(value => new
                {
                    value.Id,
                    value.Status,
                    path = RelativePath(directory, value.Path),
                    value.MediaType,
                    value.Role,
                    value.SizeBytes,
                    value.Sha256,
                    value.UnavailableReason,
                }),
            }, JsonOptions), "application/json");
    }

    private static ArtifactRecord WriteArtifactStatus(
        string directory, IReadOnlyList<ArtifactRecord> records) =>
        WriteText(directory, "artifact-status.json", "artifact-status", JsonSerializer.Serialize(new
        {
            schemaVersion = "1.0",
            artifacts = records.Select(record => new
            {
                record.Id,
                record.Status,
                path = RelativePath(directory, record.Path),
                record.MediaType,
                record.Role,
                record.SizeBytes,
                record.Sha256,
                record.UnavailableReason,
            }),
        }, JsonOptions), "application/json");

    private static string BuildArtifactIndexMarkdown(
        string scenarioId, string status, IReadOnlyList<ArtifactRecord> records)
    {
        var builder = new StringBuilder();
        builder.AppendLine($"# Evaluation artifacts: {scenarioId}").AppendLine()
            .AppendLine($"Status: `{status}`").AppendLine()
            .AppendLine("| Artifact | Status | View | Media type | Bytes | SHA-256 / reason |")
            .AppendLine("| --- | --- | --- | --- | ---: | --- |");
        foreach (var record in records.OrderBy(value => value.Id, StringComparer.Ordinal))
        {
            var relative = RelativePath(string.Empty, record.Path);
            var link = relative is null ? "—" : $"[open]({Uri.EscapeDataString(relative).Replace("%2F", "/", StringComparison.OrdinalIgnoreCase)})";
            builder.Append("| ").Append(EscapeCell(record.Id)).Append(" | ")
                .Append(EscapeCell(record.Status)).Append(" | ").Append(link).Append(" | ")
                .Append(EscapeCell(record.MediaType)).Append(" | ")
                .Append(record.SizeBytes?.ToString() ?? "—").Append(" | ")
                .Append(EscapeCell(record.Sha256 ?? record.UnavailableReason ?? "—")).AppendLine(" |");
        }
        return builder.ToString();
    }

    private static string BuildArtifactIndexHtml(
        string scenarioId, string status, IReadOnlyList<ArtifactRecord> records)
    {
        var builder = new StringBuilder();
        builder.Append("<!doctype html><html><head><meta charset=\"utf-8\">")
            .Append("<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; style-src 'unsafe-inline'; base-uri 'none'\">")
            .Append("<title>Evaluation artifacts</title><style>body{font-family:sans-serif;max-width:80rem;margin:2rem auto}table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;padding:.4rem;text-align:left}code{overflow-wrap:anywhere}</style></head><body>")
            .Append("<h1>Evaluation artifacts: ").Append(WebUtility.HtmlEncode(scenarioId))
            .Append("</h1><p>Status: <strong>").Append(WebUtility.HtmlEncode(status))
            .Append("</strong></p><table><thead><tr><th>Artifact</th><th>Status</th><th>View</th><th>Media type</th><th>Bytes</th><th>SHA-256 / reason</th></tr></thead><tbody>");
        foreach (var record in records.OrderBy(value => value.Id, StringComparer.Ordinal))
        {
            var relative = RelativePath(string.Empty, record.Path);
            builder.Append("<tr><td><code>").Append(WebUtility.HtmlEncode(record.Id))
                .Append("</code></td><td>").Append(WebUtility.HtmlEncode(record.Status)).Append("</td><td>");
            if (relative is not null)
                builder.Append("<a href=\"").Append(WebUtility.HtmlEncode(relative)).Append("\">open</a>");
            else builder.Append("&mdash;");
            builder.Append("</td><td>").Append(WebUtility.HtmlEncode(record.MediaType))
                .Append("</td><td>").Append(record.SizeBytes?.ToString() ?? "&mdash;")
                .Append("</td><td><code>").Append(WebUtility.HtmlEncode(
                    record.Sha256 ?? record.UnavailableReason ?? "—"))
                .Append("</code></td></tr>");
        }
        return builder.Append("</tbody></table></body></html>").ToString();
    }

    private static string BuildSummary(
        string scenarioId,
        ScoreKind scoreKind,
        string status,
        IReadOnlyList<MetricResult> metrics)
    {
        var builder = new StringBuilder();
        builder.AppendLine($"# Legal workflow evaluation: {scenarioId}").AppendLine()
            .AppendLine($"- Score kind: `{scoreKind}`")
            .AppendLine($"- Status: `{status}`")
            .AppendLine($"- Applicable metrics passed: {metrics.Count(value => value.Status == "passed")}/"
                + $"{metrics.Count(value => value.Status != "not_applicable")}")
            .AppendLine($"- Not applicable: {metrics.Count(value => value.Status == "not_applicable")}")
            .AppendLine("- Artifact browser: [index.html](index.html)")
            .AppendLine().AppendLine("| Metric | Category | Status | Detail |")
            .AppendLine("| --- | --- | --- | --- |");
        foreach (var metric in metrics)
            builder.Append("| ").Append(EscapeCell(metric.Id)).Append(" | ")
                .Append(EscapeCell(metric.Category)).Append(" | ")
                .Append(EscapeCell(metric.Status)).Append(" | ")
                .Append(EscapeCell(metric.Detail)).AppendLine(" |");
        return builder.ToString();
    }

    private static ArtifactRecord WritePreview(
        string directory,
        string fileName,
        string id,
        byte[]? bytes,
        string? failureReason,
        string? cached = null)
    {
        if (bytes is not null && failureReason is null)
        {
            try
            {
                return WriteText(directory, fileName, id,
                    cached ?? EvaluationScorer.RenderHtml(bytes), "text/html", "review");
            }
            catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
            {
                failureReason = $"Docxodus HTML conversion failed: {exception.Message}";
            }
        }
        var reason = failureReason ?? "document bytes were unavailable";
        return WriteText(directory, fileName, id, DiagnosticHtml(id, reason),
            "text/html", "review") with
        {
            Status = "failed",
            UnavailableReason = reason,
        };
    }

    private static string DiagnosticHtml(string artifact, string reason) =>
        "<!doctype html><html><head><meta charset=\"utf-8\"><meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; style-src 'unsafe-inline'\"><title>Unavailable preview</title></head>"
        + "<body><h1>Unavailable preview</h1><p><code>" + WebUtility.HtmlEncode(artifact)
        + "</code></p><pre>" + WebUtility.HtmlEncode(reason) + "</pre></body></html>";

    private static ArtifactRecord WriteOptionalDocx(
        string directory, string fileName, string id, byte[]? bytes, string reason,
        string role = "evidence") =>
        bytes is null ? Unavailable(id, reason, DocxMediaType, role)
            : WriteBytes(directory, fileName, id, bytes, DocxMediaType, role);

    private static ArtifactRecord WriteBytes(
        string directory, string fileName, string id, byte[] bytes, string mediaType,
        string role = "evidence")
    {
        var path = Path.Combine(directory, fileName);
        File.WriteAllBytes(path, bytes);
        return Record(id, path, mediaType, role);
    }

    private static ArtifactRecord WriteText(
        string directory, string fileName, string id, string value, string mediaType,
        string role = "evidence")
    {
        var path = Path.Combine(directory, fileName);
        File.WriteAllText(path, value, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        return Record(id, path, mediaType, role);
    }

    private static ArtifactRecord Record(
        string id, string path, string mediaType, string role = "evidence")
    {
        var info = new FileInfo(path);
        return new ArtifactRecord(id, "available", path,
            Sha256File(path), null, mediaType, info.Length, role);
    }

    private static ArtifactRecord Unavailable(
        string id, string reason, string mediaType, string role = "evidence") =>
        new(id, "unavailable", null, null, reason, mediaType, null, role);

    private static string Sha256File(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
            bufferSize: 81920, FileOptions.SequentialScan);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    private static ArtifactPublication PublishFresh(
        string directory, Func<string, ArtifactPublication> build)
    {
        directory = Path.GetFullPath(directory);
        var parent = Path.GetDirectoryName(directory)
            ?? throw new InvalidOperationException("artifact directory has no parent");
        Directory.CreateDirectory(parent);
        var stage = Path.Combine(parent, "." + Path.GetFileName(directory) + ".stage-" + Guid.NewGuid().ToString("N"));
        var backup = Path.Combine(parent, "." + Path.GetFileName(directory) + ".backup-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(stage);
        try
        {
            var stagedPublication = build(stage);
            var hadExisting = Directory.Exists(directory);
            if (hadExisting) Directory.Move(directory, backup);
            try
            {
                Directory.Move(stage, directory);
            }
            catch
            {
                if (hadExisting && Directory.Exists(backup) && !Directory.Exists(directory))
                    Directory.Move(backup, directory);
                throw;
            }
            if (Directory.Exists(backup)) Directory.Delete(backup, recursive: true);
            return stagedPublication with
            {
                Artifacts = stagedPublication.Artifacts.Select(record => record.Path is null
                        ? record
                        : record with
                        {
                            Path = Path.Combine(directory, Path.GetRelativePath(stage, record.Path)),
                        })
                    .ToList(),
            };
        }
        finally
        {
            if (Directory.Exists(stage)) Directory.Delete(stage, recursive: true);
        }
    }

    private static string? RelativePath(string directory, string? path)
    {
        if (path is null) return null;
        if (directory.Length == 0) return Path.GetFileName(path);
        return Path.GetRelativePath(directory, path).Replace(Path.DirectorySeparatorChar, '/');
    }

    private static string EscapeCell(string value) =>
        value.Replace("|", "\\|", StringComparison.Ordinal)
            .Replace("\r\n", "<br>", StringComparison.Ordinal)
            .Replace("\n", "<br>", StringComparison.Ordinal);

    private static bool IsRenderedArtifact(string id) =>
        id.EndsWith("-pdf", StringComparison.Ordinal)
        || id.EndsWith("-visual", StringComparison.Ordinal)
        || id.Contains("-page-", StringComparison.Ordinal)
        || id == "candidate-target-visual-diff";

    private static string? FindExecutable(string name)
    {
        var path = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrEmpty(path)) return null;
        foreach (var directory in path.Split(Path.PathSeparator))
        {
            var candidate = Path.Combine(directory, name);
            if (File.Exists(candidate)) return candidate;
        }
        return null;
    }

    private static string RequireRenderedPath(string directory, string path)
    {
        var root = Path.GetFullPath(directory).TrimEnd(
            Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        var fullPath = Path.GetFullPath(path);
        if (!fullPath.StartsWith(root, PathComparison))
            throw new ScenarioValidationException(
                $"renderer output escapes the score artifact directory: {path}");
        return fullPath;
    }

    private static StringComparison PathComparison =>
        OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },
    };
}
