// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Diagnostics;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace LegalEval;

internal static class ArtifactWriter
{
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

    public static IReadOnlyList<ArtifactRecord> WriteIncomplete(
        string directory,
        string scenarioId,
        ScoreKind scoreKind,
        string status,
        IReadOnlyList<MetricResult> metrics,
        string reason,
        IReadOnlyList<string>? operationLog = null,
        byte[]? input = null,
        byte[]? candidate = null,
        byte[]? expected = null,
        string? semanticDiff = null,
        bool semanticDiffSucceeded = false,
        string? targetSemanticDiff = null,
        bool targetSemanticDiffSucceeded = false,
        ArtifactRenderMode renderMode = ArtifactRenderMode.Disabled,
        bool allowExternalRenderer = false,
        IEvaluationArtifactRenderer? artifactRenderer = null,
        string? inputSafetyError = null,
        string? candidateSafetyError = null,
        string? expectedSafetyError = null)
    {
        return PublishFresh(directory, stage =>
        {
            var records = CoreMetadata(stage, scenarioId, scoreKind, status, metrics,
                operationLog ?? Array.Empty<string>());
            records.Add(WriteOptionalDocx(stage, "input.docx", "input-docx", input, reason));
            records.Add(WriteOptionalDocx(stage, "candidate.docx", "candidate-docx", candidate, reason));
            records.Add(WriteOptionalDocx(stage, "expected.docx", "expected-docx", expected, reason));
            var semanticRecord = WriteText(stage, "semantic-diff.anchor-text-v1.json", "semantic-diff",
                semanticDiff ?? EvaluationScorer.ErrorEnvelope("semantic-diff", status, reason),
                "application/json");
            records.Add(semanticDiffSucceeded ? semanticRecord : semanticRecord with
            {
                Status = "failed",
                UnavailableReason = reason,
            });
            var targetSemanticRecord = WriteText(stage, "target-semantic-diff.anchor-text-v1.json",
                "target-semantic-diff",
                targetSemanticDiff ?? EvaluationScorer.ErrorEnvelope("target-semantic-diff", status, reason),
                "application/json");
            records.Add(targetSemanticDiffSucceeded ? targetSemanticRecord : targetSemanticRecord with
            {
                Status = "failed",
                UnavailableReason = reason,
            });
            records.Add(WritePreview(stage, "before.html", "before-html", input,
                input is null ? reason : inputSafetyError));
            records.Add(WritePreview(stage, "after.html", "after-html", candidate,
                candidate is null ? reason : candidateSafetyError));
            records.Add(WritePreview(stage, "target.html", "target-html", expected,
                expected is null ? reason : expectedSafetyError));
            records.Add(Unavailable("redline-docx", reason, DocxMediaType));
            AddFoundationRecords(records);
            AddRenderArtifacts(stage, records, renderMode, allowExternalRenderer,
                input, candidate, expected, redline: null, artifactRenderer);
            FinalizeIndexes(stage, scenarioId, scoreKind, status, metrics,
                operationLog ?? Array.Empty<string>(), records);
            return records;
        });
    }

    public static IReadOnlyList<ArtifactRecord> Write(
        string directory,
        string scenarioId,
        ScoreKind scoreKind,
        string status,
        IReadOnlyList<MetricResult> metrics,
        IReadOnlyList<string> operationLog,
        byte[] input,
        byte[] candidate,
        byte[] expected,
        string semanticDiff,
        bool semanticDiffSucceeded,
        string targetSemanticDiff,
        bool targetSemanticDiffSucceeded,
        string? candidateHtml,
        byte[]? redline,
        ArtifactRenderMode renderMode,
        bool allowExternalRenderer,
        string? candidateSafetyError,
        IEvaluationArtifactRenderer? artifactRenderer = null)
    {
        return PublishFresh(directory, stage =>
        {
            var records = CoreMetadata(stage, scenarioId, scoreKind, status, metrics, operationLog);
            records.Add(WriteBytes(stage, "input.docx", "input-docx", input, DocxMediaType));
            records.Add(WriteBytes(stage, "candidate.docx", "candidate-docx", candidate, DocxMediaType));
            records.Add(WriteBytes(stage, "expected.docx", "expected-docx", expected, DocxMediaType));
            var semanticRecord = WriteText(stage, "semantic-diff.anchor-text-v1.json", "semantic-diff",
                semanticDiff, "application/json");
            records.Add(semanticDiffSucceeded ? semanticRecord : semanticRecord with
            {
                Status = "failed",
                UnavailableReason = "semantic diff generation failed; the file contains a diagnostic envelope",
            });
            var targetSemanticRecord = WriteText(stage, "target-semantic-diff.anchor-text-v1.json",
                "target-semantic-diff", targetSemanticDiff, "application/json");
            records.Add(targetSemanticDiffSucceeded ? targetSemanticRecord : targetSemanticRecord with
            {
                Status = "failed",
                UnavailableReason = "target semantic diff generation failed; the file contains a diagnostic envelope",
            });
            records.Add(WritePreview(stage, "before.html", "before-html", input, null));
            records.Add(WritePreview(stage, "after.html", "after-html", candidate,
                candidateSafetyError, candidateHtml));
            records.Add(WritePreview(stage, "target.html", "target-html", expected, null));
            records.Add(redline is null
                ? Unavailable("redline-docx", "DocxDiff redline generation failed", DocxMediaType)
                : WriteBytes(stage, "redline.docx", "redline-docx", redline, DocxMediaType));
            AddFoundationRecords(records);
            AddRenderArtifacts(stage, records, renderMode, allowExternalRenderer,
                input, candidate, expected, redline, artifactRenderer);
            FinalizeIndexes(stage, scenarioId, scoreKind, status, metrics, operationLog, records);
            return records;
        });
    }

    public static IReadOnlyList<ArtifactRecord> RewriteScoreMetadata(
        string directory,
        string scenarioId,
        ScoreKind scoreKind,
        string status,
        IReadOnlyList<MetricResult> metrics,
        IReadOnlyList<ArtifactRecord> artifacts,
        IReadOnlyList<string>? operationLog = null)
    {
        var records = artifacts.Where(value => value.Id is not
            ("metrics-json" or "scenario-summary" or "artifact-index-markdown"
                or "artifact-index-html" or "evaluation-receipt" or "artifact-status"))
            .ToList();
        records.Add(WriteMetrics(directory, scenarioId, scoreKind, status, metrics));
        records.Add(WriteText(directory, "summary.md", "scenario-summary",
            BuildSummary(scenarioId, scoreKind, status, metrics), "text/markdown"));
        FinalizeIndexes(directory, scenarioId, scoreKind, status, metrics,
            operationLog ?? ReadOperationLog(directory), records);
        return records;
    }

    private static List<ArtifactRecord> CoreMetadata(
        string directory,
        string scenarioId,
        ScoreKind scoreKind,
        string status,
        IReadOnlyList<MetricResult> metrics,
        IReadOnlyList<string> operationLog) =>
        new()
        {
            WriteMetrics(directory, scenarioId, scoreKind, status, metrics),
            WriteText(directory, "operation-log.json", "operation-log", JsonSerializer.Serialize(new
            {
                schemaVersion = "1.0",
                operations = operationLog,
            }, JsonOptions), "application/json"),
            WriteText(directory, "summary.md", "scenario-summary",
                BuildSummary(scenarioId, scoreKind, status, metrics), "text/markdown"),
        };

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

    private static void AddFoundationRecords(List<ArtifactRecord> records)
    {
        records.Add(Unavailable("package-manifest-v1",
            "foundation capability unavailable until issue #456 lands; interim safety validation is not labeled a manifest",
            "application/json"));
        records.Add(Unavailable("semantic-diff-v2",
            "expanded semantic-diff capability unavailable until issue #457 lands; anchor-text v1 is preserved separately",
            "application/json"));
        records.Add(Unavailable("delivery-receipt-v1",
            "portable delivery receipt capability unavailable until issue #458 lands; evaluation-receipt is intentionally a distinct contract",
            "application/json"));
        records.Add(Unavailable("redline-proof-v1",
            "full-surface redline proof unavailable until issue #464 lands; interim text-projection metric is not labeled a proof",
            "application/json"));
    }

    private static void AddRenderArtifacts(
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
            return;
        }
        if (!allowExternalRenderer)
        {
            AddUnavailableRenderSet(records,
                "external rendering is disabled for untrusted model candidates; sanitized HTML remains available");
            return;
        }

        artifactRenderer ??= new ExternalArtifactRenderer();

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
                    "application/pdf"));
                records.Add(Unavailable($"{document.Prefix}-visual",
                    $"{document.Prefix} DOCX is unavailable", "image/png"));
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
                    rendered.Error ?? "renderer produced no PDF", "application/pdf"));
                records.Add(Unavailable($"{document.Prefix}-visual",
                    "PDF was unavailable, so no raster pages were generated", "image/png"));
                continue;
            }
            records.Add(Record(document.PdfId, pdfPath, "application/pdf"));
            var pagePaths = rendered.PagePaths
                .Select(value => RequireRenderedPath(directory, value)).ToList();
            if (rendered.Error is not null || pagePaths.Count == 0)
            {
                records.Add(Unavailable($"{document.Prefix}-visual",
                    rendered.Error ?? "renderer produced no pages", "image/png"));
                continue;
            }
            pages[document.Prefix] = pagePaths;
            for (var index = 0; index < pagePaths.Count; index++)
            {
                var id = $"{document.Prefix}-visual-page-{index + 1:D3}";
                records.Add(Record(id, pagePaths[index], "image/png"));
            }
            records.Add(Record($"{document.Prefix}-visual", pagePaths[0], "image/png"));
        }

        if (!pages.TryGetValue("candidate", out var candidatePages)
            || !pages.TryGetValue("target", out var targetPages))
        {
            records.Add(Unavailable("candidate-target-visual-diff",
                "candidate and target raster pages are both required for visual diffs", "image/png"));
            return;
        }
        var diff = artifactRenderer.ComparePages(directory, targetPages, candidatePages);
        var diffPaths = diff.Paths
            .Select(value => RequireRenderedPath(directory, value)).ToList();
        if (diff.Error is not null || diffPaths.Count == 0)
        {
            records.Add(Unavailable("candidate-target-visual-diff",
                diff.Error ?? "visual comparator produced no output", "image/png"));
            return;
        }
        for (var index = 0; index < diffPaths.Count; index++)
            records.Add(Record($"candidate-target-visual-diff-page-{index + 1:D3}",
                diffPaths[index], "image/png"));
        records.Add(Record("candidate-target-visual-diff", diffPaths[0], "image/png"));
    }

    private static void AddUnavailableRenderSet(List<ArtifactRecord> records, string reason)
    {
        foreach (var prefix in new[] { "input", "candidate", "target", "redline" })
        {
            records.Add(Unavailable($"{prefix}-pdf", reason, "application/pdf"));
            records.Add(Unavailable($"{(prefix == "input" ? "before" : prefix)}-visual",
                reason, "image/png"));
        }
        records.Add(Unavailable("candidate-target-visual-diff", reason, "image/png"));
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
        catch (Exception exception)
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
            var paths = Directory.GetFiles(directory, prefix + "-page-*.png")
                .Order(StringComparer.Ordinal).ToList();
            return result.ExitCode == 0 && paths.Count != 0
                ? (paths, null)
                : (Array.Empty<string>(), result.Diagnostics);
        }
        catch (Exception exception)
        {
            return (Array.Empty<string>(), exception.Message);
        }
    }

    private static (IReadOnlyList<string> Paths, string? Error) TryVisualDiffs(
        string directory,
        IReadOnlyList<string> targetPages,
        IReadOnlyList<string> candidatePages)
    {
        var executable = FindExecutable("compare");
        if (executable is null)
            return (Array.Empty<string>(), "ImageMagick compare is not installed");
        if (targetPages.Count != candidatePages.Count)
            return (Array.Empty<string>(),
                $"page count differs: target={targetPages.Count}, candidate={candidatePages.Count}");
        var paths = new List<string>(targetPages.Count);
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
                return (Array.Empty<string>(), result.Diagnostics);
            paths.Add(output);
        }
        return (paths, null);
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
            return new RenderedVisualDiffEvidence(result.Paths, result.Error);
        }
    }

    private static (int ExitCode, string Diagnostics) RunProcess(
        ProcessStartInfo startInfo, TimeSpan timeout)
    {
        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException($"{startInfo.FileName} did not start");
        var outputTask = process.StandardOutput.ReadToEndAsync();
        var errorTask = process.StandardError.ReadToEndAsync();
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
            or "evaluation-receipt" or "artifact-status");
        records.Add(WriteText(directory, "index.md", "artifact-index-markdown",
            BuildArtifactIndexMarkdown(scenarioId, status, records), "text/markdown"));
        records.Add(WriteText(directory, "index.html", "artifact-index-html",
            BuildArtifactIndexHtml(scenarioId, status, records), "text/html"));
        records.Add(WriteEvaluationReceipt(directory, scenarioId, scoreKind, status,
            metrics, operationLog, records));
        records.Add(WriteArtifactStatus(directory, records));
    }

    private static ArtifactRecord WriteEvaluationReceipt(
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
        return WriteText(directory, "evaluation-receipt.json", "evaluation-receipt",
            JsonSerializer.Serialize(new
            {
                schemaVersion = "docxodus.evaluation-receipt/1.0",
                receiptKind = "legal-workflow-evaluation",
                runId,
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
            .AppendLine($"- Passed metrics: {metrics.Count(value => value.Status == "passed")}/{metrics.Count}")
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
                    cached ?? EvaluationScorer.RenderHtml(bytes), "text/html");
            }
            catch (Exception exception)
            {
                failureReason = $"Docxodus HTML conversion failed: {exception.Message}";
            }
        }
        var reason = failureReason ?? "document bytes were unavailable";
        return WriteText(directory, fileName, id, DiagnosticHtml(id, reason), "text/html") with
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
        string directory, string fileName, string id, byte[]? bytes, string reason) =>
        bytes is null ? Unavailable(id, reason, DocxMediaType)
            : WriteBytes(directory, fileName, id, bytes, DocxMediaType);

    private static ArtifactRecord WriteBytes(
        string directory, string fileName, string id, byte[] bytes, string mediaType)
    {
        var path = Path.Combine(directory, fileName);
        File.WriteAllBytes(path, bytes);
        return Record(id, path, mediaType);
    }

    private static ArtifactRecord WriteText(
        string directory, string fileName, string id, string value, string mediaType)
    {
        var path = Path.Combine(directory, fileName);
        File.WriteAllText(path, value, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        return Record(id, path, mediaType);
    }

    private static ArtifactRecord Record(string id, string path, string mediaType)
    {
        var info = new FileInfo(path);
        return new ArtifactRecord(id, "available", path,
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant(),
            null, mediaType, info.Length);
    }

    private static ArtifactRecord Unavailable(string id, string reason, string mediaType) =>
        new(id, "unavailable", null, null, reason, mediaType, null);

    private static IReadOnlyList<ArtifactRecord> PublishFresh(
        string directory, Func<string, List<ArtifactRecord>> build)
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
            var stagedRecords = build(stage);
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
            return stagedRecords.Select(record => record.Path is null
                    ? record
                    : record with { Path = Path.Combine(directory, Path.GetRelativePath(stage, record.Path)) })
                .ToList();
        }
        finally
        {
            if (Directory.Exists(stage)) Directory.Delete(stage, recursive: true);
        }
    }

    private static IReadOnlyList<string> ReadOperationLog(string directory)
    {
        var path = Path.Combine(directory, "operation-log.json");
        if (!File.Exists(path)) return Array.Empty<string>();
        try
        {
            using var document = JsonDocument.Parse(File.ReadAllBytes(path));
            return document.RootElement.GetProperty("operations").EnumerateArray()
                .Select(value => value.GetString() ?? string.Empty).ToList();
        }
        catch (Exception exception) when (exception is JsonException or IOException)
        {
            return new[] { $"operation-log-unreadable:{exception.Message}" };
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
