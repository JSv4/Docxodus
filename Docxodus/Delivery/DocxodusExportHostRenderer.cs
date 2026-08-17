// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Buffers;
using System.Buffers.Binary;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Docxodus.Internal;
using Docxodus.Verification;

namespace Docxodus.Delivery;

/// <summary>How the standalone renderer handles content it cannot faithfully project.</summary>
public enum DeliveryUnsupportedContentPolicy
{
    Warn,
    Strict,
}

/// <summary>
/// Explicit process-owned configuration for the <c>@docxodus/export</c> framed host. Executable
/// locations are never inferred from PATH and framed callers cannot choose local paths.
/// </summary>
public sealed record DocxodusExportHostRendererOptions
{
    public const string NodePathEnvironmentVariable = "DOCXODUS_NODE_PATH";
    public const string HostPathEnvironmentVariable = "DOCXODUS_EXPORT_HOST_PATH";
    public const string ChromiumPathEnvironmentVariable = "DOCXODUS_CHROMIUM_PATH";

    required public string NodeExecutablePath { get; init; }
    required public string HostScriptPath { get; init; }
    public string? ChromiumExecutablePath { get; init; }
    public TimeSpan RenderTimeout { get; init; } = TimeSpan.FromMinutes(2);
    public DeliveryUnsupportedContentPolicy UnsupportedContent { get; init; } =
        DeliveryUnsupportedContentPolicy.Warn;
    public bool StrictFonts { get; init; }
    public int MaximumFrameBytes { get; init; } = 512 * 1024 * 1024;
    public int MaximumStandardErrorCharacters { get; init; } = 64 * 1024;

    /// <summary>
    /// Resolve the process-owned adapter from environment configuration. Both Node and host paths
    /// must be present together; Chromium remains optional because the matching package can own it.
    /// </summary>
    public static DocxodusExportHostRendererOptions? FromEnvironment()
    {
        var node = Environment.GetEnvironmentVariable(NodePathEnvironmentVariable);
        var host = Environment.GetEnvironmentVariable(HostPathEnvironmentVariable);
        if (string.IsNullOrWhiteSpace(node) && string.IsNullOrWhiteSpace(host))
            return null;
        if (string.IsNullOrWhiteSpace(node) || string.IsNullOrWhiteSpace(host))
            throw new InvalidOperationException(
                $"{NodePathEnvironmentVariable} and {HostPathEnvironmentVariable} must be configured together.");
        return new DocxodusExportHostRendererOptions
        {
            NodeExecutablePath = node,
            HostScriptPath = host,
            ChromiumExecutablePath = Environment.GetEnvironmentVariable(
                ChromiumPathEnvironmentVariable),
        };
    }

    internal void Validate()
    {
        ValidateFile(NodeExecutablePath, nameof(NodeExecutablePath));
        ValidateFile(HostScriptPath, nameof(HostScriptPath));
        if (!string.IsNullOrWhiteSpace(ChromiumExecutablePath))
            ValidateFile(ChromiumExecutablePath, nameof(ChromiumExecutablePath));
        if (RenderTimeout <= TimeSpan.Zero || RenderTimeout > TimeSpan.FromMinutes(10))
            throw new ArgumentOutOfRangeException(nameof(RenderTimeout));
        if (!Enum.IsDefined(UnsupportedContent))
            throw new ArgumentOutOfRangeException(nameof(UnsupportedContent));
        if (MaximumFrameBytes is <= 0 or > 512 * 1024 * 1024)
            throw new ArgumentOutOfRangeException(nameof(MaximumFrameBytes));
        if (MaximumStandardErrorCharacters is <= 0 or > 1024 * 1024)
            throw new ArgumentOutOfRangeException(nameof(MaximumStandardErrorCharacters));
    }

    private static void ValidateFile(string? path, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(path) || !Path.IsPathFullyQualified(path))
            throw new ArgumentException("An absolute executable or host path is required.", parameterName);
        if (!File.Exists(path))
            throw new FileNotFoundException("The configured executable or host file does not exist.", path);
    }
}

/// <summary>
/// Production delivery adapter for the framed <c>@docxodus/export</c> host delivered by epic
/// #434. Every exact source/profile cohort is rendered once, then HTML, PDF, PageMap, and report
/// outputs are projected from that one verified host result.
/// </summary>
public sealed class DocxodusExportHostRenderer : IDeliveryArtifactBatchRenderer
{
    private const long MaximumSafeJavaScriptInteger = 9_007_199_254_740_991;
    private const int MaximumBatchRequests = 1_024;
    private const int MaximumRenderDiagnostics = 1_024;
    private const string RenderReportSchema =
        "https://docxodus.dev/schemas/render/render-report/v1";
    private static readonly JsonSerializerOptions HostJson = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    private readonly DocxodusExportHostRendererOptions _options;

    public DocxodusExportHostRenderer(DocxodusExportHostRendererOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _options.Validate();
    }

    public DeliveryRendererCapabilities Capabilities { get; } = new(
        "@docxodus/export/framed-host-v1",
        new[]
        {
            DeliveryArtifactKind.StandaloneHtml,
            DeliveryArtifactKind.FinalPdf,
            DeliveryArtifactKind.ReviewPdf,
            DeliveryArtifactKind.PageMap,
            DeliveryArtifactKind.RenderReport,
        },
        Enum.GetValues<DeliveryReviewProfile>(),
        Enum.GetValues<DeliveryCommentProfile>());

    public async ValueTask<DeliveryRenderResult> RenderAsync(
        DeliveryRenderRequest request,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(request);
        var results = await RenderBatchAsync(new[] { request }, cancellationToken)
            .ConfigureAwait(false);
        return results[request.ArtifactId];
    }

    public async ValueTask<IReadOnlyDictionary<string, DeliveryRenderResult>> RenderBatchAsync(
        IReadOnlyList<DeliveryRenderRequest> requests,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(requests);
        cancellationToken.ThrowIfCancellationRequested();
        if (requests.Count > MaximumBatchRequests)
            throw new InvalidDataException(
                "The export-host request exceeds the bounded batch-request limit.");
        var snapshot = requests.ToArray();
        if (snapshot.Any(request => request is null))
            throw new ArgumentException("Render requests cannot contain null entries.", nameof(requests));
        if (snapshot.GroupBy(request => request.ArtifactId, StringComparer.Ordinal)
            .Any(group => group.Count() != 1))
            throw new ArgumentException("Render request IDs must be unique.", nameof(requests));
        if (snapshot.Length == 0)
            return new Dictionary<string, DeliveryRenderResult>(StringComparer.Ordinal);
        foreach (var request in snapshot)
        {
            if (!ValidTransportString(request.ArtifactId)
                || !ValidTransportString(request.SourceDocumentName))
                throw new ArgumentException(
                    "Render request IDs and source names must be bounded, control-free strings.",
                    nameof(requests));
            if (!Capabilities.Supports(
                    request.Kind, request.ReviewProfile, request.CommentProfile))
                throw new ArgumentException(
                    $"Render request '{request.ArtifactId}' is not supported by this adapter.",
                    nameof(requests));
            if (request.SourceDocumentVersion > MaximumSafeJavaScriptInteger)
                throw new ArgumentOutOfRangeException(
                    nameof(requests),
                    "The export host requires a JavaScript-safe document version.");
        }

        var groups = snapshot
            .GroupBy(request => new RenderGroupKey(
                request.SourceDocumentName,
                request.SourcePackageDigest.Value,
                request.SourceDocumentVersion,
                request.ReviewProfile,
                request.CommentProfile))
            .OrderBy(group => group.Key.SourceDocumentName, StringComparer.Ordinal)
            .ThenBy(group => group.Key.SourceDigest, StringComparer.Ordinal)
            .ThenBy(group => group.Key.DocumentVersion)
            .ThenBy(group => group.Key.ReviewProfile)
            .ThenBy(group => group.Key.CommentProfile)
            .Select((group, index) => RenderGroup.Create(
                $"render-{index + 1:D4}", group.Key, group.ToArray()))
            .ToArray();
        var requestBytes = SerializeRequest(groups);
        var responseBytes = await InvokeHostAsync(requestBytes, groups.Length, cancellationToken)
            .ConfigureAwait(false);
        return ParseResponse(responseBytes, groups);
    }

    private byte[] SerializeRequest(IReadOnlyList<RenderGroup> groups)
    {
        long encodedDocumentCharacters = 0;
        foreach (var group in groups)
        {
            var length = group.Requests[0].SourceByteLength;
            var encodedLength = checked(((long)length + 2) / 3 * 4);
            if (encodedLength > _options.MaximumFrameBytes
                || encodedDocumentCharacters > _options.MaximumFrameBytes - encodedLength)
                throw new InvalidDataException(
                    "The export-host request exceeds the configured frame limit.");
            encodedDocumentCharacters += encodedLength;
        }

        HostRequest Envelope(Func<RenderGroup, string> document) => new(
            1, groups.Select(group => new HostBatchRequest(
                group.BatchId,
                document(group),
                new HostBatchOptions(
                    group.Outputs,
                    group.Key.DocumentVersion,
                    group.Key.SourceDigest,
                    Name(group.Key.ReviewProfile),
                    UsesProfileResolvedSource(group.Key.ReviewProfile) ? true : null,
                    Name(group.Key.CommentProfile),
                    Name(_options.UnsupportedContent),
                    _options.StrictFonts,
                    checked((int)_options.RenderTimeout.TotalMilliseconds))))
                .ToArray());
        var envelopeBytes = JsonSerializer.SerializeToUtf8Bytes(
            Envelope(_ => string.Empty), HostJson).LongLength;
        if (envelopeBytes > _options.MaximumFrameBytes
            || encodedDocumentCharacters > _options.MaximumFrameBytes - envelopeBytes)
            throw new InvalidDataException(
                "The export-host request exceeds the configured frame limit.");
        var envelope = Envelope(group =>
            Convert.ToBase64String(group.Requests[0].CopySourceBytes()));
        var bytes = JsonSerializer.SerializeToUtf8Bytes(envelope, HostJson);
        if (bytes.Length > _options.MaximumFrameBytes)
            throw new InvalidDataException("The export-host request exceeds the configured frame limit.");
        return bytes;
    }

    private async ValueTask<byte[]> InvokeHostAsync(
        byte[] payload,
        int batchCount,
        CancellationToken cancellationToken)
    {
        var start = new ProcessStartInfo
        {
            FileName = _options.NodeExecutablePath,
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };
        start.ArgumentList.Add(_options.HostScriptPath);
        if (!string.IsNullOrWhiteSpace(_options.ChromiumExecutablePath))
            start.Environment[DocxodusExportHostRendererOptions.ChromiumPathEnvironmentVariable] =
                _options.ChromiumExecutablePath;

        using var process = new Process { StartInfo = start };
        if (!process.Start())
            throw new InvalidOperationException("The configured export host could not be started.");
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var totalTimeoutTicks = Math.Min(
            TimeSpan.FromMinutes(30).Ticks,
            checked(_options.RenderTimeout.Ticks * Math.Max(1, batchCount)
                    + TimeSpan.FromSeconds(30).Ticks));
        deadline.CancelAfter(TimeSpan.FromTicks(totalTimeoutTicks));
        var token = deadline.Token;
        var stderrTask = ReadStandardErrorAsync(
            process.StandardError, _options.MaximumStandardErrorCharacters, token);
        var responseTask = ReadFrameAsync(process.StandardOutput.BaseStream, token);
        try
        {
            var prefix = new byte[4];
            BinaryPrimitives.WriteUInt32BigEndian(prefix, checked((uint)payload.Length));
            await process.StandardInput.BaseStream.WriteAsync(prefix, token).ConfigureAwait(false);
            await process.StandardInput.BaseStream.WriteAsync(payload, token).ConfigureAwait(false);
            await process.StandardInput.BaseStream.FlushAsync(token).ConfigureAwait(false);
            process.StandardInput.Close();

            await process.WaitForExitAsync(token).ConfigureAwait(false);
            var response = await responseTask.ConfigureAwait(false);
            var stderr = await stderrTask.ConfigureAwait(false);
            if (process.ExitCode != 0)
                throw new InvalidDataException(
                    $"The export host exited with code {process.ExitCode}: {stderr}");
            return response;
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            TryKill(process);
            throw new TimeoutException("The export host exceeded the bounded render deadline.");
        }
        catch
        {
            TryKill(process);
            throw;
        }
    }

    private async Task<byte[]> ReadFrameAsync(Stream stream, CancellationToken cancellationToken)
    {
        var prefix = new byte[4];
        await ReadExactlyAsync(stream, prefix, cancellationToken).ConfigureAwait(false);
        var length = BinaryPrimitives.ReadUInt32BigEndian(prefix);
        if (length == 0 || length > _options.MaximumFrameBytes)
            throw new InvalidDataException("The export host returned an invalid frame length.");
        var payload = new byte[length];
        await ReadExactlyAsync(stream, payload, cancellationToken).ConfigureAwait(false);
        var trailing = new byte[1];
        if (await stream.ReadAsync(trailing, cancellationToken).ConfigureAwait(false) != 0)
            throw new InvalidDataException("The export host returned bytes after its response frame.");
        return payload;
    }

    private static async Task ReadExactlyAsync(
        Stream stream,
        Memory<byte> destination,
        CancellationToken cancellationToken)
    {
        var offset = 0;
        while (offset < destination.Length)
        {
            var read = await stream.ReadAsync(destination[offset..], cancellationToken)
                .ConfigureAwait(false);
            if (read == 0)
                throw new EndOfStreamException("The export host closed before its frame was complete.");
            offset += read;
        }
    }

    private static async Task<string> ReadStandardErrorAsync(
        TextReader reader,
        int maximumCharacters,
        CancellationToken cancellationToken)
    {
        var buffer = new char[4096];
        var output = new StringBuilder(Math.Min(maximumCharacters, buffer.Length));
        while (true)
        {
            var read = await reader.ReadAsync(buffer.AsMemory(), cancellationToken)
                .ConfigureAwait(false);
            if (read == 0) break;
            var remaining = maximumCharacters - output.Length;
            if (remaining > 0)
                output.Append(buffer, 0, Math.Min(read, remaining));
        }
        return output.ToString();
    }

    private static void TryKill(Process process)
    {
        try
        {
            if (!process.HasExited)
                process.Kill(entireProcessTree: true);
        }
        catch (Exception exception) when (exception is InvalidOperationException
            or System.ComponentModel.Win32Exception
            or NotSupportedException
            or ObjectDisposedException)
        {
        }
    }

    private static IReadOnlyDictionary<string, DeliveryRenderResult> ParseResponse(
        byte[] responseBytes,
        IReadOnlyList<RenderGroup> groups)
    {
        using var document = JsonDocument.Parse(responseBytes, new JsonDocumentOptions
        {
            AllowTrailingCommas = false,
            CommentHandling = JsonCommentHandling.Disallow,
            MaxDepth = 128,
        });
        EnsureUniqueProperties(document.RootElement);
        var root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Object
            || RequiredInt32(root, "schemaVersion") != 1)
            throw new InvalidDataException("The export host response schema is invalid.");
        if (root.TryGetProperty("fatal", out var fatal))
            throw new InvalidDataException("The export host rejected its envelope: "
                                           + ErrorReason(fatal));
        var batchArray = RequiredArray(root, "batches");
        if (batchArray.GetArrayLength() != groups.Count)
            throw new InvalidDataException(
                "The export host response does not match the requested batch count.");
        var byId = new Dictionary<string, JsonElement>(groups.Count, StringComparer.Ordinal);
        foreach (var batch in batchArray.EnumerateArray())
        {
            if (!byId.TryAdd(RequiredString(batch, "id"), batch))
                throw new InvalidDataException(
                    "The export host response repeats a batch ID.");
        }
        if (!groups.Select(group => group.BatchId).ToHashSet(StringComparer.Ordinal)
            .SetEquals(byId.Keys))
            throw new InvalidDataException(
                "The export host response does not match the requested batch IDs.");

        var results = new Dictionary<string, DeliveryRenderResult>(StringComparer.Ordinal);
        foreach (var group in groups)
        {
            var batch = byId[group.BatchId];
            if (!RequiredBoolean(batch, "ok"))
            {
                var failure = ParseFailedResult(group, RequiredObject(batch, "error"));
                foreach (var request in group.Requests)
                    results.Add(request.ArtifactId, failure.For(request));
                continue;
            }
            var cohort = ParseSuccessfulResult(group, RequiredObject(batch, "result"));
            foreach (var request in group.Requests)
                results.Add(request.ArtifactId, cohort.For(request));
        }
        return results;
    }

    private static FailedRenderCohort ParseFailedResult(RenderGroup group, JsonElement error)
    {
        var reason = ErrorReason(error);
        if (!error.TryGetProperty("report", out var report)
            || report.ValueKind != JsonValueKind.Object)
            return new FailedRenderCohort(reason, null, null, Array.Empty<DeliverableRenderDiagnostic>());

        ValidateFailedReport(group, error, report);
        var fingerprint = report.TryGetProperty("environment", out var environment)
                          && environment.ValueKind == JsonValueKind.Object
            ? OptionalString(environment, "rendererFingerprint")
            : null;
        return new FailedRenderCohort(
            reason,
            CanonicalJson(report),
            fingerprint,
            ParseDiagnosticArray(RequiredArray(report, "warnings")));
    }

    private static void ValidateFailedReport(
        RenderGroup group,
        JsonElement error,
        JsonElement report)
    {
        if (RequiredString(report, "schema") != RenderReportSchema
            || RequiredInt32(report, "schemaVersion") != 1
            || RequiredString(report, "status") != "failed")
            throw new InvalidDataException("The export host returned an invalid failed render report.");
        var source = RequiredObject(report, "source");
        var first = group.Requests[0];
        if (!string.Equals(RequiredString(source, "rawPackageBytesDigest"),
                group.Key.SourceDigest, StringComparison.OrdinalIgnoreCase)
            || RequiredInt64(source, "documentVersion") != group.Key.DocumentVersion
            || RequiredInt64(source, "byteLength") != first.SourceByteLength)
            throw new InvalidDataException("The failed render report is bound to different source bytes.");
        ValidateProfileBinding(group, report, "failed render report");
        var failure = RequiredObject(report, "failure");
        foreach (var field in new[] { "code", "phase", "message", "remediation" })
        {
            if (!string.Equals(RequiredString(failure, field), RequiredString(error, field),
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    $"The failed render report disagrees with its host error field '{field}'.");
        }
    }

    private static RenderCohort ParseSuccessfulResult(RenderGroup group, JsonElement result)
    {
        var fingerprint = RequiredString(result, "rendererFingerprint");
        var pageCount = RequiredInt64(result, "pageCount");
        if (pageCount <= 0)
            throw new InvalidDataException("The export host returned an invalid page count.");
        var pageMapElement = RequiredObject(result, "pageMap");
        var reportElement = RequiredObject(result, "renderReport");
        var pageMap = DocxSessionJson.ParsePageMap(pageMapElement);
        var portable = PageMapContract.ValidatePortable(pageMap);
        if (!portable.Success || pageMap.Mode != PageMapMode.Paginated
            || pageMap.Availability != PageMapAvailability.Available)
            throw new InvalidDataException(portable.Message
                ?? "The export host returned a non-portable PageMap.");
        if (pageMap.DocumentVersion != group.Key.DocumentVersion
            || pageMap.Pages.Count != pageCount
            || !string.Equals(pageMap.RendererFingerprint, fingerprint,
                StringComparison.Ordinal))
            throw new InvalidDataException("The export host returned incoherent PageMap metadata.");
        var pageMapBytes = Encoding.UTF8.GetBytes(DocxSessionJson.SerializePageMap(pageMap));
        var reportBytes = CanonicalJson(reportElement);
        ValidateReport(group, reportElement, fingerprint, pageCount, pageMapBytes);

        byte[]? htmlBytes = null;
        if (group.Outputs.Contains("html", StringComparer.Ordinal))
        {
            htmlBytes = Encoding.UTF8.GetBytes(RequiredString(
                result, "html", allowEmpty: false, maximumCharacters: int.MaxValue,
                allowControlCharacters: true));
            EnsureDigest(reportElement, "htmlDigest", htmlBytes);
        }
        byte[]? pdfBytes = null;
        if (group.Outputs.Contains("pdf", StringComparer.Ordinal))
        {
            pdfBytes = DecodeCanonicalBase64(RequiredString(
                result, "pdfBase64", maximumCharacters: int.MaxValue));
            EnsureDigest(reportElement, "pdfDigest", pdfBytes);
        }
        var diagnostics = ParseDiagnostics(result, reportElement);
        return new RenderCohort(
            htmlBytes,
            pdfBytes,
            pageMapBytes,
            reportBytes,
            fingerprint,
            pageCount,
            diagnostics);
    }

    private static void ValidateReport(
        RenderGroup group,
        JsonElement report,
        string fingerprint,
        long pageCount,
        byte[] pageMapBytes)
    {
        if (RequiredString(report, "schema") != RenderReportSchema
            || RequiredInt32(report, "schemaVersion") != 1
            || RequiredString(report, "status") != "complete")
            throw new InvalidDataException("The export host returned an invalid render report.");
        var source = RequiredObject(report, "source");
        var first = group.Requests[0];
        if (!string.Equals(RequiredString(source, "rawPackageBytesDigest"),
                group.Key.SourceDigest, StringComparison.OrdinalIgnoreCase)
            || RequiredInt64(source, "documentVersion") != group.Key.DocumentVersion
            || RequiredInt64(source, "byteLength") != first.SourceByteLength)
            throw new InvalidDataException("The render report is bound to different source bytes.");
        ValidateProfileBinding(group, report, "render report");
        var environment = RequiredObject(report, "environment");
        if (!string.Equals(RequiredString(environment, "rendererFingerprint"), fingerprint,
                StringComparison.Ordinal))
            throw new InvalidDataException("The render report fingerprint is incoherent.");
        if (RequiredArray(report, "pages").GetArrayLength() != pageCount)
            throw new InvalidDataException("The render report page count is incoherent.");
        EnsureDigest(report, "pageMapDigest", pageMapBytes);
    }

    private static IReadOnlyList<DeliverableRenderDiagnostic> ParseDiagnostics(
        JsonElement result,
        JsonElement report)
    {
        var resultWarnings = RequiredArray(result, "warnings");
        var reportWarnings = RequiredArray(report, "warnings");
        if (!CanonicalJson(resultWarnings).AsSpan().SequenceEqual(CanonicalJson(reportWarnings)))
            throw new InvalidDataException("The export host returned inconsistent warning sets.");
        return ParseDiagnosticArray(reportWarnings);
    }

    private static IReadOnlyList<DeliverableRenderDiagnostic> ParseDiagnosticArray(
        JsonElement warnings)
    {
        if (warnings.GetArrayLength() > MaximumRenderDiagnostics)
            throw new InvalidDataException(
                "The export host returned too many render diagnostics.");
        return warnings.EnumerateArray().Select(warning =>
        {
            var code = RequiredString(warning, "code");
            var severity = RequiredString(warning, "severity");
            if (severity is not ("warning" or "error"))
                throw new InvalidDataException(
                    "The export host returned an invalid warning severity.");
            return new DeliverableRenderDiagnostic
            {
                Kind = code.Contains("font_substitution", StringComparison.OrdinalIgnoreCase)
                    ? DeliverableRenderDiagnosticKind.FontSubstitution
                    : code.Contains("font", StringComparison.OrdinalIgnoreCase)
                      && code.Contains("missing", StringComparison.OrdinalIgnoreCase)
                        ? DeliverableRenderDiagnosticKind.MissingFont
                        : code.Contains("unsupported", StringComparison.OrdinalIgnoreCase)
                            ? DeliverableRenderDiagnosticKind.UnsupportedContent
                            : DeliverableRenderDiagnosticKind.Warning,
                Severity = severity == "error"
                    ? VerificationFindingSeverity.Error
                    : VerificationFindingSeverity.Warning,
                Code = code,
                Phase = RequiredString(warning, "phase"),
                Message = RequiredString(warning, "message"),
                Remediation = RequiredString(warning, "remediation"),
                OwningPartUri = OptionalString(warning, "partUri"),
                AnchorId = OptionalString(warning, "anchorId"),
                Resource = OptionalString(warning, "resource"),
            };
        }).ToArray();
    }

    private static void EnsureDigest(JsonElement report, string bindingName, byte[] bytes)
    {
        var bindings = RequiredObject(report, "bindings");
        var expected = RequiredString(bindings, bindingName);
        var actual = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        if (!string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException(
                $"The render report {bindingName} does not match the returned bytes.");
    }

    private static void ValidateProfileBinding(
        RenderGroup group,
        JsonElement report,
        string label)
    {
        var options = RequiredObject(report, "options");
        if (RequiredString(options, "reviewProfile") != Name(group.Key.ReviewProfile)
            || RequiredString(options, "commentProfile") != Name(group.Key.CommentProfile))
            throw new InvalidDataException($"The {label} profile binding is invalid.");

        var expectsResolvedSource = UsesProfileResolvedSource(group.Key.ReviewProfile);
        var declaresResolvedSource = options.TryGetProperty(
            "reviewProfileAlreadyApplied", out var declaration);
        var exactSourceBindingIsValid = expectsResolvedSource
            ? declaresResolvedSource && declaration.ValueKind == JsonValueKind.True
            : !declaresResolvedSource;
        if (!exactSourceBindingIsValid)
            throw new InvalidDataException(
                $"The {label} exact-source profile binding is invalid.");
        if (expectsResolvedSource && report.TryGetProperty("derivedProfileSource", out _))
            throw new InvalidDataException(
                $"The {label} rewrote an exact profile-resolved source.");
    }

    private static byte[] DecodeCanonicalBase64(string value)
    {
        if (!IsCanonicalBase64(value))
            throw new InvalidDataException("The export host returned non-canonical base64.");
        try
        {
            return Convert.FromBase64String(value);
        }
        catch (FormatException exception)
        {
            throw new InvalidDataException("The export host returned invalid base64.", exception);
        }
    }

    private static bool IsCanonicalBase64(string value)
    {
        if (value.Length == 0 || value.Length % 4 != 0)
            return false;
        int padding = value.EndsWith("==", StringComparison.Ordinal)
            ? 2
            : value.EndsWith('=') ? 1 : 0;
        for (var index = 0; index < value.Length - padding; index++)
        {
            if (Base64Value(value[index]) < 0)
                return false;
        }
        for (var index = value.Length - padding; index < value.Length; index++)
        {
            if (value[index] != '=')
                return false;
        }
        if (padding == 2)
            return (Base64Value(value[^3]) & 0x0f) == 0;
        if (padding == 1)
            return (Base64Value(value[^2]) & 0x03) == 0;
        return true;
    }

    private static int Base64Value(char value) => value switch
    {
        >= 'A' and <= 'Z' => value - 'A',
        >= 'a' and <= 'z' => value - 'a' + 26,
        >= '0' and <= '9' => value - '0' + 52,
        '+' => 62,
        '/' => 63,
        _ => -1,
    };

    private static byte[] CanonicalJson(JsonElement element)
    {
        var buffer = new ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer))
            WriteCanonical(writer, element);
        return buffer.WrittenSpan.ToArray();
    }

    private static void WriteCanonical(Utf8JsonWriter writer, JsonElement element)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
                writer.WriteStartObject();
                foreach (var property in element.EnumerateObject()
                             .OrderBy(property => property.Name, StringComparer.Ordinal))
                {
                    writer.WritePropertyName(property.Name);
                    WriteCanonical(writer, property.Value);
                }
                writer.WriteEndObject();
                break;
            case JsonValueKind.Array:
                writer.WriteStartArray();
                foreach (var value in element.EnumerateArray())
                    WriteCanonical(writer, value);
                writer.WriteEndArray();
                break;
            case JsonValueKind.String:
                writer.WriteStringValue(element.GetString());
                break;
            case JsonValueKind.Number:
                writer.WriteRawValue(element.GetRawText());
                break;
            case JsonValueKind.True:
                writer.WriteBooleanValue(true);
                break;
            case JsonValueKind.False:
                writer.WriteBooleanValue(false);
                break;
            case JsonValueKind.Null:
                writer.WriteNullValue();
                break;
            default:
                throw new InvalidDataException("The export host returned unsupported JSON data.");
        }
    }

    private static void EnsureUniqueProperties(JsonElement element)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            var names = new HashSet<string>(StringComparer.Ordinal);
            foreach (var property in element.EnumerateObject())
            {
                if (!names.Add(property.Name))
                    throw new InvalidDataException(
                        $"The export host response repeats property '{property.Name}'.");
                EnsureUniqueProperties(property.Value);
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
                EnsureUniqueProperties(item);
        }
    }

    private static string ErrorReason(JsonElement error)
    {
        if (error.ValueKind != JsonValueKind.Object)
            return "The export host returned an unstructured failure.";
        var code = OptionalString(error, "code") ?? OptionalString(error, "name") ?? "error";
        var phase = OptionalString(error, "phase");
        var message = OptionalString(error, "message") ?? "Rendering failed.";
        var remediation = OptionalString(error, "remediation");
        var reason = $"Export host {code}{(phase is null ? string.Empty : $" at {phase}")}: {message}"
                     + (remediation is null ? string.Empty : $" Remediation: {remediation}");
        return reason.Length <= 4096 ? reason : reason[..4096];
    }

    private static JsonElement RequiredObject(JsonElement value, string name) =>
        value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.Object
            ? property
            : throw new InvalidDataException($"The export host response is missing object '{name}'.");

    private static JsonElement RequiredArray(JsonElement value, string name) =>
        value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.Array
            ? property
            : throw new InvalidDataException($"The export host response is missing array '{name}'.");

    private static string RequiredString(
        JsonElement value,
        string name,
        bool allowEmpty = false,
        int maximumCharacters = DeliveryArtifactRequestRules.MaximumStringLength,
        bool allowControlCharacters = false) =>
        value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String
        && property.GetString()!.Length <= maximumCharacters
        && (allowEmpty || !string.IsNullOrWhiteSpace(property.GetString()))
        && (allowControlCharacters || !property.GetString()!.Any(char.IsControl))
            ? property.GetString()!
            : throw new InvalidDataException($"The export host response is missing string '{name}'.");

    private static string? OptionalString(JsonElement value, string name)
    {
        if (!value.TryGetProperty(name, out var property))
            return null;
        if (property.ValueKind != JsonValueKind.String
            || property.GetString()!.Length > DeliveryArtifactRequestRules.MaximumStringLength
            || property.GetString()!.Any(char.IsControl))
            throw new InvalidDataException(
                $"The export host response has an invalid optional string '{name}'.");
        return property.GetString();
    }

    private static int RequiredInt32(JsonElement value, string name) =>
        value.TryGetProperty(name, out var property) && property.TryGetInt32(out var number)
            ? number
            : throw new InvalidDataException($"The export host response is missing integer '{name}'.");

    private static long RequiredInt64(JsonElement value, string name) =>
        value.TryGetProperty(name, out var property) && property.TryGetInt64(out var number)
            ? number
            : throw new InvalidDataException($"The export host response is missing integer '{name}'.");

    private static bool RequiredBoolean(JsonElement value, string name) =>
        value.TryGetProperty(name, out var property)
        && property.ValueKind is JsonValueKind.True or JsonValueKind.False
            ? property.GetBoolean()
            : throw new InvalidDataException($"The export host response is missing boolean '{name}'.");

    private static string MediaType(DeliveryArtifactKind kind) => kind switch
    {
        DeliveryArtifactKind.StandaloneHtml => "text/html",
        DeliveryArtifactKind.FinalPdf or DeliveryArtifactKind.ReviewPdf => "application/pdf",
        DeliveryArtifactKind.PageMap => "application/vnd.docxodus.pagemap+json",
        DeliveryArtifactKind.RenderReport => "application/vnd.docxodus.render-report+json",
        _ => throw new ArgumentOutOfRangeException(nameof(kind)),
    };

    private static string Name<T>(T value)
        where T : struct, Enum =>
        JsonNamingPolicy.CamelCase.ConvertName(value.ToString());

    private static bool ValidTransportString(string value) =>
        !string.IsNullOrWhiteSpace(value)
        && value.Length <= DeliveryArtifactRequestRules.MaximumStringLength
        && !value.Any(char.IsControl);

    private static bool UsesProfileResolvedSource(DeliveryReviewProfile profile) =>
        profile is DeliveryReviewProfile.Final or DeliveryReviewProfile.Original;

    private sealed record RenderGroupKey(
        string SourceDocumentName,
        string SourceDigest,
        long DocumentVersion,
        DeliveryReviewProfile ReviewProfile,
        DeliveryCommentProfile CommentProfile);

    private sealed record RenderGroup(
        string BatchId,
        RenderGroupKey Key,
        IReadOnlyList<DeliveryRenderRequest> Requests,
        IReadOnlyList<string> Outputs)
    {
        internal static RenderGroup Create(
            string batchId,
            RenderGroupKey key,
            IReadOnlyList<DeliveryRenderRequest> requests)
        {
            var outputs = new List<string>();
            if (requests.Any(request => request.Kind == DeliveryArtifactKind.StandaloneHtml)
                || !requests.Any(request => request.Kind is DeliveryArtifactKind.FinalPdf
                    or DeliveryArtifactKind.ReviewPdf))
                outputs.Add("html");
            if (requests.Any(request => request.Kind is DeliveryArtifactKind.FinalPdf
                    or DeliveryArtifactKind.ReviewPdf))
                outputs.Add("pdf");
            return new RenderGroup(batchId, key, requests, outputs);
        }
    }

    private sealed record RenderCohort(
        byte[]? HtmlBytes,
        byte[]? PdfBytes,
        byte[] PageMapBytes,
        byte[] ReportBytes,
        string Fingerprint,
        long PageCount,
        IReadOnlyList<DeliverableRenderDiagnostic> Diagnostics)
    {
        internal DeliveryRenderResult For(DeliveryRenderRequest request)
        {
            var bytes = request.Kind switch
            {
                DeliveryArtifactKind.StandaloneHtml => HtmlBytes,
                DeliveryArtifactKind.FinalPdf or DeliveryArtifactKind.ReviewPdf => PdfBytes,
                DeliveryArtifactKind.PageMap => PageMapBytes,
                DeliveryArtifactKind.RenderReport => ReportBytes,
                _ => null,
            } ?? throw new InvalidDataException(
                $"The export host omitted requested artifact '{request.ArtifactId}'.");
            return DeliveryRenderResult.Available(
                bytes,
                MediaType(request.Kind),
                Fingerprint,
                PageCount,
                PageMapBytes,
                ReportBytes,
                Diagnostics);
        }
    }

    private sealed record FailedRenderCohort(
        string Reason,
        byte[]? ReportBytes,
        string? RendererFingerprint,
        IReadOnlyList<DeliverableRenderDiagnostic> Diagnostics)
    {
        internal DeliveryRenderResult For(DeliveryRenderRequest request)
        {
            if (request.Kind == DeliveryArtifactKind.RenderReport && ReportBytes is not null)
                return DeliveryRenderResult.FailedReport(
                    ReportBytes, RendererFingerprint, Diagnostics);
            return DeliveryRenderResult.Unavailable(
                MediaType(request.Kind), Reason, RendererFingerprint,
                ReportBytes is null ? ReadOnlySpan<byte>.Empty : ReportBytes, Diagnostics);
        }
    }

    private sealed record HostRequest(int SchemaVersion, IReadOnlyList<HostBatchRequest> Batches);

    private sealed record HostBatchRequest(
        string Id,
        string DocumentBase64,
        HostBatchOptions Options);

    private sealed record HostBatchOptions(
        IReadOnlyList<string> Outputs,
        long DocumentVersion,
        string ExpectedSourceDigest,
        string ReviewProfile,
        bool? ReviewProfileAlreadyApplied,
        string CommentProfile,
        string UnsupportedContent,
        bool StrictFonts,
        int TimeoutMs);
}
