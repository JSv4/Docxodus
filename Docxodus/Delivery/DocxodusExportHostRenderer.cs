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
public sealed class DocxodusExportHostRenderer : IDeliveryArtifactRenderer
{
    private const long MaximumSafeJavaScriptInteger = 9_007_199_254_740_991;
    private const string RenderReportSchema =
        "https://docxodus.dev/schemas/render/render-report/v1";
    private const string DocxMediaType =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    private const string UnrepresentableVersionReason = "document_version_unrepresentable";
    private const int MaximumControlFrameBytes = 8 * 1024 * 1024;
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

    /// <summary>
    /// The pure per-pair batch identity. The layout-options digest covers exactly the option
    /// material this adapter sends the host for the pair, minus the per-source bindings
    /// (document version, source digest, artifact request IDs); the runtime-policy digest covers
    /// the fixed process/browser policy this adapter was constructed with. Both are versioned
    /// canonical-JSON materials, so an option that starts influencing layout must enter the
    /// material — and therefore the digest — rather than ride along unseen.
    /// </summary>
    public DeliveryRenderBatchContext DescribeBatch(
        DeliveryReviewProfile reviewProfile,
        DeliveryCommentProfile commentProfile)
    {
        if (!Enum.IsDefined(reviewProfile))
            throw new ArgumentOutOfRangeException(nameof(reviewProfile));
        if (!Enum.IsDefined(commentProfile))
            throw new ArgumentOutOfRangeException(nameof(commentProfile));
        return new DeliveryRenderBatchContext(
            reviewProfile,
            commentProfile,
            ComputeLayoutOptionsDigest(reviewProfile, commentProfile),
            ComputeRuntimePolicyDigest());
    }

    private VerificationDigest ComputeLayoutOptionsDigest(
        DeliveryReviewProfile reviewProfile,
        DeliveryCommentProfile commentProfile)
    {
        var buffer = new ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer))
        {
            writer.WriteStartObject();
            writer.WriteString("commentProfile", Name(commentProfile));
            writer.WriteString("contract",
                "docxodus.delivery/export-host-layout-options/v1");
            writer.WriteString("reviewProfile", Name(reviewProfile));
            writer.WriteBoolean("reviewProfileAlreadyApplied",
                UsesProfileResolvedSource(reviewProfile));
            writer.WriteBoolean("strictFonts", _options.StrictFonts);
            writer.WriteNumber("timeoutMs",
                checked((int)_options.RenderTimeout.TotalMilliseconds));
            writer.WriteString("title", string.Empty);
            writer.WriteString("unsupportedContent", Name(_options.UnsupportedContent));
            writer.WriteEndObject();
        }
        return Sha256Digest(buffer.WrittenSpan);
    }

    private VerificationDigest ComputeRuntimePolicyDigest()
    {
        var buffer = new ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer))
        {
            writer.WriteStartObject();
            writer.WriteString("browserSelection",
                string.IsNullOrWhiteSpace(_options.ChromiumExecutablePath)
                    ? "package-managed-chromium"
                    : "configured-executable");
            writer.WriteString("contract",
                "docxodus.delivery/export-host-runtime-policy/v1");
            writer.WriteStartArray("fontDirectories");
            writer.WriteEndArray();
            writer.WriteNumber("hostWireSchemaVersion", 1);
            writer.WriteNumber("maximumResponseBytes", _options.MaximumFrameBytes);
            writer.WriteNumber("renderTimeoutMs",
                checked((int)_options.RenderTimeout.TotalMilliseconds));
            writer.WriteString("resourceLimits", "defaults");
            writer.WriteEndObject();
        }
        return Sha256Digest(buffer.WrittenSpan);
    }

    private static VerificationDigest Sha256Digest(ReadOnlySpan<byte> bytes) => new()
    {
        Algorithm = "SHA-256",
        Value = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
    };

    public async ValueTask<IReadOnlyDictionary<string, DeliveryRenderResult>> RenderBatchesAsync(
        IReadOnlyList<DeliveryRenderBatch> batches,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(batches);
        cancellationToken.ThrowIfCancellationRequested();
        var snapshot = batches.ToArray();
        ValidateBatches(snapshot);
        var results = new Dictionary<string, DeliveryRenderResult>(StringComparer.Ordinal);
        var representable = new List<DeliveryRenderBatch>();
        foreach (var batch in snapshot)
        {
            // The framed host is a JavaScript runtime. A .NET version outside its safe-integer
            // range is typed per-batch unavailability decided here, before any host frame is
            // built, never a rounded number on the wire.
            if (batch.Requests[0].SourceDocumentVersion > MaximumSafeJavaScriptInteger)
            {
                foreach (var request in batch.Requests)
                    results.Add(request.ArtifactId, DeliveryRenderResult.Unavailable(
                        MediaType(request.Kind), UnrepresentableVersionReason));
            }
            else
            {
                representable.Add(batch);
            }
        }

        if (representable.Count == 0)
            return results;
        var plan = BuildHostFramePlan(representable);
        var response = await InvokeHostAsync(plan, representable.Count, cancellationToken)
            .ConfigureAwait(false);
        foreach (var pair in ParseResponse(response, plan.WireBatches))
            results.Add(pair.Key, pair.Value);
        return results;
    }

    private void ValidateBatches(IReadOnlyList<DeliveryRenderBatch> batches)
    {
        if (batches.Any(batch => batch is null))
            throw new ArgumentException("Render batches cannot contain null entries.", nameof(batches));
        if (batches.GroupBy(batch => batch.BatchId, StringComparer.Ordinal)
            .Any(group => group.Count() != 1))
            throw new ArgumentException("Render batch IDs must be unique.", nameof(batches));
        var artifactIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var batch in batches)
        {
            if (string.IsNullOrWhiteSpace(batch.BatchId))
                throw new ArgumentException("A render batch id is required.", nameof(batches));
            var context = batch.Context
                ?? throw new ArgumentException("A render batch context is required.", nameof(batches));
            if (!context.Equals(DescribeBatch(context.ReviewProfile, context.CommentProfile)))
                throw new ArgumentException(
                    $"Batch '{batch.BatchId}' carries a context this adapter did not describe.",
                    nameof(batches));
            if (batch.Requests is not { Count: > 0 })
                throw new ArgumentException(
                    $"Batch '{batch.BatchId}' carries no render requests.", nameof(batches));
            if (batch.Requests.Any(request => request is null))
                throw new ArgumentException("Render requests cannot contain null entries.", nameof(batches));
            var first = batch.Requests[0];
            foreach (var request in batch.Requests)
            {
                if (!artifactIds.Add(request.ArtifactId))
                    throw new ArgumentException("Render request IDs must be unique.", nameof(batches));
                if (!Capabilities.Supports(request.Kind, request.ReviewProfile, request.CommentProfile))
                    throw new ArgumentException(
                        $"Render request '{request.ArtifactId}' is not supported by this adapter.",
                        nameof(batches));
                if (request.ReviewProfile != context.ReviewProfile
                    || request.CommentProfile != context.CommentProfile)
                    throw new ArgumentException(
                        $"Render request '{request.ArtifactId}' does not match its batch context profiles.",
                        nameof(batches));
                if (!string.Equals(request.SourcePackageDigest.Value,
                        first.SourcePackageDigest.Value, StringComparison.Ordinal)
                    || request.SourceDocumentVersion != first.SourceDocumentVersion)
                    throw new ArgumentException(
                        $"Batch '{batch.BatchId}' mixes render requests from different exact sources.",
                        nameof(batches));
            }
        }
    }

    /// <summary>
    /// Builds the single framed host request for one <c>RenderBatchesAsync</c> call: one control
    /// frame declaring deduplicated sources and every batch, then one raw source frame per unique
    /// source. A source shared by several batches crosses the pipe once.
    /// </summary>
    internal HostFramePlan BuildHostFramePlan(IReadOnlyList<DeliveryRenderBatch> batches)
    {
        var sources = new List<(string Id, string Digest, byte[] Bytes)>();
        var sourcesByDigest = new Dictionary<string, string>(StringComparer.Ordinal);
        var wireBatches = new List<HostWireBatch>();
        foreach (var batch in batches)
        {
            var first = batch.Requests[0];
            var bytes = first.CopySourceBytes();
            var digest = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
            if (!string.Equals(digest, first.SourcePackageDigest.Value, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    $"Batch '{batch.BatchId}' declares a source digest that does not match its bytes.");
            if (!sourcesByDigest.TryGetValue(digest, out var sourceId))
            {
                sourceId = $"src-{sources.Count + 1:D4}";
                sourcesByDigest.Add(digest, sourceId);
                sources.Add((sourceId, digest, bytes));
            }
            wireBatches.Add(new HostWireBatch(
                batch,
                sourceId,
                digest,
                bytes.LongLength,
                batch.Requests.Select(request => request.ArtifactId)
                    .OrderBy(id => id, StringComparer.Ordinal)
                    .ToArray(),
                OutputsFor(batch.Requests)));
        }

        if (sources.Count > 64 || wireBatches.Count > 64
            || wireBatches.Sum(wire => wire.ArtifactRequestIds.Count) > 256)
            throw new InvalidDataException(
                "The delivery build exceeds the export host's source, batch, or artifact ceilings.");
        if (sources.Sum(source => source.Bytes.LongLength) > 536_870_912L)
            throw new InvalidDataException(
                "The delivery build exceeds the export host's aggregate source byte ceiling.");

        var envelope = new HostRequest(
            1,
            sources.Select(source => new HostSourceDescriptor(
                source.Id, source.Bytes.LongLength, source.Digest, DocxMediaType)).ToArray(),
            wireBatches.Select(wire => new HostBatchRequest(
                wire.Batch.BatchId,
                wire.SourceId,
                wire.ArtifactRequestIds,
                new HostBatchOptions(
                    wire.Outputs,
                    wire.DocumentVersion,
                    wire.SourceDigest,
                    Name(wire.ReviewProfile),
                    UsesProfileResolvedSource(wire.ReviewProfile) ? true : null,
                    Name(wire.CommentProfile),
                    Name(_options.UnsupportedContent),
                    _options.StrictFonts,
                    checked((int)_options.RenderTimeout.TotalMilliseconds))))
                .ToArray());
        var control = JsonSerializer.SerializeToUtf8Bytes(envelope, HostJson);
        if (control.Length > MaximumControlFrameBytes)
            throw new InvalidDataException(
                "The export-host control frame exceeds the host's frame limit.");
        return new HostFramePlan(
            control,
            sources.Select(source => source.Bytes).ToArray(),
            wireBatches);
    }

    private static IReadOnlyList<string> OutputsFor(IReadOnlyList<DeliveryRenderRequest> requests)
    {
        var outputs = new List<string>();
        if (requests.Any(request => request.Kind == DeliveryArtifactKind.StandaloneHtml)
            || !requests.Any(request => request.Kind is DeliveryArtifactKind.FinalPdf
                or DeliveryArtifactKind.ReviewPdf))
            outputs.Add("html");
        if (requests.Any(request => request.Kind is DeliveryArtifactKind.FinalPdf
                or DeliveryArtifactKind.ReviewPdf))
            outputs.Add("pdf");
        return outputs;
    }

    private async ValueTask<byte[]> InvokeHostAsync(
        HostFramePlan plan,
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
        var responseTask = ReadToEndAsync(process.StandardOutput.BaseStream, token);
        try
        {
            var input = process.StandardInput.BaseStream;
            await WriteFrameAsync(input, plan.ControlFrame, token).ConfigureAwait(false);
            foreach (var sourceFrame in plan.SourceFrames)
                await WriteFrameAsync(input, sourceFrame, token).ConfigureAwait(false);
            await input.FlushAsync(token).ConfigureAwait(false);
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

    private static async Task WriteFrameAsync(
        Stream stream,
        byte[] payload,
        CancellationToken cancellationToken)
    {
        var prefix = new byte[4];
        BinaryPrimitives.WriteUInt32BigEndian(prefix, checked((uint)payload.Length));
        await stream.WriteAsync(prefix, cancellationToken).ConfigureAwait(false);
        await stream.WriteAsync(payload, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads the host's entire response — one control frame followed by the artifact frames its
    /// descriptors declare — bounded by the configured response budget. The host writes its full
    /// response and exits, so end-of-stream is the frame boundary of last resort and the parser
    /// slices the individual frames afterwards.
    /// </summary>
    private async Task<byte[]> ReadToEndAsync(Stream stream, CancellationToken cancellationToken)
    {
        var budget = checked(_options.MaximumFrameBytes + MaximumControlFrameBytes);
        using var buffered = new MemoryStream();
        var chunk = new byte[81920];
        while (true)
        {
            var read = await stream.ReadAsync(chunk, cancellationToken).ConfigureAwait(false);
            if (read == 0)
                break;
            if (buffered.Length + read > budget)
                throw new InvalidDataException("The export host response exceeds the configured response budget.");
            buffered.Write(chunk, 0, read);
        }
        return buffered.ToArray();
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
        catch (InvalidOperationException)
        {
        }
    }

    private IReadOnlyDictionary<string, DeliveryRenderResult> ParseResponse(
        byte[] responseBytes,
        IReadOnlyList<HostWireBatch> wireBatches)
    {
        var cursor = new FrameCursor(responseBytes);
        var controlBytes = cursor.ReadFrame(MaximumControlFrameBytes);
        using var document = JsonDocument.Parse(controlBytes, new JsonDocumentOptions
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
            return ParseFatalResponse(root, fatal, cursor, wireBatches);

        var byId = RequiredArray(root, "batches").EnumerateArray().ToDictionary(
            batch => RequiredString(batch, "id"), StringComparer.Ordinal);
        if (!wireBatches.Select(wire => wire.Batch.BatchId).ToHashSet(StringComparer.Ordinal)
            .SetEquals(byId.Keys))
            throw new InvalidDataException(
                "The export host response does not match the requested batch IDs.");
        var artifacts = ReadArtifactFrames(RequiredArray(root, "artifacts"), cursor);
        cursor.AssertEnd();

        var results = new Dictionary<string, DeliveryRenderResult>(StringComparer.Ordinal);
        var claimed = new HashSet<string>(StringComparer.Ordinal);
        foreach (var wire in wireBatches)
        {
            var meta = byId[wire.Batch.BatchId];
            if (!string.Equals(RequiredString(meta, "sourceId"), wire.SourceId, StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The export host response rebound a batch to a different source.");
            var cohort = ParseSuccessfulBatch(wire, meta, artifacts, claimed);
            foreach (var request in wire.Batch.Requests)
                results.Add(request.ArtifactId, cohort.For(request));
        }
        if (claimed.Count != artifacts.Count)
            throw new InvalidDataException("The export host returned artifacts no batch declared.");
        return results;
    }

    /// <summary>
    /// A fatal host envelope fails every batch, but a schema-valid failed render report the host
    /// retained as a diagnostic artifact is preserved as evidence for the batch it is bound to —
    /// HTML, PDF, and PageMap stay unavailable — instead of being flattened into an exception.
    /// </summary>
    private IReadOnlyDictionary<string, DeliveryRenderResult> ParseFatalResponse(
        JsonElement root,
        JsonElement fatal,
        FrameCursor cursor,
        IReadOnlyList<HostWireBatch> wireBatches)
    {
        var reason = "Export host fatal failure: " + ErrorReason(fatal);
        var artifacts = root.TryGetProperty("diagnosticArtifacts", out var declared)
            ? ReadArtifactFrames(declared, cursor)
            : new Dictionary<string, HostArtifactFrame>(StringComparer.Ordinal);
        cursor.AssertEnd();

        HostWireBatch? reported = null;
        HostArtifactFrame? reportFrame = null;
        if (OptionalString(fatal, "reportArtifactId") is { } reportArtifactId
            && artifacts.TryGetValue(reportArtifactId, out var candidate)
            && string.Equals(candidate.Kind, "renderReport", StringComparison.Ordinal))
        {
            reported = wireBatches.FirstOrDefault(wire =>
                string.Equals(wire.Batch.BatchId, candidate.BatchId, StringComparison.Ordinal));
            reportFrame = reported is null ? null : candidate;
        }

        FailedRenderCohort? failedCohort = null;
        if (reported is not null && reportFrame is not null)
        {
            using var reportDocument = JsonDocument.Parse(reportFrame.Bytes);
            EnsureUniqueProperties(reportDocument.RootElement);
            var report = reportDocument.RootElement;
            ValidateFailedReport(reported, fatal, report);
            var fingerprint = report.TryGetProperty("environment", out var environment)
                              && environment.ValueKind == JsonValueKind.Object
                ? OptionalString(environment, "rendererFingerprint")
                : null;
            failedCohort = new FailedRenderCohort(
                reason,
                reportFrame.Bytes,
                fingerprint,
                ParseDiagnosticArray(RequiredArray(report, "warnings")));
        }

        var results = new Dictionary<string, DeliveryRenderResult>(StringComparer.Ordinal);
        foreach (var wire in wireBatches)
        {
            var isReportedBatch = failedCohort is not null && ReferenceEquals(wire, reported);
            foreach (var request in wire.Batch.Requests)
                results.Add(request.ArtifactId, isReportedBatch
                    ? failedCohort!.For(request)
                    : DeliveryRenderResult.Unavailable(MediaType(request.Kind), reason));
        }
        return results;
    }

    /// <summary>Reads the declared artifact frames in descriptor order, verifying each frame's
    /// exact length and SHA-256 against its descriptor before anything downstream trusts it.</summary>
    private Dictionary<string, HostArtifactFrame> ReadArtifactFrames(
        JsonElement descriptors,
        FrameCursor cursor)
    {
        if (descriptors.ValueKind != JsonValueKind.Array)
            throw new InvalidDataException("The export host artifact descriptors are malformed.");
        var artifacts = new Dictionary<string, HostArtifactFrame>(StringComparer.Ordinal);
        foreach (var descriptor in descriptors.EnumerateArray())
        {
            var id = RequiredString(descriptor, "id");
            var byteLength = RequiredInt64(descriptor, "byteLength");
            if (byteLength <= 0 || byteLength > _options.MaximumFrameBytes)
                throw new InvalidDataException(
                    $"The export host declared an invalid artifact length for '{id}'.");
            var bytes = cursor.ReadFrame(_options.MaximumFrameBytes, checked((int)byteLength));
            var digest = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
            if (!string.Equals(digest, RequiredString(descriptor, "sha256"), StringComparison.Ordinal))
                throw new InvalidDataException(
                    $"The export host artifact '{id}' does not match its declared digest.");
            if (!artifacts.TryAdd(id, new HostArtifactFrame(
                    RequiredString(descriptor, "batchId"),
                    RequiredString(descriptor, "kind"),
                    RequiredString(descriptor, "mediaType"),
                    bytes)))
                throw new InvalidDataException(
                    $"The export host declared artifact '{id}' more than once.");
        }
        return artifacts;
    }

    private static void ValidateFailedReport(
        HostWireBatch wire,
        JsonElement fatal,
        JsonElement report)
    {
        if (RequiredString(report, "schema") != RenderReportSchema
            || RequiredInt32(report, "schemaVersion") != 1
            || RequiredString(report, "status") != "failed")
            throw new InvalidDataException("The export host returned an invalid failed render report.");
        var source = RequiredObject(report, "source");
        if (!string.Equals(RequiredString(source, "rawPackageBytesDigest"),
                wire.SourceDigest, StringComparison.OrdinalIgnoreCase)
            || RequiredInt64(source, "documentVersion") != wire.DocumentVersion
            || RequiredInt64(source, "byteLength") != wire.SourceByteLength)
            throw new InvalidDataException("The failed render report is bound to different source bytes.");
        ValidateProfileBinding(wire, report, "failed render report");
        var failure = RequiredObject(report, "failure");
        foreach (var field in new[] { "code", "phase", "message", "remediation" })
        {
            // The host bounds long fatal text fields with a trailing ellipsis while the retained
            // report keeps the full value, so a bounded fatal field must agree as a prefix.
            var reportValue = RequiredString(failure, field);
            var fatalValue = RequiredString(fatal, field);
            var agrees = string.Equals(reportValue, fatalValue, StringComparison.Ordinal)
                || (fatalValue.EndsWith("...", StringComparison.Ordinal)
                    && reportValue.StartsWith(fatalValue[..^3], StringComparison.Ordinal));
            if (!agrees)
                throw new InvalidDataException(
                    $"The failed render report disagrees with its host error field '{field}'.");
        }
    }

    private RenderCohort ParseSuccessfulBatch(
        HostWireBatch wire,
        JsonElement meta,
        IReadOnlyDictionary<string, HostArtifactFrame> artifacts,
        ISet<string> claimed)
    {
        var fingerprint = RequiredString(meta, "rendererFingerprint");
        var pageCount = RequiredInt64(meta, "pageCount");
        if (pageCount <= 0)
            throw new InvalidDataException("The export host returned an invalid page count.");

        var expectedKinds = new List<string> { "pageMap", "renderReport" };
        if (wire.Outputs.Contains("html", StringComparer.Ordinal)) expectedKinds.Add("html");
        if (wire.Outputs.Contains("pdf", StringComparer.Ordinal)) expectedKinds.Add("pdf");
        var kindIds = RequiredObject(meta, "artifacts");
        if (!kindIds.EnumerateObject().Select(property => property.Name)
            .ToHashSet(StringComparer.Ordinal)
            .SetEquals(expectedKinds))
            throw new InvalidDataException(
                "The export host batch artifacts do not match the requested outputs.");
        var frames = new Dictionary<string, HostArtifactFrame>(StringComparer.Ordinal);
        foreach (var kind in expectedKinds)
        {
            var artifactId = RequiredString(kindIds, kind);
            if (!artifacts.TryGetValue(artifactId, out var frame)
                || !string.Equals(frame.BatchId, wire.Batch.BatchId, StringComparison.Ordinal)
                || !string.Equals(frame.Kind, kind, StringComparison.Ordinal)
                || !claimed.Add(artifactId))
                throw new InvalidDataException(
                    $"The export host batch artifact '{artifactId}' is missing or misbound.");
            frames.Add(kind, frame);
        }

        var pageMapBytes = frames["pageMap"].Bytes;
        using var pageMapDocument = JsonDocument.Parse(pageMapBytes);
        var pageMap = DocxSessionJson.ParsePageMap(pageMapDocument.RootElement);
        var portable = PageMapContract.ValidatePortable(pageMap);
        if (!portable.Success || pageMap.Mode != PageMapMode.Paginated
            || pageMap.Availability != PageMapAvailability.Available)
            throw new InvalidDataException(portable.Message
                ?? "The export host returned a non-portable PageMap.");
        if (pageMap.DocumentVersion != wire.DocumentVersion
            || pageMap.Pages.Count != pageCount
            || !string.Equals(pageMap.RendererFingerprint, fingerprint,
                StringComparison.Ordinal))
            throw new InvalidDataException("The export host returned incoherent PageMap metadata.");

        var reportBytes = frames["renderReport"].Bytes;
        using var reportDocument = JsonDocument.Parse(reportBytes);
        EnsureUniqueProperties(reportDocument.RootElement);
        var reportElement = reportDocument.RootElement;
        ValidateReport(wire, reportElement, fingerprint, pageCount, pageMapBytes);

        byte[]? htmlBytes = null;
        if (frames.TryGetValue("html", out var htmlFrame))
        {
            htmlBytes = htmlFrame.Bytes;
            EnsureDigest(reportElement, "htmlDigest", htmlBytes);
        }
        byte[]? pdfBytes = null;
        if (frames.TryGetValue("pdf", out var pdfFrame))
        {
            pdfBytes = pdfFrame.Bytes;
            EnsureDigest(reportElement, "pdfDigest", pdfBytes);
        }
        return new RenderCohort(
            htmlBytes,
            pdfBytes,
            pageMapBytes,
            reportBytes,
            fingerprint,
            pageCount,
            ParseDiagnosticArray(RequiredArray(reportElement, "warnings")));
    }

    private static void ValidateReport(
        HostWireBatch wire,
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
        if (!string.Equals(RequiredString(source, "rawPackageBytesDigest"),
                wire.SourceDigest, StringComparison.OrdinalIgnoreCase)
            || RequiredInt64(source, "documentVersion") != wire.DocumentVersion
            || RequiredInt64(source, "byteLength") != wire.SourceByteLength)
            throw new InvalidDataException("The render report is bound to different source bytes.");
        ValidateProfileBinding(wire, report, "render report");
        var environment = RequiredObject(report, "environment");
        if (!string.Equals(RequiredString(environment, "rendererFingerprint"), fingerprint,
                StringComparison.Ordinal))
            throw new InvalidDataException("The render report fingerprint is incoherent.");
        if (RequiredArray(report, "pages").GetArrayLength() != pageCount)
            throw new InvalidDataException("The render report page count is incoherent.");
        EnsureDigest(report, "pageMapDigest", pageMapBytes);
        // The host stamps the batch's artifact request IDs into the report before canonical
        // serialization; a report that does not carry exactly this batch's IDs belongs to a
        // different render and must not be projected into these artifacts.
        var bindings = RequiredObject(report, "bindings");
        var boundIds = RequiredArray(bindings, "artifactRequestIds").EnumerateArray()
            .Select(element => element.ValueKind == JsonValueKind.String
                ? element.GetString()!
                : throw new InvalidDataException("The render report artifact bindings are malformed."))
            .ToArray();
        if (!boundIds.SequenceEqual(wire.ArtifactRequestIds, StringComparer.Ordinal))
            throw new InvalidDataException(
                "The render report is not bound to this batch's artifact request IDs.");
    }

    private static IReadOnlyList<DeliverableRenderDiagnostic> ParseDiagnosticArray(
        JsonElement warnings) => warnings.EnumerateArray().Select(warning =>
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
        HostWireBatch wire,
        JsonElement report,
        string label)
    {
        var options = RequiredObject(report, "options");
        if (RequiredString(options, "reviewProfile") != Name(wire.ReviewProfile)
            || RequiredString(options, "commentProfile") != Name(wire.CommentProfile))
            throw new InvalidDataException($"The {label} profile binding is invalid.");

        var expectsResolvedSource = UsesProfileResolvedSource(wire.ReviewProfile);
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
        bool allowEmpty = false) =>
        value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String
        && (allowEmpty || !string.IsNullOrWhiteSpace(property.GetString()))
            ? property.GetString()!
            : throw new InvalidDataException($"The export host response is missing string '{name}'.");

    private static string? OptionalString(JsonElement value, string name) =>
        value.TryGetProperty(name, out var property) && property.ValueKind == JsonValueKind.String
            ? property.GetString()
            : null;

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

    private static bool UsesProfileResolvedSource(DeliveryReviewProfile profile) =>
        profile is DeliveryReviewProfile.Final or DeliveryReviewProfile.Original;

    /// <summary>One wire batch: the delivery batch plus its host source binding and request IDs.</summary>
    internal sealed record HostWireBatch(
        DeliveryRenderBatch Batch,
        string SourceId,
        string SourceDigest,
        long SourceByteLength,
        IReadOnlyList<string> ArtifactRequestIds,
        IReadOnlyList<string> Outputs)
    {
        internal long DocumentVersion => Batch.Requests[0].SourceDocumentVersion;

        internal DeliveryReviewProfile ReviewProfile => Batch.Context.ReviewProfile;

        internal DeliveryCommentProfile CommentProfile => Batch.Context.CommentProfile;
    }

    /// <summary>The single framed request one <c>RenderBatchesAsync</c> call sends the host.</summary>
    internal sealed record HostFramePlan(
        byte[] ControlFrame,
        IReadOnlyList<byte[]> SourceFrames,
        IReadOnlyList<HostWireBatch> WireBatches);

    private sealed record HostArtifactFrame(
        string BatchId,
        string Kind,
        string MediaType,
        byte[] Bytes);

    /// <summary>Slices length-prefixed frames out of the fully buffered host response.</summary>
    private sealed class FrameCursor
    {
        private readonly byte[] _bytes;
        private int _offset;

        internal FrameCursor(byte[] bytes) => _bytes = bytes;

        internal byte[] ReadFrame(int maximumBytes, int? exactBytes = null)
        {
            if (_offset + 4 > _bytes.Length)
                throw new InvalidDataException("The export host response ended inside a frame header.");
            var length = BinaryPrimitives.ReadUInt32BigEndian(_bytes.AsSpan(_offset, 4));
            _offset += 4;
            if (length == 0 || length > (uint)maximumBytes
                || (exactBytes is { } expected && length != (uint)expected))
                throw new InvalidDataException("The export host returned an invalid frame length.");
            if (_offset + length > (uint)_bytes.Length)
                throw new InvalidDataException("The export host response ended inside a frame.");
            var payload = _bytes.AsSpan(_offset, checked((int)length)).ToArray();
            _offset += checked((int)length);
            return payload;
        }

        internal void AssertEnd()
        {
            if (_offset != _bytes.Length)
                throw new InvalidDataException("The export host returned bytes after its response frames.");
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

    private sealed record HostRequest(
        int SchemaVersion,
        IReadOnlyList<HostSourceDescriptor> Sources,
        IReadOnlyList<HostBatchRequest> Batches);

    private sealed record HostSourceDescriptor(
        string Id,
        long ByteLength,
        string Sha256,
        string MediaType);

    private sealed record HostBatchRequest(
        string Id,
        string SourceId,
        IReadOnlyList<string> ArtifactRequestIds,
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
