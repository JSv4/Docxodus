#nullable enable

using System.Text.Json;
using System.Security.Cryptography;
using Docxodus.Delivery;

namespace Docxodus.DeliveryCli;

/// <summary>Thin command-line adapter over <see cref="DeliveryBundleService"/>.</summary>
internal static class DeliveryCli
{
    internal static async Task<int> RunAsync(
        string[] args,
        TextWriter output,
        TextWriter error,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(args);
        ArgumentNullException.ThrowIfNull(output);
        ArgumentNullException.ThrowIfNull(error);
        if (args.Length == 0 || args.Any(value => value is "-h" or "--help"))
        {
            await output.WriteLineAsync(Usage()).ConfigureAwait(false);
            return 0;
        }

        try
        {
            var parsed = Parse(args);
            var baselineBytes = await ReadStableInputAsync(
                parsed.BaselinePath, "Baseline",
                DeliveryArtifactRequestRules.MaximumInputPackageBytes,
                cancellationToken).ConfigureAwait(false);
            var workingBytes = await ReadStableInputAsync(
                parsed.WorkingPath, "Working",
                DeliveryArtifactRequestRules.MaximumInputPackageBytes - baselineBytes.LongLength,
                cancellationToken).ConfigureAwait(false);

            var request = new DeliveryBundleBuildRequest(
                new DeliveryDocumentSnapshot(
                    "baseline:" + Path.GetFileName(parsed.BaselinePath),
                    parsed.BaselineVersion,
                    baselineBytes),
                new DeliveryDocumentSnapshot(
                    "working:" + Path.GetFileName(parsed.WorkingPath),
                    parsed.WorkingVersion,
                    workingBytes),
                parsed.FinalName,
                parsed.FinalVersion,
                new DeliveryBundleRevisionPolicy
                {
                    PreExistingRevisions = parsed.PreExistingPolicy,
                    GeneratedRevisions = parsed.GeneratedPolicy,
                },
                parsed.Artifacts);
            var renderer = CreateRenderer(parsed);
            var bundle = await new DeliveryBundleService(renderer).BuildAsync(
                request,
                new DeliveryBundleBuildOptions
                {
                    ReturnIncompleteBundle = parsed.ReturnIncomplete,
                    FailOnDeliverableValidationFailure = parsed.FailOnValidationFailure,
                },
                cancellationToken).ConfigureAwait(false);
            var published = DeliveryBundleDirectoryPublisher.Publish(
                bundle,
                parsed.OutputDirectory,
                publicationOptions: new DeliveryBundleDirectoryPublicationOptions
                {
                    AllowDiagnosticBundle = parsed.ReturnIncomplete,
                });
            await output.WriteLineAsync(JsonSerializer.Serialize(new
            {
                status = Name(bundle.Manifest.Payload.Status),
                verified = bundle.IsVerifiedDelivery,
                manifestVerified = bundle.Verification.IsValid,
                deliverableDecision = bundle.DeliverableDecision is { } decision
                    ? Name(decision)
                    : null,
                outputDirectory = published,
                manifestPath = Path.Combine(published, DeliveryBundle.ManifestFileName),
                manifestDigest = bundle.Manifest.ManifestDigest.Value,
                artifactCount = bundle.Manifest.Payload.Artifacts.Count,
            })).ConfigureAwait(false);
            return 0;
        }
        catch (CliUsageException ex)
        {
            await error.WriteLineAsync($"usage_error: {ex.Message}").ConfigureAwait(false);
            return 2;
        }
        catch (DeliveryBundleException ex)
        {
            await error.WriteLineAsync($"{ex.Code}: {ex.Message}").ConfigureAwait(false);
            return 1;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException
                                   or ArgumentException or InvalidOperationException
                                   or UnauthorizedAccessException)
        {
            await error.WriteLineAsync($"delivery_failed: {ex.Message}").ConfigureAwait(false);
            return 1;
        }
    }

    private static ParsedRequest Parse(IReadOnlyList<string> args)
    {
        var positional = new List<string>();
        var artifacts = new List<DeliveryArtifactRequest>();
        DeliveryRevisionPolicy? preExisting = null;
        DeliveryRevisionPolicy? generated = null;
        long? baselineVersion = null;
        long? workingVersion = null;
        long? finalVersion = null;
        string? finalName = null;
        string? nodeExecutable = null;
        string? exportHost = null;
        string? chromiumExecutable = null;
        int? renderTimeoutMilliseconds = null;
        DeliveryUnsupportedContentPolicy unsupportedContent =
            DeliveryUnsupportedContentPolicy.Warn;
        bool strictFonts = false;
        bool returnIncomplete = false;
        bool failOnValidation = true;

        foreach (var argument in args)
        {
            if (!argument.StartsWith("--", StringComparison.Ordinal))
            {
                positional.Add(argument);
                continue;
            }
            if (Value(argument, "--artifact=") is { } artifact)
            {
                if (artifacts.Count >= DeliveryArtifactRequestRules.MaximumArtifactCount)
                    throw new CliUsageException(
                        $"At most {DeliveryArtifactRequestRules.MaximumArtifactCount} artifacts may be requested.");
                artifacts.Add(ParseArtifact(artifact));
            }
            else if (Value(argument, "--pre-existing=") is { } preExistingName)
                preExisting = One(preExisting, RevisionPolicy(preExistingName), "--pre-existing");
            else if (Value(argument, "--generated=") is { } generatedName)
                generated = One(generated, RevisionPolicy(generatedName), "--generated");
            else if (Value(argument, "--baseline-version=") is { } baselineVersionText)
                baselineVersion = One(baselineVersion,
                    NonNegativeLong(baselineVersionText, "--baseline-version"),
                    "--baseline-version");
            else if (Value(argument, "--working-version=") is { } workingVersionText)
                workingVersion = One(workingVersion,
                    NonNegativeLong(workingVersionText, "--working-version"),
                    "--working-version");
            else if (Value(argument, "--final-version=") is { } finalVersionText)
                finalVersion = One(finalVersion,
                    NonNegativeLong(finalVersionText, "--final-version"),
                    "--final-version");
            else if (Value(argument, "--final-name=") is { } finalNameValue)
                finalName = One(finalName, NonBlank(finalNameValue, "--final-name"), "--final-name");
            else if (Value(argument, "--node-executable=") is { } nodeValue)
                nodeExecutable = One(nodeExecutable,
                    NonBlank(nodeValue, "--node-executable"), "--node-executable");
            else if (Value(argument, "--export-host=") is { } hostValue)
                exportHost = One(exportHost,
                    NonBlank(hostValue, "--export-host"), "--export-host");
            else if (Value(argument, "--chromium-executable=") is { } chromiumValue)
                chromiumExecutable = One(chromiumExecutable,
                    NonBlank(chromiumValue, "--chromium-executable"), "--chromium-executable");
            else if (Value(argument, "--render-timeout=") is { } timeoutValue)
                renderTimeoutMilliseconds = One(
                    renderTimeoutMilliseconds,
                    PositiveInt(timeoutValue, "--render-timeout", maximum: 600_000),
                    "--render-timeout");
            else if (Value(argument, "--unsupported-content=") is { } unsupportedValue)
                unsupportedContent = EnumValue<DeliveryUnsupportedContentPolicy>(
                    unsupportedValue, "unsupported-content policy");
            else if (argument == "--strict-fonts")
                strictFonts = true;
            else if (argument == "--return-incomplete")
                returnIncomplete = true;
            else if (argument == "--allow-validation-failure")
                failOnValidation = false;
            else
                throw new CliUsageException($"Unknown option: {argument}");
        }

        if (positional.Count != 3)
            throw new CliUsageException(
                "Expected <baseline.docx> <working.docx> <new-output-directory>.");
        if (artifacts.Count == 0)
            throw new CliUsageException("At least one --artifact selection is required.");
        if (preExisting is null || generated is null)
            throw new CliUsageException(
                "Both --pre-existing and --generated revision policies are required.");
        if (baselineVersion is null || finalVersion is null)
            throw new CliUsageException(
                "Both --baseline-version and --final-version are required.");
        if (finalName is null)
            throw new CliUsageException("--final-name is required.");
        return new ParsedRequest(
            positional[0],
            positional[1],
            Path.GetFullPath(positional[2]),
            baselineVersion.Value,
            workingVersion ?? finalVersion.Value,
            finalVersion.Value,
            finalName,
            preExisting.Value,
            generated.Value,
            artifacts,
            returnIncomplete,
            failOnValidation,
            nodeExecutable,
            exportHost,
            chromiumExecutable,
            renderTimeoutMilliseconds ?? 120_000,
            unsupportedContent,
            strictFonts);
    }

    private static DeliveryArtifactRequest ParseArtifact(string value)
    {
        var parts = value.Split(':');
        DeliveryArtifactKind kind;
        DeliveryArtifactRequiredness requiredness;
        DeliveryReviewProfile? review = null;
        DeliveryCommentProfile? comments = null;
        string artifactId;
        if (parts.Length >= 5
            && TryEnumValue(parts[^4], out kind)
            && DeliveryArtifactRequestRules.IsProfiledRenderKind(kind))
        {
            artifactId = string.Join(':', parts[..^4]);
            requiredness = EnumValue<DeliveryArtifactRequiredness>(
                parts[^3], "artifact requiredness");
            review = EnumValue<DeliveryReviewProfile>(parts[^2], "review profile");
            comments = EnumValue<DeliveryCommentProfile>(parts[^1], "comment profile");
        }
        else if (parts.Length >= 3 && TryEnumValue(parts[^2], out kind))
        {
            artifactId = string.Join(':', parts[..^2]);
            requiredness = EnumValue<DeliveryArtifactRequiredness>(
                parts[^1], "artifact requiredness");
        }
        else
        {
            throw new CliUsageException(
                "--artifact must be id:kind:requiredness[:review-profile:comment-profile].");
        }

        var request = new DeliveryArtifactRequest
        {
            ArtifactId = NonBlank(artifactId, "artifact id"),
            Kind = kind,
            Requiredness = requiredness,
            ReviewProfile = review,
            CommentProfile = comments,
        };
        try
        {
            DeliveryArtifactRequestRules.ValidateProfileSelection(request);
        }
        catch (ArgumentException exception)
        {
            throw new CliUsageException(exception.Message);
        }
        return request;
    }

    private static string? Value(string argument, string prefix) =>
        argument.StartsWith(prefix, StringComparison.Ordinal)
            ? argument[prefix.Length..]
            : null;

    private static T One<T>(T? existing, T value, string option)
        where T : struct => existing is null
            ? value
            : throw new CliUsageException($"Option {option} can be supplied only once.");

    private static string One(string? existing, string value, string option) => existing is null
        ? value
        : throw new CliUsageException($"Option {option} can be supplied only once.");

    private static long NonNegativeLong(string value, string option) =>
        long.TryParse(value, out var number) && number >= 0
            ? number
            : throw new CliUsageException($"Option {option} requires a non-negative integer.");

    private static int PositiveInt(string value, string option, int maximum) =>
        int.TryParse(value, out var number) && number > 0 && number <= maximum
            ? number
            : throw new CliUsageException(
                $"Option {option} requires an integer from 1 through {maximum}.");

    private static string NonBlank(string value, string name) =>
        !string.IsNullOrWhiteSpace(value)
        && value.Length <= DeliveryArtifactRequestRules.MaximumStringLength
        && !value.Any(char.IsControl)
            ? value
            : throw new CliUsageException(
                $"{name} must be non-blank, control-free, and at most "
                + $"{DeliveryArtifactRequestRules.MaximumStringLength} characters.");

    private static DeliveryRevisionPolicy RevisionPolicy(string value) =>
        EnumValue<DeliveryRevisionPolicy>(value, "revision policy");

    private static T EnumValue<T>(string value, string name)
        where T : struct, Enum
    {
        var compact = value.Replace("-", string.Empty, StringComparison.Ordinal)
            .Replace("_", string.Empty, StringComparison.Ordinal);
        foreach (var candidate in Enum.GetValues<T>())
        {
            if (string.Equals(candidate.ToString(), compact, StringComparison.OrdinalIgnoreCase))
                return candidate;
        }
        throw new CliUsageException($"Unknown {name}: {value}");
    }

    private static bool TryEnumValue<T>(string value, out T result)
        where T : struct, Enum
    {
        var compact = value.Replace("-", string.Empty, StringComparison.Ordinal)
            .Replace("_", string.Empty, StringComparison.Ordinal);
        foreach (var candidate in Enum.GetValues<T>())
        {
            if (!string.Equals(candidate.ToString(), compact, StringComparison.OrdinalIgnoreCase))
                continue;
            result = candidate;
            return true;
        }
        result = default;
        return false;
    }

    private static string Name<T>(T value)
        where T : struct, Enum =>
        JsonNamingPolicy.CamelCase.ConvertName(value.ToString());

    private static string Usage() =>
        """
        docxodus-deliver - build one verified document delivery bundle

        Usage:
          docxodus-deliver <baseline.docx> <working.docx> <new-output-directory> \
            --baseline-version=N --final-version=N --final-name=NAME \
            --pre-existing=<preserve|accept|reject> \
            --generated=<preserve|accept|reject> \
            --artifact=id:kind:requiredness[:review-profile:comment-profile] [...]

        Options:
          --working-version=N          Defaults to final-version.
          --return-incomplete          Return an explicitly incomplete bundle when a required
                                       adapter-backed artifact is unavailable.
          --allow-validation-failure   Publish diagnostic validation failures; verified remains
                                       false and deliverableDecision reports the failed decision.
          --node-executable=PATH       Absolute Node.js executable for @docxodus/export.
          --export-host=PATH           Absolute path to the built framed host script.
          --chromium-executable=PATH   Optional explicit Chromium executable.
          --render-timeout=MS          Per-render deadline, from 1 through 600000.
          --unsupported-content=MODE   warn or strict (default: warn).
          --strict-fonts               Fail unless the export runtime proves exact font identity.

        Render profiles are final, original, or markup; comment profiles are hidden, inline,
        endnotes, or margin. Configure HTML/PDF rendering with the explicit options above or the
        DOCXODUS_NODE_PATH and DOCXODUS_EXPORT_HOST_PATH process environment variables.
        Change-receipt output requires authoritative transaction evidence through the programmatic API.

        Example:
          docxodus-deliver baseline.docx working.docx delivery \
            --baseline-version=0 --final-version=1 --final-name=final \
            --pre-existing=preserve --generated=accept \
            --artifact=final:final-docx:required \
            --artifact=validation:validation-report:required \
            --artifact=html:standalone-html:optional:final:endnotes
        """;

    private sealed record ParsedRequest(
        string BaselinePath,
        string WorkingPath,
        string OutputDirectory,
        long BaselineVersion,
        long WorkingVersion,
        long FinalVersion,
        string FinalName,
        DeliveryRevisionPolicy PreExistingPolicy,
        DeliveryRevisionPolicy GeneratedPolicy,
        IReadOnlyList<DeliveryArtifactRequest> Artifacts,
        bool ReturnIncomplete,
        bool FailOnValidationFailure,
        string? NodeExecutablePath,
        string? ExportHostPath,
        string? ChromiumExecutablePath,
        int RenderTimeoutMilliseconds,
        DeliveryUnsupportedContentPolicy UnsupportedContent,
        bool StrictFonts);

    private static IDeliveryArtifactRenderer? CreateRenderer(ParsedRequest request)
    {
        var environment = request.NodeExecutablePath is null || request.ExportHostPath is null
            ? DocxodusExportHostRendererOptions.FromEnvironment()
            : null;
        var node = request.NodeExecutablePath ?? environment?.NodeExecutablePath;
        var host = request.ExportHostPath ?? environment?.HostScriptPath;
        if (node is null && host is null && request.ChromiumExecutablePath is null)
            return null;
        if (node is null || host is null)
            throw new CliUsageException(
                "Rendering requires both --node-executable and --export-host (or their environment variables).");
        var chromium = request.ChromiumExecutablePath
            ?? environment?.ChromiumExecutablePath
            ?? Environment.GetEnvironmentVariable(
                DocxodusExportHostRendererOptions.ChromiumPathEnvironmentVariable);
        return new DocxodusExportHostRenderer(new DocxodusExportHostRendererOptions
        {
            NodeExecutablePath = Path.GetFullPath(node),
            HostScriptPath = Path.GetFullPath(host),
            ChromiumExecutablePath = chromium is null ? null : Path.GetFullPath(chromium),
            RenderTimeout = TimeSpan.FromMilliseconds(request.RenderTimeoutMilliseconds),
            UnsupportedContent = request.UnsupportedContent,
            StrictFonts = request.StrictFonts,
        });
    }

    private static async Task<byte[]> ReadStableInputAsync(
        string path,
        string label,
        long maximumBytes,
        CancellationToken cancellationToken)
    {
        try
        {
            await using var stream = new FileStream(path, new FileStreamOptions
            {
                Mode = FileMode.Open,
                Access = FileAccess.Read,
                Share = FileShare.Read,
                BufferSize = 64 * 1024,
                Options = FileOptions.Asynchronous | FileOptions.SequentialScan,
            });
            if (maximumBytes <= 0 || stream.Length > maximumBytes || stream.Length > int.MaxValue)
                throw new IOException(
                    $"{label} file exceeds the {maximumBytes}-byte remaining input budget: {path}");

            var length = checked((int)stream.Length);
            var firstDigest = await SHA256.HashDataAsync(stream, cancellationToken)
                .ConfigureAwait(false);
            if (stream.Length != length)
                throw new IOException($"{label} file changed while it was being snapshotted: {path}");

            stream.Position = 0;
            var bytes = GC.AllocateUninitializedArray<byte>(length);
            await stream.ReadExactlyAsync(bytes, cancellationToken).ConfigureAwait(false);
            if (await stream.ReadAsync(new byte[1], cancellationToken).ConfigureAwait(false) != 0
                || stream.Length != length
                || !CryptographicOperations.FixedTimeEquals(
                    firstDigest, SHA256.HashData(bytes)))
                throw new IOException($"{label} file changed while it was being snapshotted: {path}");
            return bytes;
        }
        catch (FileNotFoundException)
        {
            throw new CliUsageException($"{label} file not found: {path}");
        }
        catch (DirectoryNotFoundException)
        {
            throw new CliUsageException($"{label} file not found: {path}");
        }
    }

    private sealed class CliUsageException : Exception
    {
        internal CliUsageException(string message)
            : base(message)
        {
        }
    }
}
