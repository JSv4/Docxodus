#nullable enable

using System.Text.Json;
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
            if (!File.Exists(parsed.BaselinePath))
                throw new CliUsageException($"Baseline file not found: {parsed.BaselinePath}");
            if (!File.Exists(parsed.WorkingPath))
                throw new CliUsageException($"Working file not found: {parsed.WorkingPath}");

            var request = new DeliveryBundleBuildRequest(
                new DeliveryDocumentSnapshot(
                    "baseline:" + Path.GetFileName(parsed.BaselinePath),
                    parsed.BaselineVersion,
                    await File.ReadAllBytesAsync(parsed.BaselinePath, cancellationToken)
                        .ConfigureAwait(false)),
                new DeliveryDocumentSnapshot(
                    "working:" + Path.GetFileName(parsed.WorkingPath),
                    parsed.WorkingVersion,
                    await File.ReadAllBytesAsync(parsed.WorkingPath, cancellationToken)
                        .ConfigureAwait(false)),
                parsed.FinalName,
                parsed.FinalVersion,
                new DeliveryBundleRevisionPolicy
                {
                    PreExistingRevisions = parsed.PreExistingPolicy,
                    GeneratedRevisions = parsed.GeneratedPolicy,
                },
                parsed.Artifacts);
            var bundle = await new DeliveryBundleService().BuildAsync(
                request,
                new DeliveryBundleBuildOptions
                {
                    ReturnIncompleteBundle = parsed.ReturnIncomplete,
                    FailOnDeliverableValidationFailure = parsed.FailOnValidationFailure,
                },
                cancellationToken).ConfigureAwait(false);
            var published = DeliveryBundleDirectoryPublisher.Publish(
                bundle, parsed.OutputDirectory);
            await output.WriteLineAsync(JsonSerializer.Serialize(new
            {
                status = Name(bundle.Manifest.Payload.Status),
                verified = bundle.Verification.IsValid,
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
                                   or ArgumentException or UnauthorizedAccessException)
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
                artifacts.Add(ParseArtifact(artifact));
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
            positional[2],
            baselineVersion.Value,
            workingVersion ?? finalVersion.Value,
            finalVersion.Value,
            finalName,
            preExisting.Value,
            generated.Value,
            artifacts,
            returnIncomplete,
            failOnValidation);
    }

    private static DeliveryArtifactRequest ParseArtifact(string value)
    {
        var parts = value.Split(':');
        if (parts.Length is not (3 or 5))
            throw new CliUsageException(
                "--artifact must be id:kind:requiredness[:review-profile:comment-profile].");
        var kind = EnumValue<DeliveryArtifactKind>(parts[1], "artifact kind");
        bool isRender = kind is DeliveryArtifactKind.StandaloneHtml
            or DeliveryArtifactKind.FinalPdf
            or DeliveryArtifactKind.ReviewPdf
            or DeliveryArtifactKind.PageMap
            or DeliveryArtifactKind.RenderReport;
        if (isRender != (parts.Length == 5))
            throw new CliUsageException(
                isRender
                    ? $"Render artifact '{parts[0]}' requires review and comment profiles."
                    : $"Non-render artifact '{parts[0]}' cannot select render profiles.");
        return new DeliveryArtifactRequest
        {
            ArtifactId = NonBlank(parts[0], "artifact id"),
            Kind = kind,
            Requiredness = EnumValue<DeliveryArtifactRequiredness>(
                parts[2], "artifact requiredness"),
            ReviewProfile = isRender
                ? EnumValue<DeliveryReviewProfile>(parts[3], "review profile")
                : null,
            CommentProfile = isRender
                ? EnumValue<DeliveryCommentProfile>(parts[4], "comment profile")
                : null,
        };
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

    private static string NonBlank(string value, string name) =>
        !string.IsNullOrWhiteSpace(value)
            ? value
            : throw new CliUsageException($"{name} cannot be blank.");

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
          --allow-validation-failure   Record comprehensive validation without failing delivery.

        Render profiles are final, original, or markup; comment profiles are hidden, inline,
        endnotes, or margin. HTML/PDF rendering requires an adapter from epic #434. Change-receipt
        output requires authoritative transaction evidence through the programmatic API.

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
        bool FailOnValidationFailure);

    private sealed class CliUsageException : Exception
    {
        internal CliUsageException(string message)
            : base(message)
        {
        }
    }
}
