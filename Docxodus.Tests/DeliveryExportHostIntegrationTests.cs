// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Delivery;
using Docxodus.Verification;
using Xunit;
using BundleArtifactAvailability = Docxodus.Delivery.DeliveryArtifactAvailability;

namespace Docxodus.Tests;

/// <summary>
/// Opt-in cross-runtime acceptance test. The #434 workflow sets the environment gate and uploads
/// the retained artifact directory with <c>if: always()</c>; ordinary .NET-only runs stay fast and
/// do not acquire Node or Chromium implicitly.
/// </summary>
public sealed class DeliveryExportHostIntegrationTests
{
    [Fact]
    public async Task ExportHost_EveryArtifact_PublishesViewableIndependentlyReopenedBundle()
    {
        if (!string.Equals(Environment.GetEnvironmentVariable(
                "DOCXODUS_RUN_DELIVERY_EXPORT_HOST"), "1", StringComparison.Ordinal))
            return;

        var nodePath = RequiredEnvironment(
            DocxodusExportHostRendererOptions.NodePathEnvironmentVariable);
        var hostPath = RequiredEnvironment(
            DocxodusExportHostRendererOptions.HostPathEnvironmentVariable);
        var artifactRoot = RequiredEnvironment("DOCXODUS_DELIVERY_REAL_ARTIFACT_DIR");
        var runDirectory = Path.Combine(
            Path.GetFullPath(artifactRoot),
            $"run-{DateTime.UtcNow:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(runDirectory);
        var checkpoint = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            ["schemaVersion"] = 1,
            ["phase"] = "fixture_setup",
            ["complete"] = false,
        };
        DeliveryBundle? bundle = null;
        string? published = null;
        try
        {
            var fixture = CreateTrackedEdit();
            await File.WriteAllBytesAsync(
                Path.Combine(runDirectory, "baseline.docx"), fixture.BaselineBytes);
            await File.WriteAllBytesAsync(
                Path.Combine(runDirectory, "working.docx"), fixture.WorkingBytes);
            var requests = Enum.GetValues<DeliveryArtifactKind>()
                .Select(kind => kind switch
                {
                    DeliveryArtifactKind.StandaloneHtml => Render(
                        "standalone-html", kind, DeliveryReviewProfile.Final,
                        DeliveryCommentProfile.Endnotes),
                    DeliveryArtifactKind.FinalPdf => Render(
                        "final-pdf", kind, DeliveryReviewProfile.Final,
                        DeliveryCommentProfile.Endnotes),
                    DeliveryArtifactKind.ReviewPdf => Render(
                        "review-pdf", kind, DeliveryReviewProfile.Markup,
                        DeliveryCommentProfile.Endnotes),
                    DeliveryArtifactKind.PageMap => Render(
                        "page-map", kind, DeliveryReviewProfile.Final,
                        DeliveryCommentProfile.Endnotes),
                    DeliveryArtifactKind.RenderReport => Render(
                        "render-report", kind, DeliveryReviewProfile.Final,
                        DeliveryCommentProfile.Endnotes),
                    _ => new DeliveryArtifactRequest
                    {
                        ArtifactId = Kebab(kind),
                        Kind = kind,
                        Requiredness = DeliveryArtifactRequiredness.Required,
                    },
                })
                .Concat(
                    from review in Enum.GetValues<DeliveryReviewProfile>()
                    from comments in Enum.GetValues<DeliveryCommentProfile>()
                    where review != DeliveryReviewProfile.Final
                          || comments != DeliveryCommentProfile.Endnotes
                    select Render(
                        $"standalone-html-{Kebab(review)}-{Kebab(comments)}",
                        DeliveryArtifactKind.StandaloneHtml,
                        review,
                        comments))
                .ToArray();
            await WriteJsonAsync(Path.Combine(runDirectory, "request.json"), new
            {
                schemaVersion = 1,
                revisionPolicy = new { preExisting = "preserve", generated = "accept" },
                artifacts = requests.Select(request => new
                {
                    request.ArtifactId,
                    kind = Name(request.Kind),
                    requiredness = Name(request.Requiredness),
                    reviewProfile = request.ReviewProfile is { } review ? Name(review) : null,
                    commentProfile = request.CommentProfile is { } comments ? Name(comments) : null,
                }),
            });
            checkpoint["phase"] = "bundle_build";
            await WriteJsonAsync(Path.Combine(runDirectory, "checkpoint.json"), checkpoint);

            var request = new DeliveryBundleBuildRequest(
                new DeliveryDocumentSnapshot(
                    "baseline", fixture.BaselineVersion, fixture.BaselineBytes),
                new DeliveryDocumentSnapshot(
                    "working", fixture.WorkingVersion, fixture.WorkingBytes),
                "final",
                fixture.FinalVersion,
                new DeliveryBundleRevisionPolicy
                {
                    PreExistingRevisions = DeliveryRevisionPolicy.Preserve,
                    GeneratedRevisions = DeliveryRevisionPolicy.Accept,
                },
                requests,
                new DeliveryReceiptContext(new[]
                {
                    new DeliveryReceiptTransactionEvidence(
                        fixture.EditContribution,
                        new DeliveryDocumentSnapshot(
                            "edit-before", fixture.BaselineVersion, fixture.BaselineBytes),
                        new DeliveryDocumentSnapshot(
                            "edit-after", fixture.WorkingVersion, fixture.WorkingBytes)),
                    new DeliveryReceiptTransactionEvidence(
                        fixture.AcceptContribution,
                        new DeliveryDocumentSnapshot(
                            "accept-before", fixture.WorkingVersion, fixture.WorkingBytes),
                        new DeliveryDocumentSnapshot(
                            "accept-after", fixture.FinalVersion, fixture.FinalBytes)),
                }));
            var renderer = new DocxodusExportHostRenderer(
                new DocxodusExportHostRendererOptions
                {
                    NodeExecutablePath = nodePath,
                    HostScriptPath = hostPath,
                    ChromiumExecutablePath = Environment.GetEnvironmentVariable(
                        DocxodusExportHostRendererOptions.ChromiumPathEnvironmentVariable),
                    RenderTimeout = TimeSpan.FromMinutes(5),
                });
            bundle = await new DeliveryBundleService(renderer).BuildAsync(request);

            // Persist before semantic assertions so later failures cannot hide completed evidence.
            checkpoint["phase"] = "bundle_publication";
            await WriteJsonAsync(Path.Combine(runDirectory, "checkpoint.json"), checkpoint);
            published = DeliveryBundleDirectoryPublisher.Publish(
                bundle, Path.Combine(runDirectory, "bundle"));
            await WriteChecksumsAsync(bundle, runDirectory);
            await WriteViewerAsync(runDirectory, bundle, error: null);

            var verification = IndependentlyReopen(published);
            await WriteJsonAsync(Path.Combine(runDirectory, "independent-verification.json"),
                verification);
            checkpoint["phase"] = "complete";
            checkpoint["complete"] = true;
            await WriteJsonAsync(Path.Combine(runDirectory, "checkpoint.json"), checkpoint);

            Assert.Equal(DeliveryBundleStatus.Complete, bundle.Manifest.Payload.Status);
            Assert.True(bundle.Verification.IsValid,
                string.Join(Environment.NewLine, bundle.Verification.Findings));
            Assert.True(verification.Valid,
                string.Join(Environment.NewLine, verification.Findings));
            using (var validation = JsonDocument.Parse(
                       bundle.GetArtifactBytes("validation-report")))
            {
                Assert.Contains(
                    validation.RootElement.GetProperty("decision").GetString(),
                    new[] { "passed", "passedWithPreExistingFindings" });
                var actualProfiles = validation.RootElement.GetProperty("renderCohorts")
                    .EnumerateArray()
                    .Select(cohort =>
                        $"{cohort.GetProperty("reviewProfile").GetString()}/"
                        + cohort.GetProperty("commentProfile").GetString())
                    .Order(StringComparer.Ordinal)
                    .ToArray();
                var expectedProfiles = (
                        from review in Enum.GetValues<DeliveryReviewProfile>()
                        from comments in Enum.GetValues<DeliveryCommentProfile>()
                        select $"{Name(review)}/{Name(comments)}")
                    .Order(StringComparer.Ordinal)
                    .ToArray();
                Assert.Equal(expectedProfiles, actualProfiles);
            }
            Assert.Equal(
                Enum.GetValues<DeliveryArtifactKind>().Order(),
                bundle.Manifest.Payload.Artifacts.Select(artifact => artifact.Kind)
                    .Distinct().Order());
        }
        catch (Exception exception)
        {
            checkpoint["phase"] = "failed";
            checkpoint["complete"] = false;
            await WriteJsonAsync(Path.Combine(runDirectory, "checkpoint.json"), checkpoint);
            await WriteJsonAsync(Path.Combine(runDirectory, "error.json"), new
            {
                exception = exception.GetType().FullName,
                exception.Message,
                exception.StackTrace,
            });
            await WriteViewerAsync(runDirectory, bundle, exception);
            throw;
        }
        finally
        {
            await WriteLatestIndexAsync(Path.GetFullPath(artifactRoot), runDirectory);
        }
    }

    private static IntegrationFixture CreateTrackedEdit()
    {
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = new DocxSession(source, new DocxSessionSettings
        {
            PersistAnchorIds = false,
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Delivery Integration Test",
            EmitMarkdownPatch = false,
            CaptureInitialProjection = false,
        });
        var anchor = session.Project().AnchorIndex.Keys.First(value =>
            value.StartsWith("p:body:", StringComparison.Ordinal));
        var baseline = session.Save(persistAnchorIds: false);
        var beforeManifest = PackageManifestGenerator.Generate(baseline);
        var normalized = DeliveryNormalizedOperation.Create(
            "docx_edit",
            "replace_text",
            JsonSerializer.Serialize(new
            {
                anchorId = anchor,
                markdown = "Production delivery render integration.",
            }));
        var mutation = session.ExecuteBatch(new[]
        {
            new MutationBatchStep(
                "docx_edit",
                "replace_text",
                value => value.ReplaceText(anchor, "Production delivery render integration.")),
        });
        if (!mutation.Success)
            throw new InvalidOperationException(mutation.Failure?.Error.Message
                ?? "The tracked fixture mutation failed.");
        var working = session.Save(persistAnchorIds: false);
        var afterManifest = PackageManifestGenerator.Generate(working);
        var contribution = DeliveryTransactionContribution.FromMutationBatchResult(
            mutation, beforeManifest, afterManifest, new[] { normalized });
        var acceptOperation = DeliveryNormalizedOperation.Create(
            "docxodus_track_changes", "accept_all");
        var acceptance = session.ExecuteBatch(new[]
        {
            new MutationBatchStep(
                "docxodus_track_changes",
                "accept_all",
                value => value.AcceptAllRevisions()),
        });
        if (!acceptance.Success)
            throw new InvalidOperationException(acceptance.Failure?.Error.Message
                ?? "The fixture revision acceptance failed.");
        var final = session.Save(persistAnchorIds: false);
        var finalManifest = PackageManifestGenerator.Generate(final);
        var acceptContribution = DeliveryTransactionContribution.FromMutationBatchResult(
            acceptance, afterManifest, finalManifest, new[] { acceptOperation });
        return new IntegrationFixture(
            baseline,
            working,
            final,
            mutation.BaseVersion,
            mutation.ResultVersion,
            acceptance.ResultVersion,
            contribution,
            acceptContribution);
    }

    private static DeliveryArtifactRequest Render(
        string artifactId,
        DeliveryArtifactKind kind,
        DeliveryReviewProfile reviewProfile,
        DeliveryCommentProfile commentProfile) => new()
        {
            ArtifactId = artifactId,
            Kind = kind,
            Requiredness = DeliveryArtifactRequiredness.Required,
            ReviewProfile = reviewProfile,
            CommentProfile = commentProfile,
        };

    private static IndependentVerification IndependentlyReopen(string directory)
    {
        var manifestBytes = File.ReadAllBytes(
            Path.Combine(directory, DeliveryBundle.ManifestFileName));
        using var json = JsonDocument.Parse(manifestBytes);
        var artifacts = json.RootElement.GetProperty("payload").GetProperty("artifacts")
            .EnumerateArray().ToArray();
        var findings = new List<string>();
        var available = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var artifact in artifacts)
        {
            var id = artifact.GetProperty("artifactId").GetString()!;
            if (artifact.GetProperty("availability").GetString() != "available")
            {
                findings.Add($"unavailable:{id}");
                continue;
            }
            var relative = artifact.GetProperty("relativePath").GetString()!
                .Replace('/', Path.DirectorySeparatorChar);
            var path = Path.Combine(directory, relative);
            var bytes = File.ReadAllBytes(path);
            available.Add(id, bytes);
            var expectedLength = artifact.GetProperty("byteLength").GetInt64();
            if (bytes.LongLength != expectedLength)
                findings.Add($"length:{id}");
            var expectedDigest = artifact.GetProperty("digest").GetProperty("value").GetString();
            if (!string.Equals(Sha256(bytes), expectedDigest, StringComparison.Ordinal))
                findings.Add($"digest:{id}");
            if (artifact.GetProperty("kind").GetString()!.EndsWith("Docx",
                    StringComparison.Ordinal))
            {
                try
                {
                    using var stream = new MemoryStream(bytes, writable: false);
                    using var document = WordprocessingDocument.Open(stream, false);
                    _ = document.MainDocumentPart?.Document.Body
                        ?? throw new InvalidDataException("DOCX main body missing.");
                }
                catch (Exception exception)
                {
                    findings.Add($"docx:{id}:{exception.GetType().Name}");
                }
            }
        }
        var manifestVerification = DeliveryBundleVerifier.VerifyJson(manifestBytes, available);
        findings.AddRange(manifestVerification.Findings);
        var declaredPaths = artifacts
            .Where(artifact => artifact.GetProperty("availability").GetString() == "available")
            .Select(artifact => artifact.GetProperty("relativePath").GetString()!)
            .Append(DeliveryBundle.ManifestFileName)
            .Order(StringComparer.Ordinal)
            .ToArray();
        var diskPaths = Directory.EnumerateFiles(directory, "*", SearchOption.AllDirectories)
            .Select(path => Path.GetRelativePath(directory, path)
                .Replace(Path.DirectorySeparatorChar, '/'))
            .Order(StringComparer.Ordinal)
            .ToArray();
        if (!declaredPaths.SequenceEqual(diskPaths, StringComparer.Ordinal))
            findings.Add("disk_inventory_mismatch");
        var artifactsById = artifacts.ToDictionary(
            artifact => artifact.GetProperty("artifactId").GetString()!,
            StringComparer.Ordinal);
        var relationships = json.RootElement.GetProperty("payload").GetProperty("relationships")
            .EnumerateArray().ToArray();
        foreach (var artifact in artifacts)
        {
            if (!artifact.TryGetProperty("render", out var render)
                || render.ValueKind == JsonValueKind.Null)
                continue;
            var id = artifact.GetProperty("artifactId").GetString()!;
            var renderedFrom = relationships.Where(relationship =>
                    relationship.GetProperty("kind").GetString() == "renderedFrom"
                    && relationship.GetProperty("fromArtifactId").GetString() == id)
                .ToArray();
            if (renderedFrom.Length != 1)
            {
                findings.Add($"render_source_relationship:{id}");
                continue;
            }
            var sourceId = renderedFrom[0].GetProperty("toArtifactId").GetString()!;
            if (!artifactsById.TryGetValue(sourceId, out var source)
                || source.GetProperty("availability").GetString() != "available")
            {
                findings.Add($"render_source_unavailable:{id}");
                continue;
            }
            var renderDigest = render.GetProperty("sourcePackageDigest")
                .GetProperty("value").GetString();
            var sourceDigest = source.GetProperty("digest").GetProperty("value").GetString();
            if (!string.Equals(renderDigest, sourceDigest, StringComparison.Ordinal))
                findings.Add($"render_source_digest:{id}");
        }
        return new IndependentVerification(findings.Count == 0, findings);
    }

    private static async Task WriteViewerAsync(
        string runDirectory,
        DeliveryBundle? bundle,
        Exception? error)
    {
        var rows = bundle is null
            ? string.Empty
            : string.Join(Environment.NewLine, bundle.Manifest.Payload.Artifacts.Select(artifact =>
            {
                var destination = artifact.Availability == BundleArtifactAvailability.Available
                    ? $"<a href=\"bundle/{Html(artifact.RelativePath)}\">view</a>"
                    : Html(artifact.UnavailableReason ?? "unavailable");
                var source = artifact.Render is null
                    ? string.Empty
                    : $"{Html(artifact.Render.SourceDocumentName)}@"
                      + $"{artifact.Render.SourceDocumentVersion}<br><code>"
                      + $"{Html(artifact.Render.SourcePackageDigest.Value)}</code>";
                return $"<tr><td>{Html(artifact.ArtifactId)}</td><td>{artifact.Kind}</td>"
                       + $"<td>{artifact.Availability}</td><td>{artifact.ByteLength}</td>"
                       + $"<td>{Html(artifact.Digest?.Value ?? string.Empty)}</td>"
                       + $"<td>{source}</td><td>{destination}</td></tr>";
            }));
        var html = "<!doctype html><meta charset=\"utf-8\"><title>Docxodus #465 real delivery</title>"
                   + "<style>body{font:15px system-ui;margin:2rem;max-width:1400px}"
                   + "table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;"
                   + "padding:.4rem;text-align:left;vertical-align:top}code{word-break:break-all}</style>"
                   + "<h1>Docxodus #465 production delivery evidence</h1>"
                   + $"<p>Result: <strong>{Html(error is null ? "completed" : "failed")}</strong>. "
                   + "Artifacts and checkpoints are retained before assertions.</p>"
                   + "<p><a href=\"baseline.docx\">baseline DOCX</a> · "
                   + "<a href=\"working.docx\">working DOCX</a> · "
                   + "<a href=\"request.json\">request</a> · "
                   + "<a href=\"checkpoint.json\">checkpoint</a> · "
                   + "<a href=\"independent-verification.json\">independent verification</a></p>"
                   + (error is null ? string.Empty
                       : $"<p><a href=\"error.json\">failure details</a>: {Html(error.Message)}</p>")
                   + "<table><thead><tr><th>ID</th><th>Kind</th><th>Availability</th>"
                   + "<th>Bytes</th><th>SHA-256</th><th>Render source</th>"
                   + $"<th>Artifact</th></tr></thead><tbody>{rows}</tbody></table>";
        await File.WriteAllTextAsync(
            Path.Combine(runDirectory, "index.html"), html, Encoding.UTF8);
    }

    private static async Task WriteChecksumsAsync(DeliveryBundle bundle, string runDirectory)
    {
        var lines = bundle.Manifest.Payload.Artifacts
            .Where(artifact => artifact.Availability == BundleArtifactAvailability.Available)
            .Select(artifact => $"{artifact.Digest!.Value}  bundle/{artifact.RelativePath}")
            .Append($"{Sha256(bundle.ManifestBytes)}  bundle/{DeliveryBundle.ManifestFileName}");
        await File.WriteAllLinesAsync(
            Path.Combine(runDirectory, "SHA256SUMS"),
            lines,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    private static async Task WriteLatestIndexAsync(string root, string runDirectory)
    {
        Directory.CreateDirectory(root);
        var relative = Path.GetRelativePath(root, runDirectory)
            .Replace(Path.DirectorySeparatorChar, '/');
        await File.WriteAllTextAsync(
            Path.Combine(root, "index.html"),
            $"<!doctype html><meta charset=\"utf-8\"><title>Docxodus delivery evidence</title>"
            + $"<p><a href=\"{Html(relative)}/index.html\">View latest production delivery run</a></p>",
            Encoding.UTF8);
    }

    private static Task WriteJsonAsync(string path, object value) =>
        File.WriteAllBytesAsync(path, JsonSerializer.SerializeToUtf8Bytes(value, new JsonSerializerOptions
        {
            WriteIndented = true,
        }));

    private static string RequiredEnvironment(string name) =>
        Environment.GetEnvironmentVariable(name) is { Length: > 0 } value
            ? Path.GetFullPath(value)
            : throw new InvalidOperationException($"The integration test requires {name}.");

    private static string Html(string value) => HtmlEncoder.Default.Encode(value);

    private static string Sha256(byte[] bytes) => Convert.ToHexString(
        SHA256.HashData(bytes)).ToLowerInvariant();

    private static string Name<T>(T value)
        where T : struct, Enum => JsonNamingPolicy.CamelCase.ConvertName(value.ToString());

    private static string Kebab<T>(T value)
        where T : struct, Enum
    {
        var name = value.ToString();
        var builder = new StringBuilder(name.Length + 8);
        for (var index = 0; index < name.Length; index++)
        {
            var character = name[index];
            if (index > 0 && char.IsUpper(character)) builder.Append('-');
            builder.Append(char.ToLowerInvariant(character));
        }
        return builder.ToString();
    }

    private sealed record IntegrationFixture(
        byte[] BaselineBytes,
        byte[] WorkingBytes,
        byte[] FinalBytes,
        long BaselineVersion,
        long WorkingVersion,
        long FinalVersion,
        DeliveryTransactionContribution EditContribution,
        DeliveryTransactionContribution AcceptContribution);

    private sealed record IndependentVerification(
        bool Valid,
        IReadOnlyList<string> Findings);
}
