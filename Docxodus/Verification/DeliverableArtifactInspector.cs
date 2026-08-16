// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text;
using System.Text.Json;
using Docxodus.Internal;

namespace Docxodus.Verification;

/// <summary>Basic format validation and cross-artifact source/PageMap closure.</summary>
internal static class DeliverableArtifactInspector
{
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);

    internal static IReadOnlyList<DeliverableArtifactMetadata> Inspect(
        IReadOnlyList<DeliverableCompanionArtifactInput> artifacts,
        VerificationDigest packageDigest,
        DeliverableVerificationOptions options,
        ICollection<DeliverableFindingObservation> observations,
        out DeliverableCheckResult check)
    {
        int before = observations.Count;
        long total = 0;
        bool bounded = true;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var states = new List<ArtifactState>(artifacts.Count);
        foreach (var artifact in artifacts.OrderBy(item => item.ArtifactId, StringComparer.Ordinal))
        {
            if (!seen.Add(artifact.ArtifactId))
                Finding(observations, options.MaxFindings, artifact, "artifact.id_duplicate",
                    VerificationFindingSeverity.Error, "Companion artifact ids must be unique.",
                    "Assign a unique stable artifact id.");
            if (string.IsNullOrWhiteSpace(artifact.ArtifactId))
                Finding(observations, options.MaxFindings, artifact, "artifact.id_missing",
                    VerificationFindingSeverity.Error, "A companion artifact has no id.",
                    "Assign a non-empty stable artifact id.");
            if (string.IsNullOrWhiteSpace(artifact.MediaType))
                Finding(observations, options.MaxFindings, artifact, "artifact.media_type_missing",
                    VerificationFindingSeverity.Error, "A companion artifact has no media type.",
                    "Supply the artifact's MIME media type.");

            var state = new ArtifactState(artifact);
            states.Add(state);
            if (artifact.Availability == DeliverableArtifactAvailability.Available)
            {
                if (artifact.Bytes is null)
                    Finding(observations, options.MaxFindings, artifact, "artifact.bytes_missing",
                        VerificationFindingSeverity.Error,
                        "An available companion artifact has no bytes.",
                        "Supply bytes or mark the artifact unavailable with a reason.");
                else
                {
                    state.Length = artifact.Bytes.LongLength;
                    if (state.Length > options.MaxCompanionArtifactBytes
                        || state.Length > options.MaxTotalCompanionArtifactBytes - total)
                    {
                        bounded = false;
                        Finding(observations, options.MaxFindings, artifact,
                            "artifact.size_limit_exceeded", VerificationFindingSeverity.Error,
                            "Companion artifact bytes exceed the configured verification budget.",
                            "Reduce the artifact or deliberately raise the bounded artifact policy.");
                    }
                    else
                    {
                        total += state.Length.Value;
                        state.RawDigest = DeliverableVerificationIdentity.Digest(artifact.Bytes);
                        InspectFormat(state, observations, options.MaxFindings);
                    }
                }
            }
            else
            {
                if (artifact.Bytes is not null)
                    Finding(observations, options.MaxFindings, artifact,
                        "artifact.unavailable_has_bytes", VerificationFindingSeverity.Warning,
                        "An unavailable artifact also supplied bytes; the bytes were ignored.",
                        "Mark the artifact available or omit its bytes.");
                if (string.IsNullOrWhiteSpace(artifact.UnavailableReason))
                    Finding(observations, options.MaxFindings, artifact,
                        "artifact.unavailable_reason_missing", VerificationFindingSeverity.Warning,
                        "An unavailable artifact has no reason.",
                        "Record why the artifact could not be produced.");
            }

            InspectBindingMetadata(state, packageDigest, observations, options.MaxFindings);
            InspectRenderDiagnostics(state, observations, options.MaxFindings);
        }

        InspectPageMapClosure(states, observations, options.MaxFindings);
        check = new DeliverableCheckResult
        {
            Check = "companion_artifacts",
            Status = bounded ? DeliverableCheckStatus.Completed : DeliverableCheckStatus.UnavailableEvidence,
            FindingCount = observations.Count - before,
            Diagnostic = bounded ? null : "artifact byte budget exceeded",
        };
        return states.Select(state => state.Metadata()).ToArray();
    }

    private static void InspectFormat(
        ArtifactState state,
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings)
    {
        var artifact = state.Input;
        var bytes = artifact.Bytes!;
        switch (artifact.Role)
        {
            case DeliverableArtifactRole.Pdf:
                if (!MediaTypeIs(artifact.MediaType, "application/pdf"))
                    InvalidMediaType(observations, maximumFindings, artifact, "application/pdf");
                if (!LooksLikePdf(bytes))
                    Finding(observations, maximumFindings, artifact, "artifact.pdf_malformed",
                        VerificationFindingSeverity.Error,
                        "The supplied bytes do not satisfy basic PDF structure checks.",
                        "Supply a complete PDF with a header, catalog/page tree, xref/trailer, and EOF marker.");
                break;
            case DeliverableArtifactRole.Html:
                if (!MediaTypeIs(artifact.MediaType, "text/html"))
                    InvalidMediaType(observations, maximumFindings, artifact, "text/html");
                if (!LooksLikeHtml(bytes))
                    Finding(observations, maximumFindings, artifact, "artifact.html_malformed",
                        VerificationFindingSeverity.Error,
                        "The supplied bytes are not a complete UTF-8 HTML document.",
                        "Supply UTF-8 HTML with an html root and closing tag.");
                break;
            case DeliverableArtifactRole.PageMap:
                if (!MediaTypeIs(artifact.MediaType, "application/json")
                    && !MediaTypeIs(artifact.MediaType, "application/vnd.docxodus.pagemap+json"))
                    InvalidMediaType(observations, maximumFindings, artifact,
                        "application/vnd.docxodus.pagemap+json or application/json");
                ParsePageMap(state, observations, maximumFindings);
                break;
        }
    }

    private static void ParsePageMap(
        ArtifactState state,
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings)
    {
        try
        {
            using var json = JsonDocument.Parse(state.Input.Bytes!, new JsonDocumentOptions
            {
                AllowTrailingCommas = false,
                CommentHandling = JsonCommentHandling.Disallow,
                MaxDepth = 64,
            });
            EnsureUniqueProperties(json.RootElement);
            var map = DocxSessionJson.ParsePageMap(json.RootElement);
            var portable = PageMapContract.ValidatePortable(map);
            if (!portable.Success || map.Mode != PageMapMode.Paginated
                || map.Availability != PageMapAvailability.Available)
                throw new FormatException(portable.Message
                    ?? "PageMap artifact must be an available paginated portable map");
            var canonical = StrictUtf8.GetBytes(DocxSessionJson.SerializePageMap(map));
            state.PageMap = map;
            state.CanonicalPageMapDigest = DeliverableVerificationIdentity.Digest(canonical);
            state.EffectivePageCount = map.Pages.Count;
        }
        catch (Exception exception) when (exception is JsonException or FormatException
            or DecoderFallbackException or ArgumentException)
        {
            Finding(observations, maximumFindings, state.Input, "artifact.page_map_malformed",
                VerificationFindingSeverity.Error,
                $"The PageMap artifact is not a strict portable PageMap ({exception.GetType().Name}).",
                "Regenerate the PageMap using the portable schema-v1 contract.");
        }
    }

    private static void InspectBindingMetadata(
        ArtifactState state,
        VerificationDigest packageDigest,
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings)
    {
        var artifact = state.Input;
        bool layoutArtifact = artifact.Role is DeliverableArtifactRole.Html
            or DeliverableArtifactRole.Pdf or DeliverableArtifactRole.PageMap;
        if (artifact.Availability != DeliverableArtifactAvailability.Available) return;
        if (artifact.SourcePackageDigest is null)
            Finding(observations, maximumFindings, artifact, "artifact.source_digest_missing",
                layoutArtifact ? VerificationFindingSeverity.Error : VerificationFindingSeverity.Warning,
                "The companion artifact is not bound to source package bytes.",
                "Record the exact delivered package SHA-256 as sourcePackageDigest.");
        else if (!DeliverableVerificationIdentity.DigestEquals(
                     artifact.SourcePackageDigest, packageDigest))
            Finding(observations, maximumFindings, artifact, "artifact.source_digest_mismatch",
                VerificationFindingSeverity.Error,
                "The companion artifact names a different source package digest.",
                "Regenerate the artifact from the delivered package or correct the binding.");

        if (artifact.PageCount is < 0)
            Finding(observations, maximumFindings, artifact, "artifact.page_count_invalid",
                VerificationFindingSeverity.Error, "Companion artifact pageCount cannot be negative.",
                "Supply a non-negative page count.");
        if (layoutArtifact && artifact.PageCount is null)
            Finding(observations, maximumFindings, artifact, "artifact.page_count_missing",
                VerificationFindingSeverity.Error, "Layout-dependent evidence has no page count.",
                "Record the rendered page count and bind it to a PageMap.");
        if (layoutArtifact && string.IsNullOrWhiteSpace(artifact.RendererFingerprint))
            Finding(observations, maximumFindings, artifact,
                "artifact.renderer_fingerprint_missing", VerificationFindingSeverity.Error,
                "Layout-dependent evidence has no renderer fingerprint.",
                "Record the renderer name, version, and relevant configuration.");
        if (state.PageMap is not null)
        {
            if (artifact.PageMapDigest is not null
                && !DeliverableVerificationIdentity.DigestEquals(
                    artifact.PageMapDigest, state.CanonicalPageMapDigest))
                Finding(observations, maximumFindings, artifact,
                    "artifact.page_map_digest_mismatch", VerificationFindingSeverity.Error,
                    "PageMap metadata does not match the map's canonical bytes.",
                    "Record the canonical SHA-256 produced from the supplied portable PageMap.");
            if (!string.Equals(artifact.RendererFingerprint, state.PageMap.RendererFingerprint,
                    StringComparison.Ordinal))
                Finding(observations, maximumFindings, artifact,
                    "artifact.page_map_renderer_mismatch", VerificationFindingSeverity.Error,
                    "PageMap JSON and artifact metadata name different renderer fingerprints.",
                    "Use the exact fingerprint recorded by the PageMap renderer.");
            if (artifact.PageCount != state.PageMap.Pages.Count)
                Finding(observations, maximumFindings, artifact,
                    "artifact.page_map_count_mismatch", VerificationFindingSeverity.Error,
                    "PageMap page count does not match its artifact metadata.",
                    "Regenerate the PageMap metadata from the same render result.");
        }
    }

    private static void InspectPageMapClosure(
        IReadOnlyList<ArtifactState> states,
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings)
    {
        var maps = states.Where(state => state.PageMap is not null
                && state.CanonicalPageMapDigest is not null)
            .GroupBy(state => state.CanonicalPageMapDigest!.Value, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.OrdinalIgnoreCase);
        foreach (var state in states.Where(state => state.Input.Availability
                     == DeliverableArtifactAvailability.Available
                 && state.Input.Role is DeliverableArtifactRole.Pdf or DeliverableArtifactRole.Html))
        {
            var artifact = state.Input;
            if (artifact.PageMapDigest is null)
            {
                Finding(observations, maximumFindings, artifact, "artifact.page_map_digest_missing",
                    VerificationFindingSeverity.Error,
                    "Layout output does not reference exact portable PageMap bytes.",
                    "Attach a PageMap artifact and record its canonical SHA-256.");
                continue;
            }
            if (!maps.TryGetValue(artifact.PageMapDigest.Value, out var mapState))
            {
                Finding(observations, maximumFindings, artifact, "artifact.page_map_missing",
                    VerificationFindingSeverity.Error,
                    "No supplied PageMap artifact has the referenced canonical digest.",
                    "Supply the exact referenced PageMap bytes or correct the digest.");
                continue;
            }
            if (!DeliverableVerificationIdentity.DigestEquals(
                    artifact.SourcePackageDigest, mapState.Input.SourcePackageDigest))
                Finding(observations, maximumFindings, artifact,
                    "artifact.page_map_source_mismatch", VerificationFindingSeverity.Error,
                    "Layout output and PageMap are bound to different source packages.",
                    "Regenerate all layout artifacts from one exact package snapshot.");
            if (!string.Equals(artifact.RendererFingerprint,
                    mapState.PageMap!.RendererFingerprint, StringComparison.Ordinal))
                Finding(observations, maximumFindings, artifact,
                    "artifact.page_map_renderer_mismatch", VerificationFindingSeverity.Error,
                    "Layout output and PageMap name different renderer fingerprints.",
                    "Regenerate both artifacts with the same renderer configuration.");
            if (artifact.PageCount != mapState.PageMap.Pages.Count)
                Finding(observations, maximumFindings, artifact,
                    "artifact.page_map_count_mismatch", VerificationFindingSeverity.Error,
                    "Layout output and PageMap page counts differ.",
                    "Regenerate the output and PageMap in one render operation.");
        }
    }

    private static void InspectRenderDiagnostics(
        ArtifactState state,
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings)
    {
        foreach (var diagnostic in state.Input.RenderDiagnostics
                     .OrderBy(item => item.Kind)
                     .ThenBy(item => item.OwningPartUri, StringComparer.Ordinal)
                     .ThenBy(item => item.AnchorId, StringComparer.Ordinal)
                     .ThenBy(item => item.FontName, StringComparer.Ordinal)
                     .ThenBy(item => item.SubstitutedFontName, StringComparer.Ordinal))
        {
            if (observations.Count >= maximumFindings) break;
            var code = diagnostic.Kind switch
            {
                DeliverableRenderDiagnosticKind.MissingFont => "render.missing_font",
                DeliverableRenderDiagnosticKind.FontSubstitution => "render.font_substitution",
                DeliverableRenderDiagnosticKind.UnsupportedContent => "render.unsupported_content",
                _ => "render.warning",
            };
            var owner = string.IsNullOrWhiteSpace(diagnostic.OwningPartUri)
                ? "/" : diagnostic.OwningPartUri!;
            Add(observations, maximumFindings, DeliverableFindingObservation.Create(
                code, DeliverableFindingCategory.Render, diagnostic.Severity,
                string.IsNullOrWhiteSpace(diagnostic.Message)
                    ? "The renderer supplied a diagnostic without explanatory text."
                    : diagnostic.Message,
                owner,
                string.IsNullOrWhiteSpace(diagnostic.Remediation)
                    ? "Review the renderer diagnostic and correct or approve the visual result."
                    : diagnostic.Remediation,
                new ChangeLocation
                {
                    EntryUri = owner == "/" ? null : owner,
                    PropertyPath = "artifacts/" + state.Input.ArtifactId,
                }, diagnostic.AnchorId,
                subjectKey: string.Join("\u001f", state.Input.ArtifactId, diagnostic.Kind,
                    diagnostic.FontName, diagnostic.SubstitutedFontName)));
        }
    }

    private static bool LooksLikePdf(byte[] bytes)
    {
        if (bytes.Length < 32) return false;
        var span = bytes.AsSpan();
        int end = span.Length - 1;
        while (end >= 0 && span[end] is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n') end--;
        return span.StartsWith("%PDF-1."u8)
            && Contains(span, "/Type /Catalog"u8)
            && Contains(span, "/Type /Pages"u8)
            && Contains(span, "/Type /Page "u8)
            && Contains(span, "xref"u8)
            && Contains(span, "trailer"u8)
            && Contains(span, "startxref"u8)
            && end >= 4 && span[..(end + 1)].EndsWith("%%EOF"u8);
    }

    private static bool Contains(ReadOnlySpan<byte> bytes, ReadOnlySpan<byte> value) =>
        bytes.IndexOf(value) >= 0;

    private static bool LooksLikeHtml(byte[] bytes)
    {
        try
        {
            var text = StrictUtf8.GetString(bytes);
            if (text.IndexOf('\0') >= 0) return false;
            var trimmed = text.AsSpan().TrimStart();
            return (trimmed.StartsWith("<!doctype html", StringComparison.OrdinalIgnoreCase)
                    || trimmed.StartsWith("<html", StringComparison.OrdinalIgnoreCase))
                && text.Contains("<html", StringComparison.OrdinalIgnoreCase)
                && text.Contains("</html>", StringComparison.OrdinalIgnoreCase);
        }
        catch (DecoderFallbackException)
        {
            return false;
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
                    throw new FormatException("duplicate JSON property: " + property.Name);
                EnsureUniqueProperties(property.Value);
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
            foreach (var item in element.EnumerateArray()) EnsureUniqueProperties(item);
    }

    private static bool MediaTypeIs(string actual, string expected) => string.Equals(
        actual.Split(';', 2)[0].Trim(), expected, StringComparison.OrdinalIgnoreCase);

    private static void InvalidMediaType(
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings,
        DeliverableCompanionArtifactInput artifact,
        string expected) => Finding(observations, maximumFindings, artifact,
        "artifact.media_type_mismatch", VerificationFindingSeverity.Error,
        $"Artifact role and media type do not agree; expected {expected}.",
        "Correct the media type or artifact role.");

    private static void Finding(
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings,
        DeliverableCompanionArtifactInput artifact,
        string code,
        VerificationFindingSeverity severity,
        string message,
        string remediation) => Add(observations, maximumFindings,
        DeliverableFindingObservation.Create(
            code, DeliverableFindingCategory.Artifact, severity, message, "/", remediation,
            new ChangeLocation { PropertyPath = "artifacts/" + artifact.ArtifactId },
            subjectKey: artifact.ArtifactId));

    private static void Add(
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings,
        DeliverableFindingObservation observation)
    {
        if (observations.Count < maximumFindings) observations.Add(observation);
    }

    private static VerificationDigest? Normalize(VerificationDigest? digest) => digest is null
        ? null : new VerificationDigest { Algorithm = "SHA-256", Value = digest.Value.ToLowerInvariant() };

    private sealed class ArtifactState(DeliverableCompanionArtifactInput input)
    {
        internal DeliverableCompanionArtifactInput Input { get; } = input;
        internal long? Length { get; set; }
        internal VerificationDigest? RawDigest { get; set; }
        internal PageMap? PageMap { get; set; }
        internal VerificationDigest? CanonicalPageMapDigest { get; set; }
        internal long? EffectivePageCount { get; set; }

        internal DeliverableArtifactMetadata Metadata() => new()
        {
            ArtifactId = Input.ArtifactId,
            Role = Input.Role,
            MediaType = Input.MediaType,
            Availability = Input.Availability,
            ByteLength = Length,
            Digest = RawDigest,
            UnavailableReason = Input.UnavailableReason,
            PageCount = EffectivePageCount ?? Input.PageCount,
            RendererFingerprint = Input.RendererFingerprint,
            SourcePackageDigest = Normalize(Input.SourcePackageDigest),
            PageMapDigest = Input.Role == DeliverableArtifactRole.PageMap
                ? CanonicalPageMapDigest : Normalize(Input.PageMapDigest),
            RenderDiagnosticCount = Input.RenderDiagnostics.Count,
        };
    }
}
