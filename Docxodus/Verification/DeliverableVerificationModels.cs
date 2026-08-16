// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml;

namespace Docxodus.Verification;

/// <summary>Preset used to turn deliverable findings into a delivery decision.</summary>
public enum DeliverableVerificationMode
{
    /// <summary>Reject new defects and conditions that make the current deliverable unsafe.</summary>
    Standard,

    /// <summary>Reject every current warning or error, including pre-existing conditions.</summary>
    Strict,

    /// <summary>Collect evidence without making a pass/fail decision.</summary>
    ReportOnly,
}

/// <summary>The policy result of a deliverable verification run.</summary>
public enum DeliverableVerificationDecision
{
    Passed,
    PassedWithPreExistingFindings,
    Failed,
    NotEvaluated,
}

/// <summary>How a finding in the delivered package relates to the supplied baseline.</summary>
public enum DeliverableFindingDisposition
{
    New,
    PreExisting,
    Resolved,
    Unclassified,
}

/// <summary>Stable high-level owner of a deliverable finding.</summary>
public enum DeliverableFindingCategory
{
    Package,
    OpenXml,
    Relationship,
    Structure,
    Workflow,
    Delta,
    Render,
    Artifact,
}

/// <summary>Whether one verification stage completed and produced authoritative evidence.</summary>
public enum DeliverableCheckStatus
{
    Completed,
    SkippedPrerequisiteFailed,
    UnavailableEvidence,
}

/// <summary>Kind of package-level change observed between baseline and deliverable.</summary>
public enum DeliverablePackageChangeKind
{
    EntryAdded,
    EntryRemoved,
    EntryModified,
    RelationshipAdded,
    RelationshipRemoved,
    RelationshipModified,
}

/// <summary>Role of a companion artifact supplied with a deliverable.</summary>
public enum DeliverableArtifactRole
{
    Html,
    Pdf,
    PageMap,
    PageImage,
    RenderReport,
    Other,
}

/// <summary>Whether companion artifact bytes were produced.</summary>
public enum DeliverableArtifactAvailability
{
    Available,
    Unavailable,
}

/// <summary>Structured renderer diagnostic supplied by the renderer that observed it.</summary>
public enum DeliverableRenderDiagnosticKind
{
    Warning,
    UnsupportedContent,
    MissingFont,
    FontSubstitution,
}

/// <summary>Safety and policy options for <see cref="DeliverableVerifier.VerifyDeliverable"/>.</summary>
public sealed record DeliverableVerificationOptions
{
    public DeliverableVerificationMode Mode { get; init; } = DeliverableVerificationMode.Standard;

    /// <summary>Safety limits shared by package preflight and semantic comparison.</summary>
    public PackageManifestOptions PackageManifestOptions { get; init; } = new()
    {
        MaxEntryCount = 10_000,
        MaxEntryUncompressedBytes = 64L * 1024 * 1024,
        MaxTotalUncompressedBytes = 256L * 1024 * 1024,
        MaxXmlPartBytes = 64L * 1024 * 1024,
        MaxCompressionRatio = 1_000,
        MaxUriLength = 2_048,
    };

    /// <summary>
    /// Maximum aggregate raw bytes accepted across the deliverable and optional baseline package.
    /// This is enforced before either caller-owned array is cloned.
    /// </summary>
    public long MaxPackageBytes { get; init; } = 100L * 1024 * 1024;

    /// <summary>Office schema version used by the Open XML SDK validator.</summary>
    public FileFormatVersions OpenXmlVersion { get; init; } = FileFormatVersions.Office2019;

    /// <summary>When true, actual changes not present in the expected sets block delivery.</summary>
    public bool FailOnUnexpectedChanges { get; init; }

    /// <summary>Whether unresolved placeholders block standard-mode delivery. They are always reported.</summary>
    public bool RequireNoPlaceholders { get; init; } = true;

    /// <summary>
    /// Opt in to detecting broad square-bracket alternatives such as <c>[Buyer/Seller]</c>.
    /// These findings are advisory because ordinary legal citations also use square brackets.
    /// </summary>
    public bool DetectBracketedAlternativeClauses { get; init; }

    /// <summary>
    /// Exact, case-sensitive standalone editorial markers to scan for. Empty by default so words
    /// in other languages (for example Spanish <c>todo</c>) are not treated as template state.
    /// </summary>
    public IReadOnlyList<string> EditorialMarkers { get; init; } = Array.Empty<string>();

    /// <summary>Exact additional placeholder tokens to report as high-confidence template state.</summary>
    public IReadOnlyList<string> PlaceholderTokens { get; init; } = Array.Empty<string>();

    /// <summary>Maximum structured findings retained across every detector.</summary>
    public int MaxFindings { get; init; } = 10_000;

    /// <summary>Maximum XML elements visited by bounded semantic detectors per package.</summary>
    public int MaxDetectorNodes { get; init; } = 1_000_000;

    /// <summary>Maximum relationships traversed by bounded semantic detectors per package.</summary>
    public int MaxDetectorRelationships { get; init; } = 100_000;

    /// <summary>Maximum text characters inspected by workflow detectors per package.</summary>
    public long MaxDetectorTextCharacters { get; init; } = 16L * 1024 * 1024;

    /// <summary>Maximum regex/token matches inspected by workflow detectors per package.</summary>
    public int MaxDetectorRegexMatches { get; init; } = 100_000;

    /// <summary>Maximum miscellaneous detector operations per package.</summary>
    public long MaxDetectorSteps { get; init; } = 2_000_000;

    /// <summary>Maximum bytes accepted for one companion artifact.</summary>
    public long MaxCompanionArtifactBytes { get; init; } = 256L * 1024 * 1024;

    /// <summary>Maximum aggregate bytes accepted across companion artifacts.</summary>
    public long MaxTotalCompanionArtifactBytes { get; init; } = 512L * 1024 * 1024;

    /// <summary>Maximum companion-artifact records accepted in one request.</summary>
    public int MaxCompanionArtifacts { get; init; } = 1_024;

    /// <summary>Maximum aggregate renderer diagnostics accepted in one request.</summary>
    public int MaxRenderDiagnostics { get; init; } = 10_000;

    /// <summary>Maximum expected semantic plus package changes accepted in one request.</summary>
    public int MaxExpectedChanges { get; init; } = 25_000;

    /// <summary>Maximum semantic or package delta records returned by one comparison stage.</summary>
    public int MaxReportedDeltaChanges { get; init; } = 25_000;

    /// <summary>Maximum typed semantic-value nodes across expected semantic changes.</summary>
    public int MaxExpectedSemanticValueNodes { get; init; } = 1_000_000;

    /// <summary>Maximum configured workflow marker/token records.</summary>
    public int MaxConfiguredWorkflowMarkers { get; init; } = 10_000;

    /// <summary>Maximum aggregate characters in caller-supplied policy and evidence strings.</summary>
    public long MaxEvidenceTextCharacters { get; init; } = 4L * 1024 * 1024;

    /// <summary>Include baseline findings that no longer exist in the delivered package.</summary>
    public bool IncludeResolvedFindings { get; init; } = true;

    internal void Validate()
    {
        if (!Enum.IsDefined(Mode))
            throw new ArgumentOutOfRangeException(nameof(Mode));
        if (!Enum.IsDefined(OpenXmlVersion))
            throw new ArgumentOutOfRangeException(nameof(OpenXmlVersion));
        if (MaxPackageBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxPackageBytes));
        if (MaxFindings <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxFindings));
        if (MaxDetectorNodes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxDetectorNodes));
        if (MaxDetectorRelationships <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxDetectorRelationships));
        if (MaxDetectorTextCharacters <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxDetectorTextCharacters));
        if (MaxDetectorRegexMatches <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxDetectorRegexMatches));
        if (MaxDetectorSteps <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxDetectorSteps));
        if (MaxCompanionArtifactBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxCompanionArtifactBytes));
        if (MaxTotalCompanionArtifactBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxTotalCompanionArtifactBytes));
        if (MaxCompanionArtifacts <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxCompanionArtifacts));
        if (MaxRenderDiagnostics <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxRenderDiagnostics));
        if (MaxExpectedChanges <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxExpectedChanges));
        if (MaxReportedDeltaChanges <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxReportedDeltaChanges));
        if (MaxExpectedSemanticValueNodes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxExpectedSemanticValueNodes));
        if (MaxConfiguredWorkflowMarkers <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxConfiguredWorkflowMarkers));
        if (MaxEvidenceTextCharacters <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxEvidenceTextCharacters));
        if (PackageManifestOptions is null)
            throw new ArgumentNullException(nameof(PackageManifestOptions));
        if (EditorialMarkers is null)
            throw new ArgumentNullException(nameof(EditorialMarkers));
        if (PlaceholderTokens is null)
            throw new ArgumentNullException(nameof(PlaceholderTokens));
        if (EditorialMarkers.Count > MaxConfiguredWorkflowMarkers
            || PlaceholderTokens.Count > MaxConfiguredWorkflowMarkers
            || EditorialMarkers.Count > MaxConfiguredWorkflowMarkers - PlaceholderTokens.Count)
            throw new ArgumentException("Configured workflow markers exceed the verification budget.");
        if (EditorialMarkers.Any(string.IsNullOrWhiteSpace))
            throw new ArgumentException("EditorialMarkers cannot contain blank values.",
                nameof(EditorialMarkers));
        if (PlaceholderTokens.Any(string.IsNullOrEmpty))
            throw new ArgumentException("PlaceholderTokens cannot contain empty values.",
                nameof(PlaceholderTokens));
        long configuredCharacters = 0;
        bool configuredTextExceeded = false;
        foreach (var value in EditorialMarkers.Concat(PlaceholderTokens))
        {
            if (value.Length > MaxEvidenceTextCharacters - configuredCharacters)
            {
                configuredTextExceeded = true;
                break;
            }
            configuredCharacters += value.Length;
        }
        if (configuredTextExceeded)
            throw new ArgumentException("Configured workflow marker text exceeds the verification budget.");
        PackageManifestOptions.Validate();
    }
}

/// <summary>One exact package-change allowance used by unexpected-delta policy.</summary>
public sealed record DeliverablePackageChangeExpectation
{
    required public DeliverablePackageChangeKind Kind { get; init; }
    required public ChangeLocation Location { get; init; }
    public VerificationDigest? BeforeDigest { get; init; }
    public VerificationDigest? AfterDigest { get; init; }
    public string? BeforeValue { get; init; }
    public string? AfterValue { get; init; }
}

/// <summary>One render warning or font/content limitation reported by a companion renderer.</summary>
public sealed record DeliverableRenderDiagnostic
{
    required public DeliverableRenderDiagnosticKind Kind { get; init; }
    required public string Message { get; init; }
    public VerificationFindingSeverity Severity { get; init; } = VerificationFindingSeverity.Warning;
    /// <summary>The renderer's stable diagnostic code, when its protocol supplies one.</summary>
    public string? Code { get; init; }
    /// <summary>The renderer phase that observed the diagnostic.</summary>
    public string? Phase { get; init; }
    public string? OwningPartUri { get; init; }
    public string? AnchorId { get; init; }
    public string? Resource { get; init; }
    public string? FontName { get; init; }
    public string? SubstitutedFontName { get; init; }
    public string? Remediation { get; init; }
}

/// <summary>Bytes and renderer/document binding for one companion artifact.</summary>
public sealed record DeliverableCompanionArtifactInput
{
    required public string ArtifactId { get; init; }
    required public DeliverableArtifactRole Role { get; init; }
    required public string MediaType { get; init; }
    required public DeliverableArtifactAvailability Availability { get; init; }
    public byte[]? Bytes { get; init; }
    public string? UnavailableReason { get; init; }
    public long? PageCount { get; init; }
    public string? RendererFingerprint { get; init; }
    public VerificationDigest? SourcePackageDigest { get; init; }
    public VerificationDigest? PageMapDigest { get; init; }
    public IReadOnlyList<DeliverableRenderDiagnostic> RenderDiagnostics { get; init; } =
        Array.Empty<DeliverableRenderDiagnostic>();
}

/// <summary>Complete input to the single deliverable-verification operation.</summary>
public sealed record DeliverableVerificationRequest
{
    /// <summary>
    /// Exact delivered package bytes. Accepting bytes (instead of requiring an SDK-openable
    /// document) lets the same operation report corrupt, encrypted, and safety-limited inputs.
    /// </summary>
    required public byte[] DeliverableBytes { get; init; }

    /// <summary>Optional exact baseline package bytes used for disposition and delta checks.</summary>
    public byte[]? BaselineBytes { get; init; }

    /// <summary>
    /// Exact expected semantic changes. Generated <c>chg-*</c> IDs are ignored during matching;
    /// all semantic fields and values remain part of the identity.
    /// </summary>
    public SemanticChangeSet? ExpectedSemanticChanges { get; init; }

    public IReadOnlyList<DeliverablePackageChangeExpectation> ExpectedPackageChanges { get; init; } =
        Array.Empty<DeliverablePackageChangeExpectation>();

    public IReadOnlyList<DeliverableCompanionArtifactInput> CompanionArtifacts { get; init; } =
        Array.Empty<DeliverableCompanionArtifactInput>();
}

/// <summary>
/// One immutable delivery snapshot and the verification report produced from those exact bytes.
/// This is the preferred session-level handoff when the verified package will be written or sent.
/// </summary>
public sealed record VerifiedDeliverable
{
    /// <summary>Normal clean-save bytes, with internal projector anchor ids removed.</summary>
    required public byte[] DeliverableBytes { get; init; }

    /// <summary>Verification report whose raw package digest covers <see cref="DeliverableBytes"/>.</summary>
    required public DeliverableVerificationResult Report { get; init; }
}

/// <summary>Digest identity of one package inspected by the verification run.</summary>
public sealed record DeliverablePackageIdentity
{
    required public string PackageKind { get; init; }
    required public bool ManifestValid { get; init; }
    required public VerificationDigest RawPackageBytesDigest { get; init; }
    public VerificationDigest? OrderedOpcContentDigest { get; init; }
    public VerificationDigest? NormalizedSemanticDigest { get; init; }
}

/// <summary>Result of one named verification stage.</summary>
public sealed record DeliverableCheckResult
{
    required public string Check { get; init; }
    required public DeliverableCheckStatus Status { get; init; }
    required public int FindingCount { get; init; }
    public string? Diagnostic { get; init; }
}

/// <summary>A stable, actionable deliverable finding.</summary>
public sealed record DeliverableFinding
{
    required public string FindingId { get; init; }
    required public string Code { get; init; }
    required public DeliverableFindingCategory Category { get; init; }
    required public VerificationFindingSeverity Severity { get; init; }
    required public DeliverableFindingDisposition Disposition { get; init; }
    required public bool BlocksDelivery { get; init; }
    required public string Message { get; init; }
    required public string OwningPartUri { get; init; }
    public ChangeLocation? Location { get; init; }
    public string? AnchorId { get; init; }
    public string? Scope { get; init; }
    public string? XPath { get; init; }
    required public string Remediation { get; init; }
}

/// <summary>One deterministic package entry or relationship change.</summary>
public sealed record DeliverablePackageChange
{
    required public string ChangeId { get; init; }
    required public DeliverablePackageChangeKind Kind { get; init; }
    required public ChangeLocation Location { get; init; }
    public VerificationDigest? BeforeDigest { get; init; }
    public VerificationDigest? AfterDigest { get; init; }
    public string? BeforeValue { get; init; }
    public string? AfterValue { get; init; }
}

/// <summary>Stable summary of one semantic change; its fingerprint covers the full typed values.</summary>
public sealed record DeliverableSemanticChange
{
    required public string ChangeId { get; init; }
    required public string Fingerprint { get; init; }
    required public SemanticChangeOperation Operation { get; init; }
    required public SemanticChangeFamily Family { get; init; }
    required public string PartUri { get; init; }
    required public string Path { get; init; }
    public string? LeftAnchor { get; init; }
    public string? RightAnchor { get; init; }
}

/// <summary>Digest-covered semantic comparison metadata included in the report.</summary>
public sealed record DeliverableSemanticDelta
{
    required public string Schema { get; init; }
    required public int SchemaVersion { get; init; }
    required public int ChangeCount { get; init; }
    required public VerificationDigest CanonicalDigest { get; init; }
    public IReadOnlyList<DeliverableSemanticChange> Changes { get; init; } =
        Array.Empty<DeliverableSemanticChange>();
}

/// <summary>Digest-addressed metadata for one companion artifact.</summary>
public sealed record DeliverableArtifactMetadata
{
    required public string ArtifactId { get; init; }
    required public DeliverableArtifactRole Role { get; init; }
    required public string MediaType { get; init; }
    required public DeliverableArtifactAvailability Availability { get; init; }
    public long? ByteLength { get; init; }
    public VerificationDigest? Digest { get; init; }
    public string? UnavailableReason { get; init; }
    public long? PageCount { get; init; }
    public string? RendererFingerprint { get; init; }
    public VerificationDigest? SourcePackageDigest { get; init; }
    public VerificationDigest? PageMapDigest { get; init; }
    required public int RenderDiagnosticCount { get; init; }
}

/// <summary>
/// Versioned, deterministic report deciding whether a DOCX and supplied companion artifacts satisfy
/// the selected delivery policy.
/// </summary>
public sealed record DeliverableVerificationResult
{
    public const string SchemaId =
        "https://docxodus.dev/schemas/verification/deliverable-verification/v1";

    public string Schema { get; init; } = SchemaId;
    public int SchemaVersion { get; init; } = 1;
    required public DeliverableVerificationMode Mode { get; init; }
    required public DeliverableVerificationDecision Decision { get; init; }
    required public bool AnalysisCompleted { get; init; }
    required public bool BaselineCompared { get; init; }
    public DeliverablePackageIdentity? BaselinePackage { get; init; }
    required public DeliverablePackageIdentity DeliverablePackage { get; init; }
    public IReadOnlyList<DeliverableCheckResult> Checks { get; init; } =
        Array.Empty<DeliverableCheckResult>();
    public IReadOnlyList<DeliverableFinding> Findings { get; init; } =
        Array.Empty<DeliverableFinding>();
    public IReadOnlyList<DeliverableFinding> ResolvedFindings { get; init; } =
        Array.Empty<DeliverableFinding>();
    public DeliverableSemanticDelta? SemanticDelta { get; init; }
    public IReadOnlyList<DeliverablePackageChange> PackageChanges { get; init; } =
        Array.Empty<DeliverablePackageChange>();
    public IReadOnlyList<DeliverableArtifactMetadata> CompanionArtifacts { get; init; } =
        Array.Empty<DeliverableArtifactMetadata>();

    /// <summary>Compact UTF-8 bytes used whenever the report is hashed or attached as evidence.</summary>
    public byte[] ToCanonicalUtf8Bytes() => JsonSerializer.SerializeToUtf8Bytes(
        this, JsonOptions.Canonical.DeliverableVerificationResult);

    public string ToCanonicalJson() => Encoding.UTF8.GetString(ToCanonicalUtf8Bytes());

    public string ToJson(bool indented = true) => JsonSerializer.Serialize(
        this, (indented ? JsonOptions.Indented : JsonOptions.Canonical)
            .DeliverableVerificationResult);

    internal static bool IsExactCanonical(ReadOnlySpan<byte> bytes)
    {
        try
        {
            var result = JsonSerializer.Deserialize(
                bytes, JsonOptions.Canonical.DeliverableVerificationResult);
            return result is not null
                && string.Equals(result.Schema, SchemaId, StringComparison.Ordinal)
                && result.SchemaVersion == 1
                && bytes.SequenceEqual(result.ToCanonicalUtf8Bytes());
        }
        catch (JsonException)
        {
            return false;
        }
    }

    private static class JsonOptions
    {
        internal static readonly DeliverableVerificationJsonContext Canonical =
            Create(indented: false);
        internal static readonly DeliverableVerificationJsonContext Indented =
            Create(indented: true);

        private static DeliverableVerificationJsonContext Create(bool indented)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = indented,
            };
            // Generic enum converters are trim/AOT-safe. Register them explicitly because the
            // source-generation switch emits enum member names verbatim, while schema v1's
            // durable wire vocabulary is camelCase.
            options.Converters.Add(new JsonStringEnumConverter<DeliverableVerificationMode>(
                JsonNamingPolicy.CamelCase));
            options.Converters.Add(new JsonStringEnumConverter<DeliverableVerificationDecision>(
                JsonNamingPolicy.CamelCase));
            options.Converters.Add(new JsonStringEnumConverter<DeliverableCheckStatus>(
                JsonNamingPolicy.CamelCase));
            options.Converters.Add(new JsonStringEnumConverter<DeliverableFindingCategory>(
                JsonNamingPolicy.CamelCase));
            options.Converters.Add(new JsonStringEnumConverter<VerificationFindingSeverity>(
                JsonNamingPolicy.CamelCase));
            options.Converters.Add(new JsonStringEnumConverter<DeliverableFindingDisposition>(
                JsonNamingPolicy.CamelCase));
            options.Converters.Add(new JsonStringEnumConverter<DeliverablePackageChangeKind>(
                JsonNamingPolicy.CamelCase));
            options.Converters.Add(new JsonStringEnumConverter<SemanticChangeOperation>(
                JsonNamingPolicy.CamelCase));
            options.Converters.Add(new JsonStringEnumConverter<SemanticChangeFamily>(
                JsonNamingPolicy.CamelCase));
            options.Converters.Add(new JsonStringEnumConverter<DeliverableArtifactRole>(
                JsonNamingPolicy.CamelCase));
            options.Converters.Add(new JsonStringEnumConverter<DeliverableArtifactAvailability>(
                JsonNamingPolicy.CamelCase));
            return new DeliverableVerificationJsonContext(options);
        }
    }
}

/// <summary>Trim/AOT-safe metadata for the durable deliverable report wire contract.</summary>
[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(DeliverableVerificationResult))]
internal partial class DeliverableVerificationJsonContext : JsonSerializerContext
{
}
