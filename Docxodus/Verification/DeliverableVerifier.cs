// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Globalization;
using System.Xml;

namespace Docxodus.Verification;

/// <summary>
/// Performs bounded package, Open XML, structural, workflow, delta, and supplied-render-evidence
/// checks in one deterministic operation. Input byte arrays and documents are never mutated.
/// </summary>
public static class DeliverableVerifier
{
    /// <summary>Verify exact delivered bytes without a baseline.</summary>
    public static DeliverableVerificationResult VerifyDeliverable(
        byte[] deliverableBytes,
        DeliverableVerificationOptions? options = null) => VerifyDeliverable(
            new DeliverableVerificationRequest
            {
                DeliverableBytes = deliverableBytes,
            }, options);

    /// <summary>Verify exact delivered bytes and classify findings relative to exact baseline bytes.</summary>
    public static DeliverableVerificationResult VerifyDeliverable(
        byte[] baselineBytes,
        byte[] deliverableBytes,
        DeliverableVerificationOptions? options = null) => VerifyDeliverable(
            new DeliverableVerificationRequest
            {
                BaselineBytes = baselineBytes,
                DeliverableBytes = deliverableBytes,
            }, options);

    /// <summary>Verify an immutable document snapshot without a baseline.</summary>
    public static DeliverableVerificationResult VerifyDeliverable(
        WmlDocument deliverable,
        DeliverableVerificationOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(deliverable);
        return VerifyDeliverable(deliverable.DocumentByteArray, options);
    }

    /// <summary>Verify immutable document snapshots and classify findings relative to a baseline.</summary>
    public static DeliverableVerificationResult VerifyDeliverable(
        WmlDocument baseline,
        WmlDocument deliverable,
        DeliverableVerificationOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(baseline);
        ArgumentNullException.ThrowIfNull(deliverable);
        return VerifyDeliverable(baseline.DocumentByteArray, deliverable.DocumentByteArray, options);
    }

    /// <summary>Run the complete structured deliverable verification operation.</summary>
    public static DeliverableVerificationResult VerifyDeliverable(
        DeliverableVerificationRequest request,
        DeliverableVerificationOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(request.DeliverableBytes);
        options ??= new DeliverableVerificationOptions();
        options.Validate();
        ValidateRequest(request);

        // Snapshot every caller-owned collection/byte array before any inspection. The manifest,
        // SDK, session, and diff paths therefore all see the same exact package identity.
        var deliverableBytes = request.DeliverableBytes.ToArray();
        var baselineBytes = request.BaselineBytes?.ToArray();
        var expectedPackageChanges = request.ExpectedPackageChanges.ToArray();
        var artifacts = request.CompanionArtifacts.Select(artifact => artifact with
        {
            Bytes = artifact.Bytes?.ToArray(),
            RenderDiagnostics = artifact.RenderDiagnostics.ToArray(),
        }).ToArray();

        var deliverable = InspectPackage(deliverableBytes, options, "deliverable");
        DeliverableInspectionSnapshot? baseline = baselineBytes is null
            ? null
            : InspectPackage(baselineBytes, options, "baseline");

        var observations = deliverable.Observations.ToList();
        bool analysisCompleted = deliverable.AnalysisCompleted
            && (baseline?.AnalysisCompleted ?? true);
        var checks = baseline is null
            ? deliverable.Checks.ToList()
            : baseline.Checks.Concat(deliverable.Checks).ToList();

        var artifactMetadata = InspectArtifacts(
            artifacts,
            deliverable.Inspection.Manifest.RawPackageBytesDigest,
            options,
            observations,
            out var artifactCheck);
        checks.Add(artifactCheck);
        if (artifactCheck.Status != DeliverableCheckStatus.Completed)
            analysisCompleted = false;

        SemanticChangeSet? semanticChanges = null;
        DeliverableSemanticDelta? semanticDelta = null;
        IReadOnlyList<DeliverablePackageChange> packageChanges = Array.Empty<DeliverablePackageChange>();
        bool packageDeltaCompleted = false;
        if (baseline is not null)
        {
            if (baseline.Inspection.Manifest.IsValid && deliverable.Inspection.Manifest.IsValid)
            {
                packageChanges = ProjectPackageChanges(PackageDelta.Compare(
                    baseline.Inspection.Manifest, deliverable.Inspection.Manifest));
                packageDeltaCompleted = true;
                checks.Add(new DeliverableCheckResult
                {
                    Check = "package_delta",
                    Status = DeliverableCheckStatus.Completed,
                    FindingCount = 0,
                });
            }
            else
            {
                checks.Add(Skipped("package_delta", "baseline or deliverable manifest is invalid"));
                analysisCompleted = false;
            }
            if (CanOpenAsWord(baseline.Inspection) && CanOpenAsWord(deliverable.Inspection))
            {
                int before = observations.Count;
                try
                {
                    semanticChanges = SemanticDiff.Compare(
                        new WmlDocument("baseline.docx", baselineBytes!),
                        new WmlDocument("deliverable.docx", deliverableBytes),
                        new SemanticDiffOptions
                        {
                            IncludePackageChanges = false,
                            PackageOptions = options.PackageManifestOptions,
                        });
                    semanticDelta = ProjectSemanticDelta(semanticChanges);
                    checks.Add(new DeliverableCheckResult
                    {
                        Check = "semantic_delta",
                        Status = DeliverableCheckStatus.Completed,
                        FindingCount = observations.Count - before,
                    });
                }
                catch (Exception exception) when (exception is InvalidDataException or IOException
                    or ArgumentException or FormatException or InvalidOperationException
                    or PowerToolsDocumentException
                    or XmlException)
                {
                    AddObservation(observations, options.MaxFindings,
                        DeliverableFindingObservation.Create(
                            "delta.semantic_comparison_unavailable",
                            DeliverableFindingCategory.Delta,
                            VerificationFindingSeverity.Error,
                            $"Semantic comparison could not be completed ({exception.GetType().Name}).",
                            "/",
                            "Repair both packages so the bounded semantic comparison can complete.",
                            new ChangeLocation { PropertyPath = "semanticDelta" },
                            subjectKey: exception.GetType().FullName));
                    checks.Add(new DeliverableCheckResult
                    {
                        Check = "semantic_delta",
                        Status = DeliverableCheckStatus.UnavailableEvidence,
                        FindingCount = observations.Count - before,
                        Diagnostic = exception.GetType().Name,
                    });
                    analysisCompleted = false;
                }
            }
            else
            {
                checks.Add(Skipped("semantic_delta", "baseline or deliverable is not safely openable"));
                analysisCompleted = false;
            }
        }
        else if (request.ExpectedSemanticChanges is not null
                 || expectedPackageChanges.Length > 0
                 || options.FailOnUnexpectedChanges)
        {
            AddObservation(observations, options.MaxFindings,
                DeliverableFindingObservation.Create(
                    "delta.baseline_required",
                    DeliverableFindingCategory.Delta,
                    VerificationFindingSeverity.Error,
                    "Expected-change policy requires a baseline package.",
                    "/",
                    "Supply exact baseline package bytes or disable expected-change enforcement.",
                    new ChangeLocation { PropertyPath = "baselineBytes" }));
            checks.Add(Skipped("semantic_delta", "no baseline supplied"));
            analysisCompleted = false;
        }

        if (options.FailOnUnexpectedChanges && baseline is not null)
        {
            int before = observations.Count;
            if (semanticChanges is not null)
                CheckExpectedSemanticChanges(
                    semanticChanges, request.ExpectedSemanticChanges, observations, options.MaxFindings);
            if (packageDeltaCompleted)
                CheckExpectedPackageChanges(
                    packageChanges, expectedPackageChanges, observations, options.MaxFindings);
            bool expectedEvidenceAvailable = semanticChanges is not null && packageDeltaCompleted;
            checks.Add(new DeliverableCheckResult
            {
                Check = "expected_delta_policy",
                Status = expectedEvidenceAvailable
                    ? DeliverableCheckStatus.Completed
                    : DeliverableCheckStatus.UnavailableEvidence,
                FindingCount = observations.Count - before,
                Diagnostic = expectedEvidenceAvailable
                    ? null
                    : "semantic or package delta unavailable",
            });
            if (!expectedEvidenceAvailable) analysisCompleted = false;
        }

        // Reaching the global cap is fail-closed: another detector may have had evidence that
        // could not be retained, so a pass decision would overstate what was analyzed.
        if (observations.Count >= options.MaxFindings
            || (baseline?.Observations.Count ?? 0) >= options.MaxFindings)
            analysisCompleted = false;

        var classified = Classify(
            observations,
            baseline?.Observations,
            options,
            out var resolved);
        if (classified.Count + resolved.Count > options.MaxFindings)
        {
            // Current findings have priority over historical resolved evidence.
            resolved = resolved.Take(Math.Max(0, options.MaxFindings - classified.Count)).ToArray();
            analysisCompleted = false;
        }

        var decision = Decide(classified, options.Mode, analysisCompleted);
        return new DeliverableVerificationResult
        {
            Mode = options.Mode,
            Decision = decision,
            AnalysisCompleted = analysisCompleted,
            BaselineCompared = baseline is not null,
            BaselinePackage = baseline is null ? null : PackageIdentity(baseline.Inspection.Manifest),
            DeliverablePackage = PackageIdentity(deliverable.Inspection.Manifest),
            Checks = checks.ToArray(),
            Findings = classified,
            ResolvedFindings = options.IncludeResolvedFindings ? resolved : Array.Empty<DeliverableFinding>(),
            SemanticDelta = semanticDelta,
            PackageChanges = packageChanges,
            CompanionArtifacts = artifactMetadata,
        };
    }

    private static DeliverableInspectionSnapshot InspectPackage(
        byte[] bytes,
        DeliverableVerificationOptions options,
        string prefix)
    {
        var inspection = PackageManifestGenerator.Inspect(bytes, options.PackageManifestOptions);
        var observations = new List<DeliverableFindingObservation>();
        foreach (var finding in inspection.Manifest.Findings)
        {
            AddObservation(observations, options.MaxFindings, ManifestObservation(finding));
        }
        if (inspection.Manifest.Facts.MainDocumentUri is null)
        {
            AddObservation(observations, options.MaxFindings,
                DeliverableFindingObservation.Create(
                    "package.word_document_missing",
                    DeliverableFindingCategory.Package,
                    VerificationFindingSeverity.Error,
                    "The OPC package has no discoverable Word main-document part.",
                    "/",
                    "Restore the officeDocument relationship and its Word main-document target.",
                    new ChangeLocation { OwnerUri = "/", PropertyPath = "facts.mainDocumentUri" }));
        }

        var checks = new List<DeliverableCheckResult>
        {
            new()
            {
                Check = prefix + ".package_manifest",
                Status = DeliverableCheckStatus.Completed,
                FindingCount = observations.Count,
            },
        };
        bool completed = true;
        if (CanOpenAsWord(inspection))
        {
            var openXml = OpenXmlValidationInspector.Inspect(
                bytes, options.OpenXmlVersion, observations, options.MaxFindings);
            var closure = WordprocessingClosureInspector.Inspect(
                inspection, observations, options.MaxFindings);
            var session = DeliverableSessionInspector.Inspect(
                bytes, observations, options.MaxFindings);
            checks.Add(Prefix(openXml, prefix));
            checks.Add(Prefix(closure, prefix));
            checks.Add(Prefix(session, prefix));
            completed = openXml.Status == DeliverableCheckStatus.Completed
                && closure.Status == DeliverableCheckStatus.Completed
                && session.Status == DeliverableCheckStatus.Completed;
        }
        else
        {
            const string diagnostic = "package is not a bounded, parsed Word OPC package";
            checks.Add(Prefix(Skipped("open_xml", diagnostic), prefix));
            checks.Add(Prefix(Skipped("wordprocessing_closure", diagnostic), prefix));
            checks.Add(Prefix(Skipped("workflow_and_revision_registry", diagnostic), prefix));
            completed = false;
        }

        return new DeliverableInspectionSnapshot
        {
            Inspection = inspection,
            Observations = observations,
            Checks = checks,
            AnalysisCompleted = completed && observations.Count < options.MaxFindings,
        };
    }

    private static bool CanOpenAsWord(PackageManifestInspection inspection)
    {
        var uri = inspection.Manifest.Facts.MainDocumentUri;
        return inspection.Manifest.IsValid
            && inspection.Manifest.PackageKind == "opc"
            && uri is not null
            && inspection.Entries.Any(entry =>
                string.Equals(entry.Uri, uri, StringComparison.OrdinalIgnoreCase)
                && entry.Xml?.Root is not null);
    }

    private static DeliverableFindingObservation ManifestObservation(VerificationFinding finding)
    {
        bool relationship = finding.Code.Contains("relationship", StringComparison.Ordinal)
            || finding.Code.Contains("target", StringComparison.Ordinal);
        var owner = finding.Location?.OwnerUri
            ?? finding.Location?.EntryUri
            ?? "/";
        return DeliverableFindingObservation.Create(
            "package." + finding.Code,
            relationship ? DeliverableFindingCategory.Relationship : DeliverableFindingCategory.Package,
            finding.Severity,
            finding.Message,
            owner,
            ManifestRemediation(finding.Code),
            finding.Location,
            subjectKey: finding.Code);
    }

    private static string ManifestRemediation(string code)
    {
        if (code.Contains("relationship", StringComparison.Ordinal)
            || code.Contains("target", StringComparison.Ordinal))
            return "Repair the relationship declaration, owner, or target so the OPC graph closes.";
        if (code.Contains("content_type", StringComparison.Ordinal))
            return "Repair [Content_Types].xml so every part has one unambiguous content type.";
        if (code.Contains("limit", StringComparison.Ordinal)
            || code.Contains("ratio", StringComparison.Ordinal))
            return "Reduce the package size/expansion characteristics or deliberately raise the bounded inspection policy.";
        return "Repair or recreate the implicated package entry before delivery.";
    }

    private static IReadOnlyList<DeliverableArtifactMetadata> InspectArtifacts(
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
        var metadata = new List<DeliverableArtifactMetadata>(artifacts.Count);
        foreach (var artifact in artifacts.OrderBy(artifact => artifact.ArtifactId, StringComparer.Ordinal))
        {
            if (!seen.Add(artifact.ArtifactId))
                ArtifactFinding(observations, options.MaxFindings, artifact,
                    "artifact.id_duplicate", VerificationFindingSeverity.Error,
                    "Companion artifact ids must be unique.", "Assign a unique stable artifact id.");
            if (string.IsNullOrWhiteSpace(artifact.ArtifactId))
                ArtifactFinding(observations, options.MaxFindings, artifact,
                    "artifact.id_missing", VerificationFindingSeverity.Error,
                    "A companion artifact has no id.", "Assign a non-empty stable artifact id.");
            if (string.IsNullOrWhiteSpace(artifact.MediaType))
                ArtifactFinding(observations, options.MaxFindings, artifact,
                    "artifact.media_type_missing", VerificationFindingSeverity.Error,
                    "A companion artifact has no media type.", "Supply the artifact's MIME media type.");

            VerificationDigest? digest = null;
            long? length = null;
            if (artifact.Availability == DeliverableArtifactAvailability.Available)
            {
                if (artifact.Bytes is null)
                {
                    ArtifactFinding(observations, options.MaxFindings, artifact,
                        "artifact.bytes_missing", VerificationFindingSeverity.Error,
                        "An available companion artifact has no bytes.",
                        "Supply the artifact bytes or mark the artifact unavailable with a reason.");
                }
                else
                {
                    length = artifact.Bytes.LongLength;
                    if (length > options.MaxCompanionArtifactBytes
                        || length > options.MaxTotalCompanionArtifactBytes - total)
                    {
                        bounded = false;
                        ArtifactFinding(observations, options.MaxFindings, artifact,
                            "artifact.size_limit_exceeded", VerificationFindingSeverity.Error,
                            "Companion artifact bytes exceed the configured verification budget.",
                            "Reduce the artifact or deliberately raise the bounded artifact policy.");
                    }
                    else
                    {
                        total += length.Value;
                        digest = DeliverableVerificationIdentity.Digest(artifact.Bytes);
                    }
                }
            }
            else
            {
                if (artifact.Bytes is not null)
                    ArtifactFinding(observations, options.MaxFindings, artifact,
                        "artifact.unavailable_has_bytes", VerificationFindingSeverity.Warning,
                        "An unavailable artifact also supplied bytes; the bytes were ignored.",
                        "Mark the artifact available or omit its bytes.");
                if (string.IsNullOrWhiteSpace(artifact.UnavailableReason))
                    ArtifactFinding(observations, options.MaxFindings, artifact,
                        "artifact.unavailable_reason_missing", VerificationFindingSeverity.Warning,
                        "An unavailable artifact has no reason.",
                        "Record why the artifact could not be produced.");
            }

            if (artifact.SourcePackageDigest is null)
                ArtifactFinding(observations, options.MaxFindings, artifact,
                    "artifact.source_digest_missing", VerificationFindingSeverity.Warning,
                    "The companion artifact is not bound to source package bytes.",
                    "Record the exact delivered package SHA-256 as sourcePackageDigest.");
            else if (!DeliverableVerificationIdentity.DigestEquals(
                         artifact.SourcePackageDigest, packageDigest))
                ArtifactFinding(observations, options.MaxFindings, artifact,
                    "artifact.source_digest_mismatch", VerificationFindingSeverity.Error,
                    "The companion artifact names a different source package digest.",
                    "Regenerate the artifact from the delivered package or correct the binding.");

            if (artifact.PageCount is < 0)
                ArtifactFinding(observations, options.MaxFindings, artifact,
                    "artifact.page_count_invalid", VerificationFindingSeverity.Error,
                    "Companion artifact pageCount cannot be negative.",
                    "Supply a non-negative page count or omit it.");
            if (artifact.Role is DeliverableArtifactRole.Pdf or DeliverableArtifactRole.PageMap
                && string.IsNullOrWhiteSpace(artifact.RendererFingerprint))
                ArtifactFinding(observations, options.MaxFindings, artifact,
                    "artifact.renderer_fingerprint_missing", VerificationFindingSeverity.Warning,
                    "Layout-dependent evidence has no renderer fingerprint.",
                    "Record the renderer name/version/configuration used to produce this artifact.");

            foreach (var diagnostic in artifact.RenderDiagnostics
                         .OrderBy(diagnostic => diagnostic.Kind)
                         .ThenBy(diagnostic => diagnostic.OwningPartUri, StringComparer.Ordinal)
                         .ThenBy(diagnostic => diagnostic.AnchorId, StringComparer.Ordinal)
                         .ThenBy(diagnostic => diagnostic.FontName, StringComparer.Ordinal)
                         .ThenBy(diagnostic => diagnostic.SubstitutedFontName, StringComparer.Ordinal)
                         .ThenBy(diagnostic => diagnostic.Message, StringComparer.Ordinal))
            {
                var code = diagnostic.Kind switch
                {
                    DeliverableRenderDiagnosticKind.MissingFont => "render.missing_font",
                    DeliverableRenderDiagnosticKind.FontSubstitution => "render.font_substitution",
                    DeliverableRenderDiagnosticKind.UnsupportedContent => "render.unsupported_content",
                    _ => "render.warning",
                };
                var owner = string.IsNullOrWhiteSpace(diagnostic.OwningPartUri)
                    ? "/"
                    : diagnostic.OwningPartUri!;
                AddObservation(observations, options.MaxFindings,
                    DeliverableFindingObservation.Create(
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
                            PropertyPath = "artifacts/" + artifact.ArtifactId,
                        },
                        diagnostic.AnchorId,
                        subjectKey: string.Join("\u001f", artifact.ArtifactId, diagnostic.Kind,
                            diagnostic.FontName, diagnostic.SubstitutedFontName)));
            }

            metadata.Add(new DeliverableArtifactMetadata
            {
                ArtifactId = artifact.ArtifactId,
                Role = artifact.Role,
                MediaType = artifact.MediaType,
                Availability = artifact.Availability,
                ByteLength = length,
                Digest = digest,
                UnavailableReason = artifact.UnavailableReason,
                PageCount = artifact.PageCount,
                RendererFingerprint = artifact.RendererFingerprint,
                SourcePackageDigest = NormalizeDigest(artifact.SourcePackageDigest),
                PageMapDigest = NormalizeDigest(artifact.PageMapDigest),
                RenderDiagnosticCount = artifact.RenderDiagnostics.Count,
            });
        }

        check = new DeliverableCheckResult
        {
            Check = "companion_artifacts",
            Status = bounded ? DeliverableCheckStatus.Completed : DeliverableCheckStatus.UnavailableEvidence,
            FindingCount = observations.Count - before,
            Diagnostic = bounded ? null : "artifact byte budget exceeded",
        };
        return metadata;
    }

    private static void CheckExpectedSemanticChanges(
        SemanticChangeSet actual,
        SemanticChangeSet? expected,
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings)
    {
        var unmatched = (expected?.Changes ?? Array.Empty<SemanticChange>())
            .Select(change => DeliverableVerificationIdentity.SemanticChangeFingerprint(change))
            .ToList();
        foreach (var change in actual.Changes)
        {
            var fingerprint = DeliverableVerificationIdentity.SemanticChangeFingerprint(change);
            int match = unmatched.FindIndex(candidate => candidate == fingerprint);
            if (match >= 0)
            {
                unmatched.RemoveAt(match);
                continue;
            }
            AddObservation(observations, maximumFindings,
                DeliverableFindingObservation.Create(
                    "delta.semantic_change_unexpected", DeliverableFindingCategory.Delta,
                    VerificationFindingSeverity.Error,
                    $"Unexpected {change.Family} semantic change at '{change.Path}'.",
                    change.PartUri,
                    "Add this exact semantic change to the approved expectation or revert it.",
                    new ChangeLocation { EntryUri = change.PartUri, PropertyPath = change.Path },
                    change.RightAnchor ?? change.LeftAnchor,
                    change.RightScope ?? change.LeftScope,
                    subjectKey: fingerprint));
        }
        foreach (var fingerprint in unmatched.OrderBy(value => value, StringComparer.Ordinal))
            AddObservation(observations, maximumFindings,
                DeliverableFindingObservation.Create(
                    "delta.semantic_change_missing", DeliverableFindingCategory.Delta,
                    VerificationFindingSeverity.Error,
                    "An approved semantic change was not observed in the deliverable.",
                    "/",
                    "Update the deliverable or remove the stale expected change.",
                    new ChangeLocation { PropertyPath = "expectedSemanticChanges" },
                    subjectKey: fingerprint));
    }

    private static void CheckExpectedPackageChanges(
        IReadOnlyList<DeliverablePackageChange> actual,
        IReadOnlyList<DeliverablePackageChangeExpectation> expected,
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings)
    {
        var unmatched = expected.ToList();
        foreach (var change in actual)
        {
            int match = unmatched.FindIndex(candidate => MatchesPackageChange(change, candidate));
            if (match >= 0)
            {
                unmatched.RemoveAt(match);
                continue;
            }
            AddObservation(observations, maximumFindings,
                DeliverableFindingObservation.Create(
                    "delta.package_change_unexpected", DeliverableFindingCategory.Delta,
                    VerificationFindingSeverity.Error,
                    $"Unexpected package change '{change.Kind}' was observed.",
                    change.Location.OwnerUri ?? change.Location.EntryUri ?? "/",
                    "Add this exact package change to the approved expectation or revert it.",
                    change.Location, subjectKey: change.ChangeId));
        }
        foreach (var expectation in unmatched.OrderBy(item => item.Kind)
                     .ThenBy(item => DeliverableVerificationIdentity.LocationKey(item.Location), StringComparer.Ordinal))
            AddObservation(observations, maximumFindings,
                DeliverableFindingObservation.Create(
                    "delta.package_change_missing", DeliverableFindingCategory.Delta,
                    VerificationFindingSeverity.Error,
                    $"Approved package change '{expectation.Kind}' was not observed.",
                    expectation.Location.OwnerUri ?? expectation.Location.EntryUri ?? "/",
                    "Update the deliverable or remove the stale expected package change.",
                    expectation.Location,
                    subjectKey: string.Join("\u001f", expectation.Kind,
                        DeliverableVerificationIdentity.LocationKey(expectation.Location),
                        expectation.BeforeDigest?.Value, expectation.AfterDigest?.Value,
                        expectation.BeforeValue, expectation.AfterValue)));
    }

    private static IReadOnlyList<DeliverablePackageChange> ProjectPackageChanges(
        IReadOnlyList<PackageDeltaChange> changes) => changes.Select(change =>
        {
            var kind = change.Kind switch
            {
                PackageDeltaChangeKind.EntryAdded => DeliverablePackageChangeKind.EntryAdded,
                PackageDeltaChangeKind.EntryRemoved => DeliverablePackageChangeKind.EntryRemoved,
                PackageDeltaChangeKind.EntryModified => DeliverablePackageChangeKind.EntryModified,
                PackageDeltaChangeKind.RelationshipAdded => DeliverablePackageChangeKind.RelationshipAdded,
                PackageDeltaChangeKind.RelationshipRemoved => DeliverablePackageChangeKind.RelationshipRemoved,
                PackageDeltaChangeKind.RelationshipModified => DeliverablePackageChangeKind.RelationshipModified,
                _ => throw new ArgumentOutOfRangeException(nameof(change), change.Kind, null),
            };
            return new DeliverablePackageChange
            {
                ChangeId = "pkg-" + DeliverableVerificationIdentity.Token(
                    "docxodus.deliverable.package-change.v1",
                    ((int)kind).ToString(CultureInfo.InvariantCulture),
                    DeliverableVerificationIdentity.LocationKey(change.Location),
                    change.BeforeValue,
                    change.AfterValue),
                Kind = kind,
                Location = change.Location,
                BeforeDigest = change.BeforeDigest,
                AfterDigest = change.AfterDigest,
                BeforeValue = change.BeforeValue,
                AfterValue = change.AfterValue,
            };
        }).ToArray();

    private static bool MatchesPackageChange(
        DeliverablePackageChange actual,
        DeliverablePackageChangeExpectation expected)
    {
        if (actual.Kind != expected.Kind
            || !string.Equals(
                DeliverableVerificationIdentity.LocationKey(actual.Location),
                DeliverableVerificationIdentity.LocationKey(expected.Location),
                StringComparison.Ordinal))
            return false;
        if (expected.BeforeDigest is not null
            && !DeliverableVerificationIdentity.DigestEquals(
                actual.BeforeDigest, expected.BeforeDigest))
            return false;
        if (expected.AfterDigest is not null
            && !DeliverableVerificationIdentity.DigestEquals(actual.AfterDigest, expected.AfterDigest))
            return false;
        if (expected.BeforeValue is not null
            && !string.Equals(actual.BeforeValue, expected.BeforeValue, StringComparison.Ordinal))
            return false;
        return expected.AfterValue is null
            || string.Equals(actual.AfterValue, expected.AfterValue, StringComparison.Ordinal);
    }

    private static IReadOnlyList<DeliverableFinding> Classify(
        IReadOnlyList<DeliverableFindingObservation> current,
        IReadOnlyList<DeliverableFindingObservation>? baseline,
        DeliverableVerificationOptions options,
        out IReadOnlyList<DeliverableFinding> resolved)
    {
        var baselineGroups = (baseline ?? Array.Empty<DeliverableFindingObservation>())
            .GroupBy(item => item.IdentityKey, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => OrderedObservations(group).ToArray(), StringComparer.Ordinal);
        var currentGroups = current.GroupBy(item => item.IdentityKey, StringComparer.Ordinal)
            .OrderBy(group => group.Key, StringComparer.Ordinal);
        var findings = new List<DeliverableFinding>();
        var matched = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var group in currentGroups)
        {
            var items = OrderedObservations(group).ToArray();
            int baselineCount = baselineGroups.GetValueOrDefault(group.Key)?.Length ?? 0;
            for (int index = 0; index < items.Length; index++)
            {
                var disposition = baseline is null
                    ? DeliverableFindingDisposition.Unclassified
                    : index < baselineCount
                        ? DeliverableFindingDisposition.PreExisting
                        : DeliverableFindingDisposition.New;
                findings.Add(Materialize(items[index], disposition, index, options));
            }
            matched[group.Key] = Math.Min(items.Length, baselineCount);
        }

        var resolvedList = new List<DeliverableFinding>();
        if (baseline is not null)
        {
            foreach (var group in baselineGroups.OrderBy(pair => pair.Key, StringComparer.Ordinal))
            {
                int consumed = matched.GetValueOrDefault(group.Key);
                for (int index = consumed; index < group.Value.Length; index++)
                    resolvedList.Add(Materialize(group.Value[index],
                        DeliverableFindingDisposition.Resolved, index, options));
            }
        }

        findings.Sort(FindingComparison);
        resolvedList.Sort(FindingComparison);
        resolved = resolvedList;
        return findings;
    }

    private static IEnumerable<DeliverableFindingObservation> OrderedObservations(
        IEnumerable<DeliverableFindingObservation> observations) => observations
        .OrderBy(item => item.Code, StringComparer.Ordinal)
        .ThenBy(item => item.OwningPartUri, StringComparer.Ordinal)
        .ThenBy(item => DeliverableVerificationIdentity.LocationKey(item.Location), StringComparer.Ordinal)
        .ThenBy(item => item.Scope, StringComparer.Ordinal)
        .ThenBy(item => item.XPath, StringComparer.Ordinal)
        .ThenBy(item => item.Message, StringComparer.Ordinal)
        .ThenBy(item => item.Remediation, StringComparer.Ordinal);

    private static DeliverableFinding Materialize(
        DeliverableFindingObservation observation,
        DeliverableFindingDisposition disposition,
        int occurrence,
        DeliverableVerificationOptions options) => new()
        {
            FindingId = "fnd-" + DeliverableVerificationIdentity.Token(
                "docxodus.deliverable.finding-id.v1", observation.IdentityKey,
                occurrence.ToString(CultureInfo.InvariantCulture)),
            Code = observation.Code,
            Category = observation.Category,
            Severity = observation.Severity,
            Disposition = disposition,
            BlocksDelivery = Blocks(observation, disposition, options),
            Message = observation.Message,
            OwningPartUri = observation.OwningPartUri,
            Location = observation.Location,
            AnchorId = observation.AnchorId,
            Scope = observation.Scope,
            XPath = observation.XPath,
            Remediation = observation.Remediation,
        };

    private static bool Blocks(
        DeliverableFindingObservation observation,
        DeliverableFindingDisposition disposition,
        DeliverableVerificationOptions options)
    {
        if (disposition == DeliverableFindingDisposition.Resolved
            || options.Mode == DeliverableVerificationMode.ReportOnly
            || observation.Severity == VerificationFindingSeverity.Info)
            return false;
        if (options.Mode == DeliverableVerificationMode.Strict)
            return observation.Severity is VerificationFindingSeverity.Warning
                or VerificationFindingSeverity.Error;
        if (observation.Severity == VerificationFindingSeverity.Error)
            return !(observation.Category == DeliverableFindingCategory.OpenXml
                     && disposition == DeliverableFindingDisposition.PreExisting);
        return options.RequireNoPlaceholders
            && observation.Category == DeliverableFindingCategory.Workflow;
    }

    private static DeliverableVerificationDecision Decide(
        IReadOnlyList<DeliverableFinding> findings,
        DeliverableVerificationMode mode,
        bool analysisCompleted)
    {
        if (mode == DeliverableVerificationMode.ReportOnly)
            return DeliverableVerificationDecision.NotEvaluated;
        if (!analysisCompleted || findings.Any(finding => finding.BlocksDelivery))
            return DeliverableVerificationDecision.Failed;
        return findings.Any(finding =>
                finding.Disposition == DeliverableFindingDisposition.PreExisting
                && finding.Severity != VerificationFindingSeverity.Info)
            ? DeliverableVerificationDecision.PassedWithPreExistingFindings
            : DeliverableVerificationDecision.Passed;
    }

    private static DeliverableSemanticDelta ProjectSemanticDelta(SemanticChangeSet changes) => new()
    {
        Schema = changes.Schema,
        SchemaVersion = changes.SchemaVersion,
        ChangeCount = changes.ChangeCount,
        CanonicalDigest = DeliverableVerificationIdentity.Digest(changes.ToCanonicalUtf8Bytes()),
        Changes = changes.Changes.Select(change => new DeliverableSemanticChange
        {
            ChangeId = change.Id,
            Fingerprint = DeliverableVerificationIdentity.SemanticChangeFingerprint(change),
            Operation = change.Operation,
            Family = change.Family,
            PartUri = change.PartUri,
            Path = change.Path,
            LeftAnchor = change.LeftAnchor,
            RightAnchor = change.RightAnchor,
        }).ToArray(),
    };

    private static DeliverablePackageIdentity PackageIdentity(PackageManifest manifest) => new()
    {
        PackageKind = manifest.PackageKind,
        ManifestValid = manifest.IsValid,
        RawPackageBytesDigest = manifest.RawPackageBytesDigest,
        OrderedOpcContentDigest = manifest.OrderedOpcContentDigest,
        NormalizedSemanticDigest = manifest.NormalizedSemanticDigest,
    };

    private static void ArtifactFinding(
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings,
        DeliverableCompanionArtifactInput artifact,
        string code,
        VerificationFindingSeverity severity,
        string message,
        string remediation) => AddObservation(observations, maximumFindings,
        DeliverableFindingObservation.Create(
            code, DeliverableFindingCategory.Artifact, severity, message, "/", remediation,
            new ChangeLocation { PropertyPath = "artifacts/" + artifact.ArtifactId },
            subjectKey: artifact.ArtifactId));

    private static void AddObservation(
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings,
        DeliverableFindingObservation observation)
    {
        if (observations.Count < maximumFindings) observations.Add(observation);
    }

    private static DeliverableCheckResult Prefix(DeliverableCheckResult check, string prefix) =>
        check with { Check = prefix + "." + check.Check };

    private static DeliverableCheckResult Skipped(string check, string diagnostic) => new()
    {
        Check = check,
        Status = DeliverableCheckStatus.SkippedPrerequisiteFailed,
        FindingCount = 0,
        Diagnostic = diagnostic,
    };

    private static int FindingComparison(DeliverableFinding left, DeliverableFinding right)
    {
        int result = right.Severity.CompareTo(left.Severity);
        if (result != 0) return result;
        result = string.CompareOrdinal(left.Code, right.Code);
        if (result != 0) return result;
        result = string.CompareOrdinal(left.OwningPartUri, right.OwningPartUri);
        if (result != 0) return result;
        result = string.CompareOrdinal(
            DeliverableVerificationIdentity.LocationKey(left.Location),
            DeliverableVerificationIdentity.LocationKey(right.Location));
        return result != 0 ? result : string.CompareOrdinal(left.FindingId, right.FindingId);
    }

    private static void ValidateRequest(DeliverableVerificationRequest request)
    {
        if (request.ExpectedPackageChanges is null)
            throw new ArgumentException("ExpectedPackageChanges cannot be null.", nameof(request));
        if (request.CompanionArtifacts is null)
            throw new ArgumentException("CompanionArtifacts cannot be null.", nameof(request));
        if (request.ExpectedPackageChanges.Any(item => item is null))
            throw new ArgumentException("ExpectedPackageChanges cannot contain null.", nameof(request));
        if (request.CompanionArtifacts.Any(item => item is null))
            throw new ArgumentException("CompanionArtifacts cannot contain null.", nameof(request));
        foreach (var expectation in request.ExpectedPackageChanges)
        {
            if (!Enum.IsDefined(expectation.Kind))
                throw new ArgumentOutOfRangeException(nameof(request), "Expected package-change kind is invalid.");
            if (expectation.Location is null)
                throw new ArgumentException("Expected package-change locations cannot be null.", nameof(request));
            ValidateDigest(expectation.BeforeDigest, nameof(request));
            ValidateDigest(expectation.AfterDigest, nameof(request));
        }
        foreach (var artifact in request.CompanionArtifacts)
        {
            if (!Enum.IsDefined(artifact.Role) || !Enum.IsDefined(artifact.Availability))
                throw new ArgumentOutOfRangeException(nameof(request), "Artifact discriminator is invalid.");
            if (artifact.ArtifactId is null || artifact.MediaType is null)
                throw new ArgumentException("ArtifactId and MediaType cannot be null.", nameof(request));
            ValidateDigest(artifact.SourcePackageDigest, nameof(request));
            ValidateDigest(artifact.PageMapDigest, nameof(request));
            if (artifact.RenderDiagnostics is null || artifact.RenderDiagnostics.Any(item => item is null))
                throw new ArgumentException("RenderDiagnostics cannot be null or contain null.", nameof(request));
            foreach (var diagnostic in artifact.RenderDiagnostics)
            {
                if (!Enum.IsDefined(diagnostic.Kind) || !Enum.IsDefined(diagnostic.Severity))
                    throw new ArgumentOutOfRangeException(nameof(request), "Render diagnostic discriminator is invalid.");
                if (diagnostic.Message is null)
                    throw new ArgumentException("Render diagnostic messages cannot be null.", nameof(request));
            }
        }
    }

    private static void ValidateDigest(VerificationDigest? digest, string parameterName)
    {
        if (digest is null) return;
        if (!string.Equals(digest.Algorithm, "SHA-256", StringComparison.OrdinalIgnoreCase)
            || digest.Value is null || digest.Value.Length != 64
            || digest.Value.Any(character => !Uri.IsHexDigit(character)))
            throw new ArgumentException("Verification digests must be SHA-256 with 64 hexadecimal characters.",
                parameterName);
    }

    private static VerificationDigest? NormalizeDigest(VerificationDigest? digest) => digest is null
        ? null
        : new VerificationDigest
        {
            Algorithm = "SHA-256",
            Value = digest.Value.ToLowerInvariant(),
        };
}
