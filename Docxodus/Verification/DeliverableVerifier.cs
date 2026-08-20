// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Globalization;
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
        options = options with
        {
            PackageManifestOptions = options.PackageManifestOptions with { },
            EditorialMarkers = options.EditorialMarkers.ToArray(),
            PlaceholderTokens = options.PlaceholderTokens.ToArray(),
        };
        ValidatePackageByteBudget(request, options);
        ValidateRequest(request, options);

        // Snapshot every caller-owned input that can be admitted under the configured limits before
        // inspection. Oversized/unavailable artifact payloads are retained only for their immutable
        // array length; the artifact inspector never reads their contents or hashes them.
        var deliverableBytes = request.DeliverableBytes.ToArray();
        var baselineBytes = request.BaselineBytes?.ToArray();
        var expectedPackageChanges = request.ExpectedPackageChanges.ToArray();
        long admittedArtifactBytes = 0;
        var artifacts = request.CompanionArtifacts
            .OrderBy(artifact => artifact.ArtifactId, StringComparer.Ordinal)
            .Select(artifact =>
        {
            bool admitBytes = artifact.Availability == DeliverableArtifactAvailability.Available
                && artifact.Bytes is { } bytes
                && bytes.LongLength <= options.MaxCompanionArtifactBytes
                && bytes.LongLength <= options.MaxTotalCompanionArtifactBytes - admittedArtifactBytes;
            if (admitBytes) admittedArtifactBytes += artifact.Bytes!.LongLength;
            return artifact with
            {
                Bytes = admitBytes ? artifact.Bytes!.ToArray() : artifact.Bytes,
                RenderDiagnostics = artifact.RenderDiagnostics.ToArray(),
            };
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

        var artifactMetadata = DeliverableArtifactInspector.Inspect(
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
            if (CanContinueBoundedPackageInspection(baseline.Inspection)
                && CanContinueBoundedPackageInspection(deliverable.Inspection)
                && observations.Count < options.MaxFindings
                && baseline.Inspection.Manifest.Relationships.Count
                    <= options.MaxDetectorRelationships
                && deliverable.Inspection.Manifest.Relationships.Count
                    <= options.MaxDetectorRelationships)
            {
                var packageDelta = PackageDelta.Compare(
                    baseline.Inspection.Manifest,
                    deliverable.Inspection.Manifest,
                    options.MaxReportedDeltaChanges);
                if (packageDelta.Complete)
                {
                    packageChanges = ProjectPackageChanges(packageDelta.Changes);
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
                    int before = observations.Count;
                    AddObservation(observations, options.MaxFindings,
                        DeliverableFindingObservation.Create(
                            "delta.package_change_limit_exceeded",
                            DeliverableFindingCategory.Delta,
                            VerificationFindingSeverity.Error,
                            "Package comparison exceeded the configured delta-record budget.",
                            "/",
                            "Reduce the package delta or deliberately raise MaxReportedDeltaChanges.",
                            new ChangeLocation { PropertyPath = "packageDelta" },
                            subjectKey: options.MaxReportedDeltaChanges.ToString(
                                CultureInfo.InvariantCulture)));
                    checks.Add(new DeliverableCheckResult
                    {
                        Check = "package_delta",
                        Status = DeliverableCheckStatus.UnavailableEvidence,
                        FindingCount = observations.Count - before,
                        Diagnostic = "delta record limit exceeded",
                    });
                    analysisCompleted = false;
                }
            }
            else
            {
                checks.Add(Skipped("package_delta",
                    baseline.Inspection.Manifest.Relationships.Count
                        > options.MaxDetectorRelationships
                        || deliverable.Inspection.Manifest.Relationships.Count
                            > options.MaxDetectorRelationships
                        ? "baseline or deliverable detector resource budget was exceeded"
                        : observations.Count >= options.MaxFindings
                            ? "finding limit was reached"
                            : "baseline or deliverable manifest is invalid"));
                analysisCompleted = false;
            }
            if (CanContinueBoundedWordInspection(baseline.Inspection)
                && CanContinueBoundedWordInspection(deliverable.Inspection)
                && !baseline.DetectorBudgetExhausted
                && !deliverable.DetectorBudgetExhausted
                && baseline.Observations.Count < options.MaxFindings
                && deliverable.Observations.Count < options.MaxFindings
                && observations.Count < options.MaxFindings)
            {
                int before = observations.Count;
                try
                {
                    semanticChanges = SemanticDiff.CompareBounded(
                        new WmlDocument("baseline.docx", baselineBytes!),
                        new WmlDocument("deliverable.docx", deliverableBytes),
                        new SemanticDiffOptions
                        {
                            IncludePackageChanges = false,
                            PackageOptions = options.PackageManifestOptions,
                        },
                        options.MaxReportedDeltaChanges);
                    if (semanticChanges.ChangeCount <= options.MaxReportedDeltaChanges)
                    {
                        semanticDelta = ProjectSemanticDelta(semanticChanges);
                        checks.Add(new DeliverableCheckResult
                        {
                            Check = "semantic_delta",
                            Status = DeliverableCheckStatus.Completed,
                            FindingCount = observations.Count - before,
                        });
                    }
                    else
                    {
                        semanticChanges = null;
                        AddObservation(observations, options.MaxFindings,
                            DeliverableFindingObservation.Create(
                                "delta.semantic_change_limit_exceeded",
                                DeliverableFindingCategory.Delta,
                                VerificationFindingSeverity.Error,
                                "Semantic comparison exceeded the configured delta-record budget.",
                                "/",
                                "Reduce the semantic delta or deliberately raise MaxReportedDeltaChanges.",
                                new ChangeLocation { PropertyPath = "semanticDelta" },
                                subjectKey: options.MaxReportedDeltaChanges.ToString(
                                    CultureInfo.InvariantCulture)));
                        checks.Add(new DeliverableCheckResult
                        {
                            Check = "semantic_delta",
                            Status = DeliverableCheckStatus.UnavailableEvidence,
                            FindingCount = observations.Count - before,
                            Diagnostic = "delta record limit exceeded",
                        });
                        analysisCompleted = false;
                    }
                }
                catch (SemanticChangeLimitExceededException)
                {
                    AddObservation(observations, options.MaxFindings,
                        DeliverableFindingObservation.Create(
                            "delta.semantic_change_limit_exceeded",
                            DeliverableFindingCategory.Delta,
                            VerificationFindingSeverity.Error,
                            "Semantic comparison exceeded the configured delta-record budget.",
                            "/",
                            "Reduce the semantic delta or deliberately raise MaxReportedDeltaChanges.",
                            new ChangeLocation { PropertyPath = "semanticDelta" },
                            subjectKey: options.MaxReportedDeltaChanges.ToString(
                                CultureInfo.InvariantCulture)));
                    checks.Add(new DeliverableCheckResult
                    {
                        Check = "semantic_delta",
                        Status = DeliverableCheckStatus.UnavailableEvidence,
                        FindingCount = observations.Count - before,
                        Diagnostic = "delta record limit exceeded",
                    });
                    analysisCompleted = false;
                }
                catch (Exception exception) when (DeliverableExceptionBoundary.IsRecoverable(exception))
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
                checks.Add(Skipped("semantic_delta",
                    baseline.DetectorBudgetExhausted || deliverable.DetectorBudgetExhausted
                        ? "baseline or deliverable detector resource budget was exceeded"
                        : baseline.Observations.Count >= options.MaxFindings
                          || deliverable.Observations.Count >= options.MaxFindings
                          || observations.Count >= options.MaxFindings
                            ? "baseline or deliverable finding limit was reached"
                            : "baseline or deliverable is not safely openable"));
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
        if (options.IncludeResolvedFindings
            && classified.Count + resolved.Count > options.MaxFindings)
        {
            // Current findings have priority over historical resolved evidence.
            resolved = resolved.Take(Math.Max(0, options.MaxFindings - classified.Count)).ToArray();
            analysisCompleted = false;
        }
        else if (!options.IncludeResolvedFindings)
        {
            resolved = Array.Empty<DeliverableFinding>();
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
            ResolvedFindings = resolved,
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
        bool detectorBudgetExhausted = false;
        if (CanContinueBoundedWordInspection(inspection))
        {
            if (observations.Count >= options.MaxFindings)
            {
                const string diagnostic = "finding limit reached before downstream inspection";
                checks.Add(Prefix(Unavailable("open_xml", diagnostic), prefix));
                checks.Add(Prefix(Unavailable("wordprocessing_closure", diagnostic), prefix));
                checks.Add(Prefix(Unavailable("workflow_and_revision_registry", diagnostic), prefix));
                completed = false;
            }
            else
            {
                var openXml = OpenXmlValidationInspector.Inspect(
                    bytes, options.OpenXmlVersion, observations, options.MaxFindings);
                checks.Add(Prefix(openXml, prefix));
                if (observations.Count >= options.MaxFindings)
                {
                    const string diagnostic = "finding limit reached before downstream inspection";
                    checks.Add(Prefix(Unavailable("wordprocessing_closure", diagnostic), prefix));
                    checks.Add(Prefix(Unavailable("workflow_and_revision_registry", diagnostic), prefix));
                    completed = false;
                }
                else
                {
                    var budget = new DeliverableInspectionBudget(options);
                    var graph = WordprocessingInspectionGraph.Build(inspection, budget);
                    var closure = WordprocessingClosureInspector.Inspect(
                        inspection, graph, observations, options.MaxFindings, budget);
                    var session = DeliverableSessionInspector.Inspect(
                        graph, options, observations, budget);
                    detectorBudgetExhausted = budget.Exhausted;
                    if (budget.Exhausted)
                        AddObservation(observations, options.MaxFindings,
                            DeliverableFindingObservation.Create(
                                "verification.resource_budget_exceeded",
                                DeliverableFindingCategory.Structure,
                                VerificationFindingSeverity.Error,
                                $"Semantic detector resource budget was exceeded ({budget.ExhaustedResource}).",
                                "/",
                                "Reduce document complexity or deliberately raise the bounded detector policy.",
                                new ChangeLocation { PropertyPath = "detectorBudget/" + budget.ExhaustedResource },
                                subjectKey: budget.ExhaustedResource));
                    checks.Add(Prefix(closure, prefix));
                    checks.Add(Prefix(session, prefix));
                    completed = openXml.Status == DeliverableCheckStatus.Completed
                        && closure.Status == DeliverableCheckStatus.Completed
                        && session.Status == DeliverableCheckStatus.Completed;
                }
            }
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
            DetectorBudgetExhausted = detectorBudgetExhausted,
        };
    }

    private static bool CanContinueBoundedWordInspection(PackageManifestInspection inspection)
    {
        var uri = inspection.Manifest.Facts.MainDocumentUri;
        return CanContinueBoundedPackageInspection(inspection)
            && uri is not null
            && inspection.Entries.Count(entry =>
                string.Equals(entry.Uri, uri, StringComparison.OrdinalIgnoreCase)
                && entry.Xml?.Root is not null) == 1;
    }

    private static bool CanContinueBoundedPackageInspection(PackageManifestInspection inspection)
    {
        if (inspection.Manifest.PackageKind != "opc"
            || inspection.Manifest.OrderedOpcContentDigest is null
            || inspection.Entries.Any(entry => !entry.PayloadWasRead
                || entry.ManifestEntry.IsEncrypted != false))
            return false;
        var unsafeCodes = new HashSet<string>(StringComparer.Ordinal)
        {
            "malformed_package",
            "entry_count_limit_exceeded",
            "entry_size_limit_exceeded",
            "entry_expansion_limit_exceeded",
            "total_expansion_limit_exceeded",
            "compression_ratio_limit_exceeded",
            "xml_size_limit_exceeded",
            "entry_uri_limit_exceeded",
            "unsafe_entry_path",
            "unsupported_ole_encryption",
            "unsupported_zip_encryption",
            "zip_encryption_detection_unavailable",
            "malformed_entry",
            "unreadable_entry",
            "content_types_unreadable",
            "relationship_part_unreadable",
        };
        return !inspection.Manifest.Findings.Any(finding => unsafeCodes.Contains(finding.Code));
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


    private static void CheckExpectedSemanticChanges(
        SemanticChangeSet actual,
        SemanticChangeSet? expected,
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings)
    {
        var unmatched = (expected?.Changes ?? Array.Empty<SemanticChange>())
            .Select(change => DeliverableVerificationIdentity.SemanticChangeFingerprint(change))
            .GroupBy(value => value, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        foreach (var change in actual.Changes)
        {
            var fingerprint = DeliverableVerificationIdentity.SemanticChangeFingerprint(change);
            if (unmatched.TryGetValue(fingerprint, out int remaining) && remaining > 0)
            {
                if (remaining == 1) unmatched.Remove(fingerprint);
                else unmatched[fingerprint] = remaining - 1;
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
        foreach (var pair in unmatched.OrderBy(item => item.Key, StringComparer.Ordinal))
        for (int index = 0; index < pair.Value; index++)
            AddObservation(observations, maximumFindings,
                DeliverableFindingObservation.Create(
                    "delta.semantic_change_missing", DeliverableFindingCategory.Delta,
                    VerificationFindingSeverity.Error,
                    "An approved semantic change was not observed in the deliverable.",
                    "/",
                    "Update the deliverable or remove the stale expected change.",
                    new ChangeLocation { PropertyPath = "expectedSemanticChanges" },
                    subjectKey: pair.Key));
    }

    private static void CheckExpectedPackageChanges(
        IReadOnlyList<DeliverablePackageChange> actual,
        IReadOnlyList<DeliverablePackageChangeExpectation> expected,
        ICollection<DeliverableFindingObservation> observations,
        int maximumFindings)
    {
        var unmatched = expected
            .OrderByDescending(PackageExpectationSpecificity)
            .ThenBy(item => item.Kind)
            .ThenBy(item => DeliverableVerificationIdentity.LocationKey(item.Location),
                StringComparer.Ordinal)
            .ThenBy(item => item.BeforeDigest?.Value, StringComparer.OrdinalIgnoreCase)
            .ThenBy(item => item.AfterDigest?.Value, StringComparer.OrdinalIgnoreCase)
            .ThenBy(item => item.BeforeValue, StringComparer.Ordinal)
            .ThenBy(item => item.AfterValue, StringComparer.Ordinal)
            .GroupBy(PackageExpectationBucket, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
        foreach (var change in actual)
        {
            var bucket = PackageChangeBucket(change);
            var candidates = unmatched.GetValueOrDefault(bucket);
            int match = candidates?.FindIndex(candidate => MatchesPackageChange(change, candidate)) ?? -1;
            if (match >= 0)
            {
                candidates!.RemoveAt(match);
                if (candidates.Count == 0) unmatched.Remove(bucket);
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
        foreach (var expectation in unmatched.Values.SelectMany(items => items)
                     .OrderBy(item => item.Kind)
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

    private static string PackageExpectationBucket(DeliverablePackageChangeExpectation expectation) =>
        string.Join("\u001f", ((int)expectation.Kind).ToString(CultureInfo.InvariantCulture),
            DeliverableVerificationIdentity.LocationKey(expectation.Location));

    private static string PackageChangeBucket(DeliverablePackageChange change) =>
        string.Join("\u001f", ((int)change.Kind).ToString(CultureInfo.InvariantCulture),
            DeliverableVerificationIdentity.LocationKey(change.Location));

    private static int PackageExpectationSpecificity(
        DeliverablePackageChangeExpectation expectation) =>
        (expectation.BeforeDigest is null ? 0 : 1)
        + (expectation.AfterDigest is null ? 0 : 1)
        + (expectation.BeforeValue is null ? 0 : 1)
        + (expectation.AfterValue is null ? 0 : 1);

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
        .OrderBy(item => item.OccurrenceKey, StringComparer.Ordinal);

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
            && observation.Category == DeliverableFindingCategory.Workflow
            && observation.Code is "workflow.placeholder_remaining"
                or "workflow.blank_run_remaining"
                or "workflow.content_control_placeholder"
                or "workflow.editorial_marker";
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

    private static DeliverableCheckResult Unavailable(string check, string diagnostic) => new()
    {
        Check = check,
        Status = DeliverableCheckStatus.UnavailableEvidence,
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

    private static void ValidateRequest(
        DeliverableVerificationRequest request,
        DeliverableVerificationOptions options)
    {
        if (request.ExpectedPackageChanges is null)
            throw new ArgumentException("ExpectedPackageChanges cannot be null.", nameof(request));
        if (request.CompanionArtifacts is null)
            throw new ArgumentException("CompanionArtifacts cannot be null.", nameof(request));
        int semanticExpectationCount = request.ExpectedSemanticChanges?.Changes.Count ?? 0;
        if (request.ExpectedPackageChanges.Count > options.MaxExpectedChanges
            || semanticExpectationCount > options.MaxExpectedChanges
            || request.ExpectedPackageChanges.Count
                > options.MaxExpectedChanges - semanticExpectationCount)
            throw new ArgumentException("Expected changes exceed the verification budget.", nameof(request));
        if (request.CompanionArtifacts.Count > options.MaxCompanionArtifacts)
            throw new ArgumentException("Companion artifacts exceed the verification budget.", nameof(request));
        if (request.ExpectedPackageChanges.Any(item => item is null))
            throw new ArgumentException("ExpectedPackageChanges cannot contain null.", nameof(request));
        if (request.CompanionArtifacts.Any(item => item is null))
            throw new ArgumentException("CompanionArtifacts cannot contain null.", nameof(request));

        int diagnosticCount = 0;
        long evidenceCharacters = 0;
        void AddEvidenceText(string? value)
        {
            if (value is null) return;
            if (value.Length > options.MaxEvidenceTextCharacters - evidenceCharacters)
                throw new ArgumentException(
                    "Evidence text exceeds the verification budget.", nameof(request));
            evidenceCharacters += value.Length;
        }
        void AddLocationText(ChangeLocation location)
        {
            AddEvidenceText(location.EntryUri);
            AddEvidenceText(location.OwnerUri);
            AddEvidenceText(location.RelationshipId);
            AddEvidenceText(location.TargetUri);
            AddEvidenceText(location.PropertyPath);
        }
        int semanticValueNodes = 0;
        foreach (var change in request.ExpectedSemanticChanges?.Changes
                     ?? Array.Empty<SemanticChange>())
        {
            if (!Enum.IsDefined(change.Operation) || !Enum.IsDefined(change.Family))
                throw new ArgumentOutOfRangeException(nameof(request),
                    "Expected semantic-change discriminator is invalid.");
            AddEvidenceText(change.Id);
            AddEvidenceText(change.PartUri);
            AddEvidenceText(change.Path);
            AddEvidenceText(change.LeftAnchor);
            AddEvidenceText(change.RightAnchor);
            AddEvidenceText(change.LeftScope);
            AddEvidenceText(change.RightScope);
            AddEvidenceText(change.MoveId);
            var pendingValues = new Stack<SemanticValue>();
            pendingValues.Push(change.After);
            pendingValues.Push(change.Before);
            while (pendingValues.Count > 0)
            {
                var value = pendingValues.Pop();
                if (value is null)
                    throw new ArgumentException(
                        "Expected semantic values cannot contain null.", nameof(request));
                if (semanticValueNodes >= options.MaxExpectedSemanticValueNodes)
                    throw new ArgumentException(
                        "Expected semantic values exceed the verification budget.", nameof(request));
                semanticValueNodes++;
                AddEvidenceText(value.StringValue);
                AddEvidenceText(value.DigestAlgorithm);
                AddEvidenceText(value.DigestProfile);
                AddEvidenceText(value.DigestValue);
                for (int index = value.Properties.Count - 1; index >= 0; index--)
                {
                    AddEvidenceText(value.Properties[index].Name);
                    pendingValues.Push(value.Properties[index].Value);
                }
                for (int index = value.Items.Count - 1; index >= 0; index--)
                    pendingValues.Push(value.Items[index]);
            }
        }
        foreach (var expectation in request.ExpectedPackageChanges)
        {
            if (!Enum.IsDefined(expectation.Kind))
                throw new ArgumentOutOfRangeException(nameof(request), "Expected package-change kind is invalid.");
            if (expectation.Location is null)
                throw new ArgumentException("Expected package-change locations cannot be null.", nameof(request));
            ValidateDigest(expectation.BeforeDigest, nameof(request));
            ValidateDigest(expectation.AfterDigest, nameof(request));
            AddLocationText(expectation.Location);
            AddEvidenceText(expectation.BeforeValue);
            AddEvidenceText(expectation.AfterValue);
        }
        foreach (var artifact in request.CompanionArtifacts)
        {
            if (!Enum.IsDefined(artifact.Role) || !Enum.IsDefined(artifact.Availability))
                throw new ArgumentOutOfRangeException(nameof(request), "Artifact discriminator is invalid.");
            if (artifact.ArtifactId is null || artifact.MediaType is null)
                throw new ArgumentException("ArtifactId and MediaType cannot be null.", nameof(request));
            AddEvidenceText(artifact.ArtifactId);
            AddEvidenceText(artifact.MediaType);
            AddEvidenceText(artifact.UnavailableReason);
            AddEvidenceText(artifact.RendererFingerprint);
            ValidateDigest(artifact.SourcePackageDigest, nameof(request));
            ValidateDigest(artifact.PageMapDigest, nameof(request));
            if (artifact.RenderDiagnostics is null)
                throw new ArgumentException("RenderDiagnostics cannot be null.", nameof(request));
            if (artifact.RenderDiagnostics.Count > options.MaxRenderDiagnostics - diagnosticCount)
                throw new ArgumentException("Render diagnostics exceed the verification budget.", nameof(request));
            if (artifact.RenderDiagnostics.Any(item => item is null))
                throw new ArgumentException("RenderDiagnostics cannot contain null.", nameof(request));
            diagnosticCount += artifact.RenderDiagnostics.Count;
            foreach (var diagnostic in artifact.RenderDiagnostics)
            {
                if (!Enum.IsDefined(diagnostic.Kind) || !Enum.IsDefined(diagnostic.Severity))
                    throw new ArgumentOutOfRangeException(nameof(request), "Render diagnostic discriminator is invalid.");
                if (diagnostic.Message is null)
                    throw new ArgumentException("Render diagnostic messages cannot be null.", nameof(request));
                AddEvidenceText(diagnostic.Message);
                AddEvidenceText(diagnostic.OwningPartUri);
                AddEvidenceText(diagnostic.AnchorId);
                AddEvidenceText(diagnostic.FontName);
                AddEvidenceText(diagnostic.SubstitutedFontName);
                AddEvidenceText(diagnostic.Remediation);
            }
        }
    }

    private static void ValidatePackageByteBudget(
        DeliverableVerificationRequest request,
        DeliverableVerificationOptions options)
    {
        long deliverableLength = request.DeliverableBytes.LongLength;
        long baselineLength = request.BaselineBytes?.LongLength ?? 0;
        if (deliverableLength > options.MaxPackageBytes
            || baselineLength > options.MaxPackageBytes - deliverableLength)
            throw new ArgumentException(
                "Aggregate deliverable and baseline bytes exceed the verification budget.",
                nameof(request));
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

}
