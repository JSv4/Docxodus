// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.Verification;

/// <summary>
/// Proves the two generated-redline resolution paths without accepting or rejecting unrelated
/// review markup. Package inspection is performed before the Open XML SDK opens any input.
/// </summary>
public static class RedlineReversibilityVerifier
{
    /// <summary>
    /// Accept only generated revisions and compare with <paramref name="intendedFinalBytes"/>;
    /// reject only generated revisions and compare with <paramref name="baselineBytes"/>.
    /// </summary>
    public static RedlineReversibilityProofRun Prove(
        byte[] baselineBytes,
        byte[] intendedFinalBytes,
        byte[] redlineBytes,
        RedlineReversibilityProofOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(baselineBytes);
        ArgumentNullException.ThrowIfNull(intendedFinalBytes);
        ArgumentNullException.ThrowIfNull(redlineBytes);
        options ??= new RedlineReversibilityProofOptions();
        ValidateOptions(options);
        ValidatePackageByteBudget(
            baselineBytes, intendedFinalBytes, redlineBytes, options.MaxPackageBytes);
        var findingBudget = new ProofFindingBudget(options.MaxFindings);

        var baselineManifest = PackageManifestGenerator.Generate(
            baselineBytes, options.PackageManifestOptions);
        var finalManifest = PackageManifestGenerator.Generate(
            intendedFinalBytes, options.PackageManifestOptions);
        var redlineManifest = PackageManifestGenerator.Generate(
            redlineBytes, options.PackageManifestOptions);
        var sharedFindings = new List<RedlineProofFinding>();

        AppendManifestErrors(sharedFindings, findingBudget, "baseline", baselineManifest);
        AppendManifestErrors(sharedFindings, findingBudget, "intendedFinal", finalManifest);
        AppendManifestErrors(sharedFindings, findingBudget, "redline", redlineManifest);
        if (!baselineManifest.IsValid || !finalManifest.IsValid || !redlineManifest.IsValid)
        {
            return BuildRun(
                options,
                baselineManifest,
                finalManifest,
                redlineManifest,
                Array.Empty<RedlineRevisionClassification>(),
                null,
                null,
                sharedFindings,
                null,
                null);
        }

        bool strictRevisionMarkupSupported = AppendStrictRevisionFinding(
            sharedFindings, findingBudget, "baseline", baselineManifest);
        strictRevisionMarkupSupported &= AppendStrictRevisionFinding(
            sharedFindings, findingBudget, "intendedFinal", finalManifest);
        strictRevisionMarkupSupported &= AppendStrictRevisionFinding(
            sharedFindings, findingBudget, "redline", redlineManifest);
        if (!strictRevisionMarkupSupported)
        {
            return BuildRun(
                options,
                baselineManifest,
                finalManifest,
                redlineManifest,
                Array.Empty<RedlineRevisionClassification>(),
                null,
                null,
                sharedFindings,
                null,
                null);
        }

        bool revisionInventoryWithinLimit = AppendRevisionLimitFinding(
            sharedFindings, findingBudget, "baseline", baselineManifest, options.MaxRevisionElements);
        revisionInventoryWithinLimit &= AppendRevisionLimitFinding(
            sharedFindings, findingBudget, "intendedFinal", finalManifest,
            options.MaxRevisionElements);
        revisionInventoryWithinLimit &= AppendRevisionLimitFinding(
            sharedFindings, findingBudget, "redline", redlineManifest, options.MaxRevisionElements);
        if (!revisionInventoryWithinLimit)
        {
            return BuildRun(
                options,
                baselineManifest,
                finalManifest,
                redlineManifest,
                Array.Empty<RedlineRevisionClassification>(),
                null,
                null,
                sharedFindings,
                null,
                null);
        }

        IReadOnlyList<RevisionListEntry> baselineRevisions;
        IReadOnlyList<RevisionListEntry> finalRevisions;
        IReadOnlyList<RevisionListEntry> redlineRevisions;
        long baselineNativeRevisionElements;
        long finalNativeRevisionElements;
        long redlineNativeRevisionElements;
        bool baselineInventoryComplete;
        bool finalInventoryComplete;
        bool redlineInventoryComplete;
        bool baselineEvidenceComplete;
        bool finalEvidenceComplete;
        bool redlineEvidenceComplete;
        try
        {
            using var baseline = OpenProofSession(baselineBytes);
            (baselineRevisions, baselineNativeRevisionElements, baselineInventoryComplete,
                baselineEvidenceComplete) = baseline.GetRevisionInventory(
                    options.MaxRevisionElements,
                    options.MaxRevisionEvidenceItems,
                    options.MaxEvidenceTextCharacters);
            using var intendedFinal = OpenProofSession(intendedFinalBytes);
            (finalRevisions, finalNativeRevisionElements, finalInventoryComplete,
                finalEvidenceComplete) = intendedFinal.GetRevisionInventory(
                    options.MaxRevisionElements,
                    options.MaxRevisionEvidenceItems,
                    options.MaxEvidenceTextCharacters);
            using var redline = OpenProofSession(redlineBytes);
            (redlineRevisions, redlineNativeRevisionElements, redlineInventoryComplete,
                redlineEvidenceComplete) = redline.GetRevisionInventory(
                    options.MaxRevisionElements,
                    options.MaxRevisionEvidenceItems,
                    options.MaxEvidenceTextCharacters);
        }
        catch (Exception ex) when (DeliverableExceptionBoundary.IsRecoverable(ex))
        {
            findingBudget.Add(sharedFindings, Finding(
                "revision_inventory_failed",
                VerificationFindingSeverity.Error,
                $"A baseline, intended-final, or redline revision inventory could not be opened ({ex.GetType().Name}).",
                remediation: "Supply a valid WordprocessingML document whose tracked-change markup is supported."));
            return BuildRun(
                options,
                baselineManifest,
                finalManifest,
                redlineManifest,
                Array.Empty<RedlineRevisionClassification>(),
                null,
                null,
                sharedFindings,
                null,
                null);
        }

        bool liveInventoryWithinLimit = baselineInventoryComplete
            && finalInventoryComplete && redlineInventoryComplete;
        if (!baselineEvidenceComplete || !finalEvidenceComplete || !redlineEvidenceComplete)
        {
            liveInventoryWithinLimit = false;
            findingBudget.Add(sharedFindings, Finding(
                "revision_evidence_limit_exceeded",
                VerificationFindingSeverity.Error,
                "A revision inventory exceeded the configured evidence budget while it was being constructed.",
                new ChangeLocation { PropertyPath = "revisionClassifications" },
                remediation: "Reduce tracked-change evidence or deliberately raise the revision evidence limits."));
        }
        liveInventoryWithinLimit &= AppendLiveRevisionLimitFinding(
            sharedFindings, findingBudget, "baseline", baselineRevisions.Count,
            options.MaxRevisionElements);
        liveInventoryWithinLimit &= AppendLiveRevisionLimitFinding(
            sharedFindings, findingBudget, "intendedFinal", finalRevisions.Count,
            options.MaxRevisionElements);
        liveInventoryWithinLimit &= AppendLiveRevisionLimitFinding(
            sharedFindings, findingBudget, "redline", redlineRevisions.Count,
            options.MaxRevisionElements);
        liveInventoryWithinLimit &= AppendNativeRevisionLimitFinding(
            sharedFindings, findingBudget, "baseline", baselineNativeRevisionElements,
            options.MaxRevisionElements);
        liveInventoryWithinLimit &= AppendNativeRevisionLimitFinding(
            sharedFindings, findingBudget, "intendedFinal", finalNativeRevisionElements,
            options.MaxRevisionElements);
        liveInventoryWithinLimit &= AppendNativeRevisionLimitFinding(
            sharedFindings, findingBudget, "redline", redlineNativeRevisionElements,
            options.MaxRevisionElements);
        liveInventoryWithinLimit &= AppendRevisionInventoryCoverageFinding(
            sharedFindings, findingBudget, "baseline", baselineManifest,
            baselineNativeRevisionElements);
        liveInventoryWithinLimit &= AppendRevisionInventoryCoverageFinding(
            sharedFindings, findingBudget, "intendedFinal", finalManifest,
            finalNativeRevisionElements);
        liveInventoryWithinLimit &= AppendRevisionInventoryCoverageFinding(
            sharedFindings, findingBudget, "redline", redlineManifest,
            redlineNativeRevisionElements);
        liveInventoryWithinLimit &= AppendRevisionEvidenceLimitFinding(
            sharedFindings,
            findingBudget,
            baselineRevisions.Concat(finalRevisions).ToArray(),
            redlineRevisions,
            options);
        if (!liveInventoryWithinLimit)
        {
            return BuildRun(
                options,
                baselineManifest,
                finalManifest,
                redlineManifest,
                Array.Empty<RedlineRevisionClassification>(),
                null,
                null,
                sharedFindings,
                null,
                null);
        }

        var classifications = Classify(
            baselineRevisions,
            finalRevisions,
            redlineRevisions,
            sharedFindings,
            findingBudget);
        var generated = classifications
            .Where(item => item.Disposition == RedlineRevisionDisposition.Generated
                && item.Redline is not null)
            .Select(item => item.Redline!)
            .OrderBy(RevisionSortKey, StringComparer.Ordinal)
            .ToArray();
        var intendedFinalExclusive = classifications
            .Where(item => item.Disposition
                    == RedlineRevisionDisposition.IntendedFinalPreExisting
                && item.Redline is not null)
            .Select(item => item.Redline!)
            .OrderBy(RevisionSortKey, StringComparer.Ordinal)
            .ToArray();
        var preExisting = baselineRevisions
            .Select(ToIdentity)
            .OrderBy(RevisionSortKey, StringComparer.Ordinal)
            .ToArray();
        var acceptPreExisting = preExisting.Concat(finalRevisions.Select(ToIdentity))
            .GroupBy(RevisionSortKey, StringComparer.Ordinal)
            .Select(group => group.First())
            .OrderBy(RevisionSortKey, StringComparer.Ordinal)
            .ToArray();

        foreach (var revision in generated.Where(item =>
                     item.ResolutionStatus != RevisionResolutionStatus.Supported))
        {
            findingBudget.Add(sharedFindings, Finding(
                "generated_revision_not_resolvable",
                VerificationFindingSeverity.Error,
                $"Generated revision '{revision.Id}' is {revision.ResolutionStatus.ToString().ToLowerInvariant()}.",
                new ChangeLocation { EntryUri = revision.PartUri },
                revision.AnchorId,
                new[] { revision.Id },
                "Regenerate the redline with supported, unambiguous native revision markup."));
        }

        var classificationFailed = classifications.Any(item =>
                item.Disposition == RedlineRevisionDisposition.Conflicted)
            || generated.Any(item => item.ResolutionStatus != RevisionResolutionStatus.Supported);
        if (classificationFailed)
        {
            return BuildRun(
                options,
                baselineManifest,
                finalManifest,
                redlineManifest,
                classifications,
                null,
                null,
                sharedFindings,
                null,
                null);
        }

        var accept = EvaluatePath(
            RedlineProofDirection.AcceptToFinal,
            accept: true,
            redlineBytes,
            intendedFinalBytes,
            finalManifest,
            generated,
            acceptPreExisting,
            Array.Empty<RedlineRevisionIdentity>(),
            options,
            findingBudget);
        var reject = EvaluatePath(
            RedlineProofDirection.RejectToBaseline,
            accept: false,
            redlineBytes,
            baselineBytes,
            baselineManifest,
            generated,
            preExisting,
            intendedFinalExclusive,
            options,
            findingBudget);

        return BuildRun(
            options,
            baselineManifest,
            finalManifest,
            redlineManifest,
            classifications,
            accept.Result,
            reject.Result,
            sharedFindings,
            accept.OutputBytes,
            reject.OutputBytes);
    }

    private static PathEvaluation EvaluatePath(
        RedlineProofDirection direction,
        bool accept,
        byte[] redlineBytes,
        byte[] expectedBytes,
        PackageManifest expectedManifest,
        IReadOnlyList<RedlineRevisionIdentity> generated,
        IReadOnlyList<RedlineRevisionIdentity> preExisting,
        IReadOnlyList<RedlineRevisionIdentity> endpointExclusive,
        RedlineReversibilityProofOptions options,
        ProofFindingBudget findingBudget)
    {
        var findings = new List<RedlineProofFinding>();
        var requestedIds = generated.Select(item => item.Id).ToArray();
        var explicitlyResolved = new HashSet<string>(StringComparer.Ordinal);
        var implicitlyResolved = new HashSet<string>(StringComparer.Ordinal);
        var generatedOwnership = IndexRevisionOwnership(generated);
        var pendingOwnershipKeys = generatedOwnership.Keys.ToHashSet(StringComparer.Ordinal);
        byte[]? outputBytes = null;
        IReadOnlyList<RevisionListEntry> survivingEntries = Array.Empty<RevisionListEntry>();
        bool completed = true;
        bool stopResolution = false;

        try
        {
            using var expectedSession = OpenProofSession(expectedBytes);
            IReadOnlyList<int> protectedNumberingIds = accept
                ? Array.Empty<int>()
                : expectedSession.DefinedNumberingIds();
            IReadOnlyList<int> protectedAbstractNumberingIds = accept
                ? Array.Empty<int>()
                : expectedSession.DefinedAbstractNumberingIds();
            var protectedRelationshipKeys = expectedSession.RevisionRelationshipKeys();
            var protectedEmptyContainerKeys =
                expectedSession.EmptyRevisionPropertyContainerKeys();
            using var session = OpenProofSession(
                redlineBytes,
                proofSafeResolution: true,
                protectedNumberingIds: protectedNumberingIds,
                protectedAbstractNumberingIds: protectedAbstractNumberingIds,
                protectedRelationshipKeys: protectedRelationshipKeys,
                protectedEmptyContainerKeys: protectedEmptyContainerKeys);
            foreach (var requested in generated)
            {
                var requestedKeys = RevisionOwnershipKeys(requested)
                    .Where(pendingOwnershipKeys.Contains).ToArray();
                while (requestedKeys.Length > 0)
                {
                    var liveInventory = session.GetRevisionInventory(
                        options.MaxRevisionElements,
                        options.MaxRevisionEvidenceItems,
                        options.MaxEvidenceTextCharacters);
                    var liveEntries = liveInventory.Entries;
                    if (!liveInventory.EvidenceComplete)
                    {
                        completed = false;
                        findingBudget.Add(findings, Finding(
                            "revision_evidence_limit_exceeded",
                            VerificationFindingSeverity.Error,
                            $"The {DirectionName(direction)} live inventory exceeded the configured evidence budget.",
                            new ChangeLocation { PropertyPath = "proofPath/revisionInventory" },
                            revisionIds: new[] { requested.Id },
                            remediation: "Reduce tracked-change evidence or deliberately raise the revision evidence limits.",
                            direction: direction));
                        stopResolution = true;
                        break;
                    }
                    if (!liveInventory.Complete
                        || liveEntries.Count > options.MaxRevisionElements
                        || liveInventory.NativeElementCount > options.MaxRevisionElements)
                    {
                        completed = false;
                        findingBudget.Add(findings, Finding(
                            "revision_inventory_limit_exceeded_during_resolution",
                            VerificationFindingSeverity.Error,
                            $"The {DirectionName(direction)} live revision inventory exceeded "
                                + "the configured limit during selective resolution.",
                            new ChangeLocation { PropertyPath = "proofPath/revisionInventory" },
                            revisionIds: new[] { requested.Id },
                            remediation: "Regenerate a redline whose live revision topology remains bounded.",
                            direction: direction));
                        stopResolution = true;
                        break;
                    }

                    var liveRevisions = liveEntries.Select(ToIdentity)
                        .OrderBy(RevisionSortKey, StringComparer.Ordinal).ToArray();
                    var current = liveRevisions.FirstOrDefault(item =>
                        RevisionOwnershipKeys(item).Intersect(
                            requestedKeys, StringComparer.Ordinal).Any());
                    if (current is null)
                    {
                        // A previously resolved atomic envelope can consume nested generated
                        // carriers. They are accounted for below as implicit resolutions.
                        pendingOwnershipKeys.ExceptWith(requestedKeys);
                        break;
                    }

                    var currentKeys = RevisionOwnershipKeys(current);
                    bool exclusivelyPendingGenerated = currentKeys.All(key =>
                        pendingOwnershipKeys.Contains(key)
                        && generatedOwnership.TryGetValue(key, out var owners)
                        && owners.Count == 1
                        && OwnershipEquivalent(owners[0], current));
                    if (!exclusivelyPendingGenerated)
                    {
                        completed = false;
                        findingBudget.Add(findings, Finding(
                            "generated_revision_ownership_changed",
                            VerificationFindingSeverity.Error,
                            $"Live revision '{current.Id}' regrouped generated carriers with "
                                + "foreign or changed revision ownership during resolution.",
                            new ChangeLocation { EntryUri = current.PartUri },
                            current.AnchorId,
                            new[] { requested.Id, current.Id }
                                .Distinct(StringComparer.Ordinal).ToArray(),
                            "Regenerate a redline whose live groups contain only unchanged generated carriers.",
                            direction));
                        stopResolution = true;
                        break;
                    }

                    if (current.ResolutionStatus != RevisionResolutionStatus.Supported)
                    {
                        completed = false;
                        findingBudget.Add(findings, Finding(
                            "generated_revision_became_unresolvable",
                            VerificationFindingSeverity.Error,
                            $"Generated revision '{requested.Id}' became {current.ResolutionStatus.ToString().ToLowerInvariant()} during resolution.",
                            new ChangeLocation { EntryUri = current.PartUri },
                            current.AnchorId,
                            new[] { requested.Id, current.Id }
                                .Distinct(StringComparer.Ordinal).ToArray(),
                            "Regenerate a redline whose generated revision groups can be resolved independently.",
                            direction));
                        stopResolution = true;
                        break;
                    }

                    var edit = accept
                        ? session.AcceptRevision(current.Id)
                        : session.RejectRevision(current.Id);
                    if (!edit.Success)
                    {
                        completed = false;
                        findingBudget.Add(findings, Finding(
                            "generated_revision_resolution_failed",
                            VerificationFindingSeverity.Error,
                            $"Revision '{current.Id}' could not be {(accept ? "accepted" : "rejected")}: " +
                            (edit.Error?.Message ?? "unknown resolver failure"),
                            new ChangeLocation { EntryUri = current.PartUri },
                            current.AnchorId,
                            new[] { requested.Id, current.Id }
                                .Distinct(StringComparer.Ordinal).ToArray(),
                            "Inspect the revision topology and regenerate the comparison redline.",
                            direction));
                        stopResolution = true;
                        break;
                    }
                    explicitlyResolved.Add(requested.Id);
                    pendingOwnershipKeys.ExceptWith(currentKeys);
                    requestedKeys = RevisionOwnershipKeys(requested)
                        .Where(pendingOwnershipKeys.Contains).ToArray();
                }

                if (stopResolution) break;
                if (!explicitlyResolved.Contains(requested.Id))
                {
                    implicitlyResolved.Add(requested.Id);
                    findingBudget.Add(findings, Finding(
                        "generated_revision_consumed_by_resolution",
                        VerificationFindingSeverity.Info,
                        $"Generated revision '{requested.Id}' was consumed while resolving a linked generated revision.",
                        new ChangeLocation { EntryUri = requested.PartUri },
                        requested.AnchorId,
                        new[] { requested.Id },
                        "No action is required when the path output matches its target and no generated revisions survive.",
                        direction));
                }
            }

            var survivingInventory = session.GetRevisionInventory(
                options.MaxRevisionElements,
                options.MaxRevisionEvidenceItems,
                options.MaxEvidenceTextCharacters);
            survivingEntries = survivingInventory.Entries;
            bool survivingInventoryWithinLimit = survivingInventory.Complete
                && survivingInventory.EvidenceComplete;
            if (!survivingInventory.EvidenceComplete)
            {
                findingBudget.Add(findings, Finding(
                    "revision_evidence_limit_exceeded",
                    VerificationFindingSeverity.Error,
                    $"The {DirectionName(direction)} output inventory exceeded the configured evidence budget.",
                    new ChangeLocation { PropertyPath = "proofPath/revisionInventory" },
                    revisionIds: Array.Empty<string>(),
                    remediation: "Reduce tracked-change evidence or deliberately raise the revision evidence limits.",
                    direction: direction));
            }
            survivingInventoryWithinLimit &= AppendLiveRevisionLimitFinding(
                findings,
                findingBudget,
                "proofOutput",
                survivingEntries.Count,
                options.MaxRevisionElements,
                direction);
            survivingInventoryWithinLimit &= AppendNativeRevisionLimitFinding(
                findings,
                findingBudget,
                "proofOutput",
                survivingInventory.NativeElementCount,
                options.MaxRevisionElements,
                direction);
            survivingInventoryWithinLimit &= AppendRevisionEvidenceLimitFinding(
                findings,
                findingBudget,
                Array.Empty<RevisionListEntry>(),
                survivingEntries,
                options,
                direction);
            if (!survivingInventoryWithinLimit)
            {
                completed = false;
                survivingEntries = Array.Empty<RevisionListEntry>();
            }
            outputBytes = session.Save(persistAnchorIds: false);
            if (outputBytes.LongLength > options.MaxPackageBytes)
            {
                completed = false;
                findingBudget.Add(findings, Finding(
                    "proof_output_package_limit_exceeded",
                    VerificationFindingSeverity.Error,
                    $"The {DirectionName(direction)} output exceeds the configured raw package budget.",
                    new ChangeLocation { PropertyPath = "proofPath/outputPackage" },
                    revisionIds: Array.Empty<string>(),
                    remediation: "Reduce the package or deliberately raise MaxPackageBytes.",
                    direction: direction));
                outputBytes = null;
            }
        }
        catch (Exception ex) when (DeliverableExceptionBoundary.IsRecoverable(ex))
        {
            completed = false;
            findingBudget.Add(findings, Finding(
                "proof_path_execution_failed",
                VerificationFindingSeverity.Error,
                $"The {DirectionName(direction)} path failed ({ex.GetType().Name}).",
                revisionIds: requestedIds,
                remediation: "Inspect the redline package and its native revision topology.",
                direction: direction));
        }

        var surviving = survivingEntries
            .Select(ToIdentity)
            .OrderBy(RevisionSortKey, StringComparer.Ordinal)
            .ToArray();
        var survivingPreExisting = SelectSurvivingPreExisting(preExisting, surviving);
        var survivingGenerated = generated.Where(requested => surviving.Any(item =>
                string.Equals(item.Id, requested.Id, StringComparison.Ordinal)
                || RevisionOverlaps(requested, item)))
            .ToArray();
        if (survivingGenerated.Length > 0)
        {
            completed = false;
            findingBudget.Add(findings, Finding(
                "generated_revisions_survived_resolution",
                VerificationFindingSeverity.Error,
                $"The {DirectionName(direction)} path left {survivingGenerated.Length} generated revision(s) unresolved.",
                revisionIds: survivingGenerated.Select(item => item.Id).ToArray(),
                remediation: "Inspect the generated revision topology and selective resolver results.",
                direction: direction));
        }
        var resolvedIds = requestedIds.Where(explicitlyResolved.Contains).ToArray();
        var implicitlyResolvedIds = requestedIds.Where(implicitlyResolved.Contains).ToArray();
        if (resolvedIds.Length + implicitlyResolvedIds.Length != requestedIds.Length
            || pendingOwnershipKeys.Count > 0)
            completed = false;
        bool preExistingPreserved = VerifyPreExisting(
            direction, preExisting, surviving, findings, findingBudget);
        if (endpointExclusive.Count > 0)
        {
            var survivingOwnership = IndexRevisionOwnership(surviving);
            foreach (var revision in endpointExclusive.Where(revision =>
                         RevisionOwnershipKeys(revision).Any(survivingOwnership.ContainsKey)))
            {
                findingBudget.Add(findings, Finding(
                    "intended_final_revision_survived_reject_path",
                    VerificationFindingSeverity.Error,
                    $"Intended-final-only revision '{revision.Id}' survived the reject-to-baseline path.",
                    new ChangeLocation { EntryUri = revision.PartUri },
                    revision.AnchorId,
                    new[] { revision.Id },
                    "Generate a comparison envelope that removes final-only review state on rejection, or choose compatible endpoint review state.",
                    direction));
            }
        }

        if (outputBytes is null)
        {
            return new PathEvaluation(
                new RedlineProofPathResult
                {
                    Direction = direction,
                    Completed = false,
                    Equivalent = false,
                    RequestedRevisionIds = requestedIds,
                    ResolvedRevisionIds = resolvedIds,
                    ImplicitlyResolvedRevisionIds = implicitlyResolvedIds,
                    SurvivingPreExistingRevisions = survivingPreExisting,
                    PreExistingRevisionsPreserved = preExistingPreserved,
                    ModeledSemantic = SemanticUnavailable(),
                    NormalizedWholePackageEquivalent = false,
                    OrderedOpcContentEquivalent = false,
                    ExactPackageBytesEquivalent = false,
                    DivergenceAnalysisCompleted = false,
                    ExpectedPackage = ToPackageIdentity(expectedManifest),
                    ActualPackage = null,
                    FirstDivergence = null,
                    Divergences = Array.Empty<RedlinePackageDivergence>(),
                    Findings = findings,
                },
                null);
        }

        var actualManifest = PackageManifestGenerator.Generate(
            outputBytes, options.PackageManifestOptions);
        AppendManifestErrors(findings, findingBudget, "proofOutput", actualManifest, direction);
        if (!actualManifest.IsValid)
            completed = false;

        var semantic = CompareModeledSemantics(
            expectedBytes,
            outputBytes,
            options.PackageManifestOptions,
            options.MaxSemanticChanges);
        bool normalizedEquivalent = DigestEquals(
            expectedManifest.NormalizedSemanticDigest,
            actualManifest.NormalizedSemanticDigest);
        bool opcEquivalent = DigestEquals(
            expectedManifest.OrderedOpcContentDigest,
            actualManifest.OrderedOpcContentDigest);
        bool rawEquivalent = DigestEquals(
            expectedManifest.RawPackageBytesDigest,
            actualManifest.RawPackageBytesDigest);
        var divergenceEvaluation = CompareEntries(
            expectedManifest,
            actualManifest,
            generated,
            semantic.ModeledChanges,
            options.MaxPackageChanges);
        var divergences = divergenceEvaluation.Divergences;
        var firstDivergence = divergences.FirstOrDefault();
        var firstNormalizedDivergence = divergences.FirstOrDefault(divergence =>
            divergence.UnknownOrUnmodeled);
        if (!divergenceEvaluation.Completed)
        {
            completed = false;
            findingBudget.Add(findings, Finding(
                "package_divergence_limit_exceeded",
                VerificationFindingSeverity.Error,
                "Package comparison exceeded the configured change-record budget.",
                new ChangeLocation { PropertyPath = "proofPath/divergences" },
                revisionIds: Array.Empty<string>(),
                remediation: "Reduce the package delta or deliberately raise MaxPackageChanges.",
                direction: direction));
        }
        if (!semantic.Comparison.Available)
        {
            findingBudget.Add(findings, Finding(
                "modeled_semantic_comparison_unavailable",
                VerificationFindingSeverity.Error,
                semantic.Comparison.Diagnostic
                    ?? "The versioned modeled semantic comparer is unavailable.",
                revisionIds: requestedIds,
                remediation: "Run the proof with the semantic-diff component available.",
                direction: direction));
        }
        else if (semantic.Comparison.Equivalent != true)
        {
            var firstSemanticChange = semantic.FirstModeledChange;
            var firstSemanticAnchor = firstSemanticChange?.RightAnchor
                ?? firstSemanticChange?.LeftAnchor;
            var applicableRevisionIds = firstSemanticChange is null
                ? Array.Empty<string>()
                : ApplicableRevisionsForSemanticChange(
                        generated, firstSemanticChange, semantic.ModeledChanges)
                    .Select(revision => revision.Id)
                    .ToArray();
            findingBudget.Add(findings, Finding(
                "modeled_semantic_mismatch",
                VerificationFindingSeverity.Error,
                $"The {DirectionName(direction)} output has "
                    + $"{semantic.Comparison.ChangeCount ?? 0} modeled semantic change(s).",
                firstSemanticChange is null
                    ? null
                    : new ChangeLocation
                    {
                        EntryUri = firstSemanticChange.PartUri,
                        PropertyPath = firstSemanticChange.Path,
                    },
                firstSemanticAnchor,
                applicableRevisionIds,
                remediation: "Inspect the semantic change set and applicable generated revisions.",
                direction: direction));
        }

        if (!normalizedEquivalent)
        {
            findingBudget.Add(findings, Finding(
                "normalized_whole_package_mismatch",
                VerificationFindingSeverity.Error,
                $"The {DirectionName(direction)} output differs from its expected document after package normalization.",
                firstNormalizedDivergence is null
                    ? null
                    : new ChangeLocation { EntryUri = firstNormalizedDivergence.PartUri },
                firstNormalizedDivergence?.AnchorId,
                firstNormalizedDivergence?.ApplicableRevisionIds ?? requestedIds,
                "Inspect the first divergent part and the listed generated revisions.",
                direction));
        }
        if (!opcEquivalent)
        {
            findingBudget.Add(findings, Finding(
                "ordered_opc_content_mismatch",
                normalizedEquivalent ? VerificationFindingSeverity.Info : VerificationFindingSeverity.Error,
                "Exact uncompressed OPC entry bytes differ from the expected document.",
                firstDivergence is null ? null : new ChangeLocation { EntryUri = firstDivergence.PartUri },
                firstDivergence?.AnchorId,
                firstDivergence?.ApplicableRevisionIds ?? requestedIds,
                normalizedEquivalent
                    ? "No action is required when only documented XML serialization differs."
                    : "Inspect the divergent package entries.",
                direction));
        }
        if (!rawEquivalent)
        {
            findingBudget.Add(findings, Finding(
                "raw_package_bytes_mismatch",
                options.RequireExactPackageBytes
                    ? VerificationFindingSeverity.Error
                    : VerificationFindingSeverity.Info,
                "ZIP package bytes differ from the expected document.",
                revisionIds: Array.Empty<string>(),
                remediation: options.RequireExactPackageBytes
                    ? "Reproduce the exact expected ZIP container or disable the explicit exact-byte policy."
                    : "No action is required when normalized package identity is equal.",
                direction: direction));
        }

        bool equivalent = completed
            && preExistingPreserved
            && semantic.Comparison.Available
            && semantic.Comparison.Equivalent == true
            && normalizedEquivalent
            && divergenceEvaluation.Completed
            && (!options.RequireExactPackageBytes || rawEquivalent)
            && findings.All(item => item.Severity != VerificationFindingSeverity.Error);
        return new PathEvaluation(
            new RedlineProofPathResult
            {
                Direction = direction,
                Completed = completed,
                Equivalent = equivalent,
                RequestedRevisionIds = requestedIds,
                ResolvedRevisionIds = resolvedIds,
                ImplicitlyResolvedRevisionIds = implicitlyResolvedIds,
                SurvivingPreExistingRevisions = survivingPreExisting,
                PreExistingRevisionsPreserved = preExistingPreserved,
                ModeledSemantic = semantic.Comparison,
                NormalizedWholePackageEquivalent = normalizedEquivalent,
                OrderedOpcContentEquivalent = opcEquivalent,
                ExactPackageBytesEquivalent = rawEquivalent,
                DivergenceAnalysisCompleted = divergenceEvaluation.Completed,
                ExpectedPackage = ToPackageIdentity(expectedManifest),
                ActualPackage = ToPackageIdentity(actualManifest),
                FirstDivergence = firstNormalizedDivergence ?? firstDivergence,
                Divergences = divergences,
                Findings = findings,
            },
            outputBytes);
    }

    private static IReadOnlyList<RedlineRevisionClassification> Classify(
        IReadOnlyList<RevisionListEntry> baselineEntries,
        IReadOnlyList<RevisionListEntry> finalEntries,
        IReadOnlyList<RevisionListEntry> redlineEntries,
        List<RedlineProofFinding> findings,
        ProofFindingBudget findingBudget)
    {
        var baseline = baselineEntries.Select(ToIdentity)
            .OrderBy(RevisionSortKey, StringComparer.Ordinal).ToArray();
        var intendedFinal = finalEntries.Select(ToIdentity)
            .OrderBy(RevisionSortKey, StringComparer.Ordinal).ToArray();
        var redline = redlineEntries.Select(ToIdentity)
            .OrderBy(RevisionSortKey, StringComparer.Ordinal).ToArray();
        var baselineOwnership = IndexRevisionOwnership(baseline);
        var finalOwnership = IndexRevisionOwnership(intendedFinal);
        var redlineOwnership = IndexRevisionOwnership(redline);

        AppendDuplicateOwnershipFindings(
            findings, findingBudget, "baseline", baselineOwnership);
        AppendDuplicateOwnershipFindings(
            findings, findingBudget, "intendedFinal", finalOwnership);
        AppendDuplicateOwnershipFindings(
            findings, findingBudget, "redline", redlineOwnership);

        var results = new List<RedlineRevisionClassification>();
        var matchedBaselineKeys = new HashSet<string>(StringComparer.Ordinal);
        foreach (var redlineRevision in redline)
        {
            var keys = RevisionOwnershipKeys(redlineRevision);
            var baselineMatches = OwnershipMatches(keys, baselineOwnership);
            if (baselineMatches.Count > 0)
            {
                foreach (var key in keys.Where(baselineOwnership.ContainsKey))
                    matchedBaselineKeys.Add(key);
                bool unchanged = RevisionOwnershipPreserved(
                    redlineRevision, baselineOwnership);
                results.Add(new RedlineRevisionClassification
                {
                    Disposition = unchanged
                        ? RedlineRevisionDisposition.PreExisting
                        : RedlineRevisionDisposition.Conflicted,
                    Baseline = baselineMatches[0],
                    Redline = redlineRevision,
                    Reason = unchanged
                        ? "Every part-qualified native constituent remains owned by baseline review markup."
                        : "The redline mixes, duplicates, or rewrites native constituents owned by baseline review markup.",
                });
                if (!unchanged)
                    findingBudget.Add(findings,
                        RevisionConflictFinding(baselineMatches[0], redlineRevision));
                continue;
            }

            var finalMatches = OwnershipMatches(keys, finalOwnership);
            if (finalMatches.Count > 0)
            {
                bool unchanged = RevisionOwnershipPreserved(
                    redlineRevision, finalOwnership);
                results.Add(new RedlineRevisionClassification
                {
                    Disposition = unchanged
                        ? RedlineRevisionDisposition.IntendedFinalPreExisting
                        : RedlineRevisionDisposition.Conflicted,
                    IntendedFinal = finalMatches[0],
                    Redline = redlineRevision,
                    Reason = unchanged
                        ? "The revision belongs to intended-final review state absent from the selected baseline."
                        : "The redline mixes, duplicates, or rewrites native constituents owned by intended-final review state.",
                });
                findingBudget.Add(findings, Finding(
                    unchanged
                        ? "intended_final_revision_not_in_baseline"
                        : "intended_final_revision_identity_conflict",
                    unchanged
                        ? VerificationFindingSeverity.Info
                        : VerificationFindingSeverity.Error,
                    unchanged
                        ? $"Intended-final revision '{finalMatches[0].Id}' is review state, not a generated comparison revision."
                        : $"Redline revision '{redlineRevision.Id}' conflicts with intended-final review state.",
                    new ChangeLocation { EntryUri = redlineRevision.PartUri },
                    redlineRevision.AnchorId,
                    new[] { finalMatches[0].Id, redlineRevision.Id }
                        .Distinct(StringComparer.Ordinal).ToArray(),
                    "Choose a baseline/final policy with compatible pre-existing review state before generating the redline."));
                continue;
            }

            results.Add(new RedlineRevisionClassification
            {
                Disposition = RedlineRevisionDisposition.Generated,
                Redline = redlineRevision,
                Reason = "Every part-qualified native constituent is absent from both input review inventories.",
            });
        }

        foreach (var missing in baseline.Where(item =>
                     RevisionOwnershipKeys(item).Any(key => !matchedBaselineKeys.Contains(key))))
        {
            results.Add(new RedlineRevisionClassification
            {
                Disposition = RedlineRevisionDisposition.Conflicted,
                Baseline = missing,
                Reason = "A pre-existing baseline revision is missing from the redline.",
            });
            findingBudget.Add(findings, Finding(
                "preexisting_revision_missing_from_redline",
                VerificationFindingSeverity.Error,
                $"Baseline revision '{missing.Id}' is absent from the redline.",
                new ChangeLocation { EntryUri = missing.PartUri },
                missing.AnchorId,
                new[] { missing.Id },
                "Regenerate the redline while preserving the selected baseline review state."));
        }

        foreach (var missing in intendedFinal.Where(item =>
                     !RevisionOwnershipPreserved(item, redlineOwnership)))
        {
            results.Add(new RedlineRevisionClassification
            {
                Disposition = RedlineRevisionDisposition.Conflicted,
                IntendedFinal = missing,
                Reason = "Pre-existing intended-final review state is missing or rewritten in the redline.",
            });
            findingBudget.Add(findings, Finding(
                "intended_final_revision_missing_from_redline",
                VerificationFindingSeverity.Error,
                $"Intended-final revision '{missing.Id}' is absent or changed in the redline.",
                new ChangeLocation { EntryUri = missing.PartUri },
                missing.AnchorId,
                new[] { missing.Id },
                "Regenerate the redline without losing intended-final review state."));
        }

        return results
            .OrderBy(item => item.Redline?.PartUri ?? item.Baseline?.PartUri
                ?? item.IntendedFinal?.PartUri, StringComparer.Ordinal)
            .ThenBy(item => item.Redline?.Id ?? item.Baseline?.Id
                ?? item.IntendedFinal?.Id, StringComparer.Ordinal)
            .ToArray();
    }

    private static bool VerifyPreExisting(
        RedlineProofDirection direction,
        IReadOnlyList<RedlineRevisionIdentity> expected,
        IReadOnlyList<RedlineRevisionIdentity> actual,
        List<RedlineProofFinding> findings,
        ProofFindingBudget findingBudget)
    {
        var actualOwnership = IndexRevisionOwnership(actual);
        bool preserved = true;
        foreach (var revision in expected)
        {
            if (!RevisionOwnershipPreserved(revision, actualOwnership))
            {
                preserved = false;
                findingBudget.Add(findings, Finding(
                    "preexisting_revision_not_preserved",
                    VerificationFindingSeverity.Error,
                    $"Pre-existing revision '{revision.Id}' did not survive the {DirectionName(direction)} path unchanged.",
                    new ChangeLocation { EntryUri = revision.PartUri },
                    revision.AnchorId,
                    new[] { revision.Id },
                    "Do not resolve or rewrite revision markup owned by the baseline.",
                    direction));
            }
        }
        return preserved;
    }

    private static PackageDivergenceEvaluation CompareEntries(
        PackageManifest expected,
        PackageManifest actual,
        IReadOnlyList<RedlineRevisionIdentity> generated,
        IReadOnlyList<SemanticChange> modeledChanges,
        int maximumChanges)
    {
        var packageDelta = PackageDelta.Compare(expected, actual, maximumChanges);
        if (!packageDelta.Complete)
            return new PackageDivergenceEvaluation(
                false, Array.Empty<RedlinePackageDivergence>());

        var expectedByKey = expected.Entries.ToDictionary(
            item => (item.Uri, item.Occurrence), item => item);
        var actualByKey = actual.Entries.ToDictionary(
            item => (item.Uri, item.Occurrence), item => item);
        var divergences = new List<RedlinePackageDivergence>();
        foreach (var change in packageDelta.Changes.Where(change =>
                     change.Kind is PackageDeltaChangeKind.EntryAdded
                         or PackageDeltaChangeKind.EntryRemoved
                         or PackageDeltaChangeKind.EntryModified))
        {
            var partUri = change.Location.EntryUri
                ?? throw new InvalidOperationException(
                    "A package entry delta must identify its entry URI.");
            var key = (Uri: partUri, change.Occurrence);
            expectedByKey.TryGetValue(key, out var expectedEntry);
            actualByKey.TryGetValue(key, out var actualEntry);
            var kind = change.Kind switch
            {
                PackageDeltaChangeKind.EntryAdded => RedlinePackageDivergenceKind.Added,
                PackageDeltaChangeKind.EntryRemoved => RedlinePackageDivergenceKind.Removed,
                PackageDeltaChangeKind.EntryModified => RedlinePackageDivergenceKind.Modified,
                _ => throw new ArgumentOutOfRangeException(nameof(change), change.Kind, null),
            };

            var relevantModeledChanges = modeledChanges.Where(item =>
                    ModeledPartUris(item).Contains(partUri, StringComparer.Ordinal))
                .ToArray();
            var relevant = relevantModeledChanges
                .SelectMany(change => ApplicableRevisionsForSemanticChange(
                    generated, change, relevantModeledChanges))
                .Distinct()
                .OrderBy(RevisionSortKey, StringComparer.Ordinal).ToArray();
            bool normalizedDifferent = expectedEntry is null || actualEntry is null
                || !string.Equals(expectedEntry.ContentType, actualEntry.ContentType, StringComparison.Ordinal)
                || !DigestEquals(expectedEntry.NormalizedXmlDigest, actualEntry.NormalizedXmlDigest)
                || (!expectedEntry.IsXml && !DigestEquals(
                    expectedEntry.RawBytesDigest, actualEntry.RawBytesDigest));
            divergences.Add(new RedlinePackageDivergence
            {
                Kind = kind,
                PartUri = partUri,
                Occurrence = change.Occurrence,
                AnchorId = relevantModeledChanges.Select(item =>
                        item.RightAnchor ?? item.LeftAnchor)
                    .FirstOrDefault(item => item is not null),
                ApplicableRevisionIds = relevant.Select(item => item.Id).ToArray(),
                ExpectedRawDigest = expectedEntry?.RawBytesDigest,
                ActualRawDigest = actualEntry?.RawBytesDigest,
                ExpectedNormalizedDigest = expectedEntry?.NormalizedXmlDigest,
                ActualNormalizedDigest = actualEntry?.NormalizedXmlDigest,
                HasModeledSemanticChange = relevantModeledChanges.Length > 0,
                // A change set identifies modeled facts, not a complete residual projection for
                // an arbitrary XML part. Even when that same part has a modeled change, retain the
                // normalized divergence as potentially unmodeled instead of overclaiming coverage.
                UnknownOrUnmodeled = normalizedDifferent,
            });
        }
        return new PackageDivergenceEvaluation(true, divergences
            .OrderBy(item => item.PartUri, StringComparer.Ordinal)
            .ThenBy(item => item.Occurrence)
            .ToArray());
    }

    private static bool RevisionAppliesToPart(RedlineRevisionIdentity revision, string partUri)
    {
        if (string.Equals(revision.PartUri, partUri, StringComparison.Ordinal))
            return true;
        var slash = revision.PartUri.LastIndexOf('/');
        var ownerDirectory = slash < 0 ? "/" : revision.PartUri[..(slash + 1)];
        var ownerName = slash < 0 ? revision.PartUri : revision.PartUri[(slash + 1)..];
        return string.Equals(
            partUri,
            $"{ownerDirectory}_rels/{ownerName}.rels",
            StringComparison.Ordinal);
    }

    private static IReadOnlyList<RedlineRevisionIdentity> ApplicableRevisionsForSemanticChange(
        IReadOnlyList<RedlineRevisionIdentity> generated,
        SemanticChange change,
        IReadOnlyList<SemanticChange> allChanges)
    {
        var candidates = generated.Where(revision =>
                RevisionAppliesToPart(revision, change.PartUri))
            .ToArray();
        var changeAnchor = change.RightAnchor ?? change.LeftAnchor;
        var direct = changeAnchor is null
            ? Array.Empty<RedlineRevisionIdentity>()
            : candidates.Where(revision => RevisionAnchorIds(revision)
                    .Contains(changeAnchor, StringComparer.Ordinal))
                .ToArray();
        if (direct.Length > 0) return direct;

        // Resolved-output anchors are content-derived and therefore need not equal the redline's
        // tracked paragraph anchor. Use the semantic before/after text as a conservative seed,
        // then include the sibling del/ins carriers at that same redline location. If text points
        // at more than one location, make no attribution rather than overclaiming causality.
        var relatedChanges = allChanges.Where(candidate =>
                string.Equals(candidate.PartUri, change.PartUri, StringComparison.Ordinal)
                && SemanticChangesShareLocation(change, candidate))
            .ToArray();
        var semanticStrings = relatedChanges.SelectMany(candidate =>
                SemanticStrings(candidate.After))
            .Select(value => value.Trim())
            .Where(value => value.Length >= 2 && value.Any(char.IsLetterOrDigit))
            .Distinct(StringComparer.Ordinal).ToArray();
        var scored = candidates.Select(revision => new
            {
                Revision = revision,
                Score = semanticStrings.Where(value => revision.Text.Contains(
                        value, StringComparison.Ordinal))
                    .Select(value => value.Length).DefaultIfEmpty().Max(),
            })
            .Where(item => item.Score > 0).ToArray();
        if (scored.Length == 0)
        {
            semanticStrings = relatedChanges.SelectMany(candidate =>
                    SemanticStrings(candidate.Before))
                .Select(value => value.Trim())
                .Where(value => value.Length >= 2 && value.Any(char.IsLetterOrDigit))
                .Distinct(StringComparer.Ordinal).ToArray();
            scored = candidates.Select(revision => new
                {
                    Revision = revision,
                    Score = semanticStrings.Where(value => revision.Text.Contains(
                            value, StringComparison.Ordinal))
                        .Select(value => value.Length).DefaultIfEmpty().Max(),
                })
                .Where(item => item.Score > 0).ToArray();
        }
        int bestScore = scored.Select(item => item.Score).DefaultIfEmpty().Max();
        var seeds = scored.Where(item => item.Score == bestScore)
            .Select(item => item.Revision).ToArray();
        var seedLocations = seeds.Select(RevisionLocationKey)
            .Where(key => key is not null).Distinct(StringComparer.Ordinal).ToArray();
        if (seedLocations.Length != 1)
            return candidates.Length == 1 ? candidates : Array.Empty<RedlineRevisionIdentity>();

        var location = seedLocations[0]!;
        return candidates.Where(revision =>
                string.Equals(RevisionLocationKey(revision), location, StringComparison.Ordinal))
            .OrderBy(RevisionSortKey, StringComparer.Ordinal).ToArray();
    }

    private static bool SemanticChangesShareLocation(
        SemanticChange left, SemanticChange right)
    {
        var leftAnchors = new[] { left.RightAnchor, left.LeftAnchor }
            .Where(anchor => anchor is not null).Select(anchor => anchor!).ToArray();
        return leftAnchors.Length > 0
            && new[] { right.RightAnchor, right.LeftAnchor }
                .Where(anchor => anchor is not null).Select(anchor => anchor!)
                .Intersect(leftAnchors, StringComparer.Ordinal).Any();
    }

    private static IEnumerable<string> RevisionAnchorIds(RedlineRevisionIdentity revision) =>
        revision.AnchorId is null
            ? revision.AffectedAnchorIds
            : new[] { revision.AnchorId }.Concat(revision.AffectedAnchorIds);

    private static string? RevisionLocationKey(RedlineRevisionIdentity revision)
    {
        var anchors = RevisionAnchorIds(revision).Distinct(StringComparer.Ordinal)
            .OrderBy(value => value, StringComparer.Ordinal).ToArray();
        return anchors.Length == 0
            ? null
            : revision.PartUri + "\n" + string.Join("\n", anchors);
    }

    private static IEnumerable<string> SemanticStrings(SemanticValue root)
    {
        var pending = new Stack<SemanticValue>();
        pending.Push(root);
        int visited = 0;
        while (pending.Count > 0 && visited++ < 10_000)
        {
            var value = pending.Pop();
            if (value.Kind == SemanticValueKind.String && value.StringValue is { } text)
                yield return text;
            foreach (var item in value.Items) pending.Push(item);
            foreach (var property in value.Properties) pending.Push(property.Value);
        }
    }

    private static RedlineRevisionIdentity ToIdentity(RevisionListEntry entry) => new()
    {
        Id = entry.Id,
        PartUri = entry.PartUri,
        Scope = entry.Scope,
        Type = entry.Type,
        Family = entry.Family,
        ConstituentIds = entry.ConstituentIds.OrderBy(item => item, StringComparer.Ordinal).ToArray(),
        ConstituentKeys = entry.ConstituentKeys.OrderBy(item => item, StringComparer.Ordinal).ToArray(),
        Author = entry.Author,
        Date = entry.Date,
        DateUtc = entry.DateUtc,
        Text = entry.Text,
        AnchorId = entry.AnchorId,
        AffectedAnchorIds = entry.AffectedAnchors.Select(item => item.Id)
            .Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToArray(),
        ResolutionStatus = entry.ResolutionStatus,
        Diagnostic = entry.Diagnostic,
    };

    private static bool OwnershipEquivalent(
        RedlineRevisionIdentity left,
        RedlineRevisionIdentity right) =>
        string.Equals(left.PartUri, right.PartUri, StringComparison.Ordinal)
        && string.Equals(left.Type, right.Type, StringComparison.Ordinal)
        && string.Equals(left.Author, right.Author, StringComparison.Ordinal)
        && string.Equals(left.Date, right.Date, StringComparison.Ordinal)
        && string.Equals(left.DateUtc, right.DateUtc, StringComparison.Ordinal)
        && left.ResolutionStatus == right.ResolutionStatus;

    private static bool RevisionOwnershipPreserved(
        RedlineRevisionIdentity expected,
        IReadOnlyDictionary<string, List<RedlineRevisionIdentity>> actualOwnership)
    {
        var keys = RevisionOwnershipKeys(expected);
        if (keys.Any(key => !actualOwnership.TryGetValue(key, out var candidates)
                || candidates.Count != 1
                || !OwnershipEquivalent(expected, candidates[0])))
            return false;

        if (expected.Family is not (RevisionFamily.Move
            or RevisionFamily.ContentControlInsert
            or RevisionFamily.ContentControlDelete))
            return true;

        var actualGroups = keys.SelectMany(key => actualOwnership[key]).Distinct().ToArray();
        return actualGroups.Length == 1
            && actualGroups[0].Family == expected.Family
            && RevisionOwnershipKeys(actualGroups[0]).ToHashSet(StringComparer.Ordinal)
                .SetEquals(keys);
    }

    private static IReadOnlyList<string> RevisionOwnershipKeys(
        RedlineRevisionIdentity revision) => revision.ConstituentKeys.Count == 0
        ? new[] { revision.PartUri + "\ngroup:" + revision.Id }
        : revision.ConstituentKeys.Select(key => revision.PartUri + "\nnative:" + key)
            .Distinct(StringComparer.Ordinal)
            .OrderBy(key => key, StringComparer.Ordinal)
            .ToArray();

    private static Dictionary<string, List<RedlineRevisionIdentity>> IndexRevisionOwnership(
        IEnumerable<RedlineRevisionIdentity> revisions)
    {
        var result = new Dictionary<string, List<RedlineRevisionIdentity>>(StringComparer.Ordinal);
        foreach (var revision in revisions)
        {
            foreach (var key in RevisionOwnershipKeys(revision))
            {
                if (!result.TryGetValue(key, out var owners))
                    result[key] = owners = new List<RedlineRevisionIdentity>();
                owners.Add(revision);
            }
        }
        return result;
    }

    private static IReadOnlyList<RedlineRevisionIdentity> OwnershipMatches(
        IReadOnlyList<string> keys,
        IReadOnlyDictionary<string, List<RedlineRevisionIdentity>> ownership) =>
        keys.Where(ownership.ContainsKey)
            .SelectMany(key => ownership[key])
            .Distinct()
            .OrderBy(RevisionSortKey, StringComparer.Ordinal)
            .ToArray();

    private static void AppendDuplicateOwnershipFindings(
        List<RedlineProofFinding> findings,
        ProofFindingBudget findingBudget,
        string inputName,
        IReadOnlyDictionary<string, List<RedlineRevisionIdentity>> ownership)
    {
        foreach (var duplicate in ownership.Where(item => item.Value.Count != 1))
        {
            var identities = duplicate.Value.OrderBy(RevisionSortKey, StringComparer.Ordinal).ToArray();
            findingBudget.Add(findings, Finding(
                "duplicate_native_revision_ownership",
                VerificationFindingSeverity.Error,
                $"The {inputName} inventory assigns one native constituent to multiple revision groups.",
                new ChangeLocation { EntryUri = identities[0].PartUri },
                identities[0].AnchorId,
                identities.Select(item => item.Id).Distinct(StringComparer.Ordinal).ToArray(),
                "Repair ambiguous native revision ownership before proving reversibility."));
        }
    }

    private static bool RevisionOverlaps(
        RedlineRevisionIdentity baseline,
        RedlineRevisionIdentity redline) =>
        string.Equals(baseline.PartUri, redline.PartUri, StringComparison.Ordinal)
        && baseline.ConstituentKeys.Intersect(
            redline.ConstituentKeys, StringComparer.Ordinal).Any();

    private static string RevisionSortKey(RedlineRevisionIdentity item) =>
        item.PartUri + "\n" + item.Family + "\n" + item.Id;

    private static RedlineProofFinding RevisionConflictFinding(
        RedlineRevisionIdentity baseline,
        RedlineRevisionIdentity redline) => Finding(
        "preexisting_revision_identity_conflict",
        VerificationFindingSeverity.Error,
        $"Redline revision '{redline.Id}' conflicts with baseline revision '{baseline.Id}'.",
        new ChangeLocation { EntryUri = redline.PartUri },
        redline.AnchorId,
        new[] { baseline.Id, redline.Id }.Distinct(StringComparer.Ordinal).ToArray(),
        "Regenerate the redline without reusing or rewriting baseline revision identities.");

    private static void AppendManifestErrors(
        List<RedlineProofFinding> target,
        ProofFindingBudget findingBudget,
        string inputName,
        PackageManifest manifest,
        RedlineProofDirection? direction = null)
    {
        foreach (var finding in manifest.Findings.Where(item =>
                     item.Severity == VerificationFindingSeverity.Error))
        {
            findingBudget.Add(target, Finding(
                "package_" + finding.Code,
                finding.Severity,
                $"{inputName}: {finding.Message}",
                finding.Location,
                revisionIds: Array.Empty<string>(),
                remediation: "Repair or replace the unsafe package before running the proof.",
                direction: direction));
        }
    }

    private static bool AppendRevisionLimitFinding(
        List<RedlineProofFinding> target,
        ProofFindingBudget findingBudget,
        string inputName,
        PackageManifest manifest,
        int maximum)
    {
        long actual = manifest.Facts.NativeRevisionCarrierCount;
        if (actual <= maximum)
            return true;

        findingBudget.Add(target, Finding(
            "revision_element_limit_exceeded",
            VerificationFindingSeverity.Error,
            $"The {inputName} package contains "
                + $"{actual.ToString(System.Globalization.CultureInfo.InvariantCulture)} native revision "
                + $"element(s); the proof limit is "
                + $"{maximum.ToString(System.Globalization.CultureInfo.InvariantCulture)}.",
            new ChangeLocation { PropertyPath = inputName + "/revisions" },
            remediation: "Reduce tracked-change markup or explicitly raise MaxRevisionElements for a reviewed input."));
        return false;
    }

    private static bool AppendStrictRevisionFinding(
        List<RedlineProofFinding> target,
        ProofFindingBudget findingBudget,
        string inputName,
        PackageManifest manifest)
    {
        if (!manifest.Facts.IsStrictOoxml || !manifest.Facts.HasStrictRevisionMarkup)
            return true;

        findingBudget.Add(target, Finding(
            "unsupported_strict_revision_markup",
            VerificationFindingSeverity.Error,
            $"The {inputName} package contains strict-namespace tracked revision markup, "
                + "which the selective native resolver does not support.",
            new ChangeLocation { PropertyPath = inputName + "/revisions" },
            remediation: "Convert the document to transitional WordprocessingML before proving reversibility."));
        return false;
    }

    private static bool AppendLiveRevisionLimitFinding(
        List<RedlineProofFinding> target,
        ProofFindingBudget findingBudget,
        string inputName,
        int actual,
        int maximum,
        RedlineProofDirection? direction = null)
    {
        if (actual <= maximum)
            return true;

        findingBudget.Add(target, Finding(
            "revision_inventory_limit_exceeded",
            VerificationFindingSeverity.Error,
            $"The {inputName} live revision inventory contains "
                + $"{actual.ToString(System.Globalization.CultureInfo.InvariantCulture)} entries; "
                + $"the proof limit is "
                + $"{maximum.ToString(System.Globalization.CultureInfo.InvariantCulture)}.",
            new ChangeLocation { PropertyPath = inputName + "/revisionInventory" },
            remediation: "Reduce tracked-change markup or explicitly raise MaxRevisionElements for a reviewed input.",
            direction: direction));
        return false;
    }

    private static bool AppendNativeRevisionLimitFinding(
        List<RedlineProofFinding> target,
        ProofFindingBudget findingBudget,
        string inputName,
        long actual,
        int maximum,
        RedlineProofDirection? direction = null)
    {
        if (actual <= maximum)
            return true;

        findingBudget.Add(target, Finding(
            "revision_element_limit_exceeded",
            VerificationFindingSeverity.Error,
            $"The {inputName} live inventory contains "
                + $"{actual.ToString(System.Globalization.CultureInfo.InvariantCulture)} native revision "
                + $"element(s); the proof limit is "
                + $"{maximum.ToString(System.Globalization.CultureInfo.InvariantCulture)}.",
            new ChangeLocation { PropertyPath = inputName + "/revisions" },
            remediation: "Reduce tracked-change markup or explicitly raise MaxRevisionElements for a reviewed input.",
            direction: direction));
        return false;
    }

    private static bool AppendRevisionInventoryCoverageFinding(
        List<RedlineProofFinding> target,
        ProofFindingBudget findingBudget,
        string inputName,
        PackageManifest manifest,
        long inventoried)
    {
        var physical = manifest.Facts.NativeRevisionCarrierCount;
        if (physical == inventoried)
            return true;

        findingBudget.Add(target, Finding(
            "unsupported_revision_part",
            VerificationFindingSeverity.Error,
            $"The {inputName} package contains {physical} physical revision carrier(s), "
                + $"but the native registry inventoried {inventoried}; at least one carrier "
                + "is in an unmodeled Wordprocessing part.",
            new ChangeLocation { PropertyPath = inputName + "/revisionInventory" },
            remediation: "Remove tracked markup from unsupported parts or extend the native registry before proving reversibility."));
        return false;
    }

    private static bool AppendRevisionEvidenceLimitFinding(
        List<RedlineProofFinding> target,
        ProofFindingBudget findingBudget,
        IReadOnlyList<RevisionListEntry> baseline,
        IReadOnlyList<RevisionListEntry> redline,
        RedlineReversibilityProofOptions options,
        RedlineProofDirection? direction = null)
    {
        long textCharacters = 0;
        int evidenceItems = 0;
        foreach (var revision in baseline.Concat(redline))
        {
            if (!TryAddCount(ref evidenceItems, revision.ConstituentIds.Count,
                    options.MaxRevisionEvidenceItems)
                || !TryAddCount(ref evidenceItems, revision.ConstituentKeys.Count,
                    options.MaxRevisionEvidenceItems)
                || !TryAddCount(ref evidenceItems, revision.AffectedAnchors.Count,
                    options.MaxRevisionEvidenceItems)
                || !TryAddText(ref textCharacters, options.MaxEvidenceTextCharacters,
                    revision.Id, revision.PartUri, revision.Scope, revision.Type, revision.Author,
                    revision.Date, revision.DateUtc, revision.Text, revision.AnchorId,
                    revision.Diagnostic?.Code, revision.Diagnostic?.Message)
                || revision.ConstituentIds.Any(item =>
                    !TryAddText(ref textCharacters, options.MaxEvidenceTextCharacters, item))
                || revision.ConstituentKeys.Any(item =>
                    !TryAddText(ref textCharacters, options.MaxEvidenceTextCharacters, item))
                || revision.AffectedAnchors.Any(item =>
                    !TryAddText(ref textCharacters, options.MaxEvidenceTextCharacters, item.Id)))
            {
                findingBudget.Add(target, Finding(
                    "revision_evidence_limit_exceeded",
                    VerificationFindingSeverity.Error,
                    "The revision evidence exceeds the configured output budget.",
                    new ChangeLocation { PropertyPath = "revisionClassifications" },
                    remediation: "Reduce tracked-change evidence or deliberately raise the revision evidence limits.",
                    direction: direction));
                return false;
            }
        }
        return true;
    }

    private static bool TryAddCount(ref int current, int value, int maximum)
    {
        if (value < 0 || current > maximum - value)
            return false;
        current += value;
        return true;
    }

    private static bool TryAddText(ref long current, long maximum, params string?[] values)
    {
        foreach (var value in values)
        {
            if (value is null) continue;
            if (current > maximum - value.Length)
                return false;
            current += value.Length;
        }
        return true;
    }

    internal static IReadOnlyList<RedlineRevisionIdentity> SelectSurvivingPreExisting(
        IReadOnlyList<RedlineRevisionIdentity> preExisting,
        IReadOnlyList<RedlineRevisionIdentity> surviving)
    {
        var preExistingOwnership = preExisting.SelectMany(RevisionOwnershipKeys)
            .ToHashSet(StringComparer.Ordinal);
        return surviving.Where(item => RevisionOwnershipKeys(item)
                .Any(preExistingOwnership.Contains))
            .OrderBy(RevisionSortKey, StringComparer.Ordinal)
            .ToArray();
    }

    private static RedlineProofFinding Finding(
        string code,
        VerificationFindingSeverity severity,
        string message,
        ChangeLocation? location = null,
        string? anchorId = null,
        IReadOnlyList<string>? revisionIds = null,
        string? remediation = null,
        RedlineProofDirection? direction = null) => new()
    {
        Code = code,
        Severity = severity,
        Message = message,
        Direction = direction,
        Location = location,
        AnchorId = anchorId,
        RevisionIds = revisionIds ?? Array.Empty<string>(),
        Remediation = remediation,
    };

    private static RedlineReversibilityProofRun BuildRun(
        RedlineReversibilityProofOptions options,
        PackageManifest baseline,
        PackageManifest intendedFinal,
        PackageManifest redline,
        IReadOnlyList<RedlineRevisionClassification> classifications,
        RedlineProofPathResult? accept,
        RedlineProofPathResult? reject,
        IReadOnlyList<RedlineProofFinding> findings,
        byte[]? acceptedBytes,
        byte[]? rejectedBytes)
    {
        bool success = accept?.Equivalent == true
            && reject?.Equivalent == true
            && findings.All(item => item.Severity != VerificationFindingSeverity.Error);
        return new RedlineReversibilityProofRun
        {
            Proof = new RedlineReversibilityProof
            {
                Success = success,
                RequireExactPackageBytes = options.RequireExactPackageBytes,
                BaselinePackage = ToPackageIdentity(baseline),
                IntendedFinalPackage = ToPackageIdentity(intendedFinal),
                RedlinePackage = ToPackageIdentity(redline),
                RevisionClassifications = classifications,
                AcceptToFinal = accept,
                RejectToBaseline = reject,
                Findings = findings,
            },
            AcceptedPackageBytes = acceptedBytes,
            RejectedPackageBytes = rejectedBytes,
        };
    }

    private static RedlineProofPackageIdentity ToPackageIdentity(PackageManifest manifest) => new()
    {
        RawPackageBytesDigest = manifest.RawPackageBytesDigest,
        OrderedOpcContentDigest = manifest.OrderedOpcContentDigest,
        NormalizedWholePackageDigest = manifest.NormalizedSemanticDigest,
    };

    private static bool DigestEquals(VerificationDigest? left, VerificationDigest? right) =>
        DeliveryReceiptValidation.DigestEquals(left, right);

    private static DocxSession OpenProofSession(
        byte[] bytes,
        bool proofSafeResolution = false,
        IReadOnlyList<int>? protectedNumberingIds = null,
        IReadOnlyList<int>? protectedAbstractNumberingIds = null,
        IReadOnlyCollection<string>? protectedRelationshipKeys = null,
        IReadOnlyCollection<string>? protectedEmptyContainerKeys = null) =>
        new(bytes, new DocxSessionSettings
    {
        UndoDepth = 1,
        UndoMemoryBudgetBytes = 128L * 1024 * 1024,
        PersistAnchorIds = false,
        EmitMarkdownPatch = false,
        CaptureInitialProjection = false,
        ProofSafeRevisionResolution = proofSafeResolution,
        ProtectedRevisionNumberingIds = protectedNumberingIds ?? Array.Empty<int>(),
        ProtectedRevisionAbstractNumberingIds = protectedAbstractNumberingIds
            ?? Array.Empty<int>(),
        ProtectedRevisionRelationshipKeys = protectedRelationshipKeys
            ?? Array.Empty<string>(),
        ProtectedRevisionEmptyContainerKeys = protectedEmptyContainerKeys
            ?? Array.Empty<string>(),
    });

    private static void ValidateOptions(RedlineReversibilityProofOptions options)
    {
        ArgumentNullException.ThrowIfNull(options.PackageManifestOptions);
        options.PackageManifestOptions.Validate();
        if (options.MaxPackageBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxPackageBytes));
        if (options.MaxRevisionElements <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxRevisionElements));
        if (options.MaxSemanticChanges <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxSemanticChanges));
        if (options.MaxPackageChanges <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxPackageChanges));
        if (options.MaxFindings <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxFindings));
        if (options.MaxRevisionEvidenceItems <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxRevisionEvidenceItems));
        if (options.MaxEvidenceTextCharacters <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.MaxEvidenceTextCharacters));
    }

    private static void ValidatePackageByteBudget(
        byte[] baselineBytes,
        byte[] intendedFinalBytes,
        byte[] redlineBytes,
        long maximumBytes)
    {
        long baselineLength = baselineBytes.LongLength;
        long finalLength = intendedFinalBytes.LongLength;
        long redlineLength = redlineBytes.LongLength;
        if (baselineLength > maximumBytes
            || finalLength > maximumBytes - baselineLength
            || redlineLength > maximumBytes - baselineLength - finalLength)
            throw new ArgumentException(
                "Aggregate baseline, intended-final, and redline bytes exceed the proof budget.",
                nameof(redlineBytes));
    }

    private static ModeledSemanticEvaluation CompareModeledSemantics(
        byte[] expectedBytes,
        byte[] actualBytes,
        PackageManifestOptions packageOptions,
        int maximumChanges)
    {
        try
        {
            var changes = SemanticDiff.CompareBounded(
                new WmlDocument("expected.docx", expectedBytes),
                new WmlDocument("actual.docx", actualBytes),
                new SemanticDiffOptions { PackageOptions = packageOptions },
                maximumChanges);
            var modeledChanges = changes.Changes
                .Where(change => change.Family != SemanticChangeFamily.OpaquePackagePart)
                .ToArray();
            return new ModeledSemanticEvaluation(
                new RedlineModeledSemanticComparison
                {
                    Available = true,
                    Equivalent = modeledChanges.Length == 0,
                    Schema = changes.Schema,
                    ChangeCount = modeledChanges.Length,
                    Diagnostic = null,
                },
                modeledChanges,
                modeledChanges.FirstOrDefault(change =>
                    change.RightAnchor is not null || change.LeftAnchor is not null)
                    ?? modeledChanges.FirstOrDefault());
        }
        catch (Exception ex) when (DeliverableExceptionBoundary.IsRecoverable(ex))
        {
            return new ModeledSemanticEvaluation(
                SemanticUnavailable(
                    $"The versioned semantic comparison failed ({ex.GetType().Name})."),
                Array.Empty<SemanticChange>(),
                null);
        }
    }

    private static RedlineModeledSemanticComparison SemanticUnavailable(
        string diagnostic = "The versioned semantic comparison did not run.") => new()
    {
        Available = false,
        Equivalent = null,
        Schema = null,
        ChangeCount = null,
        Diagnostic = diagnostic,
    };

    private static IEnumerable<string> ModeledPartUris(SemanticChange change)
    {
        // Opaque-part digests make a difference visible, but do not claim that Docxodus
        // understands its semantics. Keep those divergences explicitly marked unmodeled.
        if (change.Family == SemanticChangeFamily.OpaquePackagePart)
            yield break;

        yield return change.PartUri;
        if (change.Family != SemanticChangeFamily.Relationship)
            yield break;

        if (string.Equals(change.PartUri, "/", StringComparison.Ordinal))
        {
            yield return "/_rels/.rels";
            yield break;
        }

        int slash = change.PartUri.LastIndexOf('/');
        string directory = slash <= 0 ? "/" : change.PartUri[..(slash + 1)];
        string name = slash < 0 ? change.PartUri : change.PartUri[(slash + 1)..];
        yield return $"{directory}_rels/{name}.rels";
    }

    private static string DirectionName(RedlineProofDirection direction) => direction switch
    {
        RedlineProofDirection.AcceptToFinal => "accept-to-final",
        _ => "reject-to-baseline",
    };

    private sealed record PathEvaluation(RedlineProofPathResult Result, byte[]? OutputBytes);

    private sealed record ModeledSemanticEvaluation(
        RedlineModeledSemanticComparison Comparison,
        IReadOnlyList<SemanticChange> ModeledChanges,
        SemanticChange? FirstModeledChange);

    private sealed record PackageDivergenceEvaluation(
        bool Completed,
        IReadOnlyList<RedlinePackageDivergence> Divergences);

    private sealed class ProofFindingBudget
    {
        private readonly int _maximum;
        private int _retained;
        private bool _reportedExhaustion;

        internal ProofFindingBudget(int maximum) => _maximum = maximum;

        internal void Add(
            List<RedlineProofFinding> target,
            RedlineProofFinding finding)
        {
            if (_retained < _maximum)
            {
                target.Add(finding);
                _retained++;
                return;
            }
            if (_reportedExhaustion)
                return;

            _reportedExhaustion = true;
            target.Add(Finding(
                "proof_finding_limit_exceeded",
                VerificationFindingSeverity.Error,
                "The proof exceeded the configured finding budget; later findings were suppressed.",
                new ChangeLocation { PropertyPath = "findings" },
                revisionIds: Array.Empty<string>(),
                remediation: "Reduce proof complexity or deliberately raise MaxFindings.",
                direction: finding.Direction));
        }
    }
}
