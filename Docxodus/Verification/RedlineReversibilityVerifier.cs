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
        ArgumentNullException.ThrowIfNull(options.PackageManifestOptions);

        var baselineManifest = PackageManifestGenerator.Generate(
            baselineBytes, options.PackageManifestOptions);
        var finalManifest = PackageManifestGenerator.Generate(
            intendedFinalBytes, options.PackageManifestOptions);
        var redlineManifest = PackageManifestGenerator.Generate(
            redlineBytes, options.PackageManifestOptions);
        var sharedFindings = new List<RedlineProofFinding>();

        AppendManifestErrors(sharedFindings, "baseline", baselineManifest);
        AppendManifestErrors(sharedFindings, "intendedFinal", finalManifest);
        AppendManifestErrors(sharedFindings, "redline", redlineManifest);
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

        IReadOnlyList<RevisionListEntry> baselineRevisions;
        IReadOnlyList<RevisionListEntry> redlineRevisions;
        try
        {
            using var baseline = OpenProofSession(baselineBytes);
            baselineRevisions = baseline.ListRevisions();
            using var redline = OpenProofSession(redlineBytes);
            redlineRevisions = redline.ListRevisions();
        }
        catch (Exception ex) when (!IsFatal(ex))
        {
            sharedFindings.Add(Finding(
                "revision_inventory_failed",
                VerificationFindingSeverity.Error,
                $"The baseline or redline revision inventory could not be opened ({ex.GetType().Name}).",
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

        var classifications = Classify(baselineRevisions, redlineRevisions, sharedFindings);
        var generated = classifications
            .Where(item => item.Disposition == RedlineRevisionDisposition.Generated
                && item.Redline is not null)
            .Select(item => item.Redline!)
            .OrderBy(RevisionSortKey, StringComparer.Ordinal)
            .ToArray();
        var preExisting = classifications
            .Where(item => item.Disposition == RedlineRevisionDisposition.PreExisting
                && item.Baseline is not null)
            .Select(item => item.Baseline!)
            .OrderBy(RevisionSortKey, StringComparer.Ordinal)
            .ToArray();

        foreach (var revision in generated.Where(item =>
                     item.ResolutionStatus != RevisionResolutionStatus.Supported))
        {
            sharedFindings.Add(Finding(
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
            preExisting,
            options);
        var reject = EvaluatePath(
            RedlineProofDirection.RejectToBaseline,
            accept: false,
            redlineBytes,
            baselineBytes,
            baselineManifest,
            generated,
            preExisting,
            options);

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
        RedlineReversibilityProofOptions options)
    {
        var findings = new List<RedlineProofFinding>();
        var requestedIds = generated.Select(item => item.Id).ToArray();
        var resolvedIds = new List<string>(requestedIds.Length);
        var implicitlyResolvedIds = new List<string>();
        byte[]? outputBytes = null;
        IReadOnlyList<RevisionListEntry> survivingEntries = Array.Empty<RevisionListEntry>();
        bool completed = true;

        try
        {
            using var session = OpenProofSession(redlineBytes);
            foreach (var requested in generated)
            {
                var current = session.ListRevisions()
                    .SingleOrDefault(item => string.Equals(item.Id, requested.Id, StringComparison.Ordinal));
                if (current is null)
                {
                    var liveRevisions = session.ListRevisions().Select(ToIdentity).ToArray();
                    var overlapping = liveRevisions.FirstOrDefault(item =>
                        RevisionOverlaps(requested, item));
                    if (overlapping is not null)
                    {
                        completed = false;
                        findings.Add(Finding(
                            "generated_revision_identity_changed",
                            VerificationFindingSeverity.Error,
                            $"Generated revision '{requested.Id}' changed identity to '{overlapping.Id}' during resolution.",
                            new ChangeLocation { EntryUri = requested.PartUri },
                            requested.AnchorId,
                            new[] { requested.Id, overlapping.Id },
                            "Regenerate a redline whose generated revision identities remain stable during selective resolution.",
                            direction));
                        break;
                    }

                    implicitlyResolvedIds.Add(requested.Id);
                    findings.Add(Finding(
                        "generated_revision_consumed_by_resolution",
                        VerificationFindingSeverity.Info,
                        $"Generated revision '{requested.Id}' was consumed while resolving a linked generated revision.",
                        new ChangeLocation { EntryUri = requested.PartUri },
                        requested.AnchorId,
                        new[] { requested.Id },
                        "No action is required when the path output matches its target and no generated revisions survive.",
                        direction));
                    continue;
                }

                if (current.ResolutionStatus != RevisionResolutionStatus.Supported)
                {
                    completed = false;
                    findings.Add(Finding(
                        "generated_revision_became_unresolvable",
                        VerificationFindingSeverity.Error,
                        $"Generated revision '{requested.Id}' became {current.ResolutionStatus.ToString().ToLowerInvariant()} during resolution.",
                        new ChangeLocation { EntryUri = current.PartUri },
                        current.AnchorId,
                        new[] { requested.Id },
                        "Regenerate a redline whose generated revision groups can be resolved independently.",
                        direction));
                    break;
                }

                var edit = accept
                    ? session.AcceptRevision(requested.Id)
                    : session.RejectRevision(requested.Id);
                if (!edit.Success)
                {
                    completed = false;
                    findings.Add(Finding(
                        "generated_revision_resolution_failed",
                        VerificationFindingSeverity.Error,
                        $"Revision '{requested.Id}' could not be {(accept ? "accepted" : "rejected")}: " +
                        (edit.Error?.Message ?? "unknown resolver failure"),
                        new ChangeLocation { EntryUri = current.PartUri },
                        current.AnchorId,
                        new[] { requested.Id },
                        "Inspect the revision topology and regenerate the comparison redline.",
                        direction));
                    break;
                }
                resolvedIds.Add(requested.Id);
            }

            survivingEntries = session.ListRevisions();
            outputBytes = session.Save(persistAnchorIds: false);
        }
        catch (Exception ex) when (!IsFatal(ex))
        {
            completed = false;
            findings.Add(Finding(
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
        var survivingGenerated = generated.Where(requested => surviving.Any(item =>
                string.Equals(item.Id, requested.Id, StringComparison.Ordinal)
                || RevisionOverlaps(requested, item)))
            .ToArray();
        if (survivingGenerated.Length > 0)
        {
            completed = false;
            findings.Add(Finding(
                "generated_revisions_survived_resolution",
                VerificationFindingSeverity.Error,
                $"The {DirectionName(direction)} path left {survivingGenerated.Length} generated revision(s) unresolved.",
                revisionIds: survivingGenerated.Select(item => item.Id).ToArray(),
                remediation: "Inspect the generated revision topology and selective resolver results.",
                direction: direction));
        }
        if (resolvedIds.Count + implicitlyResolvedIds.Count != requestedIds.Length)
            completed = false;
        bool preExistingPreserved = VerifyPreExisting(
            direction, preExisting, surviving, findings);

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
                    SurvivingPreExistingRevisions = surviving,
                    PreExistingRevisionsPreserved = preExistingPreserved,
                    ModeledSemantic = SemanticUnavailable(),
                    NormalizedWholePackageEquivalent = false,
                    OrderedOpcContentEquivalent = false,
                    ExactPackageBytesEquivalent = false,
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
        AppendManifestErrors(findings, "proofOutput", actualManifest, direction);
        if (!actualManifest.IsValid)
            completed = false;

        var semantic = CompareModeledSemantics(expectedBytes, outputBytes);
        if (!semantic.Available)
        {
            findings.Add(Finding(
                "modeled_semantic_comparison_unavailable",
                VerificationFindingSeverity.Error,
                semantic.Diagnostic ?? "The versioned modeled semantic comparer is unavailable.",
                revisionIds: requestedIds,
                remediation: "Run the proof with the semantic-diff component available.",
                direction: direction));
        }
        else if (semantic.Equivalent != true)
        {
            findings.Add(Finding(
                "modeled_semantic_mismatch",
                VerificationFindingSeverity.Error,
                $"The {DirectionName(direction)} output has {semantic.ChangeCount ?? 0} modeled semantic change(s).",
                revisionIds: requestedIds,
                remediation: "Inspect the semantic change set and applicable generated revisions.",
                direction: direction));
        }

        bool normalizedEquivalent = DigestEquals(
            expectedManifest.NormalizedSemanticDigest,
            actualManifest.NormalizedSemanticDigest);
        bool opcEquivalent = DigestEquals(
            expectedManifest.OrderedOpcContentDigest,
            actualManifest.OrderedOpcContentDigest);
        bool rawEquivalent = DigestEquals(
            expectedManifest.RawPackageBytesDigest,
            actualManifest.RawPackageBytesDigest);
        var divergences = CompareEntries(
            expectedManifest,
            actualManifest,
            generated,
            modeledPartUris: Array.Empty<string>());
        var firstDivergence = divergences.FirstOrDefault();

        if (!normalizedEquivalent)
        {
            findings.Add(Finding(
                "normalized_whole_package_mismatch",
                VerificationFindingSeverity.Error,
                $"The {DirectionName(direction)} output differs from its expected document after package normalization.",
                firstDivergence is null ? null : new ChangeLocation { EntryUri = firstDivergence.PartUri },
                firstDivergence?.AnchorId,
                firstDivergence?.ApplicableRevisionIds ?? requestedIds,
                "Inspect the first divergent part and the listed generated revisions.",
                direction));
        }
        if (!opcEquivalent)
        {
            findings.Add(Finding(
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
            findings.Add(Finding(
                "raw_package_bytes_mismatch",
                options.RequireExactPackageBytes
                    ? VerificationFindingSeverity.Error
                    : VerificationFindingSeverity.Info,
                "ZIP package bytes differ from the expected document.",
                revisionIds: requestedIds,
                remediation: options.RequireExactPackageBytes
                    ? "Reproduce the exact expected ZIP container or disable the explicit exact-byte policy."
                    : "No action is required when normalized package identity is equal.",
                direction: direction));
        }

        bool equivalent = completed
            && preExistingPreserved
            && semantic.Available
            && semantic.Equivalent == true
            && normalizedEquivalent
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
                SurvivingPreExistingRevisions = surviving,
                PreExistingRevisionsPreserved = preExistingPreserved,
                ModeledSemantic = semantic,
                NormalizedWholePackageEquivalent = normalizedEquivalent,
                OrderedOpcContentEquivalent = opcEquivalent,
                ExactPackageBytesEquivalent = rawEquivalent,
                ExpectedPackage = ToPackageIdentity(expectedManifest),
                ActualPackage = ToPackageIdentity(actualManifest),
                FirstDivergence = firstDivergence,
                Divergences = divergences,
                Findings = findings,
            },
            outputBytes);
    }

    private static IReadOnlyList<RedlineRevisionClassification> Classify(
        IReadOnlyList<RevisionListEntry> baselineEntries,
        IReadOnlyList<RevisionListEntry> redlineEntries,
        List<RedlineProofFinding> findings)
    {
        var baseline = baselineEntries.Select(ToIdentity)
            .OrderBy(RevisionSortKey, StringComparer.Ordinal).ToArray();
        var redline = redlineEntries.Select(ToIdentity)
            .OrderBy(RevisionSortKey, StringComparer.Ordinal).ToArray();
        var baselineById = baseline.GroupBy(item => item.Id, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
        var redlineById = redline.GroupBy(item => item.Id, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);

        foreach (var duplicate in baselineById.Where(item => item.Value.Length != 1))
            findings.Add(DuplicateIdentityFinding("baseline", duplicate.Key, duplicate.Value));
        foreach (var duplicate in redlineById.Where(item => item.Value.Length != 1))
            findings.Add(DuplicateIdentityFinding("redline", duplicate.Key, duplicate.Value));

        var results = new List<RedlineRevisionClassification>();
        var matchedBaselineIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var redlineRevision in redline)
        {
            if (redlineById[redlineRevision.Id].Length != 1)
            {
                results.Add(new RedlineRevisionClassification
                {
                    Disposition = RedlineRevisionDisposition.Conflicted,
                    Redline = redlineRevision,
                    Reason = "The redline contains a duplicate stable revision identity.",
                });
                continue;
            }

            if (baselineById.TryGetValue(redlineRevision.Id, out var sameId)
                && sameId.Length == 1)
            {
                var baselineRevision = sameId[0];
                matchedBaselineIds.Add(baselineRevision.Id);
                bool unchanged = IdentityEquivalent(baselineRevision, redlineRevision);
                results.Add(new RedlineRevisionClassification
                {
                    Disposition = unchanged
                        ? RedlineRevisionDisposition.PreExisting
                        : RedlineRevisionDisposition.Conflicted,
                    Baseline = baselineRevision,
                    Redline = redlineRevision,
                    Reason = unchanged
                        ? "The complete part-qualified revision identity is unchanged from the baseline."
                        : "A baseline revision reused its stable ID but changed identity or location.",
                });
                if (!unchanged)
                    findings.Add(RevisionConflictFinding(baselineRevision, redlineRevision));
                continue;
            }

            var overlaps = baseline.Where(item => RevisionOverlaps(item, redlineRevision)).ToArray();
            if (overlaps.Length > 0)
            {
                foreach (var overlap in overlaps)
                    matchedBaselineIds.Add(overlap.Id);
                results.Add(new RedlineRevisionClassification
                {
                    Disposition = RedlineRevisionDisposition.Conflicted,
                    Baseline = overlaps[0],
                    Redline = redlineRevision,
                    Reason = "The redline revision overlaps native constituent IDs owned by baseline review markup.",
                });
                findings.Add(RevisionConflictFinding(overlaps[0], redlineRevision));
                continue;
            }

            results.Add(new RedlineRevisionClassification
            {
                Disposition = RedlineRevisionDisposition.Generated,
                Redline = redlineRevision,
                Reason = "The part-qualified native revision identity is absent from the baseline.",
            });
        }

        foreach (var missing in baseline.Where(item => !matchedBaselineIds.Contains(item.Id)))
        {
            results.Add(new RedlineRevisionClassification
            {
                Disposition = RedlineRevisionDisposition.Conflicted,
                Baseline = missing,
                Reason = "A pre-existing baseline revision is missing from the redline.",
            });
            findings.Add(Finding(
                "preexisting_revision_missing_from_redline",
                VerificationFindingSeverity.Error,
                $"Baseline revision '{missing.Id}' is absent from the redline.",
                new ChangeLocation { EntryUri = missing.PartUri },
                missing.AnchorId,
                new[] { missing.Id },
                "Regenerate the redline while preserving the selected baseline review state."));
        }

        return results
            .OrderBy(item => item.Redline?.PartUri ?? item.Baseline?.PartUri, StringComparer.Ordinal)
            .ThenBy(item => item.Redline?.Id ?? item.Baseline?.Id, StringComparer.Ordinal)
            .ToArray();
    }

    private static bool VerifyPreExisting(
        RedlineProofDirection direction,
        IReadOnlyList<RedlineRevisionIdentity> expected,
        IReadOnlyList<RedlineRevisionIdentity> actual,
        List<RedlineProofFinding> findings)
    {
        var actualById = actual.GroupBy(item => item.Id, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
        bool preserved = true;
        foreach (var revision in expected)
        {
            if (!actualById.TryGetValue(revision.Id, out var candidates)
                || candidates.Length != 1
                || !IdentityEquivalent(revision, candidates[0]))
            {
                preserved = false;
                findings.Add(Finding(
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

    private static IReadOnlyList<RedlinePackageDivergence> CompareEntries(
        PackageManifest expected,
        PackageManifest actual,
        IReadOnlyList<RedlineRevisionIdentity> generated,
        IReadOnlyCollection<string> modeledPartUris)
    {
        var expectedByKey = expected.Entries.ToDictionary(
            item => (item.Uri, item.Occurrence), item => item);
        var actualByKey = actual.Entries.ToDictionary(
            item => (item.Uri, item.Occurrence), item => item);
        var keys = expectedByKey.Keys.Concat(actualByKey.Keys).Distinct()
            .OrderBy(item => item.Uri, StringComparer.Ordinal)
            .ThenBy(item => item.Occurrence)
            .ToArray();
        var divergences = new List<RedlinePackageDivergence>();
        foreach (var key in keys)
        {
            expectedByKey.TryGetValue(key, out var expectedEntry);
            actualByKey.TryGetValue(key, out var actualEntry);
            var kind = expectedEntry is null
                ? RedlinePackageDivergenceKind.Added
                : actualEntry is null
                    ? RedlinePackageDivergenceKind.Removed
                    : RedlinePackageDivergenceKind.Modified;
            if (expectedEntry is not null && actualEntry is not null
                && string.Equals(expectedEntry.ContentType, actualEntry.ContentType, StringComparison.Ordinal)
                && DigestEquals(expectedEntry.RawBytesDigest, actualEntry.RawBytesDigest)
                && DigestEquals(expectedEntry.NormalizedXmlDigest, actualEntry.NormalizedXmlDigest))
                continue;

            var relevant = generated.Where(item => RevisionAppliesToPart(item, key.Uri))
                .OrderBy(RevisionSortKey, StringComparer.Ordinal).ToArray();
            bool normalizedDifferent = expectedEntry is null || actualEntry is null
                || !string.Equals(expectedEntry.ContentType, actualEntry.ContentType, StringComparison.Ordinal)
                || !DigestEquals(expectedEntry.NormalizedXmlDigest, actualEntry.NormalizedXmlDigest)
                || (!expectedEntry.IsXml && !DigestEquals(
                    expectedEntry.RawBytesDigest, actualEntry.RawBytesDigest));
            divergences.Add(new RedlinePackageDivergence
            {
                Kind = kind,
                PartUri = key.Uri,
                Occurrence = key.Occurrence,
                AnchorId = relevant.Select(item => item.AnchorId)
                    .FirstOrDefault(item => item is not null),
                ApplicableRevisionIds = relevant.Select(item => item.Id).ToArray(),
                ExpectedRawDigest = expectedEntry?.RawBytesDigest,
                ActualRawDigest = actualEntry?.RawBytesDigest,
                ExpectedNormalizedDigest = expectedEntry?.NormalizedXmlDigest,
                ActualNormalizedDigest = actualEntry?.NormalizedXmlDigest,
                UnknownOrUnmodeled = normalizedDifferent
                    && !modeledPartUris.Contains(key.Uri),
            });
        }
        return divergences;
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

    private static RedlineRevisionIdentity ToIdentity(RevisionListEntry entry) => new()
    {
        Id = entry.Id,
        PartUri = entry.PartUri,
        Scope = entry.Scope,
        Type = entry.Type,
        Family = entry.Family,
        ConstituentIds = entry.ConstituentIds.OrderBy(item => item, StringComparer.Ordinal).ToArray(),
        Author = entry.Author,
        Date = entry.Date,
        Text = entry.Text,
        AnchorId = entry.AnchorId,
        AffectedAnchorIds = entry.AffectedAnchors.Select(item => item.Id)
            .Distinct(StringComparer.Ordinal).OrderBy(item => item, StringComparer.Ordinal).ToArray(),
        ResolutionStatus = entry.ResolutionStatus,
        Diagnostic = entry.Diagnostic,
    };

    private static bool IdentityEquivalent(
        RedlineRevisionIdentity left,
        RedlineRevisionIdentity right) =>
        string.Equals(left.Id, right.Id, StringComparison.Ordinal)
        && string.Equals(left.PartUri, right.PartUri, StringComparison.Ordinal)
        && string.Equals(left.Scope, right.Scope, StringComparison.Ordinal)
        && string.Equals(left.Type, right.Type, StringComparison.Ordinal)
        && left.Family == right.Family
        && left.ConstituentIds.SequenceEqual(right.ConstituentIds, StringComparer.Ordinal)
        && string.Equals(left.Author, right.Author, StringComparison.Ordinal)
        && string.Equals(left.Date, right.Date, StringComparison.Ordinal)
        && string.Equals(left.Text, right.Text, StringComparison.Ordinal)
        && string.Equals(left.AnchorId, right.AnchorId, StringComparison.Ordinal)
        && left.AffectedAnchorIds.SequenceEqual(right.AffectedAnchorIds, StringComparer.Ordinal)
        && left.ResolutionStatus == right.ResolutionStatus
        && Equals(left.Diagnostic, right.Diagnostic);

    private static bool RevisionOverlaps(
        RedlineRevisionIdentity baseline,
        RedlineRevisionIdentity redline) =>
        string.Equals(baseline.PartUri, redline.PartUri, StringComparison.Ordinal)
        && baseline.ConstituentIds.Intersect(
            redline.ConstituentIds, StringComparer.Ordinal).Any();

    private static string RevisionSortKey(RedlineRevisionIdentity item) =>
        item.PartUri + "\n" + item.Family + "\n" + item.Id;

    private static RedlineProofFinding DuplicateIdentityFinding(
        string inputName,
        string id,
        IReadOnlyList<RedlineRevisionIdentity> identities) => Finding(
        "duplicate_revision_identity",
        VerificationFindingSeverity.Error,
        $"The {inputName} revision inventory contains duplicate stable ID '{id}'.",
        new ChangeLocation { EntryUri = identities[0].PartUri },
        identities[0].AnchorId,
        new[] { id },
        "Repair ambiguous native revision IDs before proving reversibility.");

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
        string inputName,
        PackageManifest manifest,
        RedlineProofDirection? direction = null)
    {
        foreach (var finding in manifest.Findings.Where(item =>
                     item.Severity == VerificationFindingSeverity.Error))
        {
            target.Add(Finding(
                "package_" + finding.Code,
                finding.Severity,
                $"{inputName}: {finding.Message}",
                finding.Location,
                revisionIds: Array.Empty<string>(),
                remediation: "Repair or replace the unsafe package before running the proof.",
                direction: direction));
        }
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
        left is not null && right is not null
        && string.Equals(left.Algorithm, right.Algorithm, StringComparison.Ordinal)
        && string.Equals(left.Value, right.Value, StringComparison.Ordinal);

    private static DocxSession OpenProofSession(byte[] bytes) => new(bytes, new DocxSessionSettings
    {
        UndoDepth = 1,
        UndoMemoryBudgetBytes = 128L * 1024 * 1024,
        PersistAnchorIds = false,
        EmitMarkdownPatch = false,
        CaptureInitialProjection = false,
    });

    // Replaced by the schema-v1 semantic comparer when #457 is stacked into this branch. Keeping
    // the dependency unavailable (and therefore proof-failing) is deliberate: normalized package
    // equality must never be mislabeled as modeled semantic evidence.
    private static RedlineModeledSemanticComparison CompareModeledSemantics(
        byte[] expectedBytes,
        byte[] actualBytes) => SemanticUnavailable();

    private static RedlineModeledSemanticComparison SemanticUnavailable() => new()
    {
        Available = false,
        Equivalent = null,
        Schema = null,
        ChangeCount = null,
        Diagnostic = "The versioned semantic diff dependency is not present on this branch.",
    };

    private static string DirectionName(RedlineProofDirection direction) => direction switch
    {
        RedlineProofDirection.AcceptToFinal => "accept-to-final",
        _ => "reject-to-baseline",
    };

    private static bool IsFatal(Exception exception) => exception is
        OutOfMemoryException or StackOverflowException or AccessViolationException;

    private sealed record PathEvaluation(RedlineProofPathResult Result, byte[]? OutputBytes);
}
