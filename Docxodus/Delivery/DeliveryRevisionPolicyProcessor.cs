// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using Docxodus.Verification;

namespace Docxodus.Delivery;

/// <summary>
/// The isolated document states produced by applying a delivery revision policy.
/// </summary>
public sealed record DeliveryRevisionPolicyResult
{
    /// <summary>The baseline after applying the pre-existing-revision policy.</summary>
    required public byte[] PolicyBaselineBytes { get; init; }

    /// <summary>The edited document after applying both policy dimensions.</summary>
    required public byte[] FinalBytes { get; init; }

    /// <summary>
    /// Native tracked-changes comparison from the policy baseline to the final document. Present
    /// only when review proof was requested and successfully established.
    /// </summary>
    public byte[]? ReviewBytes { get; init; }

    /// <summary>The successful accept-to-final/reject-to-baseline proof for <see cref="ReviewBytes"/>.</summary>
    public RedlineReversibilityProofRun? ReviewProof { get; init; }

    /// <summary>
    /// Full-identity classification of the revisions supplied in the edited input relative to the
    /// original baseline. These identities are captured before either isolated clone is resolved.
    /// </summary>
    public IReadOnlyList<RedlineRevisionClassification> RevisionClassifications { get; init; } =
        Array.Empty<RedlineRevisionClassification>();
}

/// <summary>
/// Applies independent pre-existing and generated revision policies without mutating either input.
/// Revision ownership is determined by the same full, part-qualified native identity comparison
/// used by <see cref="RedlineReversibilityVerifier"/>.
/// </summary>
public static class DeliveryRevisionPolicyProcessor
{
    private const int MaxRevisionElements = 1_000;

    /// <summary>
    /// Produce an isolated policy baseline and final package. When <paramref name="requireReviewProof"/>
    /// is true, the generated-revision policy must be <see cref="DeliveryRevisionPolicy.Accept"/>;
    /// a native review document is generated and returned only if both selective proof paths succeed.
    /// </summary>
    public static DeliveryRevisionPolicyResult Apply(
        byte[] baselineBytes,
        byte[] editedBytes,
        DeliveryBundleRevisionPolicy policy,
        bool requireReviewProof,
        PackageManifestOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(baselineBytes);
        ArgumentNullException.ThrowIfNull(editedBytes);
        ArgumentNullException.ThrowIfNull(policy);
        options ??= new PackageManifestOptions();

        ValidateAction(policy.PreExistingRevisions, nameof(policy.PreExistingRevisions));
        ValidateAction(policy.GeneratedRevisions, nameof(policy.GeneratedRevisions));
        if (requireReviewProof && policy.GeneratedRevisions != DeliveryRevisionPolicy.Accept)
        {
            throw new ArgumentException(
                "Review proof requires the generated-revision policy to be Accept.",
                nameof(policy));
        }

        // Copy before any package consumer runs. Although the current consumers are read-only at
        // their byte-array boundary, this makes input isolation part of this API's own contract.
        var baselineInput = (byte[])baselineBytes.Clone();
        var editedInput = (byte[])editedBytes.Clone();
        Preflight("baseline", baselineInput, options);
        Preflight("edited", editedInput, options);

        IReadOnlyList<RevisionListEntry> baselineEntries;
        IReadOnlyList<RevisionListEntry> editedEntries;
        using (var baseline = OpenPolicySession(baselineInput))
            baselineEntries = baseline.ListRevisions();
        using (var edited = OpenPolicySession(editedInput))
            editedEntries = edited.ListRevisions();

        EnsureInventoryWithinLimit("baseline", baselineEntries.Count);
        EnsureInventoryWithinLimit("edited", editedEntries.Count);

        var classificationFindings = new List<RedlineProofFinding>();
        var classifications = RedlineReversibilityVerifier.Classify(
            baselineEntries, editedEntries, classificationFindings);
        if (classifications.Any(item => item.Disposition == RedlineRevisionDisposition.Conflicted))
        {
            var finding = classificationFindings.FirstOrDefault(item =>
                item.Severity == VerificationFindingSeverity.Error);
            throw new InvalidDataException(
                "The edited document does not preserve an unambiguous baseline revision identity."
                + (finding is null ? string.Empty : " " + finding.Message));
        }

        var preExistingBaseline = classifications
            .Where(item => item.Disposition == RedlineRevisionDisposition.PreExisting
                && item.Baseline is not null)
            .Select(item => item.Baseline!)
            .OrderBy(RedlineReversibilityVerifier.RevisionSortKey, StringComparer.Ordinal)
            .ToArray();
        var preExistingEdited = classifications
            .Where(item => item.Disposition == RedlineRevisionDisposition.PreExisting
                && item.Redline is not null)
            .Select(item => item.Redline!)
            .OrderBy(RedlineReversibilityVerifier.RevisionSortKey, StringComparer.Ordinal)
            .ToArray();
        var generatedEdited = classifications
            .Where(item => item.Disposition == RedlineRevisionDisposition.Generated
                && item.Redline is not null)
            .Select(item => item.Redline!)
            .OrderBy(RedlineReversibilityVerifier.RevisionSortKey, StringComparer.Ordinal)
            .ToArray();

        EnsureResolvable("pre-existing baseline", preExistingBaseline, policy.PreExistingRevisions);
        EnsureResolvable("pre-existing edited", preExistingEdited, policy.PreExistingRevisions);
        EnsureResolvable("generated edited", generatedEdited, policy.GeneratedRevisions);

        var policyBaseline = ApplyActions(
            baselineInput,
            new RevisionActionSet(policy.PreExistingRevisions, preExistingBaseline));
        var final = ApplyActions(
            editedInput,
            new RevisionActionSet(policy.PreExistingRevisions, preExistingEdited),
            new RevisionActionSet(policy.GeneratedRevisions, generatedEdited));

        byte[]? review = null;
        RedlineReversibilityProofRun? reviewProof = null;
        if (requireReviewProof)
        {
            // Reuse caller-authored native revisions when applying only the pre-existing policy
            // already makes them a complete reversible description of the final package. This
            // retains their exact identities. An untracked edit makes this candidate fail proof,
            // so the deterministic comparison below remains the complete fallback.
            var authoredReview = ApplyActions(
                editedInput,
                new RevisionActionSet(policy.PreExistingRevisions, preExistingEdited));
            var authoredProof = RedlineReversibilityVerifier.Prove(
                policyBaseline,
                final,
                authoredReview,
                new RedlineReversibilityProofOptions
                {
                    PackageManifestOptions = options,
                    MaxRevisionElements = MaxRevisionElements,
                });

            if (authoredProof.Proof.Success)
            {
                review = authoredReview;
                reviewProof = authoredProof;
            }
            else if (CanCanonicalizeAuthoredEndpoints(authoredProof))
            {
                var canonicalBaseline = authoredProof.RejectedPackageBytes!;
                var canonicalFinal = authoredProof.AcceptedPackageBytes!;
                var canonicalProof = RedlineReversibilityVerifier.Prove(
                    canonicalBaseline,
                    canonicalFinal,
                    authoredReview,
                    new RedlineReversibilityProofOptions
                    {
                        PackageManifestOptions = options,
                        MaxRevisionElements = MaxRevisionElements,
                    });
                if (canonicalProof.Proof.Success)
                {
                    // Select the resolver's concrete endpoints. They are semantically equivalent
                    // to the initially derived policy states and, unlike a separately saved clone,
                    // are the exact packages this native review accepts and rejects to.
                    policyBaseline = canonicalBaseline;
                    final = canonicalFinal;
                    review = authoredReview;
                    reviewProof = canonicalProof;
                }
            }

            if (reviewProof is null)
            {
                var generatedReview = DocxDiff.Compare(
                    new WmlDocument("policy-baseline.docx", policyBaseline),
                    new WmlDocument("final.docx", final),
                    new DocxDiffSettings
                    {
                        AuthorForRevisions = "Docxodus Delivery",
                        Deterministic = true,
                        PreAcceptInputRevisions = false,
                        PreserveInputRevisions = true,
                    }).DocumentByteArray;
                var generatedProof = RedlineReversibilityVerifier.Prove(
                    policyBaseline,
                    final,
                    generatedReview,
                    new RedlineReversibilityProofOptions
                    {
                        PackageManifestOptions = options,
                        MaxRevisionElements = MaxRevisionElements,
                    });
                if (!generatedProof.Proof.Success)
                {
                    var findingCodes = ProofFindings(authoredProof)
                        .Concat(ProofFindings(generatedProof))
                        .Select(item => item.Code)
                        .Distinct(StringComparer.Ordinal)
                        .OrderBy(item => item, StringComparer.Ordinal)
                        .ToArray();
                    throw new InvalidDataException(
                        "Neither the policy-filtered authored redline nor a deterministic "
                        + "comparison proved accept-to-final and reject-to-policy-baseline "
                        + "reversibility. Findings: "
                        + string.Join(", ", findingCodes)
                        + $". Authored accept divergence: {FirstDivergence(authoredProof, true)}"
                        + $"; authored reject divergence: {FirstDivergence(authoredProof, false)}"
                        + $"; generated accept divergence: {FirstDivergence(generatedProof, true)}"
                        + $"; generated reject divergence: {FirstDivergence(generatedProof, false)}.");
                }

                review = generatedReview;
                reviewProof = generatedProof;
            }
        }

        return new DeliveryRevisionPolicyResult
        {
            PolicyBaselineBytes = (byte[])policyBaseline.Clone(),
            FinalBytes = (byte[])final.Clone(),
            ReviewBytes = review is null ? null : (byte[])review.Clone(),
            ReviewProof = reviewProof,
            RevisionClassifications = classifications,
        };
    }

    private static IEnumerable<RedlineProofFinding> ProofFindings(
        RedlineReversibilityProofRun run) => run.Proof.Findings
        .Concat(run.Proof.AcceptToFinal?.Findings ?? Array.Empty<RedlineProofFinding>())
        .Concat(run.Proof.RejectToBaseline?.Findings ?? Array.Empty<RedlineProofFinding>());

    private static bool CanCanonicalizeAuthoredEndpoints(RedlineReversibilityProofRun run)
    {
        var accept = run.Proof.AcceptToFinal;
        var reject = run.Proof.RejectToBaseline;
        return run.AcceptedPackageBytes is not null
            && run.RejectedPackageBytes is not null
            && accept is
            {
                Completed: true,
                PreExistingRevisionsPreserved: true,
                ModeledSemantic.Available: true,
                ModeledSemantic.Equivalent: true,
            }
            && reject is
            {
                Completed: true,
                PreExistingRevisionsPreserved: true,
                ModeledSemantic.Available: true,
                ModeledSemantic.Equivalent: true,
            }
            && ProofFindings(run).Where(item =>
                    item.Severity == VerificationFindingSeverity.Error)
                .All(item => item.Code is
                    "normalized_whole_package_mismatch"
                    or "ordered_opc_content_mismatch"
                    or "raw_package_bytes_mismatch");
    }

    private static string FirstDivergence(RedlineReversibilityProofRun run, bool accept) =>
        (accept ? run.Proof.AcceptToFinal : run.Proof.RejectToBaseline)?.FirstDivergence is { } value
            ? $"{value.Kind}:{value.PartUri}"
            : "none";

    private static byte[] ApplyActions(byte[] input, params RevisionActionSet[] actionSets)
    {
        var actionable = actionSets.Where(item =>
                item.Action != DeliveryRevisionPolicy.Preserve && item.Revisions.Count > 0)
            .ToArray();
        if (actionable.Length == 0)
            return (byte[])input.Clone();

        using var session = OpenPolicySession(input);
        foreach (var actionSet in actionable)
        {
            foreach (var requested in actionSet.Revisions)
                Resolve(session, requested, actionSet.Action);
        }
        return session.Save(persistAnchorIds: false);
    }

    private static void Resolve(
        DocxSession session,
        RedlineRevisionIdentity requested,
        DeliveryRevisionPolicy action)
    {
        var currentEntries = session.ListRevisions();
        EnsureInventoryWithinLimit("working", currentEntries.Count);
        var currentIdentities = currentEntries
            .Select(RedlineReversibilityVerifier.ToIdentity)
            .ToArray();
        var exact = currentIdentities.Where(item =>
                RedlineReversibilityVerifier.IdentityEquivalent(requested, item))
            .ToArray();
        if (exact.Length == 0)
        {
            var overlap = currentIdentities.FirstOrDefault(item =>
                RedlineReversibilityVerifier.RevisionOverlaps(requested, item));
            if (overlap is not null)
            {
                throw new InvalidDataException(
                    $"Revision '{requested.Id}' changed identity to '{overlap.Id}' during policy resolution.");
            }

            // Resolving one member of a native linked family can consume another requested member.
            return;
        }
        if (exact.Length != 1)
        {
            throw new InvalidDataException(
                $"Revision '{requested.Id}' became ambiguous during policy resolution.");
        }
        if (exact[0].ResolutionStatus != RevisionResolutionStatus.Supported)
        {
            throw new InvalidDataException(
                $"Revision '{requested.Id}' is {exact[0].ResolutionStatus.ToString().ToLowerInvariant()} "
                + "and cannot be resolved by delivery policy.");
        }

        var edit = action == DeliveryRevisionPolicy.Accept
            ? session.AcceptRevision(requested.Id)
            : session.RejectRevision(requested.Id);
        if (!edit.Success)
        {
            throw new InvalidDataException(
                $"Revision '{requested.Id}' could not be {ActionVerb(action)}: "
                + (edit.Error?.Message ?? "unknown resolver failure"));
        }
    }

    private static void Preflight(
        string inputName,
        byte[] bytes,
        PackageManifestOptions options)
    {
        var manifest = PackageManifestGenerator.Generate(bytes, options);
        if (!manifest.IsValid)
        {
            var finding = manifest.Findings.FirstOrDefault(item =>
                item.Severity == VerificationFindingSeverity.Error);
            throw new InvalidDataException(
                $"The {inputName} package failed bounded validation."
                + (finding is null ? string.Empty : " " + finding.Message));
        }
        EnsureInventoryWithinLimit(inputName, manifest.Facts.Revisions.Total);
    }

    private static void EnsureInventoryWithinLimit(string inputName, int count)
    {
        if (count > MaxRevisionElements)
        {
            throw new InvalidDataException(
                $"The {inputName} revision inventory contains {count} entries; "
                + $"the delivery-policy limit is {MaxRevisionElements}.");
        }
    }

    private static void EnsureResolvable(
        string inputName,
        IReadOnlyList<RedlineRevisionIdentity> revisions,
        DeliveryRevisionPolicy action)
    {
        if (action == DeliveryRevisionPolicy.Preserve)
            return;
        var unsupported = revisions.FirstOrDefault(item =>
            item.ResolutionStatus != RevisionResolutionStatus.Supported);
        if (unsupported is not null)
        {
            throw new InvalidDataException(
                $"The {inputName} revision '{unsupported.Id}' is "
                + $"{unsupported.ResolutionStatus.ToString().ToLowerInvariant()} and cannot be "
                + $"{ActionVerb(action)}.");
        }
    }

    private static void ValidateAction(DeliveryRevisionPolicy action, string parameterName)
    {
        if (!Enum.IsDefined(action))
            throw new ArgumentOutOfRangeException(parameterName, action, null);
    }

    private static string ActionVerb(DeliveryRevisionPolicy action) => action switch
    {
        DeliveryRevisionPolicy.Accept => "accepted",
        DeliveryRevisionPolicy.Reject => "rejected",
        _ => throw new ArgumentOutOfRangeException(nameof(action), action, null),
    };

    private static DocxSession OpenPolicySession(byte[] bytes) => new(bytes, new DocxSessionSettings
    {
        UndoDepth = 1,
        UndoMemoryBudgetBytes = 128L * 1024 * 1024,
        PersistAnchorIds = false,
        EmitMarkdownPatch = false,
        CaptureInitialProjection = false,
    });

    private sealed record RevisionActionSet(
        DeliveryRevisionPolicy Action,
        IReadOnlyList<RedlineRevisionIdentity> Revisions);
}
