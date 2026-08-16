// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using Docxodus.Verification;

namespace Docxodus.Delivery;

/// <summary>Exact before/after snapshots paired with one authoritative #458 contribution.</summary>
public sealed class DeliveryReceiptTransactionEvidence
{
    public DeliveryReceiptTransactionEvidence(
        DeliveryTransactionContribution contribution,
        DeliveryDocumentSnapshot before,
        DeliveryDocumentSnapshot after)
    {
        Contribution = contribution ?? throw new ArgumentNullException(nameof(contribution));
        Before = before ?? throw new ArgumentNullException(nameof(before));
        After = after ?? throw new ArgumentNullException(nameof(after));
    }

    public DeliveryTransactionContribution Contribution { get; }
    public DeliveryDocumentSnapshot Before { get; }
    public DeliveryDocumentSnapshot After { get; }
}

/// <summary>
/// Authoritative transaction/lineage evidence needed to mint a #458 receipt. The bundle service
/// never synthesizes missing mutation history from a baseline/final comparison.
/// </summary>
public sealed class DeliveryReceiptContext
{
    private readonly DeliveryReceiptTransactionEvidence[] _transactions;
    private readonly DeliveryLineageEventInput[] _lineage;
    private readonly DeliveryChangeAttributionRule[] _attributionRules;
    private readonly string[] _warnings;

    public DeliveryReceiptContext(
        IEnumerable<DeliveryReceiptTransactionEvidence> transactions,
        IEnumerable<DeliveryLineageEventInput>? lineage = null,
        IEnumerable<DeliveryChangeAttributionRule>? attributionRules = null,
        IEnumerable<string>? warnings = null,
        DeliveryReceiptPrivacyProfile privacyProfile = DeliveryReceiptPrivacyProfile.HashAndSummary,
        bool failOnUnexpectedChanges = false)
    {
        _transactions = transactions?.ToArray()
            ?? throw new ArgumentNullException(nameof(transactions));
        _lineage = lineage?.ToArray() ?? Array.Empty<DeliveryLineageEventInput>();
        _attributionRules = attributionRules?.ToArray()
            ?? Array.Empty<DeliveryChangeAttributionRule>();
        _warnings = warnings?.ToArray() ?? Array.Empty<string>();
        if (_transactions.Any(item => item is null)
            || _lineage.Any(item => item is null)
            || _attributionRules.Any(item => item is null)
            || _warnings.Any(string.IsNullOrWhiteSpace))
            throw new ArgumentException("Receipt context collections cannot contain null or blank entries.");
        if (!Enum.IsDefined(privacyProfile))
            throw new ArgumentOutOfRangeException(nameof(privacyProfile));
        PrivacyProfile = privacyProfile;
        FailOnUnexpectedChanges = failOnUnexpectedChanges;
    }

    public DeliveryReceiptPrivacyProfile PrivacyProfile { get; }
    public bool FailOnUnexpectedChanges { get; }
    public IReadOnlyList<DeliveryReceiptTransactionEvidence> Transactions =>
        _transactions.ToArray();
    public IReadOnlyList<DeliveryLineageEventInput> Lineage => _lineage.ToArray();
    public IReadOnlyList<DeliveryChangeAttributionRule> AttributionRules =>
        _attributionRules.ToArray();
    public IReadOnlyList<string> Warnings => _warnings.ToArray();

    internal IReadOnlyList<DeliveryReceiptTransactionEvidence> TransactionSnapshot => _transactions;
    internal IReadOnlyList<DeliveryLineageEventInput> LineageSnapshot => _lineage;
    internal IReadOnlyList<DeliveryChangeAttributionRule> AttributionRuleSnapshot =>
        _attributionRules;
    internal IReadOnlyList<string> WarningSnapshot => _warnings;
}

/// <summary>Caller intent before revision policy derives the exact named final bytes.</summary>
public sealed class DeliveryBundleBuildRequest
{
    private readonly DeliveryArtifactRequest[] _artifacts;

    public DeliveryBundleBuildRequest(
        DeliveryDocumentSnapshot baseline,
        DeliveryDocumentSnapshot working,
        string finalDocumentName,
        long finalDocumentVersion,
        DeliveryBundleRevisionPolicy revisionPolicy,
        IEnumerable<DeliveryArtifactRequest> artifacts,
        DeliveryReceiptContext? receiptContext = null)
    {
        Baseline = baseline ?? throw new ArgumentNullException(nameof(baseline));
        Working = working ?? throw new ArgumentNullException(nameof(working));
        if (string.IsNullOrWhiteSpace(finalDocumentName))
            throw new ArgumentException("A final document name is required.", nameof(finalDocumentName));
        if (finalDocumentVersion < 0)
            throw new ArgumentOutOfRangeException(nameof(finalDocumentVersion));
        RevisionPolicy = revisionPolicy ?? throw new ArgumentNullException(nameof(revisionPolicy));
        _artifacts = artifacts?.ToArray() ?? throw new ArgumentNullException(nameof(artifacts));
        FinalDocumentName = finalDocumentName;
        FinalDocumentVersion = finalDocumentVersion;
        ReceiptContext = receiptContext;
    }

    public DeliveryDocumentSnapshot Baseline { get; }
    public DeliveryDocumentSnapshot Working { get; }
    public string FinalDocumentName { get; }
    public long FinalDocumentVersion { get; }
    public DeliveryBundleRevisionPolicy RevisionPolicy { get; }
    public IReadOnlyList<DeliveryArtifactRequest> Artifacts => _artifacts.ToArray();
    public DeliveryReceiptContext? ReceiptContext { get; }

    internal IReadOnlyList<DeliveryArtifactRequest> ArtifactSnapshot => _artifacts;
}

/// <summary>Policy knobs owned by the #465 orchestrator.</summary>
public sealed record DeliveryBundleBuildOptions
{
    public PackageManifestOptions PackageManifestOptions { get; init; } = new();
    public DeliverableVerificationOptions DeliverableVerificationOptions { get; init; } = new();
    public DeliveryReceiptLimits DeliveryReceiptLimits { get; init; } = new();
    public DeliveryBundleVerificationLimits BundleVerificationLimits { get; init; } = new();

    /// <summary>Reject a final DOCX when #463's selected policy returns Failed.</summary>
    public bool FailOnDeliverableValidationFailure { get; init; } = true;

    /// <summary>
    /// Permit a byte-return result whose manifest is explicitly incomplete. Directory publication
    /// still rejects incomplete/failed bundles unless its caller makes a separate diagnostic choice.
    /// </summary>
    public bool ReturnIncompleteBundle { get; init; }
}

/// <summary>Stable failure from artifact planning or bundle orchestration.</summary>
public sealed class DeliveryBundleException : InvalidOperationException
{
    public DeliveryBundleException(string code, string message)
        : base(message)
    {
        Code = code;
    }

    public string Code { get; }
}
