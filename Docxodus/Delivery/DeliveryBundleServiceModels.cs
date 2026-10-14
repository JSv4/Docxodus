// Copyright (c) John Scrudato IV. All rights reserved.
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
/// One step of a session's edit history in the order it happened: either a transaction's
/// evidence or an undo/redo lineage event. The receipt validates lineage as a single ordered
/// sequence, so an undo that happened between two transactions must be replayed between them
/// — a transactions-then-lineage split cannot express that.
/// </summary>
public sealed class DeliveryReceiptHistoryEvent
{
    private DeliveryReceiptHistoryEvent(
        DeliveryReceiptTransactionEvidence? transaction,
        DeliveryLineageEventInput? lineage)
    {
        Transaction = transaction;
        Lineage = lineage;
    }

    public DeliveryReceiptTransactionEvidence? Transaction { get; }

    public DeliveryLineageEventInput? Lineage { get; }

    public static DeliveryReceiptHistoryEvent FromTransaction(DeliveryReceiptTransactionEvidence evidence) =>
        new(evidence ?? throw new ArgumentNullException(nameof(evidence)), null);

    public static DeliveryReceiptHistoryEvent FromLineage(DeliveryLineageEventInput lineage) =>
        new(null, lineage ?? throw new ArgumentNullException(nameof(lineage)));
}

/// <summary>
/// Authoritative transaction/lineage evidence needed to mint a #458 receipt. The bundle service
/// never synthesizes missing mutation history from a baseline/final comparison.
/// </summary>
public sealed class DeliveryReceiptContext
{
    private readonly DeliveryReceiptHistoryEvent[] _history;
    private readonly DeliveryChangeAttributionRule[] _attributionRules;
    private readonly string[] _warnings;

    /// <summary>
    /// The original shape: every transaction, then every lineage event. Equivalent to the
    /// history constructor with that ordering, which is only a valid lineage when no undo or
    /// redo happened before the last transaction.
    /// </summary>
    public DeliveryReceiptContext(
        IEnumerable<DeliveryReceiptTransactionEvidence> transactions,
        IEnumerable<DeliveryLineageEventInput>? lineage = null,
        IEnumerable<DeliveryChangeAttributionRule>? attributionRules = null,
        IEnumerable<string>? warnings = null,
        DeliveryReceiptPrivacyProfile privacyProfile = DeliveryReceiptPrivacyProfile.HashAndSummary,
        bool failOnUnexpectedChanges = false)
        : this(
            InOrder(
                transactions ?? throw new ArgumentNullException(nameof(transactions)),
                lineage ?? Array.Empty<DeliveryLineageEventInput>()),
            attributionRules,
            warnings,
            privacyProfile,
            failOnUnexpectedChanges)
    {
    }

    /// <summary>Evidence in the exact order the session applied it (issue #748).</summary>
    public DeliveryReceiptContext(
        IEnumerable<DeliveryReceiptHistoryEvent> history,
        IEnumerable<DeliveryChangeAttributionRule>? attributionRules = null,
        IEnumerable<string>? warnings = null,
        DeliveryReceiptPrivacyProfile privacyProfile = DeliveryReceiptPrivacyProfile.HashAndSummary,
        bool failOnUnexpectedChanges = false)
    {
        _history = history?.ToArray() ?? throw new ArgumentNullException(nameof(history));
        _attributionRules = attributionRules?.ToArray()
            ?? Array.Empty<DeliveryChangeAttributionRule>();
        _warnings = warnings?.ToArray() ?? Array.Empty<string>();
        if (_history.Any(item => item is null)
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
    public IReadOnlyList<DeliveryReceiptHistoryEvent> History => _history.ToArray();
    public IReadOnlyList<DeliveryReceiptTransactionEvidence> Transactions =>
        _history.Where(item => item.Transaction is not null).Select(item => item.Transaction!).ToArray();
    public IReadOnlyList<DeliveryLineageEventInput> Lineage =>
        _history.Where(item => item.Lineage is not null).Select(item => item.Lineage!).ToArray();
    public IReadOnlyList<DeliveryChangeAttributionRule> AttributionRules =>
        _attributionRules.ToArray();
    public IReadOnlyList<string> Warnings => _warnings.ToArray();

    internal IReadOnlyList<DeliveryReceiptHistoryEvent> HistorySnapshot => _history;
    internal IReadOnlyList<DeliveryChangeAttributionRule> AttributionRuleSnapshot =>
        _attributionRules;
    internal IReadOnlyList<string> WarningSnapshot => _warnings;

    private static IEnumerable<DeliveryReceiptHistoryEvent> InOrder(
        IEnumerable<DeliveryReceiptTransactionEvidence> transactions,
        IEnumerable<DeliveryLineageEventInput> lineage)
    {
        foreach (var transaction in transactions)
        {
            if (transaction is null)
                throw new ArgumentException("Receipt context collections cannot contain null or blank entries.");
            yield return DeliveryReceiptHistoryEvent.FromTransaction(transaction);
        }
        foreach (var lineageEvent in lineage)
        {
            if (lineageEvent is null)
                throw new ArgumentException("Receipt context collections cannot contain null or blank entries.");
            yield return DeliveryReceiptHistoryEvent.FromLineage(lineageEvent);
        }
    }
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
