// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Docxodus.Verification;

namespace Docxodus.Delivery;

/// <summary>
/// What the session's host-owned evidence recorder holds right now (issue #748). A null
/// <see cref="UnavailableReason"/> means a complete change receipt can be minted from the
/// captured history; otherwise it names the first reason it cannot, and the receipt artifact
/// of any delivery is reported unavailable with that reason rather than claiming a history.
/// </summary>
public sealed record DeliveryEvidenceStatus
{
    public bool Enabled { get; init; }
    public int TransactionCount { get; init; }
    public int LineageEventCount { get; init; }

    /// <summary>
    /// Version steps applied outside a described transaction (typed direct calls, browser
    /// direct calls, a caller's own transaction scope). Their before/after packages are exact
    /// but their request is unknown, so they are recorded as <c>docx_session/unlabeled_mutation</c>.
    /// </summary>
    public int UnlabeledTransactionCount { get; init; }
    public int RetainedStateCount { get; init; }
    public long RetainedBytes { get; init; }
    public long SourceVersion { get; init; }
    public long CurrentVersion { get; init; }
    public string? UnavailableReason { get; init; }
}

/// <summary>Receipt policy for one delivery built from captured evidence.</summary>
public sealed record DeliveryReceiptBuildOptions
{
    public DeliveryReceiptPrivacyProfile PrivacyProfile { get; init; } =
        DeliveryReceiptPrivacyProfile.HashAndSummary;

    public bool FailOnUnexpectedChanges { get; init; }
}

/// <summary>
/// The captured history handed to a delivery operation: the exact opening package, the exact
/// current package, and — when the history is complete — the ordered transaction and undo/redo
/// evidence the bundle service turns into a change receipt.
/// </summary>
public sealed class DeliveryEvidenceExport
{
    internal DeliveryEvidenceExport(
        DeliveryReceiptContext? receiptContext,
        DeliveryDocumentSnapshot source,
        DeliveryDocumentSnapshot working,
        DeliveryEvidenceStatus status)
    {
        ReceiptContext = receiptContext;
        Source = source;
        Working = working;
        Status = status;
    }

    /// <summary>Null exactly when <see cref="Status"/> carries an unavailable reason.</summary>
    public DeliveryReceiptContext? ReceiptContext { get; }

    public DeliveryDocumentSnapshot Source { get; }

    public DeliveryDocumentSnapshot Working { get; }

    public DeliveryEvidenceStatus Status { get; }

    public string? UnavailableReason => Status.UnavailableReason;
}
