// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.Verification;

internal sealed record DeliveryLineageValidationResult
{
    required public bool IsValid { get; init; }
    public IReadOnlyList<string> Findings { get; init; } = Array.Empty<string>();
    public IReadOnlyList<DeliveryDocumentIdentity> ReachableDocuments { get; init; } =
        Array.Empty<DeliveryDocumentIdentity>();
    public IReadOnlyDictionary<long, DeliveryDocumentIdentity> ReachableDocumentsByVersion
        { get; init; } = new Dictionary<long, DeliveryDocumentIdentity>();
    public IReadOnlyList<DeliveryTransactionEntry> StateChangingTransactions { get; init; } =
        Array.Empty<DeliveryTransactionEntry>();
    public IReadOnlySet<string> AppliedTransactionEntryIds { get; init; } =
        new HashSet<string>(StringComparer.Ordinal);
}

/// <summary>
/// One deterministic state machine shared by receipt construction and independent verification.
/// Applied and redo histories are LIFO, and only state-changing committed transactions affect them.
/// </summary>
internal static class DeliveryReceiptLineageValidator
{
    public static DeliveryLineageValidationResult Validate(
        DeliveryDocumentIdentity sourceDocument,
        DeliveryDocumentIdentity deliveredDocument,
        IReadOnlyList<DeliveryTransactionEntry> transactions,
        IReadOnlyList<DeliveryLineageEvent> lineage)
    {
        var reachable = new List<DeliveryDocumentIdentity>();
        var identityByVersion = new Dictionary<long, DeliveryDocumentIdentity>();
        var stateChanging = new List<DeliveryTransactionEntry>();
        var applied = new List<DeliveryTransactionEntry>();
        var redo = new List<DeliveryTransactionEntry>();

        DeliveryLineageValidationResult Fail(string finding) => new()
        {
            IsValid = false,
            Findings = new[] { finding },
            ReachableDocuments = reachable.ToArray(),
            ReachableDocumentsByVersion = new Dictionary<long, DeliveryDocumentIdentity>(
                identityByVersion),
            StateChangingTransactions = stateChanging.ToArray(),
            AppliedTransactionEntryIds = applied
                .Select(entry => entry.EntryId)
                .ToHashSet(StringComparer.Ordinal),
        };

        bool RegisterReachable(DeliveryDocumentIdentity document)
        {
            if (identityByVersion.TryGetValue(document.DocumentVersion, out var existing))
            {
                return DocumentEquals(existing, document);
            }
            identityByVersion.Add(document.DocumentVersion, document);
            reachable.Add(document);
            return true;
        }

        if (!RegisterReachable(sourceDocument))
            return Fail("document_version_collision");

        var transactionGroups = transactions
            .GroupBy(transaction => transaction.Sequence)
            .ToDictionary(group => group.Key, group => group.ToArray());
        var lineageGroups = lineage
            .GroupBy(lineageEvent => lineageEvent.Sequence)
            .ToDictionary(group => group.Key, group => group.ToArray());
        var entriesById = transactions
            .GroupBy(transaction => transaction.EntryId, StringComparer.Ordinal)
            .Where(group => group.Count() == 1)
            .ToDictionary(group => group.Key, group => group.Single(), StringComparer.Ordinal);
        long transitionCount = transactions.Count + lineage.Count;
        var current = sourceDocument;

        for (long sequence = 0; sequence < transitionCount; sequence++)
        {
            transactionGroups.TryGetValue(sequence, out var transactionGroup);
            lineageGroups.TryGetValue(sequence, out var lineageGroup);
            if ((transactionGroup?.Length ?? 0) + (lineageGroup?.Length ?? 0) != 1)
                return Fail("lineage_sequence_mismatch");

            if (transactionGroup is { Length: 1 })
            {
                var transaction = transactionGroup[0];
                if (!DocumentEquals(current, transaction.BeforeDocument))
                    return Fail("transaction_lineage_gap");

                switch (transaction.Status)
                {
                    case DeliveryTransactionStatus.Committed:
                    case DeliveryTransactionStatus.PartiallyCommitted:
                        if (DocumentEquals(transaction.BeforeDocument, transaction.AfterDocument))
                            break;
                        if (transaction.AfterDocument.DocumentVersion
                            != transaction.BeforeDocument.DocumentVersion + 1)
                        {
                            return Fail("invalid_transaction_version");
                        }
                        current = transaction.AfterDocument;
                        if (!RegisterReachable(current))
                            return Fail("document_version_collision");
                        applied.Add(transaction);
                        redo.Clear();
                        stateChanging.Add(transaction);
                        break;
                    case DeliveryTransactionStatus.Failed:
                        if (!DocumentEquals(transaction.BeforeDocument, transaction.AfterDocument))
                            return Fail("failed_transaction_changed_document");
                        break;
                    case DeliveryTransactionStatus.Prediction:
                        if (!DocumentEquals(transaction.BeforeDocument, transaction.AfterDocument)
                            && (SameIdentityExceptVersion(
                                    transaction.BeforeDocument, transaction.AfterDocument)
                                || transaction.AfterDocument.DocumentVersion
                                != transaction.BeforeDocument.DocumentVersion + 1
                                || !transaction.Operations.Any(operation =>
                                    operation.ExecutionStatus
                                        == DeliveryOperationExecutionStatus.Succeeded)))
                        {
                            return Fail("invalid_prediction_transition");
                        }
                        break;
                    default:
                        return Fail("invalid_transaction_status");
                }
                continue;
            }

            var lineageEvent = lineageGroup![0];
            if (!DocumentEquals(current, lineageEvent.BeforeDocument))
                return Fail("lineage_gap");
            if (lineageEvent.AfterDocument.DocumentVersion
                != lineageEvent.BeforeDocument.DocumentVersion + 1)
            {
                return Fail("invalid_lineage_version");
            }
            if (!entriesById.TryGetValue(lineageEvent.AffectedEntryId, out var affected)
                || affected.Status is not (DeliveryTransactionStatus.Committed
                    or DeliveryTransactionStatus.PartiallyCommitted)
                || DocumentEquals(affected.BeforeDocument, affected.AfterDocument))
            {
                return Fail("invalid_lineage_target");
            }

            if (lineageEvent.Action == DeliveryLineageAction.Undo)
            {
                if (applied.Count == 0
                    || !string.Equals(applied[^1].EntryId, affected.EntryId,
                        StringComparison.Ordinal))
                {
                    return Fail("invalid_undo_order");
                }
                if (!SamePackageContent(affected.BeforeDocument, lineageEvent.AfterDocument))
                    return Fail("lineage_package_mismatch");
                applied.RemoveAt(applied.Count - 1);
                redo.Add(affected);
            }
            else if (lineageEvent.Action == DeliveryLineageAction.Redo)
            {
                if (redo.Count == 0
                    || !string.Equals(redo[^1].EntryId, affected.EntryId,
                        StringComparison.Ordinal))
                {
                    return Fail("invalid_redo_order");
                }
                if (!SamePackageContent(affected.AfterDocument, lineageEvent.AfterDocument))
                    return Fail("lineage_package_mismatch");
                redo.RemoveAt(redo.Count - 1);
                applied.Add(affected);
            }
            else
            {
                return Fail("unknown_lineage_action");
            }

            current = lineageEvent.AfterDocument;
            if (!RegisterReachable(current))
                return Fail("document_version_collision");
        }

        if (!DocumentEquals(current, deliveredDocument))
            return Fail("delivered_lineage_mismatch");

        return new DeliveryLineageValidationResult
        {
            IsValid = true,
            ReachableDocuments = reachable.ToArray(),
            ReachableDocumentsByVersion = new Dictionary<long, DeliveryDocumentIdentity>(
                identityByVersion),
            StateChangingTransactions = stateChanging.ToArray(),
            AppliedTransactionEntryIds = applied
                .Select(entry => entry.EntryId)
                .ToHashSet(StringComparer.Ordinal),
        };
    }

    public static bool IsReachable(
        DeliveryLineageValidationResult validation,
        long documentVersion,
        VerificationDigest packageDigest) =>
        validation.ReachableDocumentsByVersion.TryGetValue(
            documentVersion, out var document)
        && DeliveryReceiptValidation.DigestEquals(
            document.RawPackageBytesDigest, packageDigest);

    public static bool DocumentEquals(
        DeliveryDocumentIdentity left,
        DeliveryDocumentIdentity right) =>
        left.DocumentVersion == right.DocumentVersion
        && string.Equals(left.PackageKind, right.PackageKind, StringComparison.Ordinal)
        && string.Equals(left.PackageManifestSchema, right.PackageManifestSchema,
            StringComparison.Ordinal)
        && string.Equals(left.MainDocumentUri, right.MainDocumentUri, StringComparison.Ordinal)
        && DeliveryReceiptValidation.DigestEquals(
            left.RawPackageBytesDigest, right.RawPackageBytesDigest)
        && OptionalDigestEquals(left.OrderedOpcContentDigest, right.OrderedOpcContentDigest)
        && OptionalDigestEquals(left.NormalizedSemanticDigest, right.NormalizedSemanticDigest);

    private static bool SameIdentityExceptVersion(
        DeliveryDocumentIdentity left,
        DeliveryDocumentIdentity right) =>
        string.Equals(left.PackageKind, right.PackageKind, StringComparison.Ordinal)
        && string.Equals(left.PackageManifestSchema, right.PackageManifestSchema,
            StringComparison.Ordinal)
        && string.Equals(left.MainDocumentUri, right.MainDocumentUri, StringComparison.Ordinal)
        && DeliveryReceiptValidation.DigestEquals(
            left.RawPackageBytesDigest, right.RawPackageBytesDigest)
        && OptionalDigestEquals(left.OrderedOpcContentDigest, right.OrderedOpcContentDigest)
        && OptionalDigestEquals(left.NormalizedSemanticDigest, right.NormalizedSemanticDigest);

    public static bool SamePackageContent(
        DeliveryDocumentIdentity left,
        DeliveryDocumentIdentity right)
    {
        if (left.OrderedOpcContentDigest is not null
            && right.OrderedOpcContentDigest is not null)
        {
            return DeliveryReceiptValidation.DigestEquals(
                left.OrderedOpcContentDigest, right.OrderedOpcContentDigest);
        }
        return DeliveryReceiptValidation.DigestEquals(
            left.RawPackageBytesDigest, right.RawPackageBytesDigest);
    }

    private static bool OptionalDigestEquals(VerificationDigest? left, VerificationDigest? right) =>
        left is null && right is null
        || DeliveryReceiptValidation.DigestEquals(left, right);
}
