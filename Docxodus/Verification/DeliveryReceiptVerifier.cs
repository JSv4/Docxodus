// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text;
using System.Text.Json;
using Docxodus.Internal;

namespace Docxodus.Verification;

public enum DeliveryArtifactVerificationStatus
{
    Verified,
    Unavailable,
    Missing,
    LengthMismatch,
    DigestMismatch,
    InvalidRecord,
}

public sealed record DeliveryArtifactVerification
{
    required public string ArtifactId { get; init; }
    required public DeliveryArtifactVerificationStatus Status { get; init; }
    public long? ExpectedLength { get; init; }
    public long? ActualLength { get; init; }
    public VerificationDigest? ExpectedDigest { get; init; }
    public VerificationDigest? ActualDigest { get; init; }
}

public sealed record DeliveryReceiptVerificationResult
{
    required public bool IsValid { get; init; }
    required public bool ReceiptDigestValid { get; init; }
    required public bool ContractValid { get; init; }
    required public bool CitationBindingsValid { get; init; }
    public IReadOnlyList<DeliveryArtifactVerification> Artifacts { get; init; } =
        Array.Empty<DeliveryArtifactVerification>();
    public IReadOnlyList<string> Findings { get; init; } = Array.Empty<string>();
}

/// <summary>Verifies receipt integrity first, then independently verifies every available artifact.</summary>
public static class DeliveryChangeReceiptVerifier
{
    public static DeliveryReceiptVerificationResult Verify(
        DeliveryChangeReceipt receipt,
        IReadOnlyDictionary<string, byte[]> artifactBytes,
        DeliveryReceiptVerificationOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(receipt);
        ArgumentNullException.ThrowIfNull(artifactBytes);
        var limits = (options ?? new DeliveryReceiptVerificationOptions())
            .ValidateAndClone().Limits;
        try
        {
            DeliveryReceiptResourceValidator.ValidateArtifacts(artifactBytes, limits);
            DeliveryReceiptResourceValidator.ValidatePayload(receipt.Payload, limits);
            var canonicalPayload = DeliveryChangeReceiptSerializer.SerializePayload(
                receipt.Payload, limits);
            var digestValid = DeliveryReceiptCanonicalJson.FixedTimeEquals(
                receipt.ReceiptDigest, canonicalPayload);
            return VerifyCore(
                receipt.Payload, receipt.ReceiptDigest, digestValid, artifactBytes, limits);
        }
        catch (DeliveryReceiptValidationException ex) when (IsResourceLimitCode(ex.Code))
        {
            return Rejected(ex.Code);
        }
        catch (Exception ex) when (IsMalformedReceiptException(ex))
        {
            return Malformed($"malformed_receipt:{ex.GetType().Name}");
        }
    }

    /// <summary>
    /// Parse and verify a portable JSON receipt. Unknown optional properties remain covered by
    /// the raw canonical payload digest even though this v1 reader ignores them semantically.
    /// </summary>
    public static DeliveryReceiptVerificationResult VerifyJson(
        ReadOnlySpan<byte> receiptJson,
        IReadOnlyDictionary<string, byte[]> artifactBytes,
        DeliveryReceiptVerificationOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(artifactBytes);
        var limits = (options ?? new DeliveryReceiptVerificationOptions())
            .ValidateAndClone().Limits;
        try
        {
            DeliveryReceiptResourceValidator.ValidateArtifacts(artifactBytes, limits);
            var canonicalEnvelope = DeliveryReceiptCanonicalJson.CanonicalizeBounded(
                receiptJson, limits, limits.MaxReceiptJsonBytes, "receipt_resource_limit");
            using var document = JsonDocument.Parse(canonicalEnvelope, new JsonDocumentOptions
            {
                AllowTrailingCommas = false,
                CommentHandling = JsonCommentHandling.Disallow,
                MaxDepth = limits.MaxJsonDepth,
            });
            if (document.RootElement.ValueKind != JsonValueKind.Object
                || document.RootElement.EnumerateObject().Any(property =>
                    property.Name is not ("payload" or "receiptDigest"))
                || !document.RootElement.TryGetProperty("payload", out var payloadElement)
                || payloadElement.ValueKind != JsonValueKind.Object
                || !document.RootElement.TryGetProperty("receiptDigest", out var digestElement)
                || digestElement.ValueKind != JsonValueKind.Object)
            {
                return Malformed("malformed_envelope");
            }

            var digest = ReadDigest(digestElement);
            var canonicalPayload = DeliveryReceiptCanonicalJson.SerializeCanonicalBounded(
                payloadElement, limits, limits.MaxReceiptJsonBytes, "receipt_resource_limit");
            var digestValid = DeliveryReceiptCanonicalJson.FixedTimeEquals(digest, canonicalPayload);
            var jsonOptions = new JsonSerializerOptions(DeliveryReceiptCanonicalJson.JsonOptions)
            {
                MaxDepth = limits.MaxJsonDepth,
            };
            var payload = JsonSerializer.Deserialize<DeliveryChangeReceiptPayload>(
                payloadElement.GetRawText(), jsonOptions);
            if (payload is null)
                return Malformed("missing_payload");
            DeliveryReceiptResourceValidator.ValidatePayload(payload, limits);
            return VerifyCore(payload, digest, digestValid, artifactBytes, limits);
        }
        catch (DeliveryReceiptValidationException ex) when (IsResourceLimitCode(ex.Code))
        {
            return Rejected(ex.Code);
        }
        catch (Exception ex) when (IsMalformedReceiptException(ex))
        {
            return Malformed($"malformed_receipt:{ex.GetType().Name}");
        }
    }

    public static DeliveryReceiptVerificationResult VerifyJson(
        string receiptJson,
        IReadOnlyDictionary<string, byte[]> artifactBytes,
        DeliveryReceiptVerificationOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(receiptJson);
        var validatedOptions = (options ?? new DeliveryReceiptVerificationOptions())
            .ValidateAndClone();
        if (Encoding.UTF8.GetByteCount(receiptJson)
            > validatedOptions.Limits.MaxReceiptJsonBytes)
        {
            return Rejected("receipt_resource_limit");
        }
        return VerifyJson(Encoding.UTF8.GetBytes(receiptJson), artifactBytes, validatedOptions);
    }

    private static DeliveryReceiptVerificationResult VerifyCore(
        DeliveryChangeReceiptPayload payload,
        VerificationDigest receiptDigest,
        bool digestValid,
        IReadOnlyDictionary<string, byte[]> artifactBytes,
        DeliveryReceiptLimits limits)
    {
        var findings = new List<string>();
        var lineageValidation = DeliveryReceiptLineageValidator.Validate(
            payload.SourceDocument,
            payload.DeliveredDocument,
            payload.Transactions,
            payload.Lineage);
        bool contractValid = ValidateContract(
            payload, receiptDigest, artifactBytes, lineageValidation, limits, findings);
        var artifactResults = VerifyArtifacts(payload.Artifacts, artifactBytes, limits, findings);
        bool artifactsValid = artifactResults.All(result => result.Status is
            DeliveryArtifactVerificationStatus.Verified
            or DeliveryArtifactVerificationStatus.Unavailable);
        bool citationsValid = ValidateCitationBindings(
            payload, artifactBytes, lineageValidation, limits, findings);
        if (!digestValid)
            findings.Add("receipt_digest_mismatch");

        return new DeliveryReceiptVerificationResult
        {
            IsValid = digestValid && contractValid && artifactsValid && citationsValid,
            ReceiptDigestValid = digestValid,
            ContractValid = contractValid,
            CitationBindingsValid = citationsValid,
            Artifacts = artifactResults,
            Findings = findings.Distinct(StringComparer.Ordinal).ToArray(),
        };
    }

    private static bool ValidateContract(
        DeliveryChangeReceiptPayload payload,
        VerificationDigest receiptDigest,
        IReadOnlyDictionary<string, byte[]> artifactBytes,
        DeliveryLineageValidationResult lineageValidation,
        DeliveryReceiptLimits limits,
        ICollection<string> findings)
    {
        bool valid = true;
        if (!string.Equals(payload.Schema, DeliveryChangeReceiptPayload.SchemaId,
                StringComparison.Ordinal)
            || payload.SchemaVersion != 1)
        {
            findings.Add("unsupported_receipt_schema");
            valid = false;
        }
        if (!string.Equals(payload.Canonicalization,
                DeliveryChangeReceiptPayload.CanonicalizationId, StringComparison.Ordinal))
        {
            findings.Add("unsupported_canonicalization");
            valid = false;
        }
        if (!Enum.IsDefined(payload.PrivacyProfile))
        {
            findings.Add("unknown_privacy_profile");
            valid = false;
        }
        try
        {
            DeliveryReceiptValidation.ValidateDigest(receiptDigest, "receipt digest");
            ValidateDocument(payload.SourceDocument);
            ValidateDocument(payload.DeliveredDocument);
        }
        catch (DeliveryReceiptValidationException ex)
        {
            findings.Add(ex.Code);
            valid = false;
        }

        if (payload.HasUnexpectedChanges != payload.PackageChanges.Any(change =>
                change.Disposition == DeliveryChangeDisposition.Unexpected))
        {
            findings.Add("unexpected_change_flag_mismatch");
            valid = false;
        }
        if (payload.Transactions.Select(entry => entry.EntryId)
            .Distinct(StringComparer.Ordinal).Count() != payload.Transactions.Count)
        {
            findings.Add("duplicate_transaction_entry");
            valid = false;
        }
        var transactionIds = payload.Transactions.Where(entry => entry.TransactionId is not null)
            .Select(entry => entry.TransactionId!);
        if (transactionIds.Distinct(StringComparer.Ordinal).Count() != transactionIds.Count())
        {
            findings.Add("duplicate_transaction_id");
            valid = false;
        }
        if (payload.Artifacts.Select(artifact => artifact.ArtifactId)
            .Distinct(StringComparer.Ordinal).Count() != payload.Artifacts.Count)
        {
            findings.Add("duplicate_artifact_id");
            valid = false;
        }
        if (!IsStrictlyOrdered(payload.Transactions,
                static (left, right) => left.Sequence.CompareTo(right.Sequence)))
        {
            findings.Add("transaction_order_mismatch");
            valid = false;
        }
        if (!IsStrictlyOrdered(payload.Lineage,
                static (left, right) => left.Sequence.CompareTo(right.Sequence)))
        {
            findings.Add("lineage_order_mismatch");
            valid = false;
        }
        if (!IsStrictlyOrdered(payload.Artifacts,
                static (left, right) => string.CompareOrdinal(
                    left.ArtifactId, right.ArtifactId)))
        {
            findings.Add("artifact_order_mismatch");
            valid = false;
        }
        if (!IsStrictlyOrdered(payload.PackageChanges, ComparePackageChanges))
        {
            findings.Add("package_change_order_mismatch");
            valid = false;
        }
        if (!IsStrictlyOrdered(payload.PageCitations, CompareCitations))
        {
            findings.Add("citation_order_mismatch");
            valid = false;
        }
        if (!IsStrictlyOrdered(payload.Evidence, CompareEvidence))
        {
            findings.Add("evidence_order_mismatch");
            valid = false;
        }
        if (!IsStrictlyOrdered(payload.Warnings, CompareTextEvidence))
        {
            findings.Add("warning_order_mismatch");
            valid = false;
        }

        foreach (var transaction in payload.Transactions)
            valid &= ValidateTransaction(transaction, payload.PrivacyProfile, findings);
        foreach (var lineageEvent in payload.Lineage)
        {
            if (lineageEvent.Sequence < 0)
            {
                findings.Add("invalid_lineage_sequence");
                valid = false;
            }
            if (!Enum.IsDefined(lineageEvent.Action))
            {
                findings.Add("unknown_lineage_action");
                valid = false;
            }
            try
            {
                DeliveryReceiptValidation.RequireNonBlank(
                    lineageEvent.AffectedEntryId, "lineage entry id", 256);
                ValidateDocument(lineageEvent.BeforeDocument);
                ValidateDocument(lineageEvent.AfterDocument);
            }
            catch (DeliveryReceiptValidationException ex)
            {
                findings.Add(ex.Code);
                valid = false;
            }
        }
        if (!lineageValidation.IsValid)
        {
            foreach (var finding in lineageValidation.Findings)
                findings.Add(finding);
            valid = false;
        }
        foreach (var change in payload.PackageChanges)
            valid &= ValidatePackageChange(change, payload.PrivacyProfile, findings);
        if (payload.PackageChanges.Select(change => change.ChangeId)
            .Distinct(StringComparer.Ordinal).Count() != payload.PackageChanges.Count)
        {
            findings.Add("duplicate_package_change_id");
            valid = false;
        }
        var transactionsByEntryId = payload.Transactions
            .GroupBy(transaction => transaction.EntryId, StringComparer.Ordinal)
            .Where(group => group.Count() == 1)
            .ToDictionary(group => group.Key, group => group.Single(), StringComparer.Ordinal);
        foreach (var change in payload.PackageChanges)
        {
            if (change.Disposition == DeliveryChangeDisposition.UserRequested
                && (change.TransactionEntryId is null
                    || change.RequestedOperationIndex is null))
            {
                findings.Add("invalid_package_change_attribution");
                valid = false;
                continue;
            }
            if (change.TransactionEntryId is null)
            {
                if (change.RequestedOperationIndex is not null)
                {
                    findings.Add("invalid_package_change_attribution");
                    valid = false;
                }
                continue;
            }
            if (!transactionsByEntryId.TryGetValue(change.TransactionEntryId, out var transaction)
                || transaction.Status is not (DeliveryTransactionStatus.Committed
                    or DeliveryTransactionStatus.PartiallyCommitted)
                || (change.RequestedOperationIndex is { } operationIndex
                    && (operationIndex < 0 || operationIndex >= transaction.Operations.Count
                        || transaction.Operations[operationIndex].ExecutionStatus
                            != DeliveryOperationExecutionStatus.Succeeded)))
            {
                findings.Add("invalid_package_change_attribution");
                valid = false;
            }
        }
        foreach (var artifact in payload.Artifacts)
            valid &= ValidateArtifactRecord(artifact, findings);

        var artifactsById = payload.Artifacts
            .GroupBy(artifact => artifact.ArtifactId, StringComparer.Ordinal)
            .Where(group => group.Count() == 1)
            .ToDictionary(group => group.Key, group => group.Single(), StringComparer.Ordinal);
        valid &= ValidateRequiredCleanDocx(payload, artifactBytes, limits, findings);
        valid &= ValidateSemanticBindings(
            payload, artifactsById, artifactBytes, lineageValidation, limits, findings);
        foreach (var evidence in payload.Evidence)
        {
            if (!Enum.IsDefined(evidence.Kind))
            {
                findings.Add("unknown_evidence_kind");
                valid = false;
            }
            if (evidence.Kind == DeliveryEvidenceKind.SemanticChangeSet)
            {
                findings.Add("semantic_evidence_requires_typed_factory");
                valid = false;
            }
            try { DeliveryReceiptValidation.ValidateDigest(evidence.Digest, "evidence digest"); }
            catch (DeliveryReceiptValidationException ex)
            {
                findings.Add(ex.Code);
                valid = false;
            }
            try
            {
                DeliveryReceiptValidation.RequireNonBlank(
                    evidence.Schema, "evidence schema", 2048);
                if (evidence.ArtifactId is not null)
                {
                    DeliveryReceiptValidation.RequireNonBlank(
                        evidence.ArtifactId, "evidence artifact id", 256);
                }
            }
            catch (DeliveryReceiptValidationException ex)
            {
                findings.Add(ex.Code);
                valid = false;
            }
            if (evidence.ArtifactId is { } artifactId)
            {
                var expectedRole = evidence.Kind switch
                {
                    DeliveryEvidenceKind.SemanticChangeSet => DeliveryArtifactRole.SemanticDiff,
                    DeliveryEvidenceKind.ValidationResult => DeliveryArtifactRole.ValidationReport,
                    DeliveryEvidenceKind.RedlineReversibility =>
                        DeliveryArtifactRole.ReversibilityProof,
                    _ => (DeliveryArtifactRole?)null,
                };
                if (!artifactsById.TryGetValue(artifactId, out var artifact)
                    || artifact.Digest is null
                    || artifact.Availability != DeliveryArtifactAvailability.Available
                    || artifact.Role != expectedRole
                    || !DeliveryReceiptValidation.DigestEquals(artifact.Digest, evidence.Digest))
                {
                    findings.Add("evidence_artifact_binding_mismatch");
                    valid = false;
                }
            }
            if (payload.PrivacyProfile == DeliveryReceiptPrivacyProfile.HashOnly
                && evidence.Summary is not null)
            {
                findings.Add("privacy_profile_violation");
                valid = false;
            }
        }
        foreach (var warning in payload.Warnings)
            valid &= ValidateTextEvidence(warning, payload.PrivacyProfile, findings);
        return valid;
    }

    private static bool ValidateTransaction(
        DeliveryTransactionEntry transaction,
        DeliveryReceiptPrivacyProfile privacyProfile,
        ICollection<string> findings)
    {
        bool valid = true;
        if (transaction.Sequence < 0)
        {
            findings.Add("invalid_lineage_sequence");
            valid = false;
        }
        if (!Enum.IsDefined(transaction.Mode) || !Enum.IsDefined(transaction.Status))
        {
            findings.Add("unknown_transaction_enum");
            valid = false;
        }
        try
        {
            DeliveryTransactionContribution.ValidateFingerprint(transaction.EntryId);
            DeliveryTransactionContribution.ValidateFingerprint(transaction.RequestFingerprint);
            if (transaction.TransactionId is not null)
            {
                DeliveryReceiptValidation.RequireNonBlank(
                    transaction.TransactionId, "transaction id", 1024);
            }
            ValidateDocument(transaction.BeforeDocument);
            ValidateDocument(transaction.AfterDocument);
        }
        catch (DeliveryReceiptValidationException ex)
        {
            findings.Add(ex.Code);
            valid = false;
        }
        if (transaction.BaseVersion < 0 || transaction.ResultVersion < 0
            || transaction.BeforeDocument.DocumentVersion != transaction.BaseVersion
            || transaction.AfterDocument.DocumentVersion != transaction.ResultVersion)
        {
            findings.Add("transaction_version_mismatch");
            valid = false;
        }

        if (!Enum.IsDefined(transaction.Status)
            || !ValidateTransactionOutcome(transaction, findings))
        {
            valid = false;
        }
        var expectedEntryId = DeliveryReceiptCanonicalJson.DigestToken(
            DeliveryReceiptCanonicalJson.SerializeCanonical(new
            {
                afterPackage = transaction.AfterDocument.RawPackageBytesDigest,
                baseVersion = transaction.BaseVersion,
                beforePackage = transaction.BeforeDocument.RawPackageBytesDigest,
                requestFingerprint = transaction.RequestFingerprint,
                resultVersion = transaction.ResultVersion,
                transactionId = transaction.TransactionId,
                transactionSequence = transaction.TransactionId is null
                    ? transaction.Sequence
                    : (long?)null,
            }));
        if (!string.Equals(transaction.EntryId, expectedEntryId, StringComparison.Ordinal))
        {
            findings.Add("transaction_entry_id_mismatch");
            valid = false;
        }

        for (int i = 0; i < transaction.Operations.Count; i++)
        {
            var operation = transaction.Operations[i];
            if (operation.Index != i)
            {
                findings.Add("operation_order_mismatch");
                valid = false;
            }
            if (!Enum.IsDefined(operation.ExecutionStatus)
                || !OperationShapeMatchesStatus(operation))
            {
                findings.Add("invalid_operation_execution");
                valid = false;
            }
            if (operation.ExecutionStatus != DeliveryOperationExecutionStatus.NotExecuted
                && operation.Success != operation.Results.All(result => result.Success))
            {
                findings.Add("operation_success_mismatch");
                valid = false;
            }
            try
            {
                DeliveryReceiptValidation.RequireNonBlank(operation.Tool, "operation tool", 256);
                DeliveryReceiptValidation.RequireNonBlank(operation.Action, "operation action", 256);
                DeliveryReceiptValidation.ValidateDigest(
                    operation.ArgumentsDigest, "operation arguments digest");
            }
            catch (DeliveryReceiptValidationException ex)
            {
                findings.Add(ex.Code);
                valid = false;
            }
            bool requiresFullEvidence =
                privacyProfile == DeliveryReceiptPrivacyProfile.FullEvidence;
            if ((privacyProfile == DeliveryReceiptPrivacyProfile.HashOnly
                    && operation.ArgumentsSummary is not null)
                || (privacyProfile != DeliveryReceiptPrivacyProfile.HashOnly
                    && operation.ArgumentsSummary is null)
                || (!requiresFullEvidence && operation.Arguments is not null)
                || (requiresFullEvidence && operation.Arguments is null))
            {
                findings.Add("privacy_profile_violation");
                valid = false;
            }
            if (operation.Arguments is { } arguments)
            {
                if (arguments.ValueKind != JsonValueKind.Object
                    || !DeliveryReceiptValidation.DigestEquals(
                        operation.ArgumentsDigest,
                        DeliveryReceiptCanonicalJson.Digest(
                            DeliveryReceiptCanonicalJson.SerializeCanonical(arguments))))
                {
                    findings.Add("operation_arguments_digest_mismatch");
                    valid = false;
                }
            }
            foreach (var result in operation.Results)
            {
                try
                {
                    DeliveryReceiptValidation.ValidateDigest(
                        result.ResultDigest, "operation result digest");
                }
                catch (DeliveryReceiptValidationException ex)
                {
                    findings.Add(ex.Code);
                    valid = false;
                }
                if (result.ErrorMessage is not null)
                    valid &= ValidateTextEvidence(result.ErrorMessage, privacyProfile, findings);
                bool errorShapeValid = result.Success
                    ? result.ErrorCode is null && result.ErrorMessage is null
                    : result.ErrorCode is not null
                        && Enum.TryParse<EditErrorCode>(
                            result.ErrorCode, ignoreCase: false, out var errorCode)
                        && Enum.IsDefined(errorCode)
                        && result.ErrorMessage is not null;
                if (!errorShapeValid)
                {
                    findings.Add("invalid_operation_result");
                    valid = false;
                }
                if ((!requiresFullEvidence && result.FullResult is not null)
                    || (requiresFullEvidence && result.FullResult is null))
                {
                    findings.Add("privacy_profile_violation");
                    valid = false;
                }
                if (result.FullResult is { } fullResult
                    && (!DeliveryReceiptValidation.DigestEquals(
                            result.ResultDigest,
                            DeliveryReceiptCanonicalJson.Digest(
                                DeliveryReceiptCanonicalJson.SerializeCanonical(fullResult)))
                        || fullResult.ValueKind != JsonValueKind.Object
                        || !fullResult.TryGetProperty("success", out var fullSuccess)
                        || fullSuccess.ValueKind is not
                            (JsonValueKind.True or JsonValueKind.False)
                        || fullSuccess.GetBoolean() != result.Success))
                {
                    findings.Add("operation_result_digest_mismatch");
                    valid = false;
                }
                foreach (var change in result.ObjectChanges)
                {
                    if (!Enum.IsDefined(change.ChangeKind))
                    {
                        findings.Add("unknown_object_change_kind");
                        valid = false;
                    }
                    try
                    {
                        DeliveryReceiptValidation.RequireNonBlank(
                            change.AnchorId, "changed anchor id", 4096);
                        DeliveryReceiptValidation.RequireNonBlank(change.Kind, "anchor kind", 256);
                        DeliveryReceiptValidation.RequireNonBlank(change.Scope, "anchor scope", 1024);
                        DeliveryReceiptValidation.RequireNonBlank(change.Unid, "anchor unid", 2048);
                    }
                    catch (DeliveryReceiptValidationException ex)
                    {
                        findings.Add(ex.Code);
                        valid = false;
                    }
                }
                if (!IsStrictlyOrdered(
                        result.ObjectChanges,
                        CompareObjectChanges))
                {
                    findings.Add("object_change_order_mismatch");
                    valid = false;
                }
            }
        }

        foreach (var change in transaction.AuthoredChanges)
        {
            if (!Enum.IsDefined(change.EntityKind) || !Enum.IsDefined(change.ChangeKind))
            {
                findings.Add("unknown_authored_change_enum");
                valid = false;
            }
            try
            {
                DeliveryReceiptValidation.RequireNonBlank(
                    change.EntityId, "authored entity id", 2048);
                DeliveryReceiptValidation.ValidateDigest(
                    change.SourceDigest, "authored evidence digest");
            }
            catch (DeliveryReceiptValidationException ex)
            {
                findings.Add(ex.Code);
                valid = false;
            }
            if (change.Text is not null)
                valid &= ValidateTextEvidence(change.Text, privacyProfile, findings);
            bool requiresFullEvidence =
                privacyProfile == DeliveryReceiptPrivacyProfile.FullEvidence;
            if ((!requiresFullEvidence && change.FullEvidence is not null)
                || (requiresFullEvidence && change.FullEvidence is null))
            {
                findings.Add("privacy_profile_violation");
                valid = false;
            }
            if (change.FullEvidence is { } fullEvidence
                && !DeliveryReceiptValidation.DigestEquals(
                    change.SourceDigest,
                    DeliveryReceiptCanonicalJson.Digest(
                        DeliveryReceiptCanonicalJson.SerializeCanonical(fullEvidence))))
            {
                findings.Add("authored_evidence_digest_mismatch");
                valid = false;
            }
            if (!IsStrictlyOrdered(
                    change.AffectedAnchorIds,
                    static (left, right) => string.CompareOrdinal(left, right)))
            {
                findings.Add("affected_anchor_order_mismatch");
                valid = false;
            }
        }
        if (!IsStrictlyOrdered(transaction.AuthoredChanges, CompareAuthoredChanges))
        {
            findings.Add("authored_change_order_mismatch");
            valid = false;
        }
        foreach (var warning in transaction.Warnings)
            valid &= ValidateTextEvidence(warning, privacyProfile, findings);
        if (!IsStrictlyOrdered(transaction.Warnings, CompareTextEvidence))
        {
            findings.Add("transaction_warning_order_mismatch");
            valid = false;
        }
        return valid;
    }

    private static bool ValidateTransactionOutcome(
        DeliveryTransactionEntry transaction,
        ICollection<string> findings)
    {
        bool valid = transaction.Status switch
        {
            DeliveryTransactionStatus.Committed =>
                transaction.Operations.All(operation => operation.ExecutionStatus
                    == DeliveryOperationExecutionStatus.Succeeded),
            DeliveryTransactionStatus.PartiallyCommitted =>
                transaction.Mode == MutationBatchMode.BestEffort
                && transaction.Operations.Any(operation => operation.ExecutionStatus
                    == DeliveryOperationExecutionStatus.Succeeded)
                && transaction.Operations.Any(operation => operation.ExecutionStatus
                    == DeliveryOperationExecutionStatus.Failed)
                && transaction.Operations.All(operation => operation.ExecutionStatus is
                    DeliveryOperationExecutionStatus.Succeeded
                    or DeliveryOperationExecutionStatus.Failed),
            DeliveryTransactionStatus.Failed when transaction.Mode == MutationBatchMode.Atomic =>
                ValidFailedAtomicOperations(transaction.Operations),
            DeliveryTransactionStatus.Failed when transaction.Mode == MutationBatchMode.BestEffort =>
                transaction.Operations.Count > 0
                && transaction.Operations.All(operation => operation.ExecutionStatus
                    == DeliveryOperationExecutionStatus.Failed),
            DeliveryTransactionStatus.Prediction when transaction.Mode == MutationBatchMode.Atomic =>
                transaction.TransactionId is null
                && (transaction.Operations.All(operation => operation.ExecutionStatus
                        == DeliveryOperationExecutionStatus.Succeeded)
                    || ValidFailedAtomicOperations(transaction.Operations)),
            DeliveryTransactionStatus.Prediction when transaction.Mode == MutationBatchMode.BestEffort =>
                transaction.TransactionId is null
                && transaction.Operations.All(operation => operation.ExecutionStatus is
                    DeliveryOperationExecutionStatus.Succeeded
                    or DeliveryOperationExecutionStatus.Failed),
            _ => false,
        };
        if (!valid)
            findings.Add("transaction_outcome_mismatch");
        return valid;
    }

    private static bool ValidFailedAtomicOperations(
        IReadOnlyList<DeliveryOperationEvidence> operations)
    {
        var failures = operations
            .Where(operation => operation.ExecutionStatus
                == DeliveryOperationExecutionStatus.FailedRolledBack)
            .ToArray();
        if (failures.Length != 1)
            return false;
        int failedIndex = failures[0].Index;
        var preceding = operations.Take(failedIndex).ToArray();
        bool preflightFailure = preceding.All(operation => operation.ExecutionStatus
            == DeliveryOperationExecutionStatus.NotExecuted);
        bool executionFailure = preceding.All(operation => operation.ExecutionStatus
            == DeliveryOperationExecutionStatus.SucceededRolledBack);
        return (preflightFailure || executionFailure)
            && operations.Skip(failedIndex + 1).All(operation => operation.ExecutionStatus
                == DeliveryOperationExecutionStatus.NotExecuted);
    }

    private static bool OperationShapeMatchesStatus(DeliveryOperationEvidence operation) =>
        operation.ExecutionStatus switch
        {
            DeliveryOperationExecutionStatus.NotExecuted =>
                !operation.Success && !operation.RolledBack && operation.Results.Count == 0,
            DeliveryOperationExecutionStatus.Succeeded =>
                operation.Success && !operation.RolledBack,
            DeliveryOperationExecutionStatus.Failed =>
                !operation.Success && !operation.RolledBack && operation.Results.Count > 0,
            DeliveryOperationExecutionStatus.SucceededRolledBack =>
                operation.Success && operation.RolledBack,
            DeliveryOperationExecutionStatus.FailedRolledBack =>
                !operation.Success && operation.RolledBack && operation.Results.Count > 0,
            _ => false,
        };

    private static bool IsStrictlyOrdered<T>(
        IReadOnlyList<T> values,
        Comparison<T> comparison)
    {
        for (int i = 1; i < values.Count; i++)
        {
            if (comparison(values[i - 1], values[i]) >= 0)
                return false;
        }
        return true;
    }

    private static int ComparePackageChanges(
        DeliveryPackageChange left,
        DeliveryPackageChange right)
    {
        int comparison = left.Kind.CompareTo(right.Kind);
        if (comparison != 0) return comparison;
        comparison = string.CompareOrdinal(left.Location.EntryUri, right.Location.EntryUri);
        if (comparison != 0) return comparison;
        comparison = string.CompareOrdinal(left.Location.OwnerUri, right.Location.OwnerUri);
        if (comparison != 0) return comparison;
        comparison = string.CompareOrdinal(
            left.Location.RelationshipId, right.Location.RelationshipId);
        if (comparison != 0) return comparison;
        comparison = string.CompareOrdinal(
            left.Location.TargetUri, right.Location.TargetUri);
        if (comparison != 0) return comparison;
        return string.CompareOrdinal(
            left.Location.PropertyPath, right.Location.PropertyPath);
    }

    private static int CompareCitations(
        DeliveryPageCitation left,
        DeliveryPageCitation right)
    {
        int comparison = string.CompareOrdinal(left.AnchorId, right.AnchorId);
        if (comparison != 0) return comparison;
        comparison = string.CompareOrdinal(left.RenderArtifactId, right.RenderArtifactId);
        return comparison != 0
            ? comparison
            : string.CompareOrdinal(left.PageMapArtifactId, right.PageMapArtifactId);
    }

    private static int CompareEvidence(
        DeliveryEvidenceReference left,
        DeliveryEvidenceReference right)
    {
        int comparison = left.Kind.CompareTo(right.Kind);
        if (comparison != 0) return comparison;
        comparison = string.CompareOrdinal(left.Schema, right.Schema);
        return comparison != 0
            ? comparison
            : string.CompareOrdinal(left.Digest.Value, right.Digest.Value);
    }

    private static int CompareTextEvidence(
        DeliveryTextEvidence left,
        DeliveryTextEvidence right) =>
        string.CompareOrdinal(left.Digest.Value, right.Digest.Value);

    private static int CompareAuthoredChanges(
        DeliveryAuthoredChange left,
        DeliveryAuthoredChange right)
    {
        int comparison = left.EntityKind.CompareTo(right.EntityKind);
        if (comparison != 0) return comparison;
        comparison = left.ChangeKind.CompareTo(right.ChangeKind);
        if (comparison != 0) return comparison;
        comparison = string.CompareOrdinal(left.EntityId, right.EntityId);
        if (comparison != 0) return comparison;
        comparison = string.CompareOrdinal(left.PartUri, right.PartUri);
        if (comparison != 0) return comparison;
        comparison = string.CompareOrdinal(left.Scope, right.Scope);
        return comparison != 0
            ? comparison
            : string.CompareOrdinal(left.SourceDigest.Value, right.SourceDigest.Value);
    }

    private static int CompareObjectChanges(
        DeliveryObjectChange left,
        DeliveryObjectChange right)
    {
        int comparison = left.ChangeKind.CompareTo(right.ChangeKind);
        return comparison != 0
            ? comparison
            : string.CompareOrdinal(left.AnchorId, right.AnchorId);
    }

    private static bool ValidatePackageChange(
        DeliveryPackageChange change,
        DeliveryReceiptPrivacyProfile privacyProfile,
        ICollection<string> findings)
    {
        bool valid = true;
        if (!Enum.IsDefined(change.Kind) || !Enum.IsDefined(change.Disposition))
        {
            findings.Add("unknown_package_change_enum");
            valid = false;
        }
        try
        {
            DeliveryTransactionContribution.ValidateFingerprint(change.ChangeId);
            if (change.Kind is DeliveryPackageChangeKind.PartAdded
                or DeliveryPackageChangeKind.PartRemoved
                or DeliveryPackageChangeKind.PartModified)
            {
                DeliveryReceiptValidation.RequireNonBlank(
                    change.Location.EntryUri, "changed entry URI", 4096);
            }
            else
            {
                DeliveryReceiptValidation.RequireNonBlank(
                    change.Location.OwnerUri, "relationship owner URI", 4096);
                DeliveryReceiptValidation.RequireNonBlank(
                    change.Location.RelationshipId, "relationship id", 2048);
            }
            if (change.Disposition == DeliveryChangeDisposition.Derived)
                DeliveryReceiptValidation.RequireNonBlank(change.Derivation, "derivation", 4096);
        }
        catch (DeliveryReceiptValidationException ex)
        {
            findings.Add(ex.Code);
            valid = false;
        }
        bool evidenceShapeValid = change.Kind switch
        {
            DeliveryPackageChangeKind.PartAdded
                or DeliveryPackageChangeKind.RelationshipAdded =>
                change.Before is null && change.After is not null,
            DeliveryPackageChangeKind.PartRemoved
                or DeliveryPackageChangeKind.RelationshipRemoved =>
                change.Before is not null && change.After is null,
            DeliveryPackageChangeKind.PartModified
                or DeliveryPackageChangeKind.RelationshipModified =>
                change.Before is not null && change.After is not null,
            _ => false,
        };
        if (!evidenceShapeValid)
        {
            findings.Add("package_change_evidence_mismatch");
            valid = false;
        }
        if (change.Before is not null)
            valid &= ValidateTextEvidence(change.Before, privacyProfile, findings);
        if (change.After is not null)
            valid &= ValidateTextEvidence(change.After, privacyProfile, findings);
        var expectedChangeId = DeliveryReceiptCanonicalJson.DigestToken(
            DeliveryReceiptCanonicalJson.SerializeCanonical(new
            {
                after = change.After?.Digest,
                before = change.Before?.Digest,
                kind = change.Kind.ToString(),
                location = change.Location,
            }));
        if (!string.Equals(change.ChangeId, expectedChangeId, StringComparison.Ordinal))
        {
            findings.Add("package_change_id_mismatch");
            valid = false;
        }
        return valid;
    }

    private static bool ValidateArtifactRecord(
        DeliveryArtifact artifact,
        ICollection<string> findings)
    {
        bool valid = true;
        if (!Enum.IsDefined(artifact.Role) || !Enum.IsDefined(artifact.Availability))
        {
            findings.Add("unknown_artifact_enum");
            valid = false;
        }
        try
        {
            DeliveryReceiptValidation.RequireNonBlank(artifact.ArtifactId, "artifact id", 256);
            DeliveryReceiptValidation.RequireNonBlank(
                artifact.MediaType, "artifact media type", 512);
            DeliveryReceiptValidation.NormalizeRelativePath(artifact.RelativePath);
            DeliveryReceiptValidation.ValidateOptionalDigest(
                artifact.PackageDigest, "artifact package digest");
            DeliveryReceiptValidation.ValidateOptionalDigest(
                artifact.PageMapDigest, "artifact page-map digest");
            if (artifact.RendererFingerprint is not null)
            {
                DeliveryReceiptValidation.RequireNonBlank(
                    artifact.RendererFingerprint, "renderer fingerprint", 4096);
            }
            if (artifact.DocumentVersion is < 0)
            {
                throw new DeliveryReceiptValidationException(
                    "invalid_document_version", "Artifact document version cannot be negative.");
            }
        }
        catch (DeliveryReceiptValidationException ex)
        {
            findings.Add(ex.Code);
            valid = false;
        }

        if (artifact.Availability == DeliveryArtifactAvailability.Available)
        {
            try { DeliveryReceiptValidation.ValidateDigest(artifact.Digest, "artifact digest"); }
            catch (DeliveryReceiptValidationException ex)
            {
                findings.Add(ex.Code);
                valid = false;
            }
            if (artifact.ByteLength is null or < 0 || artifact.UnavailableReason is not null)
            {
                findings.Add("invalid_artifact_record");
                valid = false;
            }
            if (artifact.Role is (DeliveryArtifactRole.CleanDocx
                    or DeliveryArtifactRole.ReviewDocx)
                && artifact.PackageDigest is not null
                && !DeliveryReceiptValidation.DigestEquals(
                    artifact.Digest, artifact.PackageDigest))
            {
                findings.Add("docx_artifact_identity_mismatch");
                valid = false;
            }
        }
        else if (artifact.Availability == DeliveryArtifactAvailability.Unavailable)
        {
            if (artifact.Digest is not null || artifact.ByteLength is not null
                || string.IsNullOrWhiteSpace(artifact.UnavailableReason))
            {
                findings.Add("invalid_artifact_record");
                valid = false;
            }
        }
        return valid;
    }

    private static bool ValidateRequiredCleanDocx(
        DeliveryChangeReceiptPayload payload,
        IReadOnlyDictionary<string, byte[]> artifactBytes,
        DeliveryReceiptLimits limits,
        ICollection<string> findings)
    {
        var cleanArtifacts = payload.Artifacts
            .Where(artifact => artifact.Role == DeliveryArtifactRole.CleanDocx)
            .ToArray();
        if (cleanArtifacts.Length != 1)
        {
            findings.Add(cleanArtifacts.Length == 0
                ? "missing_clean_docx"
                : "multiple_clean_docx_artifacts");
            return false;
        }

        var clean = cleanArtifacts[0];
        if (clean.Availability != DeliveryArtifactAvailability.Available
            || clean.Digest is null
            || clean.ByteLength is null
            || !artifactBytes.TryGetValue(clean.ArtifactId, out var cleanBytes)
            || clean.ByteLength != cleanBytes.LongLength
            || clean.DocumentVersion != payload.DeliveredDocument.DocumentVersion
            || !DeliveryReceiptValidation.DigestEquals(
                clean.PackageDigest, payload.DeliveredDocument.RawPackageBytesDigest)
            || !DeliveryReceiptValidation.DigestEquals(
                clean.Digest, payload.DeliveredDocument.RawPackageBytesDigest)
            || !DeliveryReceiptValidation.DigestEquals(
                clean.Digest, DeliveryReceiptCanonicalJson.Digest(cleanBytes)))
        {
            findings.Add("clean_docx_delivery_mismatch");
            return false;
        }

        try
        {
            var actualManifest = PackageManifestGenerator.Generate(
                cleanBytes, limits.CleanDocxManifestOptions);
            if (!actualManifest.IsValid
                || !string.Equals(actualManifest.PackageKind, "opc", StringComparison.Ordinal)
                || string.IsNullOrWhiteSpace(actualManifest.Facts.MainDocumentUri)
                || !DeliveryReceiptLineageValidator.DocumentEquals(
                    DeliveryDocumentIdentity.FromManifest(
                        actualManifest, clean.DocumentVersion.Value),
                    payload.DeliveredDocument))
            {
                findings.Add("clean_docx_delivery_mismatch");
                return false;
            }
        }
        catch (DeliveryReceiptValidationException ex)
        {
            findings.Add(IsResourceLimitCode(ex.Code)
                ? ex.Code
                : "clean_docx_delivery_mismatch");
            return false;
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidDataException)
        {
            findings.Add("clean_docx_delivery_mismatch");
            return false;
        }
        return true;
    }

    private static bool ValidateSemanticBindings(
        DeliveryChangeReceiptPayload payload,
        IReadOnlyDictionary<string, DeliveryArtifact> artifactsById,
        IReadOnlyDictionary<string, byte[]> artifactBytes,
        DeliveryLineageValidationResult lineageValidation,
        DeliveryReceiptLimits limits,
        ICollection<string> findings)
    {
        bool valid = true;
        var aggregates = payload.SemanticChangeSets
            .Where(binding => binding.Scope
                == DeliverySemanticComparisonScope.SourceToDelivered)
            .ToArray();
        if (aggregates.Length != 1)
        {
            findings.Add("missing_source_to_delivered_semantic_evidence");
            valid = false;
        }

        var expectedTransactions = lineageValidation.StateChangingTransactions
            .GroupBy(transaction => transaction.EntryId, StringComparer.Ordinal)
            .Where(group => group.Count() == 1)
            .ToDictionary(group => group.Key, group => group.Single(), StringComparer.Ordinal);
        var transactionBindings = payload.SemanticChangeSets
            .Where(binding => binding.Scope == DeliverySemanticComparisonScope.Transaction
                && binding.TransactionEntryId is not null)
            .GroupBy(binding => binding.TransactionEntryId!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
        if (transactionBindings.Any(pair => pair.Value.Length != 1
                || !expectedTransactions.ContainsKey(pair.Key))
            || expectedTransactions.Keys.Any(entryId => !transactionBindings.ContainsKey(entryId)))
        {
            findings.Add("semantic_transaction_coverage_mismatch");
            valid = false;
        }

        foreach (var binding in payload.SemanticChangeSets)
        {
            if (!Enum.IsDefined(binding.Scope))
            {
                findings.Add("unknown_semantic_comparison_scope");
                valid = false;
                continue;
            }

            DeliveryDocumentIdentity? expectedBefore = null;
            DeliveryDocumentIdentity? expectedAfter = null;
            if (binding.Scope == DeliverySemanticComparisonScope.SourceToDelivered)
            {
                if (binding.TransactionEntryId is not null)
                {
                    findings.Add("semantic_binding_identity_mismatch");
                    valid = false;
                }
                expectedBefore = payload.SourceDocument;
                expectedAfter = payload.DeliveredDocument;
            }
            else if (binding.TransactionEntryId is null
                || !expectedTransactions.TryGetValue(
                    binding.TransactionEntryId, out var transaction))
            {
                findings.Add("semantic_binding_identity_mismatch");
                valid = false;
            }
            else
            {
                expectedBefore = transaction.BeforeDocument;
                expectedAfter = transaction.AfterDocument;
            }

            if (expectedBefore is not null
                && (!DeliveryReceiptLineageValidator.DocumentEquals(
                        binding.BeforeDocument, expectedBefore)
                    || !DeliveryReceiptLineageValidator.DocumentEquals(
                        binding.AfterDocument, expectedAfter!)))
            {
                findings.Add("semantic_binding_identity_mismatch");
                valid = false;
            }

            try
            {
                ValidateDocument(binding.BeforeDocument);
                ValidateDocument(binding.AfterDocument);
                DeliveryReceiptValidation.ValidateDigest(
                    binding.Digest, "semantic change-set digest");
                DeliveryReceiptValidation.RequireNonBlank(
                    binding.ArtifactId, "semantic artifact id", 256);
            }
            catch (DeliveryReceiptValidationException ex)
            {
                findings.Add(ex.Code);
                valid = false;
            }
            if (!string.Equals(binding.Schema, SemanticChangeSet.CurrentSchema,
                    StringComparison.Ordinal)
                || binding.SchemaVersion != SemanticChangeSet.CurrentSchemaVersion
                || binding.ChangeCount < 0)
            {
                findings.Add("unsupported_semantic_change_set");
                valid = false;
            }

            byte[]? bytes = null;
            bool artifactBindingValid =
                artifactsById.TryGetValue(binding.ArtifactId, out var artifact)
                && artifact.Role == DeliveryArtifactRole.SemanticDiff
                && artifact.Availability == DeliveryArtifactAvailability.Available
                && artifact.Digest is not null
                && artifact.DocumentVersion is null
                && artifact.PackageDigest is null
                && artifact.RendererFingerprint is null
                && artifact.PageMapDigest is null
                && DeliveryReceiptValidation.DigestEquals(artifact.Digest, binding.Digest)
                && artifactBytes.TryGetValue(binding.ArtifactId, out bytes);
            if (artifactBindingValid)
            {
                try
                {
                    DeliveryReceiptResourceBudget.Bytes(
                        bytes!.LongLength,
                        limits.MaxSemanticEvidenceBytes,
                        "semantic_resource_limit",
                        "Semantic artifact");
                    var projection = DeliverySemanticChangeSetAdapter.InspectExact(bytes, limits);
                    artifactBindingValid =
                        DeliveryReceiptCanonicalJson.FixedTimeEquals(binding.Digest, bytes)
                        && string.Equals(
                            projection.Schema, binding.Schema, StringComparison.Ordinal)
                        && projection.SchemaVersion == binding.SchemaVersion
                        && projection.ChangeCount == binding.ChangeCount
                        && DeliveryReceiptValidation.DigestEquals(
                            projection.Digest, binding.Digest);
                }
                catch (DeliveryReceiptValidationException ex)
                {
                    findings.Add(ex.Code);
                    artifactBindingValid = false;
                }
                catch (Exception ex) when (ex is JsonException or FormatException
                    or ArgumentException or InvalidOperationException)
                {
                    findings.Add("invalid_semantic_change_set");
                    artifactBindingValid = false;
                }
            }
            if (!artifactBindingValid)
            {
                findings.Add($"semantic_artifact_binding_mismatch:{binding.ArtifactId}");
                valid = false;
            }
        }

        var expectedOrder = new List<(DeliverySemanticComparisonScope Scope, string? EntryId)>
        {
            (DeliverySemanticComparisonScope.SourceToDelivered, null),
        };
        expectedOrder.AddRange(lineageValidation.StateChangingTransactions.Select(transaction =>
            (DeliverySemanticComparisonScope.Transaction, (string?)transaction.EntryId)));
        var actualOrder = payload.SemanticChangeSets
            .Select(binding => (binding.Scope, binding.TransactionEntryId))
            .ToArray();
        if (!actualOrder.SequenceEqual(expectedOrder))
        {
            findings.Add("semantic_binding_order_mismatch");
            valid = false;
        }
        return valid;
    }

    private static bool ValidateTextEvidence(
        DeliveryTextEvidence evidence,
        DeliveryReceiptPrivacyProfile privacyProfile,
        ICollection<string> findings)
    {
        bool valid = true;
        try { DeliveryReceiptValidation.ValidateDigest(evidence.Digest, "text digest"); }
        catch (DeliveryReceiptValidationException ex)
        {
            findings.Add(ex.Code);
            valid = false;
        }
        if (evidence.CharacterCount < 0)
        {
            findings.Add("invalid_text_character_count");
            valid = false;
        }
        if (evidence.Value is { } value
            && (value.Length != evidence.CharacterCount
                || !DeliveryReceiptValidation.DigestEquals(
                    evidence.Digest,
                    DeliveryReceiptCanonicalJson.Digest(Encoding.UTF8.GetBytes(value)))))
        {
            findings.Add("text_evidence_digest_mismatch");
            valid = false;
        }
        bool fullEvidence = privacyProfile == DeliveryReceiptPrivacyProfile.FullEvidence;
        if ((privacyProfile == DeliveryReceiptPrivacyProfile.HashOnly
                && evidence.Summary is not null)
            || (privacyProfile != DeliveryReceiptPrivacyProfile.HashOnly
                && evidence.Summary is null)
            || (fullEvidence && evidence.Value is null)
            || (!fullEvidence && evidence.Value is not null))
        {
            findings.Add("privacy_profile_violation");
            valid = false;
        }
        return valid;
    }

    private static IReadOnlyList<DeliveryArtifactVerification> VerifyArtifacts(
        IReadOnlyList<DeliveryArtifact> artifacts,
        IReadOnlyDictionary<string, byte[]> artifactBytes,
        DeliveryReceiptLimits limits,
        ICollection<string> findings)
    {
        var results = new List<DeliveryArtifactVerification>();
        foreach (var artifact in artifacts.OrderBy(value => value.ArtifactId, StringComparer.Ordinal))
        {
            if (artifact.Availability == DeliveryArtifactAvailability.Unavailable)
            {
                bool recordValid = artifact.Digest is null && artifact.ByteLength is null
                    && artifact.UnavailableReason is not null;
                results.Add(new DeliveryArtifactVerification
                {
                    ArtifactId = artifact.ArtifactId,
                    Status = recordValid
                        ? DeliveryArtifactVerificationStatus.Unavailable
                        : DeliveryArtifactVerificationStatus.InvalidRecord,
                });
                if (!recordValid)
                    findings.Add($"invalid_artifact_record:{artifact.ArtifactId}");
                continue;
            }
            if (artifact.Availability != DeliveryArtifactAvailability.Available
                || artifact.Digest is null || artifact.ByteLength is null
                || artifact.ByteLength < 0)
            {
                results.Add(new DeliveryArtifactVerification
                {
                    ArtifactId = artifact.ArtifactId,
                    Status = DeliveryArtifactVerificationStatus.InvalidRecord,
                });
                findings.Add($"invalid_artifact_record:{artifact.ArtifactId}");
                continue;
            }
            try { DeliveryReceiptValidation.ValidateDigest(artifact.Digest, "artifact digest"); }
            catch (DeliveryReceiptValidationException)
            {
                results.Add(new DeliveryArtifactVerification
                {
                    ArtifactId = artifact.ArtifactId,
                    Status = DeliveryArtifactVerificationStatus.InvalidRecord,
                    ExpectedLength = artifact.ByteLength,
                    ExpectedDigest = artifact.Digest,
                });
                findings.Add($"invalid_artifact_record:{artifact.ArtifactId}");
                continue;
            }
            if (!artifactBytes.TryGetValue(artifact.ArtifactId, out var bytes))
            {
                results.Add(new DeliveryArtifactVerification
                {
                    ArtifactId = artifact.ArtifactId,
                    Status = DeliveryArtifactVerificationStatus.Missing,
                    ExpectedLength = artifact.ByteLength,
                    ExpectedDigest = artifact.Digest,
                });
                findings.Add($"artifact_missing:{artifact.ArtifactId}");
                continue;
            }

            var roleLimit = artifact.Role switch
            {
                DeliveryArtifactRole.SemanticDiff => limits.MaxSemanticEvidenceBytes,
                DeliveryArtifactRole.PageMap => limits.MaxPageMapBytes,
                _ => limits.MaxArtifactBytes,
            };
            if (bytes.LongLength > roleLimit)
            {
                var finding = artifact.Role switch
                {
                    DeliveryArtifactRole.SemanticDiff => "semantic_resource_limit",
                    DeliveryArtifactRole.PageMap => "page_map_resource_limit",
                    _ => "artifact_resource_limit",
                };
                results.Add(new DeliveryArtifactVerification
                {
                    ArtifactId = artifact.ArtifactId,
                    Status = DeliveryArtifactVerificationStatus.InvalidRecord,
                    ExpectedLength = artifact.ByteLength,
                    ActualLength = bytes.LongLength,
                    ExpectedDigest = artifact.Digest,
                });
                findings.Add(finding);
                continue;
            }

            var actualDigest = DeliveryReceiptCanonicalJson.Digest(bytes);
            var status = bytes.LongLength != artifact.ByteLength
                ? DeliveryArtifactVerificationStatus.LengthMismatch
                : !DeliveryReceiptValidation.DigestEquals(artifact.Digest, actualDigest)
                    ? DeliveryArtifactVerificationStatus.DigestMismatch
                    : DeliveryArtifactVerificationStatus.Verified;
            results.Add(new DeliveryArtifactVerification
            {
                ArtifactId = artifact.ArtifactId,
                Status = status,
                ExpectedLength = artifact.ByteLength,
                ActualLength = bytes.LongLength,
                ExpectedDigest = artifact.Digest,
                ActualDigest = actualDigest,
            });
            if (status != DeliveryArtifactVerificationStatus.Verified)
                findings.Add($"artifact_{status.ToString().ToLowerInvariant()}:{artifact.ArtifactId}");
        }
        return results;
    }

    private static bool ValidateCitationBindings(
        DeliveryChangeReceiptPayload payload,
        IReadOnlyDictionary<string, byte[]> artifactBytes,
        DeliveryLineageValidationResult lineageValidation,
        DeliveryReceiptLimits limits,
        ICollection<string> findings)
    {
        bool valid = true;
        var artifacts = payload.Artifacts
            .GroupBy(artifact => artifact.ArtifactId, StringComparer.Ordinal)
            .Where(group => group.Count() == 1)
            .ToDictionary(group => group.Key, group => group.Single(), StringComparer.Ordinal);
        foreach (var citation in payload.PageCitations)
        {
            try
            {
                DeliveryReceiptValidation.RequireNonBlank(
                    citation.AnchorId, "citation anchor id", 4096);
                DeliveryReceiptValidation.RequireNonBlank(
                    citation.Scope, "citation scope", 1024);
                DeliveryReceiptValidation.RequireNonBlank(
                    citation.RendererFingerprint, "citation renderer fingerprint", 4096);
                DeliveryReceiptValidation.RequireNonBlank(
                    citation.RenderArtifactId, "citation artifact id", 256);
                DeliveryReceiptValidation.RequireNonBlank(
                    citation.PageMapArtifactId, "citation page-map artifact id", 256);
                DeliveryReceiptValidation.ValidateDigest(
                    citation.PackageDigest, "citation package digest");
                DeliveryReceiptValidation.ValidateDigest(
                    citation.PageMapDigest, "citation page-map digest");
                DeliveryReceiptValidation.ValidateDigest(
                    citation.RenderArtifactDigest, "citation render-artifact digest");
                if (citation.DocumentVersion < 0
                    || !string.Equals(
                        ScopeFromAnchor(citation.AnchorId), citation.Scope,
                        StringComparison.Ordinal))
                {
                    throw new DeliveryReceiptValidationException(
                        "invalid_citation_identity", "Citation identity is invalid.");
                }
            }
            catch (DeliveryReceiptValidationException ex)
            {
                findings.Add($"{ex.Code}:{citation.AnchorId}");
                valid = false;
            }
            if (!DeliveryReceiptLineageValidator.IsReachable(
                    lineageValidation, citation.DocumentVersion, citation.PackageDigest))
            {
                findings.Add($"unreachable_citation_document:{citation.AnchorId}");
                valid = false;
            }
            if (!artifacts.TryGetValue(citation.RenderArtifactId, out var artifact)
                || artifact.Availability != DeliveryArtifactAvailability.Available
                || artifact.Digest is null
                || artifact.Role is not (DeliveryArtifactRole.Pdf
                    or DeliveryArtifactRole.PageImage
                    or DeliveryArtifactRole.RenderReport)
                || artifact.DocumentVersion != citation.DocumentVersion
                || !string.Equals(artifact.RendererFingerprint,
                    citation.RendererFingerprint, StringComparison.Ordinal)
                || !DeliveryReceiptValidation.DigestEquals(
                    artifact.PackageDigest, citation.PackageDigest)
                || !DeliveryReceiptValidation.DigestEquals(
                    artifact.PageMapDigest, citation.PageMapDigest)
                || !DeliveryReceiptValidation.DigestEquals(
                    artifact.Digest, citation.RenderArtifactDigest))
            {
                findings.Add($"citation_binding_mismatch:{citation.AnchorId}");
                valid = false;
            }
            if (!artifacts.TryGetValue(citation.PageMapArtifactId, out var pageMapArtifact)
                || pageMapArtifact.Availability != DeliveryArtifactAvailability.Available
                || pageMapArtifact.Role != DeliveryArtifactRole.PageMap
                || pageMapArtifact.Digest is null
                || pageMapArtifact.DocumentVersion != citation.DocumentVersion
                || !string.Equals(pageMapArtifact.RendererFingerprint,
                    citation.RendererFingerprint, StringComparison.Ordinal)
                || !DeliveryReceiptValidation.DigestEquals(
                    pageMapArtifact.PackageDigest, citation.PackageDigest)
                || !DeliveryReceiptValidation.DigestEquals(
                    pageMapArtifact.Digest, citation.PageMapDigest))
            {
                findings.Add($"citation_page_map_binding_mismatch:{citation.AnchorId}");
                valid = false;
            }

            if (!artifactBytes.TryGetValue(citation.PageMapArtifactId, out var pageMapBytes))
            {
                findings.Add($"citation_page_map_bytes_missing:{citation.AnchorId}");
                valid = false;
                continue;
            }
            try
            {
                // Strict JSON/duplicate-property gate only; the map digest remains over raw bytes.
                _ = DeliveryReceiptCanonicalJson.CanonicalizeBounded(
                    pageMapBytes, limits, limits.MaxPageMapBytes,
                    "page_map_resource_limit");
                var pageMap = DocxSessionJson.ParsePageMap(Encoding.UTF8.GetString(pageMapBytes));
                var mapValidation = PageMapContract.ValidatePortable(
                    pageMap, citation.DocumentVersion, citation.RendererFingerprint);
                if (!mapValidation.Success)
                {
                    findings.Add($"invalid_page_map_artifact:{citation.AnchorId}");
                    valid = false;
                    continue;
                }
                var projected = PageMapContract.ProjectCitation(
                    pageMap,
                    citation.AnchorId,
                    new PageCitationRequest(
                        citation.DocumentVersion, citation.RendererFingerprint));
                if (projected.Availability != PageMapAvailability.Available
                    || !CitationCoordinatesEqual(citation, projected)
                    || projected.Fragments.Any(fragment =>
                        !PageMapContract.StoryMatchesScope(fragment.Story, citation.Scope)))
                {
                    findings.Add($"citation_page_map_projection_mismatch:{citation.AnchorId}");
                    valid = false;
                }
            }
            catch (DeliveryReceiptValidationException ex)
            {
                findings.Add(ex.Code);
                valid = false;
            }
            catch (Exception ex) when (ex is JsonException or FormatException)
            {
                findings.Add($"invalid_page_map_artifact:{citation.AnchorId}");
                valid = false;
            }
        }
        return valid;
    }

    private static bool CitationCoordinatesEqual(
        DeliveryPageCitation receiptCitation,
        PageCitation projected) =>
        string.Equals(receiptCitation.AnchorId, projected.AnchorId, StringComparison.Ordinal)
        && receiptCitation.DocumentVersion == projected.DocumentVersion
        && string.Equals(receiptCitation.RendererFingerprint,
            projected.RendererFingerprint, StringComparison.Ordinal)
        && receiptCitation.Pages.SequenceEqual(projected.Pages)
        && receiptCitation.Fragments.SequenceEqual(projected.Fragments);

    private static void ValidateDocument(DeliveryDocumentIdentity document)
    {
        ArgumentNullException.ThrowIfNull(document);
        if (document.DocumentVersion < 0)
            throw new DeliveryReceiptValidationException(
                "invalid_document_version", "Document version cannot be negative.");
        if (!DeliveryPackageManifestAdapter.IsSupportedSchema(
                document.PackageManifestSchema))
        {
            throw new DeliveryReceiptValidationException(
                "unsupported_package_manifest", "Unsupported package-manifest schema.");
        }
        DeliveryReceiptValidation.RequireNonBlank(
            document.PackageKind, "document package kind", 256);
        DeliveryReceiptValidation.ValidateDigest(
            document.RawPackageBytesDigest, "document package digest");
        DeliveryReceiptValidation.ValidateOptionalDigest(
            document.OrderedOpcContentDigest, "document content digest");
        DeliveryReceiptValidation.ValidateOptionalDigest(
            document.NormalizedSemanticDigest, "document semantic digest");
    }

    private static string ScopeFromAnchor(string anchorId)
    {
        var first = anchorId.IndexOf(':');
        var second = first < 0 ? -1 : anchorId.IndexOf(':', first + 1);
        if (first <= 0 || second <= first + 1 || second == anchorId.Length - 1)
        {
            throw new DeliveryReceiptValidationException(
                "invalid_anchor_id", "Citation anchor is not a canonical kind:scope:unid value.");
        }
        return anchorId[(first + 1)..second];
    }

    private static bool IsMalformedReceiptException(Exception exception) =>
        exception is JsonException
            or NotSupportedException
            or FormatException
            or InvalidOperationException
            or ArgumentException
            or NullReferenceException
            or KeyNotFoundException
            or OverflowException;

    private static bool IsResourceLimitCode(string code) => code is
        "receipt_resource_limit"
        or "semantic_resource_limit"
        or "page_map_resource_limit"
        or "artifact_resource_limit";

    private static VerificationDigest ReadDigest(JsonElement element)
    {
        int propertyCount = 0;
        bool hasAlgorithm = false;
        bool hasValue = false;
        foreach (var property in element.EnumerateObject())
        {
            propertyCount++;
            if (propertyCount > 2)
                break;
            hasAlgorithm |= string.Equals(
                property.Name, "algorithm", StringComparison.Ordinal);
            hasValue |= string.Equals(property.Name, "value", StringComparison.Ordinal);
        }
        if (propertyCount != 2
            || !hasAlgorithm
            || !hasValue
            || !element.TryGetProperty("algorithm", out var algorithm)
            || algorithm.ValueKind != JsonValueKind.String
            || !element.TryGetProperty("value", out var value)
            || value.ValueKind != JsonValueKind.String)
        {
            throw new DeliveryReceiptValidationException(
                "invalid_digest", "Receipt digest is malformed.");
        }
        var digest = new VerificationDigest
        {
            Algorithm = algorithm.GetString()!,
            Value = value.GetString()!,
        };
        DeliveryReceiptValidation.ValidateDigest(digest, "receipt digest");
        return digest;
    }

    private static DeliveryReceiptVerificationResult Malformed(string finding) => new()
    {
        IsValid = false,
        ReceiptDigestValid = false,
        ContractValid = false,
        CitationBindingsValid = false,
        Findings = new[] { finding },
    };

    private static DeliveryReceiptVerificationResult Rejected(string finding) => new()
    {
        IsValid = false,
        ReceiptDigestValid = false,
        ContractValid = false,
        CitationBindingsValid = false,
        Findings = new[] { finding },
    };
}
