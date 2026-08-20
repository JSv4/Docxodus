// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Globalization;
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
    private static readonly HashSet<string> EditErrorCodes = Enum
        .GetValues<EditErrorCode>()
        .Select(value => DocxSessionJson.EnumToSnake(value))
        .ToHashSet(StringComparer.Ordinal);

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
            if (!digestValid)
                return IntegrityFailure("receipt_digest_mismatch");
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
            if (!digestValid)
                return IntegrityFailure("receipt_digest_mismatch");
            var payload = payloadElement.Deserialize(
                DeliveryReceiptCanonicalJson.JsonContext.DeliveryChangeReceiptPayload);
            if (payload is null)
                return Malformed("missing_payload");
            DeliveryReceiptResourceValidator.ValidatePayload(payload, limits);
            var knownPayload = DeliveryChangeReceiptSerializer.SerializePayload(payload, limits);
            using var knownDocument = JsonDocument.Parse(knownPayload, new JsonDocumentOptions
        {
            MaxDepth = DeliveryReceiptLimits.MaxAllowedJsonDepth,
        });
            if (!DeliveryReceiptCanonicalJson.ContainsCanonicalKnownProjection(
                    payloadElement, knownDocument.RootElement))
            {
                return Malformed("noncanonical_known_payload");
            }
            if (payload.PrivacyProfile != DeliveryReceiptPrivacyProfile.FullEvidence
                && !DeliveryReceiptCanonicalJson.HasOnlyKnownProperties(
                    payloadElement, knownDocument.RootElement))
            {
                return Malformed("unknown_fields_violate_privacy_profile");
            }
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
        if (!ValidateHeader(payload, findings)
            || !ValidatePortableIntegers(payload, findings))
        {
            return new DeliveryReceiptVerificationResult
            {
                IsValid = false,
                ReceiptDigestValid = digestValid,
                ContractValid = false,
                CitationBindingsValid = false,
                Findings = findings.ToArray(),
            };
        }
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
        var verifiedArtifacts = artifactResults
            .Where(result => result.Status == DeliveryArtifactVerificationStatus.Verified)
            .Select(result => result.ArtifactId)
            .ToHashSet(StringComparer.Ordinal);
        bool citationsValid = ValidateCitationBindings(
            payload, artifactBytes, verifiedArtifacts, lineageValidation, limits, findings);
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

    private static bool ValidateHeader(
        DeliveryChangeReceiptPayload payload,
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
        return valid;
    }

    private static bool ValidatePortableIntegers(
        DeliveryChangeReceiptPayload payload,
        ICollection<string> findings)
    {
        bool valid = true;

        void Check(long value, string finding)
        {
            if (value < 0 || value > DeliveryReceiptValidation.MaxPortableInteger)
            {
                findings.Add(finding);
                valid = false;
            }
        }

        void CheckDocument(DeliveryDocumentIdentity document) =>
            Check(document.DocumentVersion, "invalid_document_version");

        CheckDocument(payload.SourceDocument);
        CheckDocument(payload.DeliveredDocument);
        foreach (var transaction in payload.Transactions)
        {
            Check(transaction.Sequence, "invalid_lineage_sequence");
            Check(transaction.BaseVersion, "invalid_document_version");
            Check(transaction.ResultVersion, "invalid_document_version");
            CheckDocument(transaction.BeforeDocument);
            CheckDocument(transaction.AfterDocument);
        }
        foreach (var lineageEvent in payload.Lineage)
        {
            Check(lineageEvent.Sequence, "invalid_lineage_sequence");
            CheckDocument(lineageEvent.BeforeDocument);
            CheckDocument(lineageEvent.AfterDocument);
        }
        foreach (var artifact in payload.Artifacts)
        {
            if (artifact.ByteLength is { } byteLength)
                Check(byteLength, "invalid_artifact_record");
            if (artifact.DocumentVersion is { } documentVersion)
                Check(documentVersion, "invalid_document_version");
        }
        foreach (var binding in payload.SemanticChangeSets)
        {
            CheckDocument(binding.BeforeDocument);
            CheckDocument(binding.AfterDocument);
        }
        foreach (var citation in payload.PageCitations)
            Check(citation.DocumentVersion, "invalid_citation_identity");

        return valid;
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
            if (lineageEvent.Sequence < 0
                || lineageEvent.Sequence > DeliveryReceiptValidation.MaxPortableInteger)
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
        var boundEvidenceArtifacts = payload.SemanticChangeSets
            .Select(binding => binding.ArtifactId)
            .Concat(payload.Evidence.Select(item => item.ArtifactId)
                .Where(artifactId => artifactId is not null)
                .Select(artifactId => artifactId!))
            .ToHashSet(StringComparer.Ordinal);
        foreach (var artifact in payload.Artifacts)
        {
            if (artifact.Role is DeliveryArtifactRole.SemanticDiff
                    or DeliveryArtifactRole.ValidationReport
                    or DeliveryArtifactRole.ReversibilityProof
                && !boundEvidenceArtifacts.Contains(artifact.ArtifactId))
            {
                findings.Add($"unbound_evidence_artifact:{artifact.ArtifactId}");
                valid = false;
            }
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
                || !lineageValidation.AppliedTransactionEntryIds.Contains(
                    change.TransactionEntryId)
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
        {
            valid &= ValidateArtifactRecord(artifact, payload.PrivacyProfile, findings);
            if (artifact.Role != DeliveryArtifactRole.ReviewDocx
                && artifact.DocumentVersion is { } documentVersion
                && artifact.PackageDigest is { } packageDigest
                && !DeliveryReceiptLineageValidator.IsArtifactDocumentReachable(
                    lineageValidation, documentVersion, packageDigest, payload.Artifacts))
            {
                findings.Add("unreachable_artifact_document");
                valid = false;
            }
        }

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
                var expectedSchema = ExpectedEvidenceSchema(evidence.Kind);
                if (!string.Equals(evidence.Schema, expectedSchema, StringComparison.Ordinal))
                {
                    findings.Add("evidence_schema_mismatch");
                    valid = false;
                }
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
                else if (!artifactBytes.TryGetValue(artifactId, out var evidenceBytes)
                    || !IsExactEvidence(evidence.Kind, evidenceBytes))
                {
                    findings.Add("invalid_evidence_artifact");
                    valid = false;
                }
            }
            if (payload.PrivacyProfile == DeliveryReceiptPrivacyProfile.HashOnly
                && evidence.Summary is not null)
            {
                findings.Add("privacy_profile_violation");
                valid = false;
            }
            else if (evidence.Summary is not null
                && !ValidateProfiledFreeText(
                    evidence.Summary, payload.PrivacyProfile, "evidence summary"))
            {
                findings.Add("privacy_profile_violation");
                valid = false;
            }
        }
        foreach (var warning in payload.Warnings)
            valid &= ValidateTextEvidence(warning, payload.PrivacyProfile, findings);
        return valid;
    }

    private static string ExpectedEvidenceSchema(DeliveryEvidenceKind kind) => kind switch
    {
        DeliveryEvidenceKind.ValidationResult => DeliverableVerificationResult.SchemaId,
        DeliveryEvidenceKind.RedlineReversibility => RedlineReversibilityProof.SchemaId,
        DeliveryEvidenceKind.SemanticChangeSet => string.Empty,
        _ => string.Empty,
    };

    private static bool IsExactEvidence(DeliveryEvidenceKind kind, ReadOnlySpan<byte> bytes) =>
        kind switch
        {
            DeliveryEvidenceKind.ValidationResult =>
                DeliverableVerificationResult.IsExactCanonical(bytes),
            DeliveryEvidenceKind.RedlineReversibility =>
                RedlineReversibilityProof.IsExactCanonical(bytes),
            _ => false,
        };

    private static bool ValidateTransaction(
        DeliveryTransactionEntry transaction,
        DeliveryReceiptPrivacyProfile privacyProfile,
        ICollection<string> findings)
    {
        bool valid = true;
        if (transaction.Sequence < 0
            || transaction.Sequence > DeliveryReceiptValidation.MaxPortableInteger)
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
            DeliveryReceiptValidation.ValidateOptionalDigest(
                transaction.ReportedPackageContentDigest,
                "reported package-content digest");
        }
        catch (DeliveryReceiptValidationException ex)
        {
            findings.Add(ex.Code);
            valid = false;
        }
        if (transaction.BaseVersion < 0
            || transaction.BaseVersion > DeliveryReceiptValidation.MaxPortableInteger
            || transaction.ResultVersion < 0
            || transaction.ResultVersion > DeliveryReceiptValidation.MaxPortableInteger
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
        var expectedEntryId = DeliveryReceiptIdentity.TransactionEntryId(
            transaction.RequestFingerprint,
            transaction.BeforeDocument,
            transaction.AfterDocument,
            transaction.BaseVersion,
            transaction.ResultVersion,
            transaction.TransactionId,
            transaction.Sequence);
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
                        && EditErrorCodes.Contains(result.ErrorCode)
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
                else if (result.FullResult is { } projectedResult
                    && !OperationResultProjectionMatches(result, projectedResult))
                {
                    findings.Add("operation_result_projection_mismatch");
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
                        if (!string.Equals(
                                change.AnchorId,
                                $"{change.Kind}:{change.Scope}:{change.Unid}",
                                StringComparison.Ordinal))
                        {
                            throw new DeliveryReceiptValidationException(
                                "invalid_anchor_id", "Changed anchor identity is inconsistent.");
                        }
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
            if (change.Diagnostic is { } diagnostic)
            {
                try
                {
                    DeliveryReceiptValidation.RequireNonBlank(
                        diagnostic.Code, "authored diagnostic code", 1024);
                }
                catch (DeliveryReceiptValidationException ex)
                {
                    findings.Add(ex.Code);
                    valid = false;
                }
                if (diagnostic.Message is null)
                {
                    findings.Add("missing_value");
                    valid = false;
                }
                else
                {
                    valid &= ValidateTextEvidence(
                        diagnostic.Message, privacyProfile, findings);
                }
            }
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
            else if (change.FullEvidence is { } projectedEvidence
                && !AuthoredEvidenceProjectionMatches(change, projectedEvidence))
            {
                findings.Add("authored_evidence_projection_mismatch");
                valid = false;
            }
            if (!IsStrictlyOrdered(
                    change.ConstituentIds,
                    static (left, right) => string.CompareOrdinal(left, right))
                || !IsStrictlyOrdered(
                    change.ConstituentKeys,
                    static (left, right) => string.CompareOrdinal(left, right)))
            {
                findings.Add("constituent_identity_order_mismatch");
                valid = false;
            }
            if (change.EntityKind == DeliveryAuthoredEntityKind.Revision)
            {
                if (change.ConstituentIds.Count == 0
                    || change.ConstituentKeys.Count == 0
                    || change.Family is not { } family
                    || !Enum.IsDefined(family)
                    || string.IsNullOrWhiteSpace(change.PartUri)
                    || string.IsNullOrWhiteSpace(change.Scope)
                    || change.ResolutionStatus is not { } resolutionStatus
                    || !Enum.IsDefined(resolutionStatus)
                    || (change.AnchorId is not null
                        && !change.AffectedAnchorIds.Contains(
                            change.AnchorId, StringComparer.Ordinal)))
                {
                    findings.Add("invalid_revision_identity");
                    valid = false;
                }
            }
            else if (change.ConstituentIds.Count != 0
                || change.ConstituentKeys.Count != 0
                || change.DateUtc is not null
                || change.Family is not null
                || change.AnchorId is not null
                || change.ResolutionStatus is not null
                || change.Diagnostic is not null)
            {
                findings.Add("invalid_authored_change_shape");
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
        if (transaction.AuthoredChanges.GroupBy(change => new
            {
                change.EntityKind,
                change.EntityId,
                change.PartUri,
                change.Scope,
            }).Any(group => group.Count() != 1))
        {
            findings.Add("duplicate_authored_change");
            valid = false;
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

    private static bool OperationResultProjectionMatches(
        DeliveryOperationResultEvidence receipt,
        JsonElement fullResult)
    {
        if (fullResult.ValueKind != JsonValueKind.Object
            || !fullResult.TryGetProperty("success", out var success)
            || success.ValueKind is not (JsonValueKind.True or JsonValueKind.False)
            || success.GetBoolean() != receipt.Success)
        {
            return false;
        }

        bool hasError = fullResult.TryGetProperty("error", out var error);
        if (receipt.Success)
        {
            if (hasError || receipt.ErrorCode is not null || receipt.ErrorMessage is not null)
                return false;
        }
        else if (!hasError
            || error.ValueKind != JsonValueKind.Object
            || !TryGetRequiredString(error, "code", out var errorCode)
            || !TryGetRequiredString(error, "message", out var errorMessage)
            || !string.Equals(receipt.ErrorCode, errorCode, StringComparison.Ordinal)
            || !string.Equals(receipt.ErrorMessage?.Value, errorMessage,
                StringComparison.Ordinal))
        {
            return false;
        }

        var changes = new List<DeliveryObjectChange>();
        if (!TryReadObjectChanges(
                fullResult, "created", DeliveryObjectChangeKind.Added, changes)
            || !TryReadObjectChanges(
                fullResult, "removed", DeliveryObjectChangeKind.Removed, changes)
            || !TryReadObjectChanges(
                fullResult, "modified", DeliveryObjectChangeKind.Modified, changes))
        {
            return false;
        }
        changes.Sort(CompareObjectChanges);
        return changes.SequenceEqual(receipt.ObjectChanges);
    }

    private static bool TryReadObjectChanges(
        JsonElement fullResult,
        string propertyName,
        DeliveryObjectChangeKind changeKind,
        ICollection<DeliveryObjectChange> destination)
    {
        if (!fullResult.TryGetProperty(propertyName, out var anchors)
            || anchors.ValueKind != JsonValueKind.Array)
        {
            return false;
        }
        foreach (var anchor in anchors.EnumerateArray())
        {
            if (anchor.ValueKind != JsonValueKind.Object
                || anchor.EnumerateObject().Count() != 4
                || !TryGetRequiredString(anchor, "id", out var id)
                || !TryGetRequiredString(anchor, "kind", out var kind)
                || !TryGetRequiredString(anchor, "scope", out var scope)
                || !TryGetRequiredString(anchor, "unid", out var unid)
                || !string.Equals(id, $"{kind}:{scope}:{unid}", StringComparison.Ordinal))
            {
                return false;
            }
            destination.Add(new DeliveryObjectChange
            {
                ChangeKind = changeKind,
                AnchorId = id,
                Kind = kind,
                Scope = scope,
                Unid = unid,
            });
        }
        return true;
    }

    private static bool AuthoredEvidenceProjectionMatches(
        DeliveryAuthoredChange receipt,
        JsonElement evidence)
    {
        if (evidence.ValueKind != JsonValueKind.Object)
            return false;
        return receipt.EntityKind switch
        {
            DeliveryAuthoredEntityKind.Revision =>
                RevisionProjectionMatches(receipt, evidence),
            DeliveryAuthoredEntityKind.Comment =>
                CommentProjectionMatches(receipt, evidence),
            DeliveryAuthoredEntityKind.Annotation =>
                AnnotationProjectionMatches(receipt, evidence),
            _ => false,
        };
    }

    private static bool RevisionProjectionMatches(
        DeliveryAuthoredChange receipt,
        JsonElement evidence)
    {
        if (!TryGetRequiredString(evidence, "id", out var id)
            || !TryGetRequiredString(evidence, "type", out var type)
            || !TryGetRequiredString(evidence, "family", out var family)
            || !TryGetRequiredString(evidence, "author", out var author)
            || !TryGetRequiredString(evidence, "text", out var text)
            || !TryGetRequiredString(evidence, "partUri", out var partUri)
            || !TryGetRequiredString(evidence, "scope", out var scope)
            || !TryGetOptionalString(evidence, "date", out var date)
            || !TryGetOptionalString(evidence, "dateUtc", out var dateUtc)
            || !TryGetOptionalString(evidence, "anchorId", out var anchorId)
            || !TryGetRequiredString(
                evidence, "resolutionStatus", out var resolutionStatus)
            || !TryParseOwnerEnum(family, out RevisionFamily parsedFamily)
            || !TryParseOwnerEnum(
                resolutionStatus, out RevisionResolutionStatus parsedResolutionStatus)
            || !TryGetStringArray(evidence, "constituentIds", out var constituentIds)
            || !TryGetStringArray(evidence, "constituentKeys", out var constituentKeys)
            || !TryGetAffectedAnchorIds(evidence, out var affectedAnchors))
        {
            return false;
        }
        DeliveryAuthoredDiagnostic? diagnostic = null;
        if (evidence.TryGetProperty("diagnostic", out var diagnosticElement))
        {
            if (diagnosticElement.ValueKind != JsonValueKind.Object
                || !TryGetRequiredString(diagnosticElement, "code", out var code)
                || !TryGetRequiredString(diagnosticElement, "message", out var message))
            {
                return false;
            }
            diagnostic = new DeliveryAuthoredDiagnostic
            {
                Code = code,
                Message = new DeliveryTextEvidence
                {
                    Digest = DeliveryReceiptCanonicalJson.Digest(
                        Encoding.UTF8.GetBytes(message)),
                    CharacterCount = message.Length,
                    Value = message,
                },
            };
        }
        return string.Equals(receipt.EntityId, id, StringComparison.Ordinal)
            && string.Equals(receipt.Type, type, StringComparison.Ordinal)
            && receipt.Family == parsedFamily
            && string.Equals(receipt.Author, author, StringComparison.Ordinal)
            && string.Equals(receipt.Date, date, StringComparison.Ordinal)
            && string.Equals(receipt.DateUtc, dateUtc, StringComparison.Ordinal)
            && string.Equals(receipt.PartUri, partUri, StringComparison.Ordinal)
            && string.Equals(receipt.Scope, scope, StringComparison.Ordinal)
            && string.Equals(receipt.AnchorId, anchorId, StringComparison.Ordinal)
            && receipt.ResolutionStatus == parsedResolutionStatus
            && AuthoredDiagnosticProjectionMatches(receipt.Diagnostic, diagnostic)
            && string.Equals(receipt.Text?.Value, text, StringComparison.Ordinal)
            && SortedDistinct(constituentIds).SequenceEqual(receipt.ConstituentIds)
            && SortedDistinct(constituentKeys).SequenceEqual(receipt.ConstituentKeys)
            && SortedDistinct(affectedAnchors).SequenceEqual(receipt.AffectedAnchorIds);
    }

    private static bool TryParseOwnerEnum<TEnum>(string wireValue, out TEnum value)
        where TEnum : struct, Enum
    {
        foreach (var candidate in Enum.GetValues<TEnum>())
        {
            if (string.Equals(
                    DocxSessionJson.EnumToSnake(candidate), wireValue,
                    StringComparison.Ordinal))
            {
                value = candidate;
                return true;
            }
        }
        value = default;
        return false;
    }

    private static bool AuthoredDiagnosticProjectionMatches(
        DeliveryAuthoredDiagnostic? receipt,
        DeliveryAuthoredDiagnostic? owner) =>
        receipt is null && owner is null
        || receipt is not null
            && owner is not null
            && string.Equals(receipt.Code, owner.Code, StringComparison.Ordinal)
            && receipt.Message is not null
            && owner.Message is not null
            && string.Equals(
                receipt.Message.Value, owner.Message.Value, StringComparison.Ordinal);

    private static bool CommentProjectionMatches(
        DeliveryAuthoredChange receipt,
        JsonElement evidence)
    {
        if (!TryGetRequiredString(evidence, "anchorId", out var anchorId)
            || !TryGetRequiredString(evidence, "author", out var author)
            || !TryGetRequiredString(evidence, "text", out var text)
            || !TryGetOptionalString(evidence, "date", out var date)
            || !TryGetOptionalString(evidence, "parentAnchorId", out var parentAnchorId))
        {
            return false;
        }
        string scope;
        try { scope = ScopeFromAnchor(anchorId); }
        catch (DeliveryReceiptValidationException) { return false; }
        var affected = parentAnchorId is null
            ? new[] { anchorId }
            : new[] { anchorId, parentAnchorId };
        return string.Equals(receipt.EntityId, anchorId, StringComparison.Ordinal)
            && string.Equals(receipt.Author, author, StringComparison.Ordinal)
            && string.Equals(receipt.Date, date, StringComparison.Ordinal)
            && receipt.DateUtc is null
            && string.Equals(receipt.Type,
                parentAnchorId is null ? "comment" : "commentReply",
                StringComparison.Ordinal)
            && receipt.PartUri is null
            && string.Equals(receipt.Scope, scope, StringComparison.Ordinal)
            && string.Equals(receipt.Text?.Value, text, StringComparison.Ordinal)
            && SortedDistinct(affected).SequenceEqual(receipt.AffectedAnchorIds)
            && receipt.ConstituentIds.Count == 0
            && receipt.ConstituentKeys.Count == 0;
    }

    private static bool AnnotationProjectionMatches(
        DeliveryAuthoredChange receipt,
        JsonElement evidence)
    {
        if (!TryGetRequiredString(evidence, "id", out var id)
            || !TryGetRequiredString(evidence, "labelId", out var labelId)
            || !TryGetOptionalString(evidence, "author", out var author)
            || !TryGetOptionalString(evidence, "created", out var created)
            || !TryGetOptionalString(evidence, "annotatedText", out var annotatedText))
        {
            return false;
        }
        return string.Equals(receipt.EntityId, id, StringComparison.Ordinal)
            && string.Equals(receipt.Author, author, StringComparison.Ordinal)
            && string.Equals(receipt.Date, created, StringComparison.Ordinal)
            && receipt.DateUtc is null
            && string.Equals(receipt.Type, labelId, StringComparison.Ordinal)
            && receipt.PartUri is null
            && receipt.Scope is null
            && string.Equals(receipt.Text?.Value, annotatedText ?? string.Empty,
                StringComparison.Ordinal)
            && receipt.AffectedAnchorIds.Count == 0
            && receipt.ConstituentIds.Count == 0
            && receipt.ConstituentKeys.Count == 0;
    }

    private static bool TryGetRequiredString(
        JsonElement element,
        string propertyName,
        out string value)
    {
        value = string.Empty;
        if (!element.TryGetProperty(propertyName, out var property)
            || property.ValueKind != JsonValueKind.String)
        {
            return false;
        }
        value = property.GetString()!;
        return true;
    }

    private static bool TryGetOptionalString(
        JsonElement element,
        string propertyName,
        out string? value)
    {
        value = null;
        if (!element.TryGetProperty(propertyName, out var property))
            return true;
        if (property.ValueKind != JsonValueKind.String)
            return false;
        value = property.GetString();
        return true;
    }

    private static bool TryGetStringArray(
        JsonElement element,
        string propertyName,
        out IReadOnlyList<string> values)
    {
        values = Array.Empty<string>();
        if (!element.TryGetProperty(propertyName, out var array)
            || array.ValueKind != JsonValueKind.Array)
        {
            return false;
        }
        var parsed = new List<string>();
        foreach (var item in array.EnumerateArray())
        {
            if (item.ValueKind != JsonValueKind.String)
                return false;
            parsed.Add(item.GetString()!);
        }
        values = parsed;
        return true;
    }

    private static bool TryGetAffectedAnchorIds(
        JsonElement evidence,
        out IReadOnlyList<string> values)
    {
        values = Array.Empty<string>();
        if (!evidence.TryGetProperty("affectedAnchors", out var array)
            || array.ValueKind != JsonValueKind.Array)
        {
            return false;
        }
        var parsed = new List<string>();
        foreach (var anchor in array.EnumerateArray())
        {
            if (anchor.ValueKind != JsonValueKind.Object
                || !TryGetRequiredString(anchor, "id", out var id))
            {
                return false;
            }
            parsed.Add(id);
        }
        values = parsed;
        return true;
    }

    private static IReadOnlyList<string> SortedDistinct(IEnumerable<string> values) => values
        .Distinct(StringComparer.Ordinal)
        .OrderBy(value => value, StringComparer.Ordinal)
        .ToArray();

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
        return string.CompareOrdinal(left.Scope, right.Scope);
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
            {
                DeliveryReceiptValidation.RequireNonBlank(change.Derivation, "derivation", 4096);
                if (!ValidateProfiledFreeText(
                        change.Derivation!, privacyProfile, "derivation"))
                {
                    throw new DeliveryReceiptValidationException(
                        "privacy_profile_violation", "Derivation text violates the privacy profile.");
                }
            }
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
        var expectedChangeId = DeliveryReceiptIdentity.PackageChangeId(
            change.Kind,
            change.Location,
            change.Before?.Digest,
            change.After?.Digest);
        if (!string.Equals(change.ChangeId, expectedChangeId, StringComparison.Ordinal))
        {
            findings.Add("package_change_id_mismatch");
            valid = false;
        }
        return valid;
    }

    private static bool ValidateArtifactRecord(
        DeliveryArtifact artifact,
        DeliveryReceiptPrivacyProfile privacyProfile,
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
            if (!string.Equals(
                    DeliveryReceiptValidation.NormalizeRelativePath(artifact.RelativePath),
                    artifact.RelativePath, StringComparison.Ordinal))
            {
                throw new DeliveryReceiptValidationException(
                    "unsafe_artifact_path",
                    "Artifact display paths must be stored in normalized form.");
            }
            DeliveryReceiptValidation.ValidateOptionalDigest(
                artifact.PackageDigest, "artifact package digest");
            DeliveryReceiptValidation.ValidateOptionalDigest(
                artifact.PageMapDigest, "artifact page-map digest");
            if (artifact.RendererFingerprint is not null)
            {
                DeliveryReceiptValidation.RequireNonBlank(
                    artifact.RendererFingerprint, "renderer fingerprint", 4096);
            }
            if (artifact.DocumentVersion is < 0
                or > DeliveryReceiptValidation.MaxPortableInteger)
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
        if ((artifact.DocumentVersion is null) != (artifact.PackageDigest is null))
        {
            findings.Add("invalid_artifact_document_binding");
            valid = false;
        }
        if (artifact.Availability == DeliveryArtifactAvailability.Unavailable)
        {
            if (artifact.Digest is not null || artifact.ByteLength is not null
                || string.IsNullOrWhiteSpace(artifact.UnavailableReason))
            {
                findings.Add("invalid_artifact_record");
                valid = false;
            }
            else if (!ValidateProfiledFreeText(
                artifact.UnavailableReason!, privacyProfile, "artifact unavailable reason"))
            {
                findings.Add("privacy_profile_violation");
                valid = false;
            }
        }
        return valid;
    }

    private static bool ValidateProfiledFreeText(
        string value,
        DeliveryReceiptPrivacyProfile privacyProfile,
        string label)
    {
        if (privacyProfile == DeliveryReceiptPrivacyProfile.FullEvidence)
            return true;
        if (privacyProfile == DeliveryReceiptPrivacyProfile.HashOnly)
            return IsDigestToken(value);
        var prefix = $"{label}; ";
        const string separator = " characters; ";
        if (!value.StartsWith(prefix, StringComparison.Ordinal))
            return false;
        int separatorIndex = value.IndexOf(separator, prefix.Length, StringComparison.Ordinal);
        return separatorIndex > prefix.Length
            && int.TryParse(
                value.AsSpan(prefix.Length, separatorIndex - prefix.Length),
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out var count)
            && count >= 0
            && IsDigestToken(value[(separatorIndex + separator.Length)..]);
    }

    private static bool IsDigestToken(string value)
    {
        if (value.Length != 71 || !value.StartsWith("sha256:", StringComparison.Ordinal))
            return false;
        foreach (var character in value.AsSpan(7))
        {
            if (character is not (>= '0' and <= '9')
                and not (>= 'a' and <= 'f'))
            {
                return false;
            }
        }
        return true;
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
                findings.Add($"artifact_{DocxSessionJson.EnumToSnake(status)}:{artifact.ArtifactId}");
        }
        return results;
    }

    private static bool ValidateCitationBindings(
        DeliveryChangeReceiptPayload payload,
        IReadOnlyDictionary<string, byte[]> artifactBytes,
        IReadOnlySet<string> verifiedArtifacts,
        DeliveryLineageValidationResult lineageValidation,
        DeliveryReceiptLimits limits,
        ICollection<string> findings)
    {
        bool valid = true;
        var artifacts = payload.Artifacts
            .GroupBy(artifact => artifact.ArtifactId, StringComparer.Ordinal)
            .Where(group => group.Count() == 1)
            .ToDictionary(group => group.Key, group => group.Single(), StringComparer.Ordinal);
        var pageMapCache = new Dictionary<string, (PageMap? Map, string? Finding)>(
            StringComparer.Ordinal);
        var projectionCache = new Dictionary<
            (string ArtifactId, string AnchorId, long DocumentVersion, string Renderer),
            PageCitation>();
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
                    || citation.DocumentVersion > DeliveryReceiptValidation.MaxPortableInteger
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
            else if (!verifiedArtifacts.Contains(citation.RenderArtifactId))
            {
                findings.Add($"citation_render_artifact_unverified:{citation.AnchorId}");
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

            if (!verifiedArtifacts.Contains(citation.PageMapArtifactId))
            {
                findings.Add($"citation_page_map_artifact_unverified:{citation.AnchorId}");
                valid = false;
                continue;
            }

            if (!artifactBytes.TryGetValue(citation.PageMapArtifactId, out var pageMapBytes))
            {
                findings.Add($"citation_page_map_bytes_missing:{citation.AnchorId}");
                valid = false;
                continue;
            }
            if (!pageMapCache.TryGetValue(citation.PageMapArtifactId, out var cachedMap))
            {
                try
                {
                    // Strict JSON/duplicate-property gate only; the map digest remains over raw
                    // bytes already authenticated by VerifyArtifacts.
                    _ = DeliveryReceiptCanonicalJson.CanonicalizeBounded(
                        pageMapBytes, limits, limits.MaxPageMapBytes,
                        "page_map_resource_limit");
                    // Cache only the citation-independent parse; ValidatePortable is
                    // parameterized per citation and must not share a cached verdict.
                    cachedMap = (DocxSessionJson.ParsePageMap(
                        Encoding.UTF8.GetString(pageMapBytes)), null);
                }
                catch (DeliveryReceiptValidationException ex)
                {
                    cachedMap = (null, ex.Code);
                }
                catch (Exception ex) when (ex is JsonException or FormatException)
                {
                    cachedMap = (null, "invalid_page_map_artifact");
                }
                pageMapCache.Add(citation.PageMapArtifactId, cachedMap);
            }
            if (cachedMap.Map is null)
            {
                findings.Add($"{cachedMap.Finding}:{citation.AnchorId}");
                valid = false;
                continue;
            }
            if (!PageMapContract.ValidatePortable(
                    cachedMap.Map, citation.DocumentVersion,
                    citation.RendererFingerprint).Success)
            {
                findings.Add($"invalid_page_map_artifact:{citation.AnchorId}");
                valid = false;
                continue;
            }

            var projectionKey = (citation.PageMapArtifactId, citation.AnchorId,
                citation.DocumentVersion, citation.RendererFingerprint);
            if (!projectionCache.TryGetValue(projectionKey, out var projected))
            {
                projected = PageMapContract.ProjectCitation(
                    cachedMap.Map,
                    citation.AnchorId,
                    new PageCitationRequest(
                        citation.DocumentVersion, citation.RendererFingerprint));
                projectionCache.Add(projectionKey, projected);
            }
            if (projected.Availability != PageMapAvailability.Available
                || !CitationCoordinatesEqual(citation, projected)
                || projected.Fragments.Any(fragment =>
                    !PageMapContract.StoryMatchesScope(fragment.Story, citation.Scope)))
            {
                findings.Add($"citation_page_map_projection_mismatch:{citation.AnchorId}");
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
        DeliveryReceiptValidation.ValidatePortableNonNegativeInteger(
            document.DocumentVersion, "invalid_document_version", "Document version");
        if (!DeliveryPackageManifestAdapter.IsSupportedSchema(
                document.PackageManifestSchema))
        {
            throw new DeliveryReceiptValidationException(
                "unsupported_package_manifest", "Unsupported package-manifest schema.");
        }
        DeliveryReceiptValidation.RequireNonBlank(
            document.PackageKind, "document package kind", 256);
        if (!string.Equals(document.PackageKind, "opc", StringComparison.Ordinal))
        {
            throw new DeliveryReceiptValidationException(
                "not_wordprocessing_package", "Document identity must be an OPC package.");
        }
        DeliveryReceiptValidation.RequireOpcMainDocumentUri(
            document.MainDocumentUri, "main document URI");
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

    private static DeliveryReceiptVerificationResult IntegrityFailure(string finding) => new()
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
