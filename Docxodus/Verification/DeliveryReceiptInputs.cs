// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text;
using System.Text.Json;

namespace Docxodus.Verification;

/// <summary>A caller-normalized document operation paired with one batch step.</summary>
public sealed class DeliveryNormalizedOperation
{
    private DeliveryNormalizedOperation(string tool, string action, JsonElement arguments)
    {
        Tool = tool;
        Action = action;
        Arguments = arguments;
        CanonicalArguments = DeliveryReceiptCanonicalJson.SerializeCanonical(arguments);
        ArgumentsDigest = DeliveryReceiptCanonicalJson.Digest(CanonicalArguments);
    }

    public string Tool { get; }
    public string Action { get; }
    public JsonElement Arguments { get; }
    public VerificationDigest ArgumentsDigest { get; }
    internal byte[] CanonicalArguments { get; }

    public static DeliveryNormalizedOperation Create(
        string tool,
        string action,
        string argumentsJson = "{}")
    {
        ArgumentNullException.ThrowIfNull(argumentsJson);
        if (argumentsJson.Length > new DeliveryReceiptLimits().MaxStringLength)
        {
            throw new DeliveryReceiptValidationException(
                "receipt_resource_limit", "Operation arguments exceed the string-length limit.");
        }
        return new DeliveryNormalizedOperation(
            DeliveryReceiptValidation.RequireNonBlank(tool, nameof(tool), 256),
            DeliveryReceiptValidation.RequireNonBlank(action, nameof(action), 256),
            DeliveryReceiptCanonicalJson.ParseCanonicalObject(argumentsJson, nameof(argumentsJson)));
    }
}

/// <summary>Optional transport identity supplied alongside a core MutationBatchResult.</summary>
public sealed record DeliveryTransactionIdentity
{
    public int SchemaVersion { get; init; } = 1;
    required public string TransactionId { get; init; }
    required public string RequestFingerprint { get; init; }
}

/// <summary>
/// The immutable inputs needed to turn one MutationBatchResult into receipt evidence. Operation
/// arguments and transport identity are explicit because the core result intentionally does not
/// retain either one.
/// </summary>
public sealed class DeliveryTransactionContribution
{
    private DeliveryTransactionContribution(
        MutationBatchResult result,
        DeliveryDocumentIdentity beforeDocument,
        DeliveryDocumentIdentity afterDocument,
        IReadOnlyList<DeliveryNormalizedOperation> operations,
        DeliveryTransactionIdentity? identity)
    {
        Result = result;
        BeforeDocument = beforeDocument;
        AfterDocument = afterDocument;
        Operations = operations;
        Identity = identity;
    }

    public MutationBatchResult Result { get; }
    public DeliveryDocumentIdentity BeforeDocument { get; }
    public DeliveryDocumentIdentity AfterDocument { get; }
    public IReadOnlyList<DeliveryNormalizedOperation> Operations { get; }
    public DeliveryTransactionIdentity? Identity { get; }

    public static DeliveryTransactionContribution FromMutationBatchResult(
        MutationBatchResult result,
        PackageManifest beforeManifest,
        PackageManifest afterManifest,
        IEnumerable<DeliveryNormalizedOperation> operations,
        DeliveryTransactionIdentity? identity = null)
    {
        ArgumentNullException.ThrowIfNull(result);
        ArgumentNullException.ThrowIfNull(beforeManifest);
        ArgumentNullException.ThrowIfNull(afterManifest);
        ArgumentNullException.ThrowIfNull(operations);
        if (!Enum.IsDefined(result.Mode))
            throw new DeliveryReceiptValidationException("invalid_batch_mode", "Unknown batch mode.");
        if (result.BaseVersion < 0 || result.ResultVersion < 0)
            throw new DeliveryReceiptValidationException(
                "invalid_document_version", "Batch versions cannot be negative.");

        var materialized = operations.ToArray();
        if (materialized.Any(operation => operation is null))
        {
            throw new DeliveryReceiptValidationException(
                "null_operation", "Normalized operations cannot contain null.");
        }
        if (materialized.Length > new DeliveryReceiptLimits().MaxOperationsPerTransaction)
        {
            throw new DeliveryReceiptValidationException(
                "receipt_resource_limit",
                "Normalized operations exceed the default per-transaction limit.");
        }
        var stepsByIndex = new Dictionary<int, MutationBatchStepResult>();
        foreach (var step in result.Steps)
        {
            if (step is null || step.Index < 0 || step.Index >= materialized.Length
                || !stepsByIndex.TryAdd(step.Index, step)
                || !string.Equals(step.Tool, materialized[step.Index].Tool,
                    StringComparison.Ordinal)
                || !string.Equals(step.Action, materialized[step.Index].Action,
                    StringComparison.Ordinal)
                || step.Results is null
                || step.Results.Any(stepResult => stepResult is null))
            {
                throw new DeliveryReceiptValidationException(
                    "operation_step_mismatch",
                    "Sparse batch results must have unique in-range indices and match requested operations.");
            }
        }
        if ((result.Mode == MutationBatchMode.BestEffort || result.Success)
            && stepsByIndex.Count != materialized.Length)
        {
            throw new DeliveryReceiptValidationException(
                "operation_step_mismatch",
                "Successful and best-effort batches require one result for every requested operation.");
        }
        if (!result.Success && result.Mode == MutationBatchMode.Atomic
            && stepsByIndex.Count == 0)
        {
            throw new DeliveryReceiptValidationException(
                "operation_step_mismatch", "A failed atomic batch requires failure evidence.");
        }
        ValidateBatchOutcome(
            result, stepsByIndex.Values.OrderBy(step => step.Index).ToArray());

        if (identity is not null)
        {
            if (result.Preview)
            {
                throw new DeliveryReceiptValidationException(
                    "preview_transaction_identity",
                    "Predictions cannot carry applying transaction identities.");
            }
            if (identity.SchemaVersion != 1)
            {
                throw new DeliveryReceiptValidationException(
                    "unsupported_transaction_identity", "Only transaction identity v1 is supported.");
            }
            DeliveryReceiptValidation.RequireNonBlank(
                identity.TransactionId, "transaction id", 1024);
            ValidateFingerprint(identity.RequestFingerprint);
        }

        return new DeliveryTransactionContribution(
            result,
            DeliveryDocumentIdentity.FromManifest(beforeManifest, result.BaseVersion),
            DeliveryDocumentIdentity.FromManifest(afterManifest, result.ResultVersion),
            materialized,
            identity);
    }

    internal static void ValidateFingerprint(string fingerprint)
    {
        if (fingerprint is null || fingerprint.Length != 71
            || !fingerprint.StartsWith("sha256:", StringComparison.Ordinal)
            || fingerprint.AsSpan(7).ToString().Any(
                c => !((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f'))))
        {
            throw new DeliveryReceiptValidationException(
                "invalid_request_fingerprint",
                "Request fingerprints must be 'sha256:' followed by 64 lower-case hex characters.");
        }
    }

    private static void ValidateBatchOutcome(
        MutationBatchResult result,
        IReadOnlyList<MutationBatchStepResult> steps)
    {
        bool derivedSuccess = steps.All(step => step.Success);
        if (result.Success != derivedSuccess
            || (result.Mode == MutationBatchMode.BestEffort && result.RolledBack)
            || (result.Mode == MutationBatchMode.Atomic
                && result.RolledBack == result.Success))
        {
            throw new DeliveryReceiptValidationException(
                "invalid_batch_result", "Batch success and rollback flags are inconsistent.");
        }

        if (result.Mode == MutationBatchMode.BestEffort
            && steps.Any(step => step.RolledBack))
        {
            throw new DeliveryReceiptValidationException(
                "invalid_batch_result", "Best-effort steps cannot be rolled back.");
        }
        if (result.Mode == MutationBatchMode.Atomic && result.Success
            && steps.Any(step => !step.Success || step.RolledBack))
        {
            throw new DeliveryReceiptValidationException(
                "invalid_batch_result", "A committed atomic batch requires retained successful steps.");
        }
        if (result.Mode == MutationBatchMode.Atomic && !result.Success)
        {
            var failedSteps = steps.Where(step => !step.Success).ToArray();
            if (failedSteps.Length != 1 || steps.Any(step => !step.RolledBack))
            {
                throw new DeliveryReceiptValidationException(
                    "invalid_batch_result", "A failed atomic batch requires one rolled-back failure.");
            }
            int failedIndex = failedSteps[0].Index;
            bool preflightShape = steps.Count == 1;
            bool executionShape = steps.Count == failedIndex + 1
                && steps.Select(step => step.Index).SequenceEqual(
                    Enumerable.Range(0, failedIndex + 1))
                && steps.Take(failedIndex).All(step => step.Success);
            if (!preflightShape && !executionShape)
            {
                throw new DeliveryReceiptValidationException(
                    "invalid_batch_result",
                    "Failed atomic steps must be one preflight failure or an executed prefix.");
            }
        }

        var firstFailure = steps.FirstOrDefault(step => !step.Success);
        var firstFailureResult = firstFailure?.Results.FirstOrDefault(value => !value.Success);
        if (result.Success != (result.Failure is null)
            || (!result.Success && (firstFailure is null
                || firstFailureResult?.Error is null
                || result.Failure!.Index != firstFailure.Index
                || !string.Equals(result.Failure.Tool, firstFailure.Tool,
                    StringComparison.Ordinal)
                || !string.Equals(result.Failure.Action, firstFailure.Action,
                    StringComparison.Ordinal)
                || result.Failure.RolledBack != firstFailure.RolledBack
                || result.Failure.Error != firstFailureResult.Error)))
        {
            throw new DeliveryReceiptValidationException(
                "invalid_batch_result", "Batch failure metadata is inconsistent.");
        }
    }
}

/// <summary>Bytes and optional render/document binding for one output artifact.</summary>
public sealed record DeliveryArtifactInput
{
    required public string ArtifactId { get; init; }
    required public DeliveryArtifactRole Role { get; init; }
    required public string MediaType { get; init; }
    required public DeliveryArtifactAvailability Availability { get; init; }
    public byte[]? Bytes { get; init; }
    public string? RelativePath { get; init; }
    public string? UnavailableReason { get; init; }
    public DeliveryDocumentIdentity? Document { get; init; }
    public string? RendererFingerprint { get; init; }
    public VerificationDigest? PageMapDigest { get; init; }

    public static DeliveryArtifactInput Available(
        string artifactId,
        DeliveryArtifactRole role,
        string mediaType,
        ReadOnlySpan<byte> bytes)
    {
        var defaults = new DeliveryReceiptLimits();
        var maximum = role switch
        {
            DeliveryArtifactRole.SemanticDiff => defaults.MaxSemanticEvidenceBytes,
            DeliveryArtifactRole.PageMap => defaults.MaxPageMapBytes,
            _ => defaults.MaxArtifactBytes,
        };
        var code = role switch
        {
            DeliveryArtifactRole.SemanticDiff => "semantic_resource_limit",
            DeliveryArtifactRole.PageMap => "page_map_resource_limit",
            _ => "artifact_resource_limit",
        };
        DeliveryReceiptResourceBudget.Bytes(
            bytes.Length,
            maximum,
            code,
            "Artifact");
        return new DeliveryArtifactInput
        {
            ArtifactId = artifactId,
            Role = role,
            MediaType = mediaType,
            Availability = DeliveryArtifactAvailability.Available,
            Bytes = bytes.ToArray(),
        };
    }

    public static DeliveryArtifactInput Unavailable(
        string artifactId,
        DeliveryArtifactRole role,
        string mediaType,
        string reason) => new()
    {
        ArtifactId = artifactId,
        Role = role,
        MediaType = mediaType,
        Availability = DeliveryArtifactAvailability.Unavailable,
        UnavailableReason = reason,
    };
}

/// <summary>
/// Typed registration of exact <see cref="SemanticChangeSet.ToCanonicalUtf8Bytes"/> output for
/// either the complete delivery or one state-changing transaction.
/// </summary>
public sealed class DeliverySemanticChangeSetInput
{
    private DeliverySemanticChangeSetInput(
        DeliverySemanticComparisonScope scope,
        string? transactionEntryId,
        SemanticChangeSet changeSet,
        string artifactId,
        string? relativePath)
    {
        Scope = scope;
        TransactionEntryId = transactionEntryId;
        ArtifactId = DeliveryReceiptValidation.RequireNonBlank(
            artifactId, "semantic artifact id", 256);
        RelativePath = DeliveryReceiptValidation.NormalizeRelativePath(relativePath);
        ChangeSet = changeSet ?? throw new ArgumentNullException(nameof(changeSet));
    }

    public DeliverySemanticComparisonScope Scope { get; }
    public string? TransactionEntryId { get; }
    public string ArtifactId { get; }
    public string? RelativePath { get; }
    internal SemanticChangeSet ChangeSet { get; }

    public static DeliverySemanticChangeSetInput ForSourceToDelivered(
        SemanticChangeSet changeSet,
        string artifactId = "semantic-source-to-delivered",
        string? relativePath = null) =>
        new(DeliverySemanticComparisonScope.SourceToDelivered, null,
            changeSet, artifactId, relativePath);

    public static DeliverySemanticChangeSetInput ForTransaction(
        string transactionEntryId,
        SemanticChangeSet changeSet,
        string artifactId,
        string? relativePath = null) =>
        new(DeliverySemanticComparisonScope.Transaction,
            DeliveryReceiptValidation.RequireNonBlank(
                transactionEntryId, "semantic transaction entry id", 256),
            changeSet, artifactId, relativePath);
}

/// <summary>Existing PageCitation plus the package, page map, and artifact that made it true.</summary>
public sealed record DeliveryPageCitationInput
{
    required public PageCitation Citation { get; init; }
    required public string Scope { get; init; }
    required public DeliveryDocumentIdentity Document { get; init; }
    required public VerificationDigest PageMapDigest { get; init; }
    required public string PageMapArtifactId { get; init; }
    required public string RenderArtifactId { get; init; }
}

/// <summary>One explicit undo/redo contribution.</summary>
public sealed record DeliveryLineageEventInput
{
    required public DeliveryLineageAction Action { get; init; }
    required public string AffectedEntryId { get; init; }
    required public DeliveryDocumentIdentity BeforeDocument { get; init; }
    required public DeliveryDocumentIdentity AfterDocument { get; init; }

    public static DeliveryLineageEventInput FromManifests(
        DeliveryLineageAction action,
        string affectedEntryId,
        PackageManifest beforeManifest,
        long beforeVersion,
        PackageManifest afterManifest,
        long afterVersion) => new()
    {
        Action = action,
        AffectedEntryId = affectedEntryId,
        BeforeDocument = DeliveryDocumentIdentity.FromManifest(beforeManifest, beforeVersion),
        AfterDocument = DeliveryDocumentIdentity.FromManifest(afterManifest, afterVersion),
    };
}

/// <summary>An exact attribution rule for one class of observed manifest changes.</summary>
public sealed record DeliveryChangeAttributionRule
{
    required public DeliveryPackageChangeKind Kind { get; init; }
    public string? EntryUri { get; init; }
    public string? OwnerUri { get; init; }
    public string? RelationshipId { get; init; }
    required public DeliveryChangeDisposition Disposition { get; init; }
    public string? TransactionEntryId { get; init; }
    public int? RequestedOperationIndex { get; init; }
    public string? Derivation { get; init; }
}
