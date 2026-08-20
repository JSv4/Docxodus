// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json;

namespace Docxodus.Verification;

/// <summary>Pre-serialization resource checks for the public object verification surface.</summary>
internal static class DeliveryReceiptResourceValidator
{
    public static void ValidateArtifacts(
        IReadOnlyDictionary<string, byte[]> artifactBytes,
        DeliveryReceiptLimits limits)
    {
        if (artifactBytes.Count > limits.MaxArtifacts)
            Fail("Artifact dictionary exceeds the artifact-count limit.", "artifact_resource_limit");

        var budget = new DeliveryReceiptResourceBudget(limits);
        budget.AddItems(artifactBytes.Count, "artifact dictionary entries");
        long total = 0;
        foreach (var pair in artifactBytes)
        {
            budget.String(pair.Key, "artifact dictionary key");
            if (pair.Value is null)
                throw new ArgumentException("Artifact byte values cannot be null.", nameof(artifactBytes));
            DeliveryReceiptResourceBudget.Bytes(
                pair.Value.LongLength, limits.MaxArtifactBytes,
                "artifact_resource_limit", $"Artifact '{pair.Key}'");
            if (pair.Value.LongLength > limits.MaxTotalArtifactBytes - total)
                Fail("Artifact dictionary exceeds the aggregate byte limit.", "artifact_resource_limit");
            total += pair.Value.LongLength;
        }
    }

    public static void ValidatePayload(
        DeliveryChangeReceiptPayload payload,
        DeliveryReceiptLimits limits)
    {
        ArgumentNullException.ThrowIfNull(payload);
        if (payload.Transactions.Count > limits.MaxTransactions)
            Fail("Receipt exceeds the transaction limit.");
        if (payload.Artifacts.Count > limits.MaxArtifacts)
            Fail("Receipt exceeds the artifact-count limit.", "artifact_resource_limit");

        var budget = new DeliveryReceiptResourceBudget(limits);
        budget.AddSerializedBytes(512, "receipt structure");
        budget.String(payload.Schema, "receipt schema");
        budget.String(payload.Canonicalization, "canonicalization id");
        ValidateDocument(payload.SourceDocument, budget);
        ValidateDocument(payload.DeliveredDocument, budget);

        Add(payload.Transactions, budget, "transactions");
        foreach (var transaction in payload.Transactions)
        {
            if (transaction.Operations.Count > limits.MaxOperationsPerTransaction)
                Fail("Transaction exceeds the operation limit.");
            budget.String(transaction.EntryId, "transaction entry id");
            budget.String(transaction.TransactionId, "transaction id");
            budget.String(transaction.RequestFingerprint, "request fingerprint");
            ValidateDocument(transaction.BeforeDocument, budget);
            ValidateDocument(transaction.AfterDocument, budget);
            ValidateDigest(transaction.ReportedPackageContentDigest, budget);
            Add(transaction.Operations, budget, "transaction operations");
            foreach (var operation in transaction.Operations)
            {
                budget.String(operation.Tool, "operation tool");
                budget.String(operation.Action, "operation action");
                budget.String(operation.ArgumentsSummary, "operation argument summary");
                ValidateDigest(operation.ArgumentsDigest, budget);
                if (operation.Arguments is { } arguments)
                    ValidateJson(arguments, limits, budget, 6);
                Add(operation.Results, budget, "operation results");
                foreach (var result in operation.Results)
                {
                    ValidateDigest(result.ResultDigest, budget);
                    budget.String(result.ErrorCode, "operation error code");
                    ValidateText(result.ErrorMessage, budget);
                    Add(result.ObjectChanges, budget, "object changes");
                    foreach (var change in result.ObjectChanges)
                    {
                        budget.String(change.AnchorId, "object-change anchor");
                        budget.String(change.Kind, "object-change kind");
                        budget.String(change.Scope, "object-change scope");
                        budget.String(change.Unid, "object-change unid");
                    }
                    if (result.FullResult is { } fullResult)
                        ValidateJson(fullResult, limits, budget, 8);
                }
            }
            Add(transaction.AuthoredChanges, budget, "authored changes");
            foreach (var change in transaction.AuthoredChanges)
            {
                budget.String(change.EntityId, "authored entity id");
                budget.String(change.Author, "authored author");
                budget.String(change.Date, "authored date");
                budget.String(change.DateUtc, "authored UTC date");
                budget.String(change.Type, "authored type");
                budget.String(change.PartUri, "authored part URI");
                budget.String(change.Scope, "authored scope");
                budget.String(change.AnchorId, "authored primary anchor");
                if (change.Diagnostic is { } diagnostic)
                {
                    budget.String(diagnostic.Code, "authored diagnostic code");
                    ValidateText(diagnostic.Message, budget);
                }
                ValidateDigest(change.SourceDigest, budget);
                Add(change.ConstituentIds, budget, "authored constituent ids");
                foreach (var id in change.ConstituentIds)
                    budget.String(id, "authored constituent id");
                Add(change.ConstituentKeys, budget, "authored constituent keys");
                foreach (var key in change.ConstituentKeys)
                    budget.String(key, "authored constituent key");
                Add(change.AffectedAnchorIds, budget, "affected anchor ids");
                foreach (var anchor in change.AffectedAnchorIds)
                    budget.String(anchor, "affected anchor id");
                ValidateText(change.Text, budget);
                if (change.FullEvidence is { } fullEvidence)
                    ValidateJson(fullEvidence, limits, budget, 6);
            }
            Add(transaction.Warnings, budget, "transaction warnings");
            foreach (var warning in transaction.Warnings)
                ValidateText(warning, budget);
        }

        Add(payload.Lineage, budget, "lineage events");
        foreach (var lineage in payload.Lineage)
        {
            budget.String(lineage.AffectedEntryId, "lineage entry id");
            ValidateDocument(lineage.BeforeDocument, budget);
            ValidateDocument(lineage.AfterDocument, budget);
        }

        Add(payload.PackageChanges, budget, "package changes");
        foreach (var change in payload.PackageChanges)
        {
            budget.String(change.ChangeId, "package change id");
            budget.String(change.Location?.EntryUri, "package entry URI");
            budget.String(change.Location?.OwnerUri, "relationship owner URI");
            budget.String(change.Location?.RelationshipId, "relationship id");
            budget.String(change.Location?.TargetUri, "relationship target URI");
            budget.String(change.Location?.PropertyPath, "package property path");
            ValidateText(change.Before, budget);
            ValidateText(change.After, budget);
            budget.String(change.TransactionEntryId, "attribution entry id");
            budget.String(change.Derivation, "derivation");
        }

        Add(payload.Evidence, budget, "evidence references");
        foreach (var evidence in payload.Evidence)
        {
            budget.String(evidence.Schema, "evidence schema");
            budget.String(evidence.ArtifactId, "evidence artifact id");
            budget.String(evidence.Summary, "evidence summary");
            ValidateDigest(evidence.Digest, budget);
        }

        Add(payload.SemanticChangeSets, budget, "semantic bindings");
        foreach (var binding in payload.SemanticChangeSets)
        {
            budget.String(binding.TransactionEntryId, "semantic transaction entry id");
            budget.String(binding.Schema, "semantic schema");
            budget.String(binding.ArtifactId, "semantic artifact id");
            ValidateDocument(binding.BeforeDocument, budget);
            ValidateDocument(binding.AfterDocument, budget);
            ValidateDigest(binding.Digest, budget);
        }

        Add(payload.Artifacts, budget, "artifacts");
        foreach (var artifact in payload.Artifacts)
        {
            budget.String(artifact.ArtifactId, "artifact id");
            budget.String(artifact.MediaType, "artifact media type");
            budget.String(artifact.RelativePath, "artifact path");
            budget.String(artifact.UnavailableReason, "artifact unavailable reason");
            budget.String(artifact.RendererFingerprint, "renderer fingerprint");
            ValidateDigest(artifact.Digest, budget);
            ValidateDigest(artifact.PackageDigest, budget);
            ValidateDigest(artifact.PageMapDigest, budget);
        }

        Add(payload.PageCitations, budget, "page citations");
        foreach (var citation in payload.PageCitations)
        {
            budget.String(citation.AnchorId, "citation anchor");
            budget.String(citation.Scope, "citation scope");
            budget.String(citation.RendererFingerprint, "renderer fingerprint");
            budget.String(citation.PageMapArtifactId, "PageMap artifact id");
            budget.String(citation.RenderArtifactId, "render artifact id");
            ValidateDigest(citation.PackageDigest, budget);
            ValidateDigest(citation.PageMapDigest, budget);
            ValidateDigest(citation.RenderArtifactDigest, budget);
            Add(citation.Pages, budget, "citation pages");
            foreach (var page in citation.Pages)
                budget.String(page?.PageName, "page name");
            Add(citation.Fragments, budget, "citation fragments");
            foreach (var fragment in citation.Fragments)
            {
                budget.String(fragment?.FragmentId, "fragment id");
                budget.String(fragment?.AnchorId, "fragment anchor id");
            }
        }

        Add(payload.Warnings, budget, "receipt warnings");
        foreach (var warning in payload.Warnings)
            ValidateText(warning, budget);
    }

    private static void ValidateDocument(
        DeliveryDocumentIdentity? document,
        DeliveryReceiptResourceBudget budget)
    {
        if (document is null)
            return;
        budget.String(document.PackageKind, "package kind");
        budget.String(document.PackageManifestSchema, "manifest schema");
        budget.String(document.MainDocumentUri, "main document URI");
        ValidateDigest(document.RawPackageBytesDigest, budget);
        ValidateDigest(document.OrderedOpcContentDigest, budget);
        ValidateDigest(document.NormalizedSemanticDigest, budget);
    }

    private static void ValidateText(
        DeliveryTextEvidence? evidence,
        DeliveryReceiptResourceBudget budget)
    {
        if (evidence is null)
            return;
        ValidateDigest(evidence.Digest, budget);
        budget.String(evidence.Summary, "text summary");
        budget.String(evidence.Value, "text value");
    }

    private static void ValidateDigest(
        VerificationDigest? digest,
        DeliveryReceiptResourceBudget budget)
    {
        if (digest is null)
            return;
        budget.String(digest.Algorithm, "digest algorithm");
        budget.String(digest.Value, "digest value");
    }

    private static void ValidateJson(
        JsonElement value,
        DeliveryReceiptLimits limits,
        DeliveryReceiptResourceBudget budget,
        int depth)
    {
        if (depth > limits.MaxJsonDepth)
            Fail("Embedded JSON exceeds the depth limit.");
        switch (value.ValueKind)
        {
            case JsonValueKind.Object:
                foreach (var property in value.EnumerateObject())
                {
                    budget.AddItems(1, "embedded JSON properties");
                    budget.String(property.Name, "JSON property name");
                    ValidateJson(property.Value, limits, budget, depth + 1);
                }
                break;
            case JsonValueKind.Array:
                foreach (var item in value.EnumerateArray())
                {
                    budget.AddItems(1, "embedded JSON items");
                    ValidateJson(item, limits, budget, depth + 1);
                }
                break;
            case JsonValueKind.String:
                budget.String(value.GetString(), "JSON string");
                break;
        }
    }

    private static void Add<T>(
        IReadOnlyCollection<T> values,
        DeliveryReceiptResourceBudget budget,
        string name)
    {
        budget.AddItems(values.Count, name);
        budget.AddSerializedBytes(checked(values.Count * 2L + 2L), name);
    }

    private static void Fail(string message, string code = "receipt_resource_limit") =>
        throw new DeliveryReceiptValidationException(code, message);
}
