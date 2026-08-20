// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Globalization;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using Docxodus.Internal;

namespace Docxodus.Verification;

/// <summary>
/// Composes immutable transaction, package, render, and artifact evidence into one receipt. This
/// type performs no mutation or rendering. It does independently re-inspect the exact clean-DOCX
/// bytes and strictly reconstruct typed semantic evidence before binding either artifact.
/// </summary>
public sealed class DeliveryChangeReceiptBuilder
{
    private readonly PackageManifest _sourceManifest;
    private readonly DeliveryDocumentIdentity _sourceDocument;
    private readonly DeliveryReceiptPrivacyProfile _privacyProfile;
    private readonly DeliveryReceiptLimits _limits;
    private readonly List<DeliveryTransactionEntry> _transactions = new();
    private readonly Dictionary<string, DeliveryTransactionEntry> _entriesById =
        new(StringComparer.Ordinal);
    private readonly Dictionary<string, DeliveryTransactionEntry> _entriesByTransactionId =
        new(StringComparer.Ordinal);
    private readonly List<(long Sequence, DeliveryLineageEventInput Input)> _lineage = new();
    private readonly Dictionary<string, DeliveryArtifact> _artifacts = new(StringComparer.Ordinal);
    private readonly Dictionary<string, byte[]> _artifactBytes = new(StringComparer.Ordinal);
    private readonly Dictionary<string, PageMap> _validatedPageMaps = new(StringComparer.Ordinal);
    private readonly Dictionary<(string ArtifactId, string AnchorId), PageCitation>
        _pageMapProjections = new();
    private readonly List<DeliveryPageCitationInput> _citations = new();
    private readonly List<DeliveryEvidenceReference> _evidence = new();
    private readonly List<(DeliverySemanticChangeSetInput Input,
        DeliverySemanticChangeSetProjection Projection)> _semanticChangeSets = new();
    private readonly Dictionary<DeliveryAttributionKey, DeliveryChangeAttributionRule>
        _attributionRules = new();
    private readonly List<string> _warnings = new();
    private long _nextSequence;
    private DeliveryDocumentIdentity? _deliveredDocument;
    private long _totalArtifactBytes;

    public DeliveryChangeReceiptBuilder(
        PackageManifest sourceManifest,
        long sourceDocumentVersion,
        DeliveryReceiptPrivacyProfile privacyProfile =
            DeliveryReceiptPrivacyProfile.HashAndSummary,
        DeliveryReceiptLimits? limits = null)
    {
        ArgumentNullException.ThrowIfNull(sourceManifest);
        if (!Enum.IsDefined(privacyProfile))
            throw new ArgumentOutOfRangeException(nameof(privacyProfile));
        _limits = (limits ?? new DeliveryReceiptLimits()).ValidateAndClone();
        ValidateManifestResources(sourceManifest);
        _sourceManifest = CloneManifest(sourceManifest);
        _sourceDocument = DeliveryDocumentIdentity.FromManifest(
            sourceManifest, sourceDocumentVersion);
        _privacyProfile = privacyProfile;
    }

    /// <summary>When true, Build rejects rather than merely reporting unexpected package changes.</summary>
    public bool FailOnUnexpectedChanges { get; set; }

    public DeliveryChangeReceiptBuilder SetDeliveredDocument(
        PackageManifest manifest,
        long documentVersion)
    {
        ArgumentNullException.ThrowIfNull(manifest);
        _deliveredDocument = DeliveryDocumentIdentity.FromManifest(manifest, documentVersion);
        return this;
    }

    /// <summary>
    /// Add one batch result. An identical idempotent retry returns the original entry id and does
    /// not append a second edit. Conflicting transaction reuse fails closed.
    /// </summary>
    public string AddTransaction(DeliveryTransactionContribution contribution)
    {
        ArgumentNullException.ThrowIfNull(contribution);
        ValidateContributionResources(contribution);
        var entry = BuildTransactionEntry(contribution, _nextSequence);

        if (contribution.Identity is { } identity
            && _entriesByTransactionId.TryGetValue(identity.TransactionId, out var prior))
        {
            if (!string.Equals(prior.RequestFingerprint, entry.RequestFingerprint,
                    StringComparison.Ordinal))
            {
                throw new DeliveryReceiptValidationException(
                    "transaction_conflict",
                    $"Transaction id '{identity.TransactionId}' has a different request fingerprint.");
            }
            if (!string.Equals(prior.EntryId, entry.EntryId, StringComparison.Ordinal))
            {
                throw new DeliveryReceiptValidationException(
                    "retry_result_conflict",
                    $"Transaction id '{identity.TransactionId}' no longer identifies the original result.");
            }
            if (!TransactionEvidenceEquals(prior, entry))
            {
                throw new DeliveryReceiptValidationException(
                    "retry_result_conflict",
                    $"Transaction id '{identity.TransactionId}' returned different result evidence.");
            }
            return prior.EntryId;
        }

        if (_entriesById.TryGetValue(entry.EntryId, out var duplicate))
        {
            if (!TransactionEvidenceEquals(duplicate, entry))
            {
                throw new DeliveryReceiptValidationException(
                    "retry_result_conflict",
                    $"Transaction entry '{entry.EntryId}' has conflicting result evidence.");
            }
            return duplicate.EntryId;
        }

        if (_transactions.Count >= _limits.MaxTransactions)
        {
            throw new DeliveryReceiptValidationException(
                "receipt_resource_limit", "Receipt transaction limit exceeded.");
        }

        _transactions.Add(entry);
        _entriesById.Add(entry.EntryId, entry);
        if (entry.TransactionId is not null)
            _entriesByTransactionId.Add(entry.TransactionId, entry);
        _nextSequence++;
        return entry.EntryId;
    }

    public DeliveryChangeReceiptBuilder AddLineageEvent(DeliveryLineageEventInput lineageEvent)
    {
        ArgumentNullException.ThrowIfNull(lineageEvent);
        if (!Enum.IsDefined(lineageEvent.Action))
            throw new ArgumentOutOfRangeException(nameof(lineageEvent));
        DeliveryReceiptValidation.RequireNonBlank(
            lineageEvent.AffectedEntryId, "affected entry id", 256);
        EnsureCollectionCapacity(_lineage.Count, "lineage events");
        _lineage.Add((_nextSequence++, lineageEvent));
        return this;
    }

    public DeliveryChangeReceiptBuilder AddArtifact(DeliveryArtifactInput input)
    {
        ArgumentNullException.ThrowIfNull(input);
        RegisterArtifact(input);
        return this;
    }

    /// <summary>
    /// Add exact canonical #457 evidence. One source-to-delivered binding and one binding for every
    /// state-changing transaction are required at Build time.
    /// </summary>
    public DeliveryChangeReceiptBuilder AddSemanticChangeSet(
        DeliverySemanticChangeSetInput input)
    {
        ArgumentNullException.ThrowIfNull(input);
        if (!Enum.IsDefined(input.Scope))
            throw new ArgumentOutOfRangeException(nameof(input));

        EnsureCollectionCapacity(_semanticChangeSets.Count, "semantic change sets");
        var projection = DeliverySemanticChangeSetAdapter.Project(input.ChangeSet, _limits);
        var artifactInput = DeliveryArtifactInput.Available(
            input.ArtifactId,
            DeliveryArtifactRole.SemanticDiff,
            "application/json",
            projection.CanonicalBytes,
            _limits) with
        {
            RelativePath = input.RelativePath,
        };
        var artifact = BuildArtifact(artifactInput);
        if (_artifacts.TryGetValue(artifact.ArtifactId, out var existing))
        {
            if (!ArtifactEquals(existing, artifact)
                || !_artifactBytes.TryGetValue(artifact.ArtifactId, out var existingBytes)
                || !existingBytes.AsSpan().SequenceEqual(projection.CanonicalBytes))
            {
                throw new DeliveryReceiptValidationException(
                    "semantic_artifact_conflict",
                    $"Semantic artifact id '{artifact.ArtifactId}' identifies different bytes or metadata.");
            }
        }
        else
        {
            RegisterArtifact(artifactInput);
        }
        _semanticChangeSets.Add((input, projection));
        return this;
    }

    public DeliveryChangeReceiptBuilder AddPageCitation(DeliveryPageCitationInput citation)
    {
        ArgumentNullException.ThrowIfNull(citation);
        EnsureCollectionCapacity(_citations.Count, "page citations");
        _citations.Add(citation);
        return this;
    }

    public DeliveryChangeReceiptBuilder AddEvidence(DeliveryEvidenceReference evidence)
    {
        ArgumentNullException.ThrowIfNull(evidence);
        EnsureCollectionCapacity(_evidence.Count, "evidence references");
        if (!Enum.IsDefined(evidence.Kind))
            throw new ArgumentOutOfRangeException(nameof(evidence));
        if (evidence.Kind == DeliveryEvidenceKind.SemanticChangeSet)
        {
            throw new DeliveryReceiptValidationException(
                "semantic_evidence_requires_typed_factory",
                "Semantic evidence must be registered from a SemanticChangeSet instance.");
        }
        DeliveryReceiptValidation.RequireNonBlank(evidence.Schema, "evidence schema", 2048);
        var expectedSchema = ExpectedEvidenceSchema(evidence.Kind);
        if (!string.Equals(evidence.Schema, expectedSchema, StringComparison.Ordinal))
        {
            throw new DeliveryReceiptValidationException(
                "evidence_schema_mismatch",
                $"Evidence kind '{evidence.Kind}' requires schema '{expectedSchema}'.");
        }
        DeliveryReceiptValidation.ValidateDigest(evidence.Digest, "evidence digest");
        if (evidence.ArtifactId is not null)
            DeliveryReceiptValidation.RequireNonBlank(evidence.ArtifactId, "evidence artifact id", 256);
        _evidence.Add(evidence with
        {
            Digest = DeliveryReceiptValidation.CloneDigest(evidence.Digest),
            Summary = evidence.Summary is null
                || _privacyProfile == DeliveryReceiptPrivacyProfile.HashOnly
                ? null
                : ProfiledFreeText(evidence.Summary, "evidence summary"),
        });
        return this;
    }

    public DeliveryChangeReceiptBuilder AddAttributionRule(DeliveryChangeAttributionRule rule)
    {
        ArgumentNullException.ThrowIfNull(rule);
        EnsureCollectionCapacity(_attributionRules.Count, "attribution rules");
        ValidateAttributionRule(rule);
        if (!_attributionRules.TryAdd(AttributionKey(rule), rule))
        {
            throw new DeliveryReceiptValidationException(
                "ambiguous_attribution",
                "More than one attribution rule identifies the same package change.");
        }
        return this;
    }

    public DeliveryChangeReceiptBuilder AddWarning(string warning)
    {
        EnsureCollectionCapacity(_warnings.Count, "warnings");
        _warnings.Add(DeliveryReceiptValidation.RequireNonBlank(
            warning, "warning", Math.Min(16_384, _limits.MaxStringLength)));
        return this;
    }

    public DeliveryChangeReceipt Build()
    {
        if (_deliveredDocument is null)
        {
            throw new DeliveryReceiptValidationException(
                "missing_delivered_document", "SetDeliveredDocument must be called before Build.");
        }

        var lineage = MaterializeLineage();
        var lineageValidation = DeliveryReceiptLineageValidator.Validate(
            _sourceDocument, _deliveredDocument, _transactions, lineage);
        if (!lineageValidation.IsValid)
        {
            throw new DeliveryReceiptValidationException(
                lineageValidation.Findings[0],
                "Transaction and undo/redo history is not a valid delivery lineage.");
        }
        var deliveredManifest = ValidateRequiredCleanDocx(_deliveredDocument);
        var semanticChangeSets = BuildSemanticChangeSetBindings(lineageValidation);
        var packageChanges = BuildPackageChanges(
            _sourceManifest, deliveredManifest, lineageValidation);
        if (FailOnUnexpectedChanges
            && packageChanges.Any(change =>
                change.Disposition == DeliveryChangeDisposition.Unexpected))
        {
            throw new DeliveryReceiptValidationException(
                "unexpected_package_change",
                "At least one package change was not covered by an attribution rule.");
        }

        var artifacts = _artifacts.Values
            .OrderBy(artifact => artifact.ArtifactId, StringComparer.Ordinal)
            .ToArray();
        var evidence = ValidateEvidence().ToArray();
        if (evidence.GroupBy(value => new
            {
                value.Kind,
                value.Schema,
                Digest = value.Digest.Value,
            }).Any(group => group.Count() != 1))
        {
            throw new DeliveryReceiptValidationException(
                "duplicate_evidence", "Receipt evidence identities must be unique.");
        }
        foreach (var artifact in artifacts)
        {
            if (artifact.Role != DeliveryArtifactRole.ReviewDocx
                && artifact.DocumentVersion is { } artifactVersion
                && artifact.PackageDigest is { } artifactPackageDigest
                && !DeliveryReceiptLineageValidator.IsArtifactDocumentReachable(
                    lineageValidation, artifactVersion, artifactPackageDigest, artifacts))
            {
                throw new DeliveryReceiptValidationException(
                    "unreachable_artifact_document",
                    $"Artifact '{artifact.ArtifactId}' is bound to a document identity that "
                    + "is neither a lineage state nor an in-receipt review copy.");
            }
        }
        var boundEvidenceArtifacts = semanticChangeSets
            .Select(binding => binding.ArtifactId)
            .Concat(evidence.Select(item => item.ArtifactId)
                .Where(artifactId => artifactId is not null)
                .Select(artifactId => artifactId!))
            .ToHashSet(StringComparer.Ordinal);
        foreach (var artifact in artifacts)
        {
            if (artifact.Role is DeliveryArtifactRole.SemanticDiff
                    or DeliveryArtifactRole.ValidationReport
                    or DeliveryArtifactRole.ReversibilityProof
                && !boundEvidenceArtifacts.Contains(artifact.ArtifactId))
            {
                throw new DeliveryReceiptValidationException(
                    "unbound_evidence_artifact",
                    $"Artifact '{artifact.ArtifactId}' claims a typed-evidence role but no "
                    + "semantic binding or evidence reference attests it.");
            }
        }
        var citations = _citations
            .Select(citation => BuildPageCitation(citation, lineageValidation))
            .OrderBy(citation => citation.AnchorId, StringComparer.Ordinal)
            .ThenBy(citation => citation.RenderArtifactId, StringComparer.Ordinal)
            .ThenBy(citation => citation.PageMapArtifactId, StringComparer.Ordinal)
            .ToArray();
        if (citations.GroupBy(value => new
            {
                value.AnchorId,
                value.RenderArtifactId,
                value.PageMapArtifactId,
            }).Any(group => group.Count() != 1))
        {
            throw new DeliveryReceiptValidationException(
                "duplicate_page_citation", "Page-citation identities must be unique.");
        }
        var warnings = _warnings
            .Distinct(StringComparer.Ordinal)
            .Select(value => TextEvidence(value, "warning"))
            .OrderBy(value => value.Digest.Value, StringComparer.Ordinal)
            .ToArray();

        var payload = new DeliveryChangeReceiptPayload
        {
            PrivacyProfile = _privacyProfile,
            SourceDocument = _sourceDocument,
            DeliveredDocument = _deliveredDocument,
            Transactions = _transactions.ToArray(),
            Lineage = lineage,
            PackageChanges = packageChanges,
            HasUnexpectedChanges = packageChanges.Any(change =>
                change.Disposition == DeliveryChangeDisposition.Unexpected),
            Evidence = evidence,
            SemanticChangeSets = semanticChangeSets,
            Artifacts = artifacts,
            PageCitations = citations,
            Warnings = warnings,
        };
        DeliveryReceiptResourceValidator.ValidatePayload(payload, _limits);
        return DeliveryChangeReceiptSerializer.Create(payload, _limits);
    }

    private DeliveryTransactionEntry BuildTransactionEntry(
        DeliveryTransactionContribution contribution,
        long sequence)
    {
        var result = contribution.Result;
        var requestFingerprint = contribution.Identity?.RequestFingerprint
            ?? DeriveRequestFingerprint(result.Mode, contribution.Operations);
        DeliveryTransactionContribution.ValidateFingerprint(requestFingerprint);
        var entryId = DeriveEntryId(
            requestFingerprint,
            contribution.BeforeDocument,
            contribution.AfterDocument,
            result.BaseVersion,
            result.ResultVersion,
            contribution.Identity?.TransactionId,
            sequence);

        var stepsByIndex = result.Steps.ToDictionary(step => step.Index);
        var operations = contribution.Operations.Select((operation, index) =>
            BuildOperationEvidence(
                index,
                operation,
                stepsByIndex.TryGetValue(index, out var step) ? step : null)).ToArray();
        var authoredChanges = BuildAuthoredChanges(result).ToArray();
        if (authoredChanges.GroupBy(value => new
            {
                value.EntityKind,
                value.EntityId,
                value.PartUri,
                value.Scope,
            }).Any(group => group.Count() != 1))
        {
            throw new DeliveryReceiptValidationException(
                "duplicate_authored_change", "Authored-change identities must be unique.");
        }
        var warnings = result.Warnings
            .Distinct(StringComparer.Ordinal)
            .Select(value => TextEvidence(value, "transaction warning"))
            .OrderBy(value => value.Digest.Value, StringComparer.Ordinal)
            .ToArray();

        return new DeliveryTransactionEntry
        {
            Sequence = sequence,
            EntryId = entryId,
            TransactionId = contribution.Identity?.TransactionId,
            RequestFingerprint = requestFingerprint,
            Mode = result.Mode,
            Status = TransactionStatus(result),
            BaseVersion = result.BaseVersion,
            ResultVersion = result.ResultVersion,
            BeforeDocument = contribution.BeforeDocument,
            AfterDocument = contribution.AfterDocument,
            ReportedPackageContentDigest = result.PackageHash is null
                ? null
                : new VerificationDigest
                {
                    Algorithm = DeliveryReceiptValidation.Sha256Algorithm,
                    Value = result.PackageHash,
                },
            Operations = operations,
            AuthoredChanges = authoredChanges,
            Warnings = warnings,
        };
    }

    private DeliveryOperationEvidence BuildOperationEvidence(
        int index,
        DeliveryNormalizedOperation operation,
        MutationBatchStepResult? step)
    {
        var propertyNames = new List<string>();
        foreach (var property in operation.Arguments.EnumerateObject())
        {
            if (propertyNames.Count >= _limits.MaxCollectionItems)
            {
                throw new DeliveryReceiptValidationException(
                    "receipt_resource_limit",
                    "Operation argument properties exceed the item limit.");
            }
            propertyNames.Add(property.Name);
        }
        return new DeliveryOperationEvidence
        {
            Index = index,
            Tool = operation.Tool,
            Action = operation.Action,
            ArgumentsDigest = DeliveryReceiptValidation.CloneDigest(operation.ArgumentsDigest),
            ArgumentsSummary = _privacyProfile == DeliveryReceiptPrivacyProfile.HashOnly
                ? null
                : propertyNames.Count == 0
                    ? "no argument properties"
                    : $"{propertyNames.Count} argument properties",
            Arguments = _privacyProfile == DeliveryReceiptPrivacyProfile.FullEvidence
                ? operation.Arguments.Clone()
                : null,
            ExecutionStatus = OperationExecutionStatus(step),
            Success = step?.Success ?? false,
            RolledBack = step?.RolledBack ?? false,
            Results = step?.Results.Select(BuildOperationResult).ToArray()
                ?? Array.Empty<DeliveryOperationResultEvidence>(),
        };
    }

    private static DeliveryOperationExecutionStatus OperationExecutionStatus(
        MutationBatchStepResult? step) => step switch
        {
            null => DeliveryOperationExecutionStatus.NotExecuted,
            { Success: true, RolledBack: false } =>
                DeliveryOperationExecutionStatus.Succeeded,
            { Success: false, RolledBack: false } =>
                DeliveryOperationExecutionStatus.Failed,
            { Success: true, RolledBack: true } =>
                DeliveryOperationExecutionStatus.SucceededRolledBack,
            _ => DeliveryOperationExecutionStatus.FailedRolledBack,
        };

    private DeliveryOperationResultEvidence BuildOperationResult(EditResult result)
    {
        var serialized = SerializeEditResultBounded(result);
        var canonical = DeliveryReceiptCanonicalJson.CanonicalizeBounded(
            serialized, _limits, _limits.MaxReceiptJsonBytes,
            "receipt_resource_limit");
        var changes = new List<DeliveryObjectChange>();
        changes.AddRange(result.Created.Select(anchor => ObjectChange(
            DeliveryObjectChangeKind.Added, anchor)));
        changes.AddRange(result.Removed.Select(anchor => ObjectChange(
            DeliveryObjectChangeKind.Removed, anchor)));
        changes.AddRange(result.Modified.Select(anchor => ObjectChange(
            DeliveryObjectChangeKind.Modified, anchor)));
        if (changes.GroupBy(value => new { value.ChangeKind, value.AnchorId })
            .Any(group => group.Count() != 1))
        {
            throw new DeliveryReceiptValidationException(
                "duplicate_object_change", "Object-change identities must be unique.");
        }

        return new DeliveryOperationResultEvidence
        {
            ResultDigest = DeliveryReceiptCanonicalJson.Digest(canonical),
            Success = result.Success,
            ErrorCode = result.Error is null
                ? null
                : DocxSessionJson.EnumToSnake(result.Error.Code),
            ErrorMessage = result.Error is null
                ? null
                : TextEvidence(result.Error.Message, "structured edit error"),
            ObjectChanges = changes
                .OrderBy(change => change.ChangeKind)
                .ThenBy(change => change.AnchorId, StringComparer.Ordinal)
                .ToArray(),
            FullResult = _privacyProfile == DeliveryReceiptPrivacyProfile.FullEvidence
                ? ParseCanonical(canonical)
                : null,
        };
    }

    private byte[] SerializeEditResultBounded(EditResult result)
    {
        using var stream = new DeliveryReceiptBoundedMemoryStream(
            _limits.MaxReceiptJsonBytes,
            "receipt_resource_limit",
            "Operation result JSON");
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions
        {
            Encoder = JavaScriptEncoder.Default,
            Indented = false,
            MaxDepth = _limits.MaxJsonDepth,
        }))
        {
            WriteEditResult(writer, result);
        }
        return stream.ToArray();
    }

    private static void WriteEditResult(Utf8JsonWriter writer, EditResult result)
    {
        writer.WriteStartObject();
        writer.WriteBoolean("success", result.Success);
        if (result.Error is { } error)
        {
            writer.WriteStartObject("error");
            writer.WriteString("code", DocxSessionJson.EnumToSnake(error.Code));
            writer.WriteString("message", error.Message);
            if (error.AnchorId is not null)
                writer.WriteString("anchorId", error.AnchorId);
            if (error.Precondition is { } precondition)
            {
                writer.WriteStartObject("precondition");
                writer.WriteString("condition", precondition.Condition);
                writer.WritePropertyName("expected");
                WritePreconditionValue(writer, precondition.Expected);
                writer.WritePropertyName("actual");
                WritePreconditionValue(writer, precondition.Actual);
                writer.WriteNumber("currentVersion", precondition.CurrentVersion);
                if (precondition.CurrentTarget is { } target)
                {
                    writer.WriteStartObject("currentTarget");
                    writer.WriteBoolean("exists", target.Exists);
                    WriteOptionalString(writer, "anchorId", target.AnchorId);
                    WriteOptionalString(writer, "kind", target.Kind);
                    WriteOptionalString(writer, "scope", target.Scope);
                    WriteOptionalString(writer, "contentHash", target.ContentHash);
                    WriteOptionalString(writer, "visibleText", target.VisibleText);
                    writer.WriteEndObject();
                }
                writer.WriteEndObject();
            }
            writer.WriteEndObject();
        }
        WriteAnchors(writer, "created", result.Created);
        WriteAnchors(writer, "removed", result.Removed);
        WriteAnchors(writer, "modified", result.Modified);
        if (result.TableAnchors is { } tableAnchors)
            WriteTableAnchorMapping(writer, tableAnchors);
        WriteOptionalString(writer, "annotationId", result.AnnotationId);
        WriteOptionalString(writer, "hyperlinkId", result.HyperlinkId);
        WriteOptionalString(writer, "bookmarkName", result.BookmarkName);
        WriteOptionalString(writer, "imageId", result.ImageId);
        if (result.Patch is { } patch)
        {
            writer.WriteStartObject("patch");
            writer.WriteString("scopeAnchorId", patch.ScopeAnchorId);
            writer.WriteString("markdown", patch.Markdown);
            writer.WriteEndObject();
        }
        writer.WriteEndObject();
    }

    private static void WritePreconditionValue(Utf8JsonWriter writer, object? value)
    {
        switch (value)
        {
            case null:
                writer.WriteNullValue();
                break;
            case string text:
                writer.WriteStringValue(text);
                break;
            case bool boolean:
                writer.WriteBooleanValue(boolean);
                break;
            case byte number:
                writer.WriteNumberValue(number);
                break;
            case sbyte number:
                writer.WriteNumberValue(number);
                break;
            case short number:
                writer.WriteNumberValue(number);
                break;
            case ushort number:
                writer.WriteNumberValue(number);
                break;
            case int number:
                writer.WriteNumberValue(number);
                break;
            case uint number:
                writer.WriteNumberValue(number);
                break;
            case long number:
                writer.WriteNumberValue(number);
                break;
            case ulong number:
                writer.WriteNumberValue(number);
                break;
            case float number:
                writer.WriteNumberValue(number);
                break;
            case double number:
                writer.WriteNumberValue(number);
                break;
            case decimal number:
                writer.WriteNumberValue(number);
                break;
            case JsonElement json:
                json.WriteTo(writer);
                break;
            default:
                throw new DeliveryReceiptValidationException(
                    "invalid_batch_evidence",
                    "Precondition evidence must be a JSON scalar or JsonElement.");
        }
    }

    private static void WriteAnchors(
        Utf8JsonWriter writer,
        string propertyName,
        IReadOnlyList<Anchor> anchors)
    {
        writer.WriteStartArray(propertyName);
        foreach (var anchor in anchors)
            WriteAnchor(writer, anchor);
        writer.WriteEndArray();
    }

    private static void WriteAnchor(Utf8JsonWriter writer, Anchor anchor)
    {
        writer.WriteStartObject();
        writer.WriteString("id", anchor.Id);
        writer.WriteString("kind", anchor.Kind);
        writer.WriteString("scope", anchor.Scope);
        writer.WriteString("unid", anchor.Unid);
        writer.WriteEndObject();
    }

    private static void WriteTableAnchorMapping(
        Utf8JsonWriter writer,
        TableAnchorMapping mapping)
    {
        writer.WriteStartObject("tableAnchors");
        writer.WriteStartArray("retained");
        foreach (var retained in mapping.Retained)
        {
            writer.WriteStartObject();
            writer.WritePropertyName("before");
            WriteTableAnchorLocation(writer, retained.Before);
            writer.WritePropertyName("after");
            WriteTableAnchorLocation(writer, retained.After);
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
        writer.WriteStartArray("added");
        foreach (var added in mapping.Added)
            WriteTableAnchorLocation(writer, added);
        writer.WriteEndArray();
        writer.WriteStartArray("invalidated");
        foreach (var invalidated in mapping.Invalidated)
            WriteTableAnchorLocation(writer, invalidated);
        writer.WriteEndArray();
        writer.WriteEndObject();
    }

    private static void WriteTableAnchorLocation(
        Utf8JsonWriter writer,
        TableAnchorLocation location)
    {
        writer.WriteStartObject();
        writer.WritePropertyName("anchor");
        WriteAnchor(writer, location.Anchor);
        writer.WriteString(
            "entityKind", location.EntityKind.ToString().ToLowerInvariant());
        if (location.RowIndex is { } rowIndex)
            writer.WriteNumber("rowIndex", rowIndex);
        if (location.ColumnIndex is { } columnIndex)
            writer.WriteNumber("columnIndex", columnIndex);
        if (location.RowSpan is { } rowSpan)
            writer.WriteNumber("rowSpan", rowSpan);
        if (location.ColumnSpan is { } columnSpan)
            writer.WriteNumber("columnSpan", columnSpan);
        if (location.IsVirtual)
            writer.WriteBoolean("isVirtual", true);
        writer.WriteEndObject();
    }

    private static void WriteOptionalString(
        Utf8JsonWriter writer,
        string propertyName,
        string? value)
    {
        if (value is not null)
            writer.WriteString(propertyName, value);
    }

    private IEnumerable<DeliveryAuthoredChange> BuildAuthoredChanges(MutationBatchResult result)
    {
        var changes = new List<DeliveryAuthoredChange>();
        AddAuthoredChanges(changes, result.RevisionChanges, BuildRevisionChange);
        AddAuthoredChanges(changes, result.CommentChanges, BuildCommentChange);
        AddAuthoredChanges(changes, result.AnnotationChanges, BuildAnnotationChange);
        return changes
            .OrderBy(change => change.EntityKind)
            .ThenBy(change => change.ChangeKind)
            .ThenBy(change => change.EntityId, StringComparer.Ordinal)
            .ThenBy(change => change.PartUri, StringComparer.Ordinal)
            .ThenBy(change => change.Scope, StringComparer.Ordinal);
    }

    private static void AddAuthoredChanges<T>(
        ICollection<DeliveryAuthoredChange> destination,
        MutationBatchChangeSet<T> source,
        Func<T, DeliveryObjectChangeKind, DeliveryAuthoredChange> map)
    {
        foreach (var item in source.Added)
            destination.Add(map(item, DeliveryObjectChangeKind.Added));
        foreach (var item in source.Removed)
            destination.Add(map(item, DeliveryObjectChangeKind.Removed));
        foreach (var item in source.Modified)
            destination.Add(map(item, DeliveryObjectChangeKind.Modified));
    }

    private DeliveryAuthoredChange BuildRevisionChange(
        RevisionListEntry revision,
        DeliveryObjectChangeKind changeKind)
    {
        var canonical = CanonicalListItem(DocxSessionJson.SerializeRevisionList(new[] { revision }));
        return new DeliveryAuthoredChange
        {
            EntityKind = DeliveryAuthoredEntityKind.Revision,
            ChangeKind = changeKind,
            EntityId = DeliveryReceiptValidation.RequireNonBlank(
                revision.Id, "revision id", 2048),
            SourceDigest = DeliveryReceiptCanonicalJson.Digest(canonical.Bytes),
            Author = revision.Author,
            Date = revision.Date,
            DateUtc = revision.DateUtc,
            Type = revision.Type,
            Family = revision.Family,
            PartUri = revision.PartUri,
            Scope = revision.Scope,
            AnchorId = revision.AnchorId,
            ResolutionStatus = revision.ResolutionStatus,
            Diagnostic = revision.Diagnostic is null
                ? null
                : new DeliveryAuthoredDiagnostic
                {
                    Code = DeliveryReceiptValidation.RequireNonBlank(
                        revision.Diagnostic.Code, "revision diagnostic code", 1024),
                    Message = TextEvidence(
                        revision.Diagnostic.Message, "revision diagnostic message"),
                },
            ConstituentIds = revision.ConstituentIds
                .Distinct(StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray(),
            ConstituentKeys = revision.ConstituentKeys
                .Distinct(StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray(),
            AffectedAnchorIds = revision.AffectedAnchors.Select(anchor => anchor.Id)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray(),
            Text = TextEvidence(revision.Text ?? string.Empty, "revision text"),
            FullEvidence = _privacyProfile == DeliveryReceiptPrivacyProfile.FullEvidence
                ? canonical.Element
                : null,
        };
    }

    private DeliveryAuthoredChange BuildCommentChange(
        CommentListEntry comment,
        DeliveryObjectChangeKind changeKind)
    {
        var canonical = CanonicalListItem(DocxSessionJson.SerializeCommentList(new[] { comment }));
        return new DeliveryAuthoredChange
        {
            EntityKind = DeliveryAuthoredEntityKind.Comment,
            ChangeKind = changeKind,
            EntityId = DeliveryReceiptValidation.RequireNonBlank(
                comment.DefAnchorId, "comment anchor id", 2048),
            SourceDigest = DeliveryReceiptCanonicalJson.Digest(canonical.Bytes),
            Author = comment.Author,
            Date = comment.Date,
            Type = comment.ParentAnchorId is null ? "comment" : "commentReply",
            Scope = ScopeFromAnchor(comment.DefAnchorId),
            AffectedAnchorIds = new[] { comment.DefAnchorId }
                .Concat(comment.ParentAnchorId is null
                    ? Array.Empty<string>()
                    : new[] { comment.ParentAnchorId })
                .Distinct(StringComparer.Ordinal)
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray(),
            Text = TextEvidence(comment.Text ?? string.Empty, "comment text"),
            FullEvidence = _privacyProfile == DeliveryReceiptPrivacyProfile.FullEvidence
                ? canonical.Element
                : null,
        };
    }

    private DeliveryAuthoredChange BuildAnnotationChange(
        DocumentAnnotation annotation,
        DeliveryObjectChangeKind changeKind)
    {
        var canonical = CanonicalListItem(DocxSessionJson.SerializeAnnotations(new[] { annotation }));
        return new DeliveryAuthoredChange
        {
            EntityKind = DeliveryAuthoredEntityKind.Annotation,
            ChangeKind = changeKind,
            EntityId = DeliveryReceiptValidation.RequireNonBlank(
                annotation.Id, "annotation id", 2048),
            SourceDigest = DeliveryReceiptCanonicalJson.Digest(canonical.Bytes),
            Author = annotation.Author,
            Date = annotation.Created is { } created
                ? DocxSessionJson.UtcRoundtrip(created)
                : null,
            Type = annotation.LabelId,
            AffectedAnchorIds = Array.Empty<string>(),
            Text = TextEvidence(annotation.AnnotatedText ?? string.Empty, "annotation text"),
            FullEvidence = _privacyProfile == DeliveryReceiptPrivacyProfile.FullEvidence
                ? canonical.Element
                : null,
        };
    }

    private void RegisterArtifact(DeliveryArtifactInput input)
    {
        if (_artifacts.Count >= _limits.MaxArtifacts)
        {
            throw new DeliveryReceiptValidationException(
                "artifact_resource_limit", "Receipt artifact count limit exceeded.");
        }
        ValidateArtifactResourceLimits(input);
        var artifact = BuildArtifact(input);
        if (!_artifacts.TryAdd(artifact.ArtifactId, artifact))
        {
            throw new DeliveryReceiptValidationException(
                "duplicate_artifact_id", $"Artifact id '{artifact.ArtifactId}' is duplicated.");
        }
        if (input.Bytes is not null)
        {
            _totalArtifactBytes = checked(_totalArtifactBytes + input.Bytes.LongLength);
            _artifactBytes.Add(artifact.ArtifactId, input.Bytes.ToArray());
        }
    }

    private void ValidateArtifactResourceLimits(DeliveryArtifactInput input)
    {
        CheckString(input.ArtifactId, "artifact id");
        CheckString(input.MediaType, "artifact media type");
        CheckString(input.RelativePath, "artifact path");
        CheckString(input.UnavailableReason, "artifact unavailable reason");
        CheckString(input.RendererFingerprint, "renderer fingerprint");
        if (input.Bytes is null)
            return;
        DeliveryReceiptResourceBudget.Bytes(
            input.Bytes.LongLength,
            _limits.MaxArtifactBytes,
            "artifact_resource_limit",
            "Artifact");
        if (input.Role == DeliveryArtifactRole.SemanticDiff)
        {
            DeliveryReceiptResourceBudget.Bytes(
                input.Bytes.LongLength,
                _limits.MaxSemanticEvidenceBytes,
                "semantic_resource_limit",
                "Semantic artifact");
        }
        if (input.Role == DeliveryArtifactRole.PageMap)
        {
            DeliveryReceiptResourceBudget.Bytes(
                input.Bytes.LongLength,
                _limits.MaxPageMapBytes,
                "page_map_resource_limit",
                "PageMap artifact");
        }
        if (input.Bytes.LongLength > _limits.MaxTotalArtifactBytes - _totalArtifactBytes)
        {
            throw new DeliveryReceiptValidationException(
                "artifact_resource_limit", "Total receipt artifact byte limit exceeded.");
        }
    }

    private DeliveryArtifact BuildArtifact(DeliveryArtifactInput input)
    {
        if (!Enum.IsDefined(input.Role) || !Enum.IsDefined(input.Availability))
            throw new DeliveryReceiptValidationException("invalid_artifact_enum", "Unknown artifact value.");
        var artifactId = DeliveryReceiptValidation.RequireNonBlank(
            input.ArtifactId, "artifact id", 256);
        var mediaType = DeliveryReceiptValidation.RequireNonBlank(
            input.MediaType, "artifact media type", 512);
        var path = DeliveryReceiptValidation.NormalizeRelativePath(input.RelativePath);
        DeliveryReceiptValidation.ValidateOptionalDigest(input.PageMapDigest, "artifact page-map digest");
        ValidateInputDocument(input.Document);
        if (input.RendererFingerprint is not null)
            DeliveryReceiptValidation.RequireNonBlank(
                input.RendererFingerprint, "renderer fingerprint", 4096);

        if (input.Availability == DeliveryArtifactAvailability.Available)
        {
            if (input.Bytes is null)
            {
                throw new DeliveryReceiptValidationException(
                    "missing_artifact_bytes", $"Available artifact '{artifactId}' has no bytes.");
            }
            if (input.UnavailableReason is not null)
            {
                throw new DeliveryReceiptValidationException(
                    "inconsistent_artifact_availability",
                    $"Available artifact '{artifactId}' cannot have an unavailable reason.");
            }
            var digest = DeliveryReceiptCanonicalJson.Digest(input.Bytes);
            var document = input.Document;
            if (input.Role == DeliveryArtifactRole.CleanDocx)
            {
                if (document is null)
                {
                    throw new DeliveryReceiptValidationException(
                        "clean_docx_delivery_mismatch",
                        "The clean DOCX requires a document identity and exact bytes.");
                }
                var manifest = PackageManifestGenerator.Generate(
                    input.Bytes, _limits.CleanDocxManifestOptions);
                if (!manifest.IsValid
                    || !string.Equals(manifest.PackageKind, "opc", StringComparison.Ordinal)
                    || string.IsNullOrWhiteSpace(manifest.Facts.MainDocumentUri))
                {
                    throw new DeliveryReceiptValidationException(
                        "invalid_clean_docx",
                        "Clean DOCX bytes must be a valid bounded WordprocessingML OPC package.");
                }
                var actual = DeliveryDocumentIdentity.FromManifest(
                    manifest, document.DocumentVersion);
                if (!DeliveryReceiptLineageValidator.DocumentEquals(actual, document))
                {
                    throw new DeliveryReceiptValidationException(
                        "clean_docx_delivery_mismatch",
                        "Caller-supplied clean DOCX identity does not match recomputed bytes.");
                }
                document = actual;
            }
            if (document is not null
                && input.Role is DeliveryArtifactRole.CleanDocx or DeliveryArtifactRole.ReviewDocx
                && !DeliveryReceiptValidation.DigestEquals(
                    digest, document.RawPackageBytesDigest))
            {
                throw new DeliveryReceiptValidationException(
                    "docx_artifact_identity_mismatch",
                    $"DOCX artifact '{artifactId}' does not match its package identity.");
            }
            return new DeliveryArtifact
            {
                ArtifactId = artifactId,
                Role = input.Role,
                MediaType = mediaType,
                Availability = input.Availability,
                ByteLength = input.Bytes.LongLength,
                Digest = digest,
                RelativePath = path,
                DocumentVersion = document?.DocumentVersion,
                PackageDigest = DeliveryReceiptValidation.CloneOptionalDigest(
                    document?.RawPackageBytesDigest),
                RendererFingerprint = input.RendererFingerprint,
                PageMapDigest = DeliveryReceiptValidation.CloneOptionalDigest(input.PageMapDigest),
            };
        }

        if (input.Bytes is not null)
        {
            throw new DeliveryReceiptValidationException(
                "inconsistent_artifact_availability",
                $"Unavailable artifact '{artifactId}' cannot carry bytes.");
        }
        return new DeliveryArtifact
        {
            ArtifactId = artifactId,
            Role = input.Role,
            MediaType = mediaType,
            Availability = input.Availability,
            UnavailableReason = ProfiledFreeText(
                DeliveryReceiptValidation.RequireNonBlank(
                    input.UnavailableReason, "artifact unavailable reason", 4096),
                "artifact unavailable reason"),
            RelativePath = path,
            DocumentVersion = input.Document?.DocumentVersion,
            PackageDigest = DeliveryReceiptValidation.CloneOptionalDigest(
                input.Document?.RawPackageBytesDigest),
            RendererFingerprint = input.RendererFingerprint,
            PageMapDigest = DeliveryReceiptValidation.CloneOptionalDigest(input.PageMapDigest),
        };
    }

    private static void ValidateInputDocument(DeliveryDocumentIdentity? document)
    {
        if (document is null)
            return;
        DeliveryReceiptValidation.ValidatePortableNonNegativeInteger(
            document.DocumentVersion, "invalid_document_version", "Artifact document version");
        DeliveryReceiptValidation.RequireNonBlank(
            document.PackageKind, "artifact package kind", 256);
        if (!string.Equals(document.PackageKind, "opc", StringComparison.Ordinal))
        {
            throw new DeliveryReceiptValidationException(
                "not_wordprocessing_package", "Artifact document must be an OPC package.");
        }
        DeliveryReceiptValidation.RequireOpcMainDocumentUri(
            document.MainDocumentUri, "artifact main document URI");
        if (!DeliveryPackageManifestAdapter.IsSupportedSchema(document.PackageManifestSchema))
        {
            throw new DeliveryReceiptValidationException(
                "unsupported_package_manifest", "Artifact package-manifest schema is unsupported.");
        }
        DeliveryReceiptValidation.ValidateDigest(
            document.RawPackageBytesDigest, "artifact package digest");
        DeliveryReceiptValidation.ValidateOptionalDigest(
            document.OrderedOpcContentDigest, "artifact content digest");
        DeliveryReceiptValidation.ValidateOptionalDigest(
            document.NormalizedSemanticDigest, "artifact semantic digest");
    }

    private static PackageManifest CloneManifest(PackageManifest manifest) => manifest with
    {
        RawPackageBytesDigest = DeliveryReceiptValidation.CloneDigest(
            manifest.RawPackageBytesDigest),
        OrderedOpcContentDigest = DeliveryReceiptValidation.CloneOptionalDigest(
            manifest.OrderedOpcContentDigest),
        NormalizedSemanticDigest = DeliveryReceiptValidation.CloneOptionalDigest(
            manifest.NormalizedSemanticDigest),
        Entries = manifest.Entries.Select(entry => entry with
        {
            RawBytesDigest = DeliveryReceiptValidation.CloneOptionalDigest(entry.RawBytesDigest),
            NormalizedXmlDigest = DeliveryReceiptValidation.CloneOptionalDigest(
                entry.NormalizedXmlDigest),
        }).ToArray(),
        ContentTypes = manifest.ContentTypes.Select(value => value with { }).ToArray(),
        Relationships = manifest.Relationships.Select(value => value with { }).ToArray(),
        Facts = manifest.Facts with
        {
            Revisions = manifest.Facts.Revisions with { },
            Annotations = manifest.Facts.Annotations with { },
        },
        Findings = manifest.Findings.Select(finding => finding with
        {
            Location = finding.Location is null ? null : finding.Location with { },
        }).ToArray(),
    };

    private IReadOnlyList<DeliveryLineageEvent> MaterializeLineage()
    {
        return _lineage.Select(value =>
            new DeliveryLineageEvent
            {
                Sequence = value.Sequence,
                Action = value.Input.Action,
                AffectedEntryId = value.Input.AffectedEntryId,
                BeforeDocument = value.Input.BeforeDocument,
                AfterDocument = value.Input.AfterDocument,
            })
            .OrderBy(value => value.Sequence)
            .ToArray();
    }

    private PackageManifest ValidateRequiredCleanDocx(
        DeliveryDocumentIdentity deliveredDocument)
    {
        var cleanArtifacts = _artifacts.Values
            .Where(artifact => artifact.Role == DeliveryArtifactRole.CleanDocx)
            .ToArray();
        if (cleanArtifacts.Length != 1)
        {
            throw new DeliveryReceiptValidationException(
                cleanArtifacts.Length == 0 ? "missing_clean_docx" : "multiple_clean_docx_artifacts",
                "A delivery receipt requires exactly one clean DOCX artifact.");
        }

        var clean = cleanArtifacts[0];
        if (clean.Availability != DeliveryArtifactAvailability.Available
            || clean.Digest is null
            || !_artifactBytes.TryGetValue(clean.ArtifactId, out var cleanBytes)
            || clean.DocumentVersion != deliveredDocument.DocumentVersion
            || !DeliveryReceiptValidation.DigestEquals(
                clean.PackageDigest, deliveredDocument.RawPackageBytesDigest)
            || !DeliveryReceiptValidation.DigestEquals(
                clean.Digest, deliveredDocument.RawPackageBytesDigest))
        {
            throw new DeliveryReceiptValidationException(
                "clean_docx_delivery_mismatch",
                "The clean DOCX artifact must be available and exactly match DeliveredDocument.");
        }
        var actualManifest = PackageManifestGenerator.Generate(
            cleanBytes, _limits.CleanDocxManifestOptions);
        if (!actualManifest.IsValid
            || !string.Equals(actualManifest.PackageKind, "opc", StringComparison.Ordinal)
            || string.IsNullOrWhiteSpace(actualManifest.Facts.MainDocumentUri)
            || !DeliveryReceiptLineageValidator.DocumentEquals(
                DeliveryDocumentIdentity.FromManifest(
                    actualManifest, deliveredDocument.DocumentVersion),
                deliveredDocument))
        {
            throw new DeliveryReceiptValidationException(
                "clean_docx_delivery_mismatch",
                "Recomputed clean DOCX package identity does not match DeliveredDocument.");
        }
        ValidateManifestResources(actualManifest);
        return actualManifest;
    }

    private IReadOnlyList<DeliverySemanticChangeSetBinding> BuildSemanticChangeSetBindings(
        DeliveryLineageValidationResult lineageValidation)
    {
        var aggregate = _semanticChangeSets
            .Where(entry => entry.Input.Scope
                == DeliverySemanticComparisonScope.SourceToDelivered)
            .ToArray();
        if (aggregate.Length != 1)
        {
            throw new DeliveryReceiptValidationException(
                "missing_source_to_delivered_semantic_evidence",
                "Exactly one source-to-delivered SemanticChangeSet is required.");
        }

        var transactionInputs = _semanticChangeSets
            .Where(entry => entry.Input.Scope == DeliverySemanticComparisonScope.Transaction)
            .GroupBy(entry => entry.Input.TransactionEntryId!, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.Ordinal);
        var expectedTransactions = lineageValidation.StateChangingTransactions
            .ToDictionary(transaction => transaction.EntryId, StringComparer.Ordinal);
        if (transactionInputs.Any(pair => pair.Value.Length != 1
                || !expectedTransactions.ContainsKey(pair.Key))
            || expectedTransactions.Keys.Any(entryId => !transactionInputs.ContainsKey(entryId)))
        {
            throw new DeliveryReceiptValidationException(
                "semantic_transaction_coverage_mismatch",
                "Every state-changing transaction requires exactly one typed semantic binding.");
        }

        var output = new List<DeliverySemanticChangeSetBinding>
        {
            BuildSemanticBinding(aggregate[0], _sourceDocument, _deliveredDocument!, null),
        };
        output.AddRange(lineageValidation.StateChangingTransactions.Select(transaction =>
            BuildSemanticBinding(
                transactionInputs[transaction.EntryId][0],
                transaction.BeforeDocument,
                transaction.AfterDocument,
                transaction.EntryId)));
        return output;
    }

    private DeliverySemanticChangeSetBinding BuildSemanticBinding(
        (DeliverySemanticChangeSetInput Input,
            DeliverySemanticChangeSetProjection Projection) entry,
        DeliveryDocumentIdentity before,
        DeliveryDocumentIdentity after,
        string? transactionEntryId)
    {
        var input = entry.Input;
        var projection = entry.Projection;
        if (!_artifacts.TryGetValue(input.ArtifactId, out var artifact)
            || artifact.Role != DeliveryArtifactRole.SemanticDiff
            || artifact.Availability != DeliveryArtifactAvailability.Available
            || artifact.Digest is null
            || !DeliveryReceiptValidation.DigestEquals(artifact.Digest, projection.Digest)
            || !_artifactBytes.TryGetValue(input.ArtifactId, out var bytes)
            || !bytes.AsSpan().SequenceEqual(projection.CanonicalBytes))
        {
            throw new DeliveryReceiptValidationException(
                "semantic_artifact_binding_mismatch",
                $"Semantic artifact '{input.ArtifactId}' does not contain the exact #457 canonical bytes.");
        }

        return new DeliverySemanticChangeSetBinding
        {
            Scope = input.Scope,
            TransactionEntryId = transactionEntryId,
            BeforeDocument = before,
            AfterDocument = after,
            Schema = projection.Schema,
            SchemaVersion = projection.SchemaVersion,
            ChangeCount = projection.ChangeCount,
            Digest = DeliveryReceiptValidation.CloneDigest(projection.Digest),
            ArtifactId = input.ArtifactId,
        };
    }

    private IReadOnlyList<DeliveryPackageChange> BuildPackageChanges(
        PackageManifest before,
        PackageManifest after,
        DeliveryLineageValidationResult lineageValidation)
    {
        return DeliveryPackageManifestAdapter.Compare(
                before, after, _limits.MaxCollectionItems)
            .Select(candidate => BuildPackageChange(candidate, lineageValidation))
            .ToArray();
    }

    private DeliveryPackageChange BuildPackageChange(
        DeliveryPackageChangeObservation candidate,
        DeliveryLineageValidationResult lineageValidation)
    {
        _attributionRules.TryGetValue(AttributionKey(candidate), out var rule);
        if (rule is not null && rule.TransactionEntryId is not null)
        {
            if (!_entriesById.TryGetValue(rule.TransactionEntryId, out var entry))
            {
                throw new DeliveryReceiptValidationException(
                    "unknown_attribution_transaction",
                    $"Attribution references unknown entry '{rule.TransactionEntryId}'.");
            }
            if (entry.Status is not (DeliveryTransactionStatus.Committed
                or DeliveryTransactionStatus.PartiallyCommitted))
            {
                throw new DeliveryReceiptValidationException(
                    "invalid_attribution_transaction",
                    "Package-change attribution must reference a committed transaction.");
            }
            if (!lineageValidation.AppliedTransactionEntryIds.Contains(entry.EntryId))
            {
                throw new DeliveryReceiptValidationException(
                    "unapplied_attribution_transaction",
                    "Package-change attribution must reference a transaction applied in the delivered state.");
            }
            if (rule.RequestedOperationIndex is { } index
                && (index < 0 || index >= entry.Operations.Count))
            {
                throw new DeliveryReceiptValidationException(
                    "invalid_attribution_operation",
                    "Attribution operation index is outside the referenced transaction.");
            }
            if (rule.RequestedOperationIndex is { } requestedIndex
                && entry.Operations[requestedIndex].ExecutionStatus
                    != DeliveryOperationExecutionStatus.Succeeded)
            {
                throw new DeliveryReceiptValidationException(
                    "invalid_attribution_operation",
                    "Attribution may reference only a successful, retained operation.");
            }
        }

        var beforeEvidence = candidate.Before is null
            ? null
            : TextEvidence(candidate.Before, "prior package record");
        var afterEvidence = candidate.After is null
            ? null
            : TextEvidence(candidate.After, "result package record");
        return new DeliveryPackageChange
        {
            ChangeId = DeliveryReceiptIdentity.PackageChangeId(
                candidate.Kind,
                candidate.Location,
                beforeEvidence?.Digest,
                afterEvidence?.Digest),
            Kind = candidate.Kind,
            Location = candidate.Location,
            Before = beforeEvidence,
            After = afterEvidence,
            Disposition = rule?.Disposition ?? DeliveryChangeDisposition.Unexpected,
            TransactionEntryId = rule?.TransactionEntryId,
            RequestedOperationIndex = rule?.RequestedOperationIndex,
            Derivation = rule?.Derivation is null
                ? null
                : ProfiledFreeText(rule.Derivation, "derivation"),
        };
    }

    private IEnumerable<DeliveryEvidenceReference> ValidateEvidence()
    {
        foreach (var reference in _evidence
                     .OrderBy(value => value.Kind)
                     .ThenBy(value => value.Schema, StringComparer.Ordinal)
                     .ThenBy(value => value.Digest.Value, StringComparer.Ordinal))
        {
            if (reference.ArtifactId is { } artifactId)
            {
                if (!_artifacts.TryGetValue(artifactId, out var artifact)
                    || artifact.Availability != DeliveryArtifactAvailability.Available
                    || artifact.Digest is null)
                {
                    throw new DeliveryReceiptValidationException(
                        "missing_evidence_artifact",
                        $"Evidence artifact '{artifactId}' is not available.");
                }
                if (!DeliveryReceiptValidation.DigestEquals(reference.Digest, artifact.Digest))
                {
                    throw new DeliveryReceiptValidationException(
                        "evidence_artifact_digest_mismatch",
                        $"Evidence digest does not match artifact '{artifactId}'.");
                }
                var expectedRole = reference.Kind switch
                {
                    DeliveryEvidenceKind.ValidationResult => DeliveryArtifactRole.ValidationReport,
                    DeliveryEvidenceKind.RedlineReversibility => DeliveryArtifactRole.ReversibilityProof,
                    DeliveryEvidenceKind.SemanticChangeSet =>
                        throw new DeliveryReceiptValidationException(
                            "semantic_evidence_requires_typed_factory",
                            "Semantic evidence must use the typed #457 binding."),
                    _ => throw new DeliveryReceiptValidationException(
                        "unknown_evidence_kind", "Unknown evidence kind."),
                };
                if (artifact.Role != expectedRole)
                {
                    throw new DeliveryReceiptValidationException(
                        "evidence_artifact_role_mismatch",
                        $"Evidence '{reference.Kind}' requires artifact role '{expectedRole}'.");
                }
                if (!_artifactBytes.TryGetValue(artifactId, out var evidenceBytes)
                    || !IsExactEvidence(reference.Kind, evidenceBytes))
                {
                    throw new DeliveryReceiptValidationException(
                        "invalid_evidence_artifact",
                        $"Evidence artifact '{artifactId}' is not the exact canonical owner contract.");
                }
            }
            yield return reference;
        }
    }

    private static string ExpectedEvidenceSchema(DeliveryEvidenceKind kind) => kind switch
    {
        DeliveryEvidenceKind.ValidationResult => DeliverableVerificationResult.SchemaId,
        DeliveryEvidenceKind.RedlineReversibility => RedlineReversibilityProof.SchemaId,
        DeliveryEvidenceKind.SemanticChangeSet =>
            throw new DeliveryReceiptValidationException(
                "semantic_evidence_requires_typed_factory",
                "Semantic evidence must use the typed #457 binding."),
        _ => throw new DeliveryReceiptValidationException(
            "unknown_evidence_kind", "Unknown evidence kind."),
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

    private DeliveryPageCitation BuildPageCitation(
        DeliveryPageCitationInput input,
        DeliveryLineageValidationResult lineageValidation)
    {
        ArgumentNullException.ThrowIfNull(input.Citation);
        ArgumentNullException.ThrowIfNull(input.Document);
        if (!lineageValidation.ReachableDocumentsByVersion.TryGetValue(
                input.Document.DocumentVersion, out var reachableDocument)
            || !DeliveryReceiptLineageValidator.DocumentEquals(
                reachableDocument, input.Document))
        {
            throw new DeliveryReceiptValidationException(
                "unreachable_citation_document",
                "Citation document identity is not a reachable receipt-owned state.");
        }
        DeliveryReceiptValidation.ValidateDigest(input.PageMapDigest, "page-map digest");
        var pageMapArtifactId = DeliveryReceiptValidation.RequireNonBlank(
            input.PageMapArtifactId, "page-map artifact id", 256);
        var artifactId = DeliveryReceiptValidation.RequireNonBlank(
            input.RenderArtifactId, "render artifact id", 256);
        if (!_artifacts.TryGetValue(artifactId, out var artifact)
            || artifact.Availability != DeliveryArtifactAvailability.Available
            || artifact.Digest is null)
        {
            throw new DeliveryReceiptValidationException(
                "missing_citation_artifact", $"Citation artifact '{artifactId}' is unavailable.");
        }
        if (artifact.Role is not (DeliveryArtifactRole.Pdf
            or DeliveryArtifactRole.RenderReport))
        {
            throw new DeliveryReceiptValidationException(
                "non_paginated_citation_artifact",
                "Page citations require PDF or paginated render-report evidence.");
        }
        if (!_artifacts.TryGetValue(pageMapArtifactId, out var pageMapArtifact)
            || pageMapArtifact.Availability != DeliveryArtifactAvailability.Available
            || pageMapArtifact.Role != DeliveryArtifactRole.PageMap
            || pageMapArtifact.Digest is null
            || !DeliveryReceiptValidation.DigestEquals(
                pageMapArtifact.Digest, input.PageMapDigest))
        {
            throw new DeliveryReceiptValidationException(
                "missing_page_map_artifact",
                $"Page-map artifact '{pageMapArtifactId}' is unavailable or has the wrong digest.");
        }
        if (!_artifactBytes.TryGetValue(pageMapArtifactId, out var pageMapBytes))
        {
            throw new DeliveryReceiptValidationException(
                "missing_page_map_artifact_bytes",
                $"Page-map artifact '{pageMapArtifactId}' has no verifiable bytes.");
        }
        var citation = input.Citation;
        if (citation.Availability != PageMapAvailability.Available
            || citation.UnavailableReason is not null
            || citation.Pages.Count == 0
            || citation.Fragments.Count == 0)
        {
            throw new DeliveryReceiptValidationException(
                "unavailable_page_citation",
                "Only available citations from a paginated page map can enter a receipt.");
        }
        var scope = DeliveryReceiptValidation.RequireNonBlank(input.Scope, "citation scope", 1024);
        if (!string.Equals(ScopeFromAnchor(citation.AnchorId), scope, StringComparison.Ordinal))
        {
            throw new DeliveryReceiptValidationException(
                "citation_scope_mismatch", "Citation scope does not match its canonical anchor.");
        }
        if (citation.DocumentVersion != input.Document.DocumentVersion)
        {
            throw new DeliveryReceiptValidationException(
                "citation_document_version_mismatch", "Citation document version is stale.");
        }
        var rendererFingerprint = DeliveryReceiptValidation.RequireNonBlank(
            citation.RendererFingerprint, "citation renderer fingerprint", 4096);

        if (artifact.DocumentVersion != input.Document.DocumentVersion
            || !DeliveryReceiptValidation.DigestEquals(
                artifact.PackageDigest, input.Document.RawPackageBytesDigest)
            || !string.Equals(artifact.RendererFingerprint, rendererFingerprint,
                StringComparison.Ordinal)
            || !DeliveryReceiptValidation.DigestEquals(
                artifact.PageMapDigest, input.PageMapDigest))
        {
            throw new DeliveryReceiptValidationException(
                "citation_render_binding_mismatch",
                "Citation package, renderer, or page-map identity does not match its artifact.");
        }
        if (pageMapArtifact.DocumentVersion != input.Document.DocumentVersion
            || !DeliveryReceiptValidation.DigestEquals(
                pageMapArtifact.PackageDigest, input.Document.RawPackageBytesDigest)
            || !string.Equals(pageMapArtifact.RendererFingerprint, rendererFingerprint,
                StringComparison.Ordinal))
        {
            throw new DeliveryReceiptValidationException(
                "citation_page_map_binding_mismatch",
                "Page-map artifact does not match the citation document and renderer.");
        }
        if (!_validatedPageMaps.TryGetValue(pageMapArtifactId, out var pageMap))
        {
            try
            {
                // Canonicalization is only a strict UTF-8/JSON/duplicate-property gate. The
                // artifact digest above remains over the renderer's original bytes.
                _ = DeliveryReceiptCanonicalJson.CanonicalizeBounded(
                    pageMapBytes, _limits, _limits.MaxPageMapBytes,
                    "page_map_resource_limit");
                pageMap = DocxSessionJson.ParsePageMap(Encoding.UTF8.GetString(pageMapBytes));
            }
            catch (Exception ex) when (ex is JsonException or FormatException)
            {
                throw new DeliveryReceiptValidationException(
                    "invalid_page_map_artifact",
                    $"Page-map artifact '{pageMapArtifactId}' is not a valid PageMap: {ex.Message}");
            }
            var mapValidation = PageMapContract.ValidatePortable(
                pageMap, input.Document.DocumentVersion, rendererFingerprint);
            if (!mapValidation.Success)
            {
                throw new DeliveryReceiptValidationException(
                    "invalid_page_map_artifact",
                    mapValidation.Message ?? "Page-map artifact is invalid.");
            }
            _validatedPageMaps.Add(pageMapArtifactId, pageMap);
        }
        var projectionKey = (pageMapArtifactId, citation.AnchorId);
        if (!_pageMapProjections.TryGetValue(projectionKey, out var projected))
        {
            projected = PageMapContract.ProjectCitation(
                pageMap,
                citation.AnchorId,
                new PageCitationRequest(citation.DocumentVersion, rendererFingerprint));
            _pageMapProjections.Add(projectionKey, projected);
        }
        if (projected.Availability != PageMapAvailability.Available
            || !CitationCoordinatesEqual(citation, projected))
        {
            throw new DeliveryReceiptValidationException(
                "citation_page_map_projection_mismatch",
                "Citation pages and fragments are not the exact projection of its PageMap bytes.");
        }
        if (projected.Fragments.Any(fragment =>
                !PageMapContract.StoryMatchesScope(fragment.Story, scope)))
        {
            throw new DeliveryReceiptValidationException(
                "citation_scope_mismatch",
                "Citation fragment story does not match its canonical anchor scope.");
        }
        return new DeliveryPageCitation
        {
            AnchorId = citation.AnchorId,
            Scope = scope,
            DocumentVersion = citation.DocumentVersion,
            PackageDigest = DeliveryReceiptValidation.CloneDigest(
                input.Document.RawPackageBytesDigest),
            RendererFingerprint = rendererFingerprint,
            PageMapDigest = DeliveryReceiptValidation.CloneDigest(input.PageMapDigest),
            PageMapArtifactId = pageMapArtifactId,
            RenderArtifactId = artifactId,
            RenderArtifactDigest = DeliveryReceiptValidation.CloneDigest(artifact.Digest),
            Pages = projected.Pages.ToArray(),
            Fragments = projected.Fragments.ToArray(),
        };
    }

    private static bool CitationCoordinatesEqual(PageCitation left, PageCitation right) =>
        string.Equals(left.AnchorId, right.AnchorId, StringComparison.Ordinal)
        && left.Availability == right.Availability
        && left.UnavailableReason == right.UnavailableReason
        && left.DocumentVersion == right.DocumentVersion
        && string.Equals(left.RendererFingerprint, right.RendererFingerprint,
            StringComparison.Ordinal)
        && left.Pages.SequenceEqual(right.Pages)
        && left.Fragments.SequenceEqual(right.Fragments);

    private DeliveryTextEvidence TextEvidence(string value, string structuralSummary)
    {
        value ??= string.Empty;
        return new DeliveryTextEvidence
        {
            Digest = DeliveryReceiptCanonicalJson.DigestText(value),
            CharacterCount = value.Length,
            Summary = _privacyProfile == DeliveryReceiptPrivacyProfile.HashOnly
                ? null
                : $"{structuralSummary}; {value.Length.ToString(CultureInfo.InvariantCulture)} characters",
            Value = _privacyProfile == DeliveryReceiptPrivacyProfile.FullEvidence ? value : null,
        };
    }

    private string ProfiledFreeText(string value, string label)
    {
        var digest = DeliveryReceiptCanonicalJson.DigestToken(Encoding.UTF8.GetBytes(value));
        return _privacyProfile switch
        {
            DeliveryReceiptPrivacyProfile.HashOnly => digest,
            DeliveryReceiptPrivacyProfile.HashAndSummary =>
                $"{label}; {value.Length.ToString(CultureInfo.InvariantCulture)} characters; {digest}",
            _ => value,
        };
    }

    private static DeliveryObjectChange ObjectChange(
        DeliveryObjectChangeKind changeKind,
        Anchor anchor)
    {
        var id = DeliveryReceiptValidation.RequireNonBlank(anchor.Id, "anchor id", 4096);
        var kind = DeliveryReceiptValidation.RequireNonBlank(anchor.Kind, "anchor kind", 256);
        var scope = DeliveryReceiptValidation.RequireNonBlank(anchor.Scope, "anchor scope", 1024);
        var unid = DeliveryReceiptValidation.RequireNonBlank(anchor.Unid, "anchor unid", 2048);
        if (!string.Equals(id, $"{kind}:{scope}:{unid}", StringComparison.Ordinal))
        {
            throw new DeliveryReceiptValidationException(
                "invalid_anchor_id", "Changed anchor identity is inconsistent.");
        }
        return new DeliveryObjectChange
        {
            ChangeKind = changeKind,
            AnchorId = id,
            Kind = kind,
            Scope = scope,
            Unid = unid,
        };
    }

    private static DeliveryTransactionStatus TransactionStatus(MutationBatchResult result)
    {
        if (result.Preview)
            return DeliveryTransactionStatus.Prediction;
        if (result.Mode == MutationBatchMode.Atomic)
            return result.Success && !result.RolledBack
                ? DeliveryTransactionStatus.Committed
                : DeliveryTransactionStatus.Failed;
        if (result.Success)
            return DeliveryTransactionStatus.Committed;
        return !result.RolledBack && result.Steps.Any(step => step.Success)
            ? DeliveryTransactionStatus.PartiallyCommitted
            : DeliveryTransactionStatus.Failed;
    }

    private string DeriveRequestFingerprint(
        MutationBatchMode mode,
        IReadOnlyList<DeliveryNormalizedOperation> operations)
    {
        return DeliveryReceiptIdentity.RequestFingerprint(mode, operations, _limits);
    }

    private static string DeriveEntryId(
        string requestFingerprint,
        DeliveryDocumentIdentity before,
        DeliveryDocumentIdentity after,
        long baseVersion,
        long resultVersion,
        string? transactionId,
        long sequence)
    {
        return DeliveryReceiptIdentity.TransactionEntryId(
            requestFingerprint,
            before,
            after,
            baseVersion,
            resultVersion,
            transactionId,
            sequence);
    }

    private bool TransactionEvidenceEquals(
        DeliveryTransactionEntry left,
        DeliveryTransactionEntry right)
    {
        var leftBytes = DeliveryReceiptIdentity.TransactionEvidence(left, _limits);
        var rightBytes = DeliveryReceiptIdentity.TransactionEvidence(right, _limits);
        return leftBytes.AsSpan().SequenceEqual(rightBytes);
    }

    private void ValidateContributionResources(DeliveryTransactionContribution contribution)
    {
        if (contribution.Operations.Count > _limits.MaxOperationsPerTransaction)
        {
            throw new DeliveryReceiptValidationException(
                "receipt_resource_limit", "Transaction operation limit exceeded.");
        }
        var budget = new DeliveryReceiptResourceBudget(_limits);
        budget.AddItems(contribution.Operations.Count, "transaction operations");
        foreach (var operation in contribution.Operations)
        {
            budget.String(operation.Tool, "operation tool");
            budget.String(operation.Action, "operation action");
            DeliveryReceiptResourceBudget.Bytes(
                operation.CanonicalArguments.LongLength,
                _limits.MaxStringLength,
                "receipt_resource_limit",
                "Canonical operation arguments");
            budget.AddSerializedBytes(
                operation.CanonicalArguments.LongLength,
                "Canonical operation arguments");
            AccountJsonElement(
                operation.Arguments, budget, depth: 1, "operation arguments");
        }
        budget.AddItems(contribution.Result.Steps.Count, "batch step results");
        foreach (var step in contribution.Result.Steps)
        {
            budget.String(step.Tool, "batch step tool");
            budget.String(step.Action, "batch step action");
            budget.AddItems(step.Results.Count, "operation results");
            foreach (var result in step.Results)
            {
                budget.AddSerializedBytes(256, "operation result structure");
                budget.AddItems(result.Created.Count, "created anchors");
                budget.AddItems(result.Removed.Count, "removed anchors");
                budget.AddItems(result.Modified.Count, "modified anchors");
                foreach (var anchor in result.Created
                    .Concat(result.Removed)
                    .Concat(result.Modified))
                {
                    AccountAnchor(anchor, budget);
                }
                budget.String(result.Error?.Message, "edit error message");
                budget.String(result.Error?.AnchorId, "edit error anchor");
                if (result.Error?.Precondition is { } precondition)
                {
                    budget.AddSerializedBytes(128, "precondition evidence structure");
                    budget.String(precondition.Condition, "precondition condition");
                    AccountPreconditionValue(
                        precondition.Expected, budget, "precondition expected value");
                    AccountPreconditionValue(
                        precondition.Actual, budget, "precondition actual value");
                    if (precondition.CurrentTarget is { } target)
                    {
                        budget.String(target.AnchorId, "precondition target anchor");
                        budget.String(target.Kind, "precondition target kind");
                        budget.String(target.Scope, "precondition target scope");
                        budget.String(target.ContentHash, "precondition target content hash");
                        budget.String(target.VisibleText, "precondition target text");
                    }
                }
                budget.String(result.AnnotationId, "annotation id");
                budget.String(result.HyperlinkId, "hyperlink id");
                budget.String(result.BookmarkName, "bookmark name");
                budget.String(result.ImageId, "image id");
                if (result.Patch is { } patch)
                {
                    budget.String(patch.ScopeAnchorId, "patch scope anchor");
                    budget.String(patch.Markdown, "patch markdown");
                }
                if (result.TableAnchors is { } tableAnchors)
                    AccountTableAnchors(tableAnchors, budget);
            }
        }
        budget.AddItems(contribution.Result.RevisionChanges.Added.Count, "revision changes");
        budget.AddItems(contribution.Result.RevisionChanges.Removed.Count, "revision changes");
        budget.AddItems(contribution.Result.RevisionChanges.Modified.Count, "revision changes");
        budget.AddItems(contribution.Result.CommentChanges.Added.Count, "comment changes");
        budget.AddItems(contribution.Result.CommentChanges.Removed.Count, "comment changes");
        budget.AddItems(contribution.Result.CommentChanges.Modified.Count, "comment changes");
        budget.AddItems(contribution.Result.AnnotationChanges.Added.Count, "annotation changes");
        budget.AddItems(contribution.Result.AnnotationChanges.Removed.Count, "annotation changes");
        budget.AddItems(contribution.Result.AnnotationChanges.Modified.Count, "annotation changes");
        foreach (var revision in contribution.Result.RevisionChanges.Added
            .Concat(contribution.Result.RevisionChanges.Removed)
            .Concat(contribution.Result.RevisionChanges.Modified))
        {
            budget.AddSerializedBytes(512, "revision evidence structure");
            budget.String(revision.Id, "revision id");
            budget.String(revision.Type, "revision type");
            budget.String(revision.Author, "revision author");
            budget.String(revision.Date, "revision date");
            budget.String(revision.DateUtc, "revision UTC date");
            budget.String(revision.Text, "revision text");
            budget.String(revision.PartUri, "revision part URI");
            budget.String(revision.Scope, "revision scope");
            budget.String(revision.AnchorId, "revision anchor");
            budget.AddItems(revision.ConstituentIds.Count, "revision constituent ids");
            foreach (var id in revision.ConstituentIds)
                budget.String(id, "revision constituent id");
            budget.AddItems(revision.ConstituentKeys.Count, "revision constituent keys");
            foreach (var key in revision.ConstituentKeys)
                budget.String(key, "revision constituent key");
            budget.AddItems(revision.AffectedAnchors.Count, "revision affected anchors");
            foreach (var anchor in revision.AffectedAnchors)
                AccountAnchor(anchor, budget);
            budget.String(revision.Diagnostic?.Code, "revision diagnostic code");
            budget.String(revision.Diagnostic?.Message, "revision diagnostic message");
        }
        foreach (var comment in contribution.Result.CommentChanges.Added
            .Concat(contribution.Result.CommentChanges.Removed)
            .Concat(contribution.Result.CommentChanges.Modified))
        {
            budget.AddSerializedBytes(256, "comment evidence structure");
            budget.String(comment.DefAnchorId, "comment anchor");
            budget.String(comment.Author, "comment author");
            budget.String(comment.Initials, "comment initials");
            budget.String(comment.Date, "comment date");
            budget.String(comment.Text, "comment text");
            budget.String(comment.ParentAnchorId, "comment parent anchor");
        }
        foreach (var annotation in contribution.Result.AnnotationChanges.Added
            .Concat(contribution.Result.AnnotationChanges.Removed)
            .Concat(contribution.Result.AnnotationChanges.Modified))
        {
            budget.AddSerializedBytes(384, "annotation evidence structure");
            budget.String(annotation.Id, "annotation id");
            budget.String(annotation.LabelId, "annotation label id");
            budget.String(annotation.Label, "annotation label");
            budget.String(annotation.Color, "annotation color");
            budget.String(annotation.Author, "annotation author");
            budget.String(annotation.BookmarkName, "annotation bookmark");
            budget.String(annotation.AnnotatedText, "annotation text");
            budget.AddItems(annotation.Metadata.Count, "annotation metadata");
            foreach (var pair in annotation.Metadata)
            {
                budget.String(pair.Key, "annotation metadata key");
                budget.String(pair.Value, "annotation metadata value");
            }
        }
        budget.AddItems(contribution.Result.Warnings.Count, "transaction warnings");
        foreach (var warning in contribution.Result.Warnings)
            budget.String(warning, "transaction warning");
    }

    private static void AccountAnchor(
        Anchor anchor,
        DeliveryReceiptResourceBudget budget)
    {
        budget.AddSerializedBytes(48, "anchor structure");
        budget.String(anchor.Id, "anchor id");
        budget.String(anchor.Kind, "anchor kind");
        budget.String(anchor.Scope, "anchor scope");
        budget.String(anchor.Unid, "anchor unid");
    }

    private static void AccountTableAnchors(
        TableAnchorMapping mapping,
        DeliveryReceiptResourceBudget budget)
    {
        budget.AddSerializedBytes(96, "table-anchor mapping structure");
        budget.AddItems(mapping.Retained.Count, "retained table anchors");
        budget.AddItems(mapping.Added.Count, "added table anchors");
        budget.AddItems(mapping.Invalidated.Count, "invalidated table anchors");
        foreach (var retained in mapping.Retained)
        {
            AccountTableAnchorLocation(retained.Before, budget);
            AccountTableAnchorLocation(retained.After, budget);
        }
        foreach (var location in mapping.Added.Concat(mapping.Invalidated))
            AccountTableAnchorLocation(location, budget);
    }

    private static void AccountTableAnchorLocation(
        TableAnchorLocation location,
        DeliveryReceiptResourceBudget budget)
    {
        budget.AddSerializedBytes(128, "table-anchor location structure");
        AccountAnchor(location.Anchor, budget);
    }

    private static void AccountPreconditionValue(
        object? value,
        DeliveryReceiptResourceBudget budget,
        string name)
    {
        switch (value)
        {
            case null:
            case bool:
            case byte:
            case sbyte:
            case short:
            case ushort:
            case int:
            case uint:
            case long:
            case ulong:
            case float:
            case double:
            case decimal:
                budget.AddSerializedBytes(32, name);
                break;
            case string text:
                budget.String(text, name);
                break;
            case JsonElement json:
                AccountJsonElement(json, budget, depth: 1, name);
                break;
            default:
                throw new DeliveryReceiptValidationException(
                    "invalid_batch_evidence",
                    "Precondition evidence must be a JSON scalar or JsonElement.");
        }
    }

    private static void AccountJsonElement(
        JsonElement value,
        DeliveryReceiptResourceBudget budget,
        int depth,
        string name)
    {
        budget.Depth(depth, name);
        budget.AddSerializedBytes(8, name);
        switch (value.ValueKind)
        {
            case JsonValueKind.Object:
            {
                foreach (var property in value.EnumerateObject())
                {
                    budget.AddItems(1, name);
                    budget.String(property.Name, $"{name} property");
                    AccountJsonElement(property.Value, budget, depth + 1, name);
                }
                break;
            }
            case JsonValueKind.Array:
            {
                foreach (var item in value.EnumerateArray())
                {
                    budget.AddItems(1, name);
                    AccountJsonElement(item, budget, depth + 1, name);
                }
                break;
            }
            case JsonValueKind.String:
                budget.String(value.GetString(), name);
                break;
            case JsonValueKind.Number:
                budget.AddSerializedBytes(value.GetRawText().Length, name);
                break;
            case JsonValueKind.True:
            case JsonValueKind.False:
            case JsonValueKind.Null:
                break;
            default:
                throw new DeliveryReceiptValidationException(
                    "invalid_batch_evidence", "Precondition JSON is incomplete.");
        }
    }

    private void ValidateManifestResources(PackageManifest manifest)
    {
        if (manifest.Entries.Count > _limits.CleanDocxManifestOptions.MaxEntryCount)
        {
            throw new DeliveryReceiptValidationException(
                "receipt_resource_limit", "Source manifest exceeds the package entry limit.");
        }
        var budget = new DeliveryReceiptResourceBudget(_limits);
        budget.String(manifest.Schema, "manifest schema");
        budget.String(manifest.PackageKind, "manifest package kind");
        AccountDigest(manifest.RawPackageBytesDigest, budget, "manifest raw digest");
        AccountDigest(manifest.OrderedOpcContentDigest, budget, "manifest content digest");
        AccountDigest(manifest.NormalizedSemanticDigest, budget, "manifest semantic digest");
        budget.AddItems(manifest.Entries.Count, "source manifest entries");
        budget.AddSerializedBytes(
            checked(manifest.Entries.Count * 256L), "source manifest entry structure");
        foreach (var entry in manifest.Entries)
        {
            budget.String(entry.Uri, "source entry URI");
            budget.String(entry.ContentType, "source entry content type");
            budget.String(entry.ContentTypeSource, "source content-type source");
            AccountDigest(entry.RawBytesDigest, budget, "source entry raw digest");
            AccountDigest(entry.NormalizedXmlDigest, budget, "source entry XML digest");
        }
        budget.AddItems(manifest.Relationships.Count, "source relationships");
        budget.AddSerializedBytes(
            checked(manifest.Relationships.Count * 192L), "source relationship structure");
        foreach (var relationship in manifest.Relationships)
        {
            budget.String(relationship.OwnerUri, "source relationship owner");
            budget.String(relationship.Id, "source relationship id");
            budget.String(relationship.Type, "source relationship type");
            budget.String(relationship.Target, "source relationship target");
            budget.String(relationship.TargetMode, "source relationship target mode");
            budget.String(relationship.ResolvedTargetUri, "source relationship target URI");
        }
        budget.AddItems(manifest.ContentTypes.Count, "source content types");
        budget.AddSerializedBytes(
            checked(manifest.ContentTypes.Count * 96L), "source content-type structure");
        foreach (var contentType in manifest.ContentTypes)
        {
            budget.String(contentType.Kind, "source content-type kind");
            budget.String(contentType.Key, "source content-type key");
            budget.String(contentType.ContentType, "source content type");
        }
        budget.AddItems(manifest.Findings.Count, "source manifest findings");
        budget.AddSerializedBytes(
            checked(manifest.Findings.Count * 192L), "source finding structure");
        foreach (var finding in manifest.Findings)
        {
            budget.String(finding.Code, "source manifest finding code");
            budget.String(finding.Message, "source manifest finding message");
            budget.String(finding.Location?.EntryUri, "source finding entry URI");
            budget.String(finding.Location?.OwnerUri, "source finding owner URI");
            budget.String(finding.Location?.RelationshipId, "source finding relationship id");
            budget.String(finding.Location?.TargetUri, "source finding target URI");
            budget.String(finding.Location?.PropertyPath, "source finding property path");
        }
        budget.String(manifest.Facts.MainDocumentUri, "source main-document URI");
    }

    private static void AccountDigest(
        VerificationDigest? digest,
        DeliveryReceiptResourceBudget budget,
        string name)
    {
        if (digest is null)
            return;
        budget.String(digest.Algorithm, $"{name} algorithm");
        budget.String(digest.Value, $"{name} value");
    }

    private void EnsureCollectionCapacity(int currentCount, string name)
    {
        if (currentCount >= _limits.MaxCollectionItems)
        {
            throw new DeliveryReceiptValidationException(
                "receipt_resource_limit", $"Receipt {name} limit exceeded.");
        }
    }

    private void CheckString(string? value, string name)
    {
        if (value is not null && value.Length > _limits.MaxStringLength)
        {
            throw new DeliveryReceiptValidationException(
                "receipt_resource_limit", $"{name} exceeds the string-length limit.");
        }
    }

    private static bool ArtifactEquals(DeliveryArtifact left, DeliveryArtifact right) =>
        left == right;

    private static string ScopeFromAnchor(string anchorId)
    {
        DeliveryReceiptValidation.RequireNonBlank(anchorId, "anchor id", 4096);
        var first = anchorId.IndexOf(':');
        var second = first < 0 ? -1 : anchorId.IndexOf(':', first + 1);
        if (first <= 0 || second <= first + 1 || second == anchorId.Length - 1)
        {
            throw new DeliveryReceiptValidationException(
                "invalid_anchor_id", $"'{anchorId}' is not a canonical kind:scope:unid anchor.");
        }
        return anchorId[(first + 1)..second];
    }

    private static JsonElement ParseCanonical(byte[] canonical)
    {
        using var document = JsonDocument.Parse(canonical, new JsonDocumentOptions
        {
            MaxDepth = DeliveryReceiptLimits.MaxAllowedJsonDepth,
        });
        return document.RootElement.Clone();
    }

    private (byte[] Bytes, JsonElement Element) CanonicalListItem(string json)
    {
        DeliveryReceiptResourceBudget.Bytes(
            Encoding.UTF8.GetByteCount(json),
            _limits.MaxReceiptJsonBytes,
            "receipt_resource_limit",
            "Authorship evidence JSON");
        var canonicalArray = DeliveryReceiptCanonicalJson.CanonicalizeBounded(
            Encoding.UTF8.GetBytes(json), _limits, _limits.MaxReceiptJsonBytes,
            "receipt_resource_limit");
        using var document = JsonDocument.Parse(canonicalArray, new JsonDocumentOptions
        {
            MaxDepth = DeliveryReceiptLimits.MaxAllowedJsonDepth,
        });
        if (document.RootElement.ValueKind != JsonValueKind.Array
            || document.RootElement.GetArrayLength() != 1)
        {
            throw new DeliveryReceiptValidationException(
                "invalid_batch_evidence", "Expected one serialized batch evidence item.");
        }
        var element = document.RootElement[0].Clone();
        return (DeliveryReceiptCanonicalJson.SerializeCanonicalBounded(
            element, _limits, _limits.MaxReceiptJsonBytes,
            "receipt_resource_limit"), element);
    }

    private static DeliveryAttributionKey AttributionKey(
        DeliveryChangeAttributionRule rule) => new(
            rule.Kind,
            rule.EntryUri ?? string.Empty,
            rule.OwnerUri ?? string.Empty,
            rule.RelationshipId ?? string.Empty);

    private static DeliveryAttributionKey AttributionKey(
        DeliveryPackageChangeObservation candidate) => new(
            candidate.Kind,
            candidate.Location.EntryUri ?? string.Empty,
            candidate.Location.OwnerUri ?? string.Empty,
            candidate.Location.RelationshipId ?? string.Empty);

    private static void ValidateAttributionRule(DeliveryChangeAttributionRule rule)
    {
        if (!Enum.IsDefined(rule.Kind) || !Enum.IsDefined(rule.Disposition))
            throw new DeliveryReceiptValidationException("invalid_attribution", "Unknown attribution value.");
        if (rule.Kind is DeliveryPackageChangeKind.PartAdded
            or DeliveryPackageChangeKind.PartRemoved
            or DeliveryPackageChangeKind.PartModified)
        {
            DeliveryReceiptValidation.RequireNonBlank(rule.EntryUri, "attribution entry URI", 4096);
            if (rule.OwnerUri is not null || rule.RelationshipId is not null)
                throw new DeliveryReceiptValidationException(
                    "invalid_attribution", "Part attribution cannot contain relationship fields.");
        }
        else
        {
            DeliveryReceiptValidation.RequireNonBlank(rule.OwnerUri, "attribution owner URI", 4096);
            DeliveryReceiptValidation.RequireNonBlank(
                rule.RelationshipId, "attribution relationship id", 2048);
            if (rule.EntryUri is not null)
                throw new DeliveryReceiptValidationException(
                    "invalid_attribution", "Relationship attribution cannot contain an entry URI.");
        }
        if (rule.Disposition == DeliveryChangeDisposition.UserRequested)
        {
            DeliveryReceiptValidation.RequireNonBlank(
                rule.TransactionEntryId, "attribution transaction entry id", 256);
            if (rule.RequestedOperationIndex is null or < 0)
                throw new DeliveryReceiptValidationException(
                    "invalid_attribution", "User-requested attribution requires an operation index.");
        }
        if (rule.TransactionEntryId is null && rule.RequestedOperationIndex is not null)
        {
            throw new DeliveryReceiptValidationException(
                "invalid_attribution",
                "An attribution operation index requires a transaction entry id.");
        }
        if (rule.Disposition == DeliveryChangeDisposition.Derived)
            DeliveryReceiptValidation.RequireNonBlank(rule.Derivation, "derivation", 4096);
    }

    private readonly record struct DeliveryAttributionKey(
        DeliveryPackageChangeKind Kind,
        string EntryUri,
        string OwnerUri,
        string RelationshipId);
}
