// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.Verification;
using DocumentFormat.OpenXml.Packaging;
using GridCell = Docxodus.Internal.TableGridCell;

namespace Docxodus;

public sealed partial class DocxSession
{
    /// <summary>
    /// Evaluate optimistic guards without mutating the document, consuming undo history, or
    /// advancing <see cref="Version"/>. Returns null when every supplied guard matches.
    /// <paramref name="actualMatchCount"/> is supplied by find/replace after it has enumerated
    /// the live matches; other callers leave it null.
    /// </summary>
    public EditError? EvaluatePreconditions(
        MutationPreconditions? preconditions,
        int? actualMatchCount = null)
    {
        if (preconditions is null) return null;
        if (_disposed) return new EditError(EditErrorCode.SessionDisposed, "session disposed");

        var target = !string.IsNullOrEmpty(preconditions.AnchorId)
            ? CurrentPreconditionTarget(preconditions.AnchorId)
            : null;
        if (preconditions.ExpectedVersion is { } expectedVersion && expectedVersion != _version)
            return PreconditionError("document_version", expectedVersion, _version,
                preconditions.AnchorId, target);

        bool hasAnchorGuard = preconditions.ExpectedContentHash is not null
            || preconditions.ExpectedText is not null
            || preconditions.ExpectedTextRange is not null
            || preconditions.ExpectedKind is not null
            || preconditions.ExpectedScope is not null;
        if (hasAnchorGuard && string.IsNullOrEmpty(preconditions.AnchorId))
            return PreconditionError("anchor_id", "present", "missing", null);
        if (hasAnchorGuard && target is not { Exists: true })
            return PreconditionError("anchor_exists", true, false, preconditions.AnchorId, target);

        if (preconditions.ExpectedKind is { } expectedKind
            && !string.Equals(expectedKind, target!.Kind, StringComparison.Ordinal))
            return PreconditionError("anchor_kind", expectedKind, target.Kind,
                preconditions.AnchorId, target);
        if (preconditions.ExpectedScope is { } expectedScope
            && !string.Equals(expectedScope, target!.Scope, StringComparison.Ordinal))
            return PreconditionError("anchor_scope", expectedScope, target.Scope,
                preconditions.AnchorId, target);
        if (preconditions.ExpectedContentHash is { } expectedHash
            && !string.Equals(expectedHash, target!.ContentHash, StringComparison.OrdinalIgnoreCase))
            return PreconditionError("anchor_content_hash", expectedHash, target.ContentHash,
                preconditions.AnchorId, target);
        if (preconditions.ExpectedText is { } expectedText
            && !string.Equals(expectedText, target!.VisibleText, StringComparison.Ordinal))
            return PreconditionError("anchor_text", expectedText, target.VisibleText,
                preconditions.AnchorId, target);
        if (preconditions.ExpectedTextRange is { } range)
        {
            var visible = target!.VisibleText ?? string.Empty;
            object actual;
            if (range.Start < 0 || range.Length < 0 || range.Start > visible.Length - range.Length)
                actual = new { start = range.Start, length = range.Length, availableLength = visible.Length };
            else
                actual = visible.Substring(range.Start, range.Length);
            if (actual is not string actualText
                || !string.Equals(range.Text, actualText, StringComparison.Ordinal))
                return PreconditionError("anchor_text_range", range.Text, actual,
                    preconditions.AnchorId, target);
        }
        if (preconditions.ExpectedMatchCount is { } expectedCount
            && actualMatchCount is { } count && expectedCount != count)
            return PreconditionError("match_count", expectedCount, count,
                preconditions.AnchorId, target);
        return null;
    }

    /// <summary>
    /// Atomically evaluate guards and invoke one synchronous mutation. This is the direct .NET
    /// primitive shared facades use; future atomic batches can evaluate their guards under the
    /// same gate before taking their aggregate snapshot.
    /// </summary>
    public EditResult ExecuteMutation(
        MutationPreconditions? preconditions,
        Func<DocxSession, EditResult> mutation)
    {
        ArgumentNullException.ThrowIfNull(mutation);
        lock (_mutationGate)
        {
            var error = EvaluatePreconditions(preconditions);
            return error is null
                ? mutation(this)
                : new EditResult { Success = false, Error = error };
        }
    }

    /// <summary>
    /// Begin a complete-package transaction under the session-wide mutation gate. Transactions
    /// nest in strict LIFO order: an inner commit remains speculative until its outer transaction
    /// commits, while an inner rollback restores the state visible at the inner begin boundary.
    /// </summary>
    public DocxSessionTransaction BeginTransaction() => BeginTransaction(fullPackage: true);

    // Only fixed, internal compositions whose mutations are covered by TakeSnapshot may use
    // the selective snapshot. Arbitrary caller callbacks always require the complete package.
    private DocxSessionTransaction BeginTransaction(bool fullPackage)
    {
        ThrowIfDisposed();
        System.Threading.Monitor.Enter(_mutationGate);
        try
        {
            if (!fullPackage && _transactions.Count == 0)
                _deliveryEvidence?.Reconcile();
            var id = checked(++_nextTransactionId);
            var state = new TransactionState(
                id,
                Environment.CurrentManagedThreadId,
                fullPackage ? TakePackageSnapshot() : TakeSnapshot() with
                {
                    RevisionCounter = _revisionCounter,
                    LastFormatRevisionTicks = _lastFormatRevisionTicks,
                },
                _history.CaptureState(),
                _transactionPendingMutations,
                _transactionMutationEpoch,
                _trackedChanges,
                _revisionAuthor,
                LastInternalError,
                LastRollbackError);
            _transactions.Push(state);
            return new DocxSessionTransaction(this, id);
        }
        catch
        {
            System.Threading.Monitor.Exit(_mutationGate);
            throw;
        }
    }

    internal void CompleteTransaction(long id, bool commit)
    {
        if (_transactions.Count == 0 || _transactions.Peek().Id != id)
            throw new InvalidOperationException("transactions must complete in strict LIFO order");
        var state = _transactions.Peek();
        if (state.OwnerThreadId != Environment.CurrentManagedThreadId)
            throw new InvalidOperationException("a transaction must complete on the thread that began it");

        bool completed = false;
        try
        {
            if (!commit)
            {
                RestoreTransactionState(state);
            }
            else if (_transactions.Count == 1)
            {
                // An inner commit stays represented by its ordinary speculative history entries.
                // The outermost commit is the only boundary that squashes them into one
                // pre-batch snapshot and advances caller-visible version once.
                // The epoch, not the pending count, decides: a scope whose ops all self-rolled
                // back nets the pending count to baseline yet may still have left the package
                // dirty, and squashing that into "nothing happened" would strand it with no
                // undo entry and no version bump.
                bool mutated = _transactionMutationEpoch > state.MutationEpoch;
                _history.RestoreState(state.History);
                _transactionPendingMutations = state.PendingMutations;
                _version = state.PackageSnapshot.Version;
                if (mutated)
                {
                    _history.RecordPreOp(state.PackageSnapshot);
                    // The transaction remains on the stack until every completion operation has
                    // succeeded, so the history callback records this speculatively. Convert that
                    // callback effect into the single caller-visible root commit advancement.
                    _transactionPendingMutations = state.PendingMutations;
                    _version = checked(state.PackageSnapshot.Version + 1);
                }
                // Only the pending count is rewound here. It is a live history depth that must
                // agree with the history state just restored; the epoch is deliberately left
                // monotonic because it is only ever read against a baseline captured at some
                // enclosing begin — and after this pop there is no enclosing scope left. An
                // inner commit rewinds neither: the outer scope must still see that its nested
                // work touched the package.
            }

            // Do not orphan the transaction if validation or restoration failed. The caller may
            // retry in the correct LIFO/thread context, or roll the still-active scope back.
            _transactions.Pop();
            completed = true;
        }
        finally
        {
            if (completed)
                System.Threading.Monitor.Exit(_mutationGate);
        }
    }

    private void RestoreTransactionState(TransactionState state)
    {
        try
        {
            // Every session mutation records before touching package state, so an unchanged
            // epoch means the scope performed only reads/configuration work and the live package
            // is still byte-identical to the checkpoint. Reopening it in that case is observably
            // worse: OPC rewrites ZIP timestamps even for a no-op rollback. Keep the live package
            // byte-pure while still restoring history/configuration below.
            //
            // The epoch, not the pending count, is the witness. RollbackFailedOp pops the op's own
            // pre-op entry, so ~40 ops return the pending count to baseline on their failure path
            // while their selective restore leaves parts the snapshot excludes (the numbering part)
            // dirty. Skipping the package restore there reported a rollback that never happened.
            if (_transactionMutationEpoch > state.MutationEpoch)
                RestoreSnapshot(state.PackageSnapshot);
            _history.RestoreState(state.History);
            _transactionPendingMutations = state.PendingMutations;
            _transactionMutationEpoch = state.MutationEpoch;
            _trackedChanges = state.TrackedChanges;
            _revisionAuthor = state.RevisionAuthor;
            LastInternalError = state.LastInternalError;
            LastRollbackError = state.LastRollbackError;
        }
        catch (Exception rollbackEx)
        {
            LastRollbackError = rollbackEx;
            throw;
        }
    }

    /// <summary>
    /// Execute a core mutation batch. Atomic is the default; callers must choose
    /// <see cref="MutationBatchMode.BestEffort"/> explicitly to retain partial successes.
    /// </summary>
    public MutationBatchResult ExecuteBatch(
        IEnumerable<MutationBatchStep> steps,
        MutationBatchMode mode = MutationBatchMode.Atomic) => ExecuteBatch(steps, mode, null);

    /// <summary>
    /// <see cref="ExecuteBatch(IEnumerable{MutationBatchStep}, MutationBatchMode)"/> with the
    /// transaction identity a retry journal bound the request to, so delivery evidence records
    /// the batch under that identity (issue #748).
    /// </summary>
    internal MutationBatchResult ExecuteBatch(
        IEnumerable<MutationBatchStep> steps,
        MutationBatchMode mode,
        Verification.DeliveryTransactionIdentity? identity)
    {
        lock (_mutationGate)
        {
            var materialized = MaterializeBatchSteps(steps, mode);
            var before = ObserveBatchSemantics();
            var baseVersion = _version;
            var evidence = _deliveryEvidence?.Begin(materialized, mode, identity);
            MutationBatchResult result;
            try
            {
                result = mode == MutationBatchMode.Atomic
                    ? ExecuteAtomicBatch(materialized)
                    : ExecuteBestEffortBatch(materialized, evidence);
                result = CompleteBatchResult(result, before, baseVersion);
            }
            catch
            {
                _deliveryEvidence?.Abandon(evidence);
                throw;
            }
            if (evidence is not null)
            {
                if (mode == MutationBatchMode.Atomic) _deliveryEvidence!.CompleteAtomic(evidence, result);
                else _deliveryEvidence!.CompleteBestEffort(evidence);
            }
            return result;
        }
    }

    /// <summary>
    /// Execute the identical batch delegates on a complete isolated clone. The live session is
    /// used only long enough to clone its current logical package and scalar configuration under
    /// the mutation gate; guards, mutations, history writes, semantic inspection, package hashing,
    /// and optional HTML rendering all target the shadow. Abandoning or disposing the shadow can
    /// therefore never require live rollback.
    /// </summary>
    /// <remarks>
    /// <para><b>Caller contract.</b> Each step's callbacks receive the shadow session as their
    /// <c>DocxSession</c> argument — <em>address that argument</em>. A callback written as
    /// <c>s =&gt; liveSession.ReplaceText(...)</c>, closing over the live session instead of using
    /// <c>s</c>, mutates the LIVE document, and this overload cannot prevent it: a delegate may
    /// call anything it can reach.</para>
    /// <para>The handle-shaped seams are intrinsically safe by construction, because a step
    /// factory there is handed only the temporary shadow handle and never sees the live one:
    /// <c>DocxSessionOps.PreviewBatch</c> for stdio/MCP, and the <c>OpenPreviewSession</c> bridge
    /// export for the browser client. Prefer those when the steps are not written by the same
    /// author as the call.</para>
    /// <para>Enrichment cost is not free and has no opt-out: see the receipt cost note in
    /// <c>docs/architecture/docx_mutation_api.md</c>.</para>
    /// </remarks>
    public MutationBatchResult PreviewBatch(
        IEnumerable<MutationBatchStep> steps,
        MutationBatchMode mode = MutationBatchMode.Atomic,
        MutationBatchPreviewOptions? options = null)
    {
        ValidatePreviewOptions(options);
        var materialized = MaterializeBatchSteps(steps, mode);
        using var shadow = CreateShadowSession();
        return shadow.FinalizePreviewResult(shadow.ExecuteBatch(materialized, mode), options);
    }

    internal static void ValidatePreviewOptions(MutationBatchPreviewOptions? options)
    {
        if (options is not null && !Enum.IsDefined(options.HtmlMode))
            throw new ArgumentOutOfRangeException(
                nameof(options), options.HtmlMode, "unknown preview HTML mode");
    }

    private static MutationBatchStep[] MaterializeBatchSteps(
        IEnumerable<MutationBatchStep> steps,
        MutationBatchMode mode)
    {
        ArgumentNullException.ThrowIfNull(steps);
        if (!Enum.IsDefined(mode))
            throw new ArgumentOutOfRangeException(nameof(mode), mode, "unknown mutation batch mode");
        var materialized = steps.ToArray();
        if (materialized.Any(step => step is null))
            throw new ArgumentException("batch steps cannot contain null", nameof(steps));
        return materialized;
    }

    /// <summary>
    /// Receipt warnings that say a fresh execution may regenerate ids or timestamps. Attached by
    /// <see cref="CompleteBatchResult"/>; removed by <see cref="CommitPreview"/>, which does not
    /// execute afresh but restores the previewed bytes.
    /// </summary>
    private static class BatchReplayCaveats
    {
        public const string RevisionDates =
            "Tracked-revision date attributes may use the execution clock; compare revision " +
            "ids, authors, types, text, and anchors across separate executions.";

        public const string CommentDates =
            "Comment date attributes may be generated from the execution clock; supply dates " +
            "explicitly when byte-identical replay is required.";

        public const string AnnotationMetadata =
            "Auto-generated annotation ids or creation timestamps are execution metadata; " +
            "supply id and created explicitly when byte-identical replay is required.";

        public const string CreatedAnchors =
            "Created anchors and related OOXML ids may be generated independently on replay; " +
            "preview/apply equivalence is semantic and packageHash or anchor ids may differ.";

        public static bool Contains(string warning) =>
            warning is RevisionDates or CommentDates or AnnotationMetadata or CreatedAnchors;
    }

    private sealed record BatchSemanticObservation(
        IReadOnlyList<RevisionListEntry>? Revisions,
        IReadOnlyList<CommentListEntry>? Comments,
        IReadOnlyList<DocumentAnnotation>? Annotations,
        IReadOnlyList<string> Warnings);

    private BatchSemanticObservation ObserveBatchSemantics()
    {
        IReadOnlyList<RevisionListEntry>? revisions = null;
        IReadOnlyList<CommentListEntry>? comments = null;
        IReadOnlyList<DocumentAnnotation>? annotations = null;
        var warnings = new List<string>();
        try { revisions = ListRevisions(); }
        catch (Exception ex) { warnings.Add($"Revision delta inspection unavailable: {ex.Message}"); }
        try { comments = ListComments(); }
        catch (Exception ex) { warnings.Add($"Comment delta inspection unavailable: {ex.Message}"); }
        try { annotations = ListAnnotations(); }
        catch (Exception ex) { warnings.Add($"Annotation delta inspection unavailable: {ex.Message}"); }
        return new BatchSemanticObservation(revisions, comments, annotations, warnings);
    }

    private MutationBatchResult CompleteBatchResult(
        MutationBatchResult result,
        BatchSemanticObservation before,
        long baseVersion)
    {
        var after = ObserveBatchSemantics();
        var warnings = before.Warnings.Concat(after.Warnings).ToList();
        // Equivalence is decided on the SERIALIZED projection of each entry, never on CLR
        // equality. That is the shape every transport actually publishes, it is exactly what
        // npm's `mutationBatchChangeSet` compares (JSON.stringify of the same wire objects), and
        // it stays correct if an entry type ever grows a collection member — record `==` would
        // then fall back to reference equality per element and report every surviving object as
        // modified.
        var revisionChanges = SafeChangeSet(
            before.Revisions, after.Revisions, revision => revision.Id,
            static (left, right) => string.Equals(
                Internal.DocxSessionJson.SerializeRevisionList(new[] { left }),
                Internal.DocxSessionJson.SerializeRevisionList(new[] { right }),
                StringComparison.Ordinal),
            "revision", warnings);
        var commentChanges = SafeChangeSet(
            before.Comments, after.Comments, comment => comment.DefAnchorId,
            static (left, right) => string.Equals(
                Internal.DocxSessionJson.SerializeCommentList(new[] { left }),
                Internal.DocxSessionJson.SerializeCommentList(new[] { right }),
                StringComparison.Ordinal),
            "comment", warnings);
        var annotationChanges = SafeChangeSet(
            before.Annotations, after.Annotations, annotation => annotation.Id,
            static (left, right) => string.Equals(
                Internal.DocxSessionJson.SerializeAnnotations(new[] { left }),
                Internal.DocxSessionJson.SerializeAnnotations(new[] { right }),
                StringComparison.Ordinal),
            "annotation", warnings);
        if (revisionChanges.Added.Concat(revisionChanges.Modified).Any(revision => revision.Date is not null))
            warnings.Add(BatchReplayCaveats.RevisionDates);
        if (commentChanges.Added.Concat(commentChanges.Modified).Any(comment => comment.Date is not null))
            warnings.Add(BatchReplayCaveats.CommentDates);
        if (annotationChanges.Added.Any(annotation => annotation.Created.HasValue))
            warnings.Add(BatchReplayCaveats.AnnotationMetadata);
        try
        {
            if (result.Steps.SelectMany(step => step.Results).Any(edit => edit.Created.Count > 0))
                warnings.Add(BatchReplayCaveats.CreatedAnchors);
        }
        catch (Exception ex)
        {
            warnings.Add($"Generated-field warning inspection unavailable: {ex.Message}");
        }
        if (result.Mode == MutationBatchMode.BestEffort && !result.Success)
            warnings.Add("Best-effort execution retains every successful step despite later failures.");

        string? packageHash = null;
        try { packageHash = GetPackageContentHash(); }
        catch (Exception ex) { warnings.Add($"Package equivalence hash unavailable: {ex.Message}"); }

        return result with
        {
            BaseVersion = baseVersion,
            ResultVersion = _version,
            PackageHash = packageHash,
            RevisionChanges = revisionChanges,
            CommentChanges = commentChanges,
            AnnotationChanges = annotationChanges,
            Warnings = warnings,
        };
    }

    private static MutationBatchChangeSet<T> SafeChangeSet<T>(
        IReadOnlyList<T>? before,
        IReadOnlyList<T>? after,
        Func<T, string?> key,
        Func<T, T, bool> equivalent,
        string kind,
        List<string> warnings)
    {
        if (before is null || after is null)
            return MutationBatchChangeSet<T>.Empty;
        try { return ChangeSet(before, after, key, equivalent); }
        catch (Exception ex)
        {
            warnings.Add($"{kind} delta comparison unavailable: {ex.Message}");
            return MutationBatchChangeSet<T>.Empty;
        }
    }

    private static MutationBatchChangeSet<T> ChangeSet<T>(
        IReadOnlyList<T> before,
        IReadOnlyList<T> after,
        Func<T, string?> key,
        Func<T, T, bool> equivalent)
    {
        static Dictionary<string, List<int>> IndexByKey(
            IReadOnlyList<T> items,
            Func<T, string?> selectKey)
        {
            var groups = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            for (var index = 0; index < items.Count; index++)
            {
                var itemKey = selectKey(items[index]) ?? string.Empty;
                if (!groups.TryGetValue(itemKey, out var indices))
                    groups[itemKey] = indices = new List<int>();
                indices.Add(index);
            }
            return groups;
        }

        // Real-world packages occasionally contain duplicate revision ids across story parts.
        // Treat each identity as a multiset: match equivalent occurrences first, classify paired
        // leftovers as modified, then classify cardinality differences as added/removed. This is
        // deterministic and cannot throw merely because a package is malformed or unconventional.
        var beforeGroups = IndexByKey(before, key);
        var afterGroups = IndexByKey(after, key);
        var beforeMatched = new bool[before.Count];
        var afterMatched = new bool[after.Count];
        var afterModified = new bool[after.Count];

        foreach (var group in afterGroups)
        {
            if (!beforeGroups.TryGetValue(group.Key, out var beforeIndices))
                continue;

            foreach (var afterIndex in group.Value)
            {
                var beforeIndex = beforeIndices.FirstOrDefault(
                    candidate => !beforeMatched[candidate]
                        && equivalent(before[candidate], after[afterIndex]),
                    -1);
                if (beforeIndex < 0) continue;
                beforeMatched[beforeIndex] = true;
                afterMatched[afterIndex] = true;
            }

            var remainingBefore = beforeIndices.Where(index => !beforeMatched[index]).ToArray();
            var remainingAfter = group.Value.Where(index => !afterMatched[index]).ToArray();
            var modifiedCount = Math.Min(remainingBefore.Length, remainingAfter.Length);
            for (var index = 0; index < modifiedCount; index++)
            {
                beforeMatched[remainingBefore[index]] = true;
                afterMatched[remainingAfter[index]] = true;
                afterModified[remainingAfter[index]] = true;
            }
        }

        return new MutationBatchChangeSet<T>(
            after.Where((_, index) => !afterMatched[index]).ToArray(),
            before.Where((_, index) => !beforeMatched[index]).ToArray(),
            after.Where((_, index) => afterModified[index]).ToArray());
    }

    /// <summary>
    /// What the host-owned evidence recorder holds (issue #748). <c>Enabled</c> is false, with
    /// the reason, when the session was opened without
    /// <see cref="DocxSessionSettings.CaptureDeliveryEvidence"/>.
    /// </summary>
    public Delivery.DeliveryEvidenceStatus GetDeliveryEvidenceStatus()
    {
        ThrowIfDisposed();
        lock (_mutationGate)
        {
            if (_deliveryEvidence is null)
            {
                return new Delivery.DeliveryEvidenceStatus
                {
                    Enabled = false,
                    CurrentVersion = _version,
                    UnavailableReason = NotCapturingDeliveryEvidence,
                };
            }
            _deliveryEvidence.Reconcile();
            return _deliveryEvidence.Status();
        }
    }

    /// <summary>
    /// Hand the captured history to a delivery operation: the exact opening package as the
    /// source, the recorder's exact current package as the working document, and — when the
    /// history is complete — the ordered receipt context. When it is not, the context is null
    /// and the status names the reason; a delivery then reports its change-receipt artifact
    /// unavailable with that reason instead of minting a receipt that claims a history.
    /// </summary>
    public Delivery.DeliveryEvidenceExport ExportDeliveryEvidence(
        Delivery.DeliveryReceiptBuildOptions? options = null)
    {
        ThrowIfDisposed();
        options ??= new Delivery.DeliveryReceiptBuildOptions();
        if (!Enum.IsDefined(options.PrivacyProfile))
            throw new ArgumentOutOfRangeException(nameof(options), options.PrivacyProfile, "unknown privacy profile");
        lock (_mutationGate)
        {
            if (_deliveryEvidence is not null) return _deliveryEvidence.Export(options);
            if (_initialPackageBytes is null)
                throw new InvalidOperationException(
                    "Delivery evidence needs the opening package; open the session with CaptureInitialProjection.");
            return new Delivery.DeliveryEvidenceExport(
                null,
                new Delivery.DeliveryDocumentSnapshot("source", 0, _initialPackageBytes),
                new Delivery.DeliveryDocumentSnapshot("working", _version, SerializeCleanCheckpoint()),
                new Delivery.DeliveryEvidenceStatus
                {
                    Enabled = false,
                    CurrentVersion = _version,
                    UnavailableReason = NotCapturingDeliveryEvidence,
                });
        }
    }

    /// <summary>
    /// Build the receipt-bearing delivery of this session through the shared bundle service:
    /// the clean current package, the source-to-delivered semantic delta, and the change
    /// receipt minted from the captured evidence (issue #748). Revisions are preserved as they
    /// are; the receipt attests the session's own edits. A history that cannot be attested
    /// yields an <c>Incomplete</c> bundle whose receipt artifact carries the reason.
    /// </summary>
    public Delivery.DeliveryBundle BuildDeliveryReceipt(Delivery.DeliveryReceiptBuildOptions? options = null)
    {
        var export = ExportDeliveryEvidence(options);
        var request = new Delivery.DeliveryBundleBuildRequest(
            export.Source,
            export.Working,
            "delivered",
            export.Working.DocumentVersion,
            new Delivery.DeliveryBundleRevisionPolicy
            {
                PreExistingRevisions = Delivery.DeliveryRevisionPolicy.Preserve,
                GeneratedRevisions = Delivery.DeliveryRevisionPolicy.Preserve,
            },
            new[]
            {
                new Delivery.DeliveryArtifactRequest
                {
                    ArtifactId = "final-docx",
                    Kind = Delivery.DeliveryArtifactKind.FinalDocx,
                    Requiredness = Delivery.DeliveryArtifactRequiredness.Required,
                },
                new Delivery.DeliveryArtifactRequest
                {
                    ArtifactId = "semantic-source-to-delivered",
                    Kind = Delivery.DeliveryArtifactKind.SemanticDelta,
                    Requiredness = Delivery.DeliveryArtifactRequiredness.Required,
                },
                new Delivery.DeliveryArtifactRequest
                {
                    ArtifactId = "change-receipt",
                    Kind = Delivery.DeliveryArtifactKind.ChangeReceipt,
                    Requiredness = Delivery.DeliveryArtifactRequiredness.Required,
                },
            },
            export.ReceiptContext);
        var bundleOptions = new Delivery.DeliveryBundleBuildOptions
        {
            ReturnIncompleteBundle = true,
            FailOnDeliverableValidationFailure = false,
        };
        return new Delivery.DeliveryBundleService().BuildAsync(request, bundleOptions)
            .AsTask().GetAwaiter().GetResult();
    }

    /// <summary>
    /// A transport that composes its batch from individual calls (the browser client, or a
    /// direct tool call on the stdio/MCP hosts) describes it before running it; the version
    /// steps it produces are then attributed to that description at completion. Returns false
    /// when nothing will be attributed (capture off, unavailable, or inside a transaction scope).
    /// </summary>
    internal bool BeginClientDeliveryEvidence(
        IReadOnlyList<(string Tool, string Action, string? ArgumentsJson)> operations,
        MutationBatchMode mode,
        Verification.DeliveryTransactionIdentity? identity)
    {
        lock (_mutationGate)
        {
            if (_clientEvidence is not null) _deliveryEvidence?.Abandon(_clientEvidence);
            _clientEvidence = _deliveryEvidence?.Begin(operations, mode, identity);
            return _clientEvidence is not null;
        }
    }

    internal void CompleteClientDeliveryEvidence(IReadOnlyList<MutationBatchStepResult> steps)
    {
        lock (_mutationGate)
        {
            if (_clientEvidence is { } pending) _deliveryEvidence!.CompleteClient(pending, steps);
            _clientEvidence = null;
        }
    }

    internal void CompleteClientDeliveryEvidence(IReadOnlyList<EditResult> directResults)
    {
        lock (_mutationGate)
        {
            if (_clientEvidence is { } pending) _deliveryEvidence!.CompleteDirect(pending, directResults);
            _clientEvidence = null;
        }
    }

    internal void AbandonClientDeliveryEvidence()
    {
        lock (_mutationGate)
        {
            _deliveryEvidence?.Abandon(_clientEvidence);
            _clientEvidence = null;
        }
    }

    internal DocxSession CreateShadowSession()
    {
        lock (_mutationGate)
        {
            ThrowIfDisposed();
            var snapshot = TakePackageSnapshot();
            var shadow = new DocxSession(
                snapshot.PackageBytes!,
                CloneSettingsForShadow(),
                skipInitialProjectionCapture: true)
            {
                _version = _version,
                _revisionCounter = snapshot.RevisionCounter ?? _revisionCounter,
                // _revisionCounterSeeded is deliberately NOT copied: the shadow re-seeds from
                // its own clone on first use, which can only raise the counter it inherited.
                _revisionCounterSeeded = false,
                _lastFormatRevisionTicks = snapshot.LastFormatRevisionTicks ?? _lastFormatRevisionTicks,
                _nextTransactionId = _nextTransactionId,
                _initialProjection = CloneProjection(_initialProjection),
                // The baseline arrays are never mutated after capture, so the throwaway shadow
                // shares the references instead of copying a full package per preview.
                _initialPackageBytes = _initialPackageBytes,
                _initialCheckpointBytes = _initialCheckpointBytes,
                _trackedChanges = _trackedChanges,
                _revisionAuthor = _revisionAuthor,
                _shadowOrigin = new ShadowOrigin(
                    this, _version, snapshot.PackageBytes!, _trackedChanges, _revisionAuthor),
            };
            return shadow;
        }
    }

    private static MarkdownProjection? CloneProjection(MarkdownProjection? source)
    {
        if (source is null) return null;
        return new MarkdownProjection
        {
            Markdown = source.Markdown,
            AnchorIndex = source.AnchorIndex.ToDictionary(
                pair => pair.Key,
                pair => new AnchorTarget
                {
                    Anchor = pair.Value.Anchor,
                    PartUri = pair.Value.PartUri,
                    Unid = pair.Value.Unid,
                    TextPreview = pair.Value.TextPreview,
                    AutoNumberPrefix = pair.Value.AutoNumberPrefix,
                },
                StringComparer.Ordinal),
        };
    }

    private DocxSessionSettings CloneSettingsForShadow()
    {
        var projection = _settings.ProjectionSettings;
        return new DocxSessionSettings
        {
            UndoDepth = _settings.UndoDepth,
            UndoMemoryBudgetBytes = _settings.UndoMemoryBudgetBytes,
            ValidateRawOps = _settings.ValidateRawOps,
            TrackedChanges = _settings.TrackedChanges,
            RevisionAuthor = _settings.RevisionAuthor,
            PersistAnchorIds = _settings.PersistAnchorIds,
            SmartQuotes = _settings.SmartQuotes,
            EmitMarkdownPatch = _settings.EmitMarkdownPatch,
            CaptureInitialProjection = _settings.CaptureInitialProjection,
            ProjectionSettings = new WmlToMarkdownConverterSettings
            {
                Scopes = projection.Scopes,
                HeadingLevelOffset = projection.HeadingLevelOffset,
                AnchorMode = projection.AnchorMode,
                TableMode = projection.TableMode,
                TableInlineCellMax = projection.TableInlineCellMax,
                TrackedChanges = projection.TrackedChanges,
                ResolveNumbering = projection.ResolveNumbering,
                ImageUriBuilder = projection.ImageUriBuilder,
                EmptyParagraphs = projection.EmptyParagraphs,
                AnchorIdRendering = projection.AnchorIdRendering,
            },
        };
    }

    /// <summary>Mark a result produced on this shadow and optionally render shadow-only HTML.</summary>
    internal MutationBatchResult FinalizePreviewResult(
        MutationBatchResult result,
        MutationBatchPreviewOptions? options)
    {
        var warnings = result.Warnings.ToList();
        string? html = null;
        try
        {
            switch (options?.HtmlMode ?? MutationPreviewHtmlMode.None)
            {
                case MutationPreviewHtmlMode.None:
                    break;
                case MutationPreviewHtmlMode.Scoped when string.IsNullOrWhiteSpace(options?.HtmlAnchorId):
                    warnings.Add("Scoped HTML was requested without htmlAnchorId; no HTML was generated.");
                    break;
                case MutationPreviewHtmlMode.Scoped:
                    html = Internal.HtmlConversionOps.RenderBlockHtml(
                        this,
                        options!.HtmlAnchorId!,
                        Internal.HtmlConversionOps.PreviewBlockOptions());
                    break;
                case MutationPreviewHtmlMode.Full:
                    html = Internal.HtmlConversionOps.ConvertToHtml(
                        this,
                        Internal.HtmlConversionOps.PreviewDocumentOptions());
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(options), "unknown preview HTML mode");
            }
        }
        catch (Exception ex)
        {
            warnings.Add($"Preview HTML could not be generated: {ex.Message}");
        }

        MutationPreviewRetention? retention = null;
        if (options?.Retain == true)
        {
            if (!result.Success)
                warnings.Add("The preview did not succeed, so it was not retained for commit.");
            else
                retention = RetainPreview(result, warnings);
        }

        return result with
        {
            Preview = true,
            Warnings = warnings,
            Html = html,
            Retention = retention,
        };
    }

    /// <summary>
    /// Keep this shadow's final package for a guarded commit on the live session it was cloned
    /// from (issue #760). Runs on the shadow; the store lives on the owner. <paramref name="result"/>
    /// is the receipt to return on commit, or null when the transport composes its own receipt
    /// (the browser client). Returns null, with a warning, when the store refuses the package.
    /// </summary>
    internal MutationPreviewRetention? RetainPreview(
        MutationBatchResult? result,
        System.Collections.Generic.List<string> warnings)
    {
        ThrowIfDisposed();
        var origin = _shadowOrigin
            ?? throw new InvalidOperationException("only a preview shadow can retain its result for commit");
        var store = origin.Owner.RetainedPreviews;
        var snapshot = TakePackageSnapshot();
        var retention = new MutationPreviewRetention(
            Internal.RetainedPreviews.NewPreviewId(),
            origin.BaseVersion,
            HashPackageBytes(origin.BaseBytes),
            store.ExpiryFromNow());
        var retained = new Internal.RetainedPreview(
            retention,
            snapshot,
            HashPackageBytes(snapshot.PackageBytes!),
            origin.TrackedChanges,
            origin.RevisionAuthor,
            result);
        if (store.Add(retained)) return retention;
        warnings.Add(
            $"The preview package ({retained.ApproximateBytes} bytes) exceeds the retained-preview " +
            $"byte budget ({store.ByteBudget} bytes), so it was not retained for commit.");
        return null;
    }

    /// <summary>
    /// Make a retained preview the live document, exactly as previewed (issue #760). The commit
    /// is guarded: it refuses, changing nothing, unless the session is still at the preview's base
    /// version with the same package content, tracked-changes mode and revision author. On
    /// success the previewed package — generated ids, timestamps and all — replaces the live one
    /// as one undoable history step, the version becomes the previewed <c>ResultVersion</c>, and
    /// the previewed receipt is returned with <see cref="MutationBatchResult.Preview"/> false.
    /// A committed preview is consumed; retry deduplication belongs to the transaction journal.
    /// </summary>
    public MutationBatchResult CommitPreview(string previewId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(previewId);
        lock (_mutationGate)
        {
            if (_disposed)
                return CommitPreviewFailure(EditErrorCode.SessionDisposed, "session disposed", null);
            if (!RetainedPreviews.TryGet(previewId, out var retained))
            {
                return CommitPreviewFailure(
                    EditErrorCode.PreviewNotFound,
                    $"no retained preview {previewId}: it was never retained, has expired, was evicted, " +
                    "or was already committed",
                    null);
            }

            var retention = retained.Retention;
            if (_version != retention.BaseVersion)
            {
                return CommitPreviewFailure(
                    EditErrorCode.PreviewStale,
                    $"preview {previewId} was predicted from version {retention.BaseVersion} but the " +
                    $"session is at version {_version}",
                    retention);
            }
            if (_trackedChanges != retained.TrackedChanges
                || !string.Equals(_revisionAuthor, retained.RevisionAuthor, StringComparison.Ordinal))
            {
                return CommitPreviewFailure(
                    EditErrorCode.PreviewStale,
                    $"the tracked-changes mode or revision author changed since preview {previewId}",
                    retention);
            }
            string liveHash;
            try
            {
                liveHash = GetPackageContentHash();
            }
            catch (Exception ex)
            {
                LastInternalError = ex;
                return CommitPreviewFailure(EditErrorCode.InternalError, ex.Message, retention);
            }
            if (!string.Equals(liveHash, retention.BasePackageHash, StringComparison.Ordinal))
            {
                return CommitPreviewFailure(
                    EditErrorCode.PreviewStale,
                    $"the package content changed since preview {previewId} " +
                    $"(predicted from {retention.BasePackageHash}, now {liveHash})",
                    retention);
            }

            RetainedPreviews.Remove(previewId);
            var evidence = _deliveryEvidence?.Begin(
                new[] { ("docx_session", "commit_preview", (string?)("{\"previewId\":" + Internal.DocxSessionJson.JsonString(previewId) + "}")) },
                MutationBatchMode.Atomic, null);
            _history.RecordPreOp(TakePackageSnapshot());
            try
            {
                // The retained snapshot IS the shadow's final state: its package bytes, the
                // predicted version, and the revision generators the batch advanced.
                RestoreSnapshot(retained.Snapshot);
            }
            catch (Exception ex)
            {
                LastInternalError = ex;
                RollbackFailedOp();
                _deliveryEvidence?.Abandon(evidence);
                return CommitPreviewFailure(EditErrorCode.InternalError, ex.Message, retention);
            }
            if (evidence is not null)
                _deliveryEvidence!.CompleteDirect(evidence, new[] { new EditResult { Success = true } });

            var result = retained.Result
                ?? new MutationBatchResult { Mode = MutationBatchMode.Atomic, Success = true };
            return result with
            {
                Preview = false,
                BaseVersion = retention.BaseVersion,
                ResultVersion = _version,
                PackageHash = retained.PackageHash,
                // The generated-value caveats describe a fresh execution; this commit restored
                // the previewed bytes, so every previewed id, timestamp and hash is now exact.
                Warnings = result.Warnings.Where(warning => !BatchReplayCaveats.Contains(warning)).ToArray(),
                Html = null,
                Retention = retention,
            };
        }
    }

    private MutationBatchResult CommitPreviewFailure(
        EditErrorCode code,
        string message,
        MutationPreviewRetention? retention)
    {
        var step = new MutationBatchStepResult(
            0, "docx_session", "commit_preview",
            new[] { new EditResult { Success = false, Error = new EditError(code, message) } },
            false);
        return new MutationBatchResult
        {
            Mode = MutationBatchMode.Atomic,
            Success = false,
            RolledBack = false,
            BaseVersion = _version,
            ResultVersion = _version,
            Steps = new[] { step },
            Failure = BatchFailure(step, rolledBack: false),
            Retention = retention,
        };
    }

    private MutationBatchResult ExecuteAtomicBatch(IReadOnlyList<MutationBatchStep> steps)
    {
        using var transaction = BeginTransaction();

        // Run every available read-only preflight before the first mutation.
        for (int i = 0; i < steps.Count; i++)
        {
            var error = RunBatchPreflight(steps[i]);
            if (error is null) continue;
            transaction.Rollback();
            var failed = new MutationBatchStepResult(
                i, steps[i].Tool, steps[i].Action,
                new[] { new EditResult { Success = false, Error = error } }, true);
            return FailedAtomicBatch(new[] { failed }, failed);
        }

        var results = new List<MutationBatchStepResult>(steps.Count);
        for (int i = 0; i < steps.Count; i++)
        {
            var stepResults = RunBatchMutation(steps[i]);
            var step = new MutationBatchStepResult(
                i, steps[i].Tool, steps[i].Action, stepResults, false);
            results.Add(step);
            if (step.Success) continue;

            transaction.Rollback();
            var rolledBack = results.Select(r => r with { RolledBack = true }).ToArray();
            return FailedAtomicBatch(rolledBack, rolledBack[^1]);
        }

        transaction.Commit();
        return new MutationBatchResult
        {
            Mode = MutationBatchMode.Atomic,
            Success = true,
            RolledBack = false,
            Steps = results,
        };
    }

    private MutationBatchResult ExecuteBestEffortBatch(
        IReadOnlyList<MutationBatchStep> steps,
        Internal.DeliveryEvidenceRecorder.PendingBatch? evidence = null)
    {
        var results = new List<MutationBatchStepResult>(steps.Count);
        MutationBatchStepResult? firstFailure = null;

        for (int i = 0; i < steps.Count; i++)
        {
            // Best-effort preserves sequential semantics: a later preflight may intentionally
            // inspect state produced by an earlier successful step. Only atomic mode preflights
            // the complete batch before its first mutation.
            var stepResults = RunBatchPreflight(steps[i]) is { } error
                ? new[] { new EditResult { Success = false, Error = error } }
                : RunBatchMutation(steps[i]);
            var step = new MutationBatchStepResult(
                i, steps[i].Tool, steps[i].Action, stepResults, false);
            results.Add(step);
            if (!step.Success && firstFailure is null) firstFailure = step;
            // Each best-effort step is its own version step, so it is its own receipt entry.
            if (evidence is not null) _deliveryEvidence!.CompleteStep(evidence, i, step);
        }

        return new MutationBatchResult
        {
            Mode = MutationBatchMode.BestEffort,
            Success = firstFailure is null,
            RolledBack = false,
            Steps = results,
            Failure = firstFailure is null ? null : BatchFailure(firstFailure, rolledBack: false),
        };
    }

    private EditError? RunBatchPreflight(MutationBatchStep step)
    {
        if (step.Preflight is null) return null;
        try
        {
            return step.Preflight(this);
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            return new EditError(EditErrorCode.InternalError, ex.Message);
        }
    }

    private IReadOnlyList<EditResult> RunBatchMutation(MutationBatchStep step)
    {
        try
        {
            var results = step.Mutation(this);
            if (results is null || results.Any(result => result is null))
                return new[]
                {
                    EditResult.Fail(EditErrorCode.InternalError,
                        "batch mutation returned invalid edit results"),
                };
            return results;
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            return new[] { EditResult.Fail(EditErrorCode.InternalError, ex.Message) };
        }
    }

    private static MutationBatchResult FailedAtomicBatch(
        IReadOnlyList<MutationBatchStepResult> results,
        MutationBatchStepResult failed) => new()
    {
        Mode = MutationBatchMode.Atomic,
        Success = false,
        RolledBack = true,
        Steps = results,
        Failure = BatchFailure(failed, rolledBack: true),
    };

    private static MutationBatchFailure BatchFailure(
        MutationBatchStepResult failed,
        bool rolledBack) => new(
            failed.Index,
            failed.Tool,
            failed.Action,
            failed.Results.FirstOrDefault(r => !r.Success)?.Error
                ?? new EditError(EditErrorCode.InternalError, "batch step failed without an error"),
            rolledBack);
}
