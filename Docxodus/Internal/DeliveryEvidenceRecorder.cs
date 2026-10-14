// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using Docxodus.Delivery;
using Docxodus.Verification;

namespace Docxodus.Internal;

/// <summary>
/// Host-owned capture of the evidence a delivery change receipt needs (issue #748): the exact
/// package before and after every version step of a live session, the normalized request that
/// produced it when a transport described one, the transaction identity a retry journal bound
/// it to, and every undo/redo as a lineage event. Evidence is captured when edits execute, not
/// reconstructed from response summaries.
/// </summary>
/// <remarks>
/// <para>
/// The unit of evidence is one session version step, which is also the receipt's unit: the
/// lineage validator requires each committed transaction to advance the version by exactly one
/// and each undo/redo to restore a recorded state. An atomic batch is one version step. A
/// best-effort batch advances the version once per successful step, so it is recorded as one
/// entry per step and the batch's transaction identity rides on the first of them. A version
/// step nobody described — a typed or browser direct call, a caller's own transaction scope —
/// is recorded as <c>docx_session/unlabeled_mutation</c> with exact packages but no request.
/// </para>
/// <para>
/// States are serialized the way a clean save serializes them (projector bookkeeping stripped)
/// from a package clone, so recording never perturbs the live caches an in-flight operation
/// holds. Retention is bounded by a state count and a byte budget; once either is exceeded,
/// or once the chain cannot be continued truthfully, the recorder releases its history and
/// reports the first reason. A receipt is then unavailable with that reason — it never claims
/// a partial history.
/// </para>
/// </remarks>
internal sealed class DeliveryEvidenceRecorder
{
    public const int DefaultMaxStates = 256;
    public const long DefaultByteBudget = 512L * 1024 * 1024;
    public const string UnlabeledTool = "docx_session";
    public const string UnlabeledAction = "unlabeled_mutation";

    private readonly DocxSession _session;
    private int _maxStates;
    private long _byteBudget;
    private readonly DeliveryReceiptLimits _limits = new();
    private readonly PackageManifestOptions _manifestOptions = new();
    private readonly List<HistoryItem> _history = new();
    private readonly List<string> _warnings = new();
    private readonly List<string> _applied = new();
    private readonly List<string> _redo = new();
    private readonly State _source;
    private State _current;
    private long _nextSequence;
    private int _stateCount;
    private long _retainedBytes;
    private int _unlabeled;
    private string? _unavailableReason;
    private PendingBatch? _pending;

    public DeliveryEvidenceRecorder(
        DocxSession session,
        byte[] sourceBytes,
        long sourceVersion,
        int maxStates = DefaultMaxStates,
        long byteBudget = DefaultByteBudget)
    {
        _session = session ?? throw new ArgumentNullException(nameof(session));
        ArgumentNullException.ThrowIfNull(sourceBytes);
        if (maxStates < 2) throw new ArgumentOutOfRangeException(nameof(maxStates));
        if (byteBudget < 1) throw new ArgumentOutOfRangeException(nameof(byteBudget));
        _maxStates = maxStates;
        _byteBudget = byteBudget;
        _source = MakeState(sourceBytes, sourceVersion);
        _current = _source;
        _stateCount = 1;
        _retainedBytes = sourceBytes.LongLength;
    }

    /// <summary>A described batch in flight; version steps observed meanwhile belong to it.</summary>
    internal sealed class PendingBatch
    {
        internal PendingBatch(
            IReadOnlyList<DeliveryNormalizedOperation> operations,
            MutationBatchMode mode,
            DeliveryTransactionIdentity? identity)
        {
            Operations = operations;
            Mode = mode;
            Identity = identity;
        }

        internal IReadOnlyList<DeliveryNormalizedOperation> Operations { get; }
        internal MutationBatchMode Mode { get; }
        internal DeliveryTransactionIdentity? Identity { get; }
        internal List<(State Before, State After)> Transitions { get; } = new();
        internal bool IdentityAssigned { get; set; }
    }

    internal sealed record State(
        long Version,
        byte[] Bytes,
        PackageManifest Manifest,
        DeliveryDocumentIdentity Identity);

    private abstract record HistoryItem(long Sequence);

    private sealed record TransactionItem(
        long Sequence,
        string EntryId,
        DeliveryTransactionContribution Contribution,
        State Before,
        State After) : HistoryItem(Sequence);

    private sealed record LineageItem(long Sequence, DeliveryLineageEventInput Event)
        : HistoryItem(Sequence);

    public bool IsAvailable => _unavailableReason is null;

    /// <summary>Retention bounds; tests tighten them to exercise eviction.</summary>
    internal void SetLimits(int maxStates, long byteBudget)
    {
        if (maxStates < 2) throw new ArgumentOutOfRangeException(nameof(maxStates));
        if (byteBudget < 1) throw new ArgumentOutOfRangeException(nameof(byteBudget));
        _maxStates = maxStates;
        _byteBudget = byteBudget;
    }

    /// <summary>
    /// Bring the recorded chain up to the live package. Called before every recorded mutation
    /// (from the history pre-op hook, outside transaction scopes), before undo/redo, and before
    /// export, so a version step nobody described is still captured exactly, as an unlabeled
    /// transaction, before anything else happens.
    /// </summary>
    public void Reconcile()
    {
        if (_unavailableReason is not null) return;
        var live = _session.Version;
        if (live == _current.Version) return;
        if (live != _current.Version + 1)
        {
            Break($"the session moved from version {_current.Version} to {live} without a recordable step");
            return;
        }
        if (!TryCapture(live, out var next)) return;
        if (_pending is not null)
            _pending.Transitions.Add((_current, next));
        else
            RecordUnlabeled(_current, next);
        _current = next;
    }

    /// <summary>
    /// Open a described batch. Returns null when nothing will be labeled: capture is already
    /// unavailable, or the batch runs inside a caller's transaction scope (the scope's commit is
    /// the version step, and it is recorded unlabeled when it lands).
    /// </summary>
    public PendingBatch? Begin(
        IReadOnlyList<MutationBatchStep> steps,
        MutationBatchMode mode,
        DeliveryTransactionIdentity? identity)
    {
        Reconcile();
        if (_unavailableReason is not null) return null;
        if (_pending is not null)
        {
            Break("a described batch began while another was still in flight");
            return null;
        }
        if (_session.InTransactionScope)
        {
            _warnings.Add(
                "A batch executed inside a caller transaction scope; the scope's commit is recorded as one unlabeled mutation.");
            return null;
        }
        var operations = new List<DeliveryNormalizedOperation>(steps.Count);
        foreach (var step in steps)
            operations.Add(Describe(step.Tool, step.Action, step.ArgumentsJson));
        _pending = new PendingBatch(operations, mode, identity);
        return _pending;
    }

    /// <summary>Open a described batch from a transport's operation descriptors.</summary>
    public PendingBatch? Begin(
        IReadOnlyList<(string Tool, string Action, string? ArgumentsJson)> operations,
        MutationBatchMode mode,
        DeliveryTransactionIdentity? identity)
    {
        Reconcile();
        if (_unavailableReason is not null) return null;
        if (_pending is not null)
        {
            Break("a described batch began while another was still in flight");
            return null;
        }
        if (_session.InTransactionScope)
        {
            _warnings.Add(
                "A batch executed inside a caller transaction scope; the scope's commit is recorded as one unlabeled mutation.");
            return null;
        }
        var described = operations.Select(op => Describe(op.Tool, op.Action, op.ArgumentsJson)).ToArray();
        _pending = new PendingBatch(described, mode, identity);
        return _pending;
    }

    /// <summary>Record an atomic batch: exactly one version step, or none when it rolled back.</summary>
    public void CompleteAtomic(PendingBatch pending, MutationBatchResult result)
    {
        if (!ReferenceEquals(pending, _pending)) return;
        Reconcile();
        _pending = null;
        if (_unavailableReason is not null) return;
        var transitions = pending.Transitions;
        if (transitions.Count > 1)
        {
            Break($"an atomic batch advanced the version by {transitions.Count}; a receipt transaction is one version step");
            return;
        }
        var (before, after) = transitions.Count == 1 ? transitions[0] : (_current, _current);
        if (result.BaseVersion != before.Version || result.ResultVersion != after.Version)
        {
            Break(
                $"a batch result reported versions {result.BaseVersion}->{result.ResultVersion} but the " +
                $"session moved {before.Version}->{after.Version}");
            return;
        }
        RecordTransaction(result, pending.Operations, pending.Identity, before, after);
    }

    /// <summary>
    /// Record one best-effort step as its own entry: a failed step is a same-state failure,
    /// a successful step is the version step it produced. A step that produced more than one
    /// version step keeps its label on the first and the rest are recorded unlabeled.
    /// </summary>
    public void CompleteStep(PendingBatch pending, int index, MutationBatchStepResult step)
    {
        if (!ReferenceEquals(pending, _pending)) return;
        Reconcile();
        if (_unavailableReason is not null) return;
        if (index < 0 || index >= pending.Operations.Count)
        {
            Break($"a best-effort step index {index} has no described operation");
            return;
        }
        var transitions = pending.Transitions.ToArray();
        pending.Transitions.Clear();
        var identity = pending.IdentityAssigned ? null : pending.Identity;
        pending.IdentityAssigned = true;
        var operation = new[] { pending.Operations[index] };
        var indexed = step with { Index = 0 };
        if (transitions.Length == 0)
        {
            RecordTransaction(
                SingleStepResult(MutationBatchMode.BestEffort, indexed, _current.Version, _current.Version),
                operation, identity, _current, _current);
            return;
        }
        var (before, after) = transitions[0];
        RecordTransaction(
            SingleStepResult(MutationBatchMode.BestEffort, indexed, before.Version, after.Version),
            operation, identity, before, after);
        if (transitions.Length > 1)
        {
            _warnings.Add(
                $"Best-effort step {index} ({step.Tool}/{step.Action}) advanced the version by " +
                $"{transitions.Length}; the additional steps are recorded unlabeled.");
            foreach (var extra in transitions.Skip(1))
                RecordUnlabeled(extra.Before, extra.After);
        }
    }

    /// <summary>Close a best-effort batch whose steps were recorded through <see cref="CompleteStep"/>.</summary>
    public void CompleteBestEffort(PendingBatch pending)
    {
        if (!ReferenceEquals(pending, _pending)) return;
        Reconcile();
        _pending = null;
        if (_unavailableReason is not null) return;
        foreach (var (before, after) in pending.Transitions)
            RecordUnlabeled(before, after);
        pending.Transitions.Clear();
        if (pending.Identity is { } identity && pending.Operations.Count > 1)
        {
            _warnings.Add(
                $"Best-effort batch '{identity.TransactionId}' is recorded as {pending.Operations.Count} " +
                "per-step transactions; its transaction identity is attached to the first.");
        }
    }

    /// <summary>
    /// Record a transport-composed batch from its step results: atomic batches as one entry,
    /// best-effort batches step by step when each successful step maps to one observed version
    /// step, and otherwise unlabeled with a warning.
    /// </summary>
    public void CompleteClient(PendingBatch pending, IReadOnlyList<MutationBatchStepResult> steps)
    {
        if (!ReferenceEquals(pending, _pending)) return;
        Reconcile();
        if (_unavailableReason is not null)
        {
            _pending = null;
            return;
        }
        if (pending.Mode == MutationBatchMode.Atomic)
        {
            var transitions = pending.Transitions;
            var (before, after) = transitions.Count == 1 ? transitions[0] : (_current, _current);
            var success = steps.Count == pending.Operations.Count && steps.All(step => step.Success);
            var result = new MutationBatchResult
            {
                Mode = MutationBatchMode.Atomic,
                Success = success,
                RolledBack = !success,
                BaseVersion = before.Version,
                ResultVersion = after.Version,
                Steps = steps,
                Failure = success ? null : FirstFailure(steps, rolledBack: true),
            };
            CompleteAtomic(pending, result);
            return;
        }

        var successful = steps.Count(step => step.Success);
        if (pending.Transitions.Count != successful)
        {
            _warnings.Add(
                $"A client-composed best-effort batch produced {pending.Transitions.Count} version steps for " +
                $"{successful} successful steps; its version steps are recorded unlabeled.");
            CompleteBestEffort(pending);
            return;
        }
        var queue = new Queue<(State Before, State After)>(pending.Transitions);
        pending.Transitions.Clear();
        foreach (var step in steps)
        {
            if (step.Success) pending.Transitions.Add(queue.Dequeue());
            CompleteStep(pending, step.Index, step);
            if (_unavailableReason is not null) return;
        }
        CompleteBestEffort(pending);
    }

    /// <summary>
    /// Record a described single operation (a direct tool call, a preview commit) from its edit
    /// results: one atomic entry whose version step is whatever the session produced.
    /// </summary>
    public void CompleteDirect(PendingBatch pending, IReadOnlyList<EditResult> results)
    {
        if (!ReferenceEquals(pending, _pending)) return;
        Reconcile();
        if (_unavailableReason is not null)
        {
            _pending = null;
            return;
        }
        var operation = pending.Operations[0];
        var success = results.All(result => result.Success);
        var (before, after) = pending.Transitions.Count == 1 ? pending.Transitions[0] : (_current, _current);
        var step = new MutationBatchStepResult(0, operation.Tool, operation.Action, results, !success);
        CompleteAtomic(pending, SingleStepResult(MutationBatchMode.Atomic, step, before.Version, after.Version));
    }

    /// <summary>A described batch that threw: whatever version steps landed are recorded unlabeled.</summary>
    public void Abandon(PendingBatch? pending)
    {
        if (pending is null || !ReferenceEquals(pending, _pending)) return;
        Reconcile();
        _pending = null;
        if (_unavailableReason is not null) return;
        foreach (var (before, after) in pending.Transitions)
            RecordUnlabeled(before, after);
        pending.Transitions.Clear();
    }

    /// <summary>
    /// Record an undo or redo after the session restored the state. The restored package must
    /// be the recorded before-state (undo) or after-state (redo) of the transaction the history
    /// cursor moved over; anything else means the history reached a state this recorder never
    /// saw, and the chain stops.
    /// </summary>
    public void RecordLineage(DeliveryLineageAction action)
    {
        if (_unavailableReason is not null) return;
        if (_pending is not null)
        {
            Break($"{action} ran while a described batch was in flight");
            return;
        }
        var live = _session.Version;
        if (live != _current.Version + 1)
        {
            Break($"{action} moved the session from version {_current.Version} to {live}");
            return;
        }
        var stack = action == DeliveryLineageAction.Undo ? _applied : _redo;
        if (stack.Count == 0)
        {
            Break($"{action} restored a state that was not recorded as a transaction");
            return;
        }
        var affectedId = stack[^1];
        var affected = _history.OfType<TransactionItem>().First(item => item.EntryId == affectedId);
        if (!TryCapture(live, out var next)) return;
        var expected = action == DeliveryLineageAction.Undo ? affected.Before : affected.After;
        if (!DeliveryReceiptLineageValidator.SamePackageContent(expected.Identity, next.Identity))
        {
            Break($"{action} restored a package that differs from the recorded state it should have reached");
            return;
        }
        _history.Add(new LineageItem(_nextSequence++, new DeliveryLineageEventInput
        {
            Action = action,
            AffectedEntryId = affectedId,
            BeforeDocument = _current.Identity,
            AfterDocument = next.Identity,
        }));
        stack.RemoveAt(stack.Count - 1);
        (action == DeliveryLineageAction.Undo ? _redo : _applied).Add(affectedId);
        _current = next;
    }

    public DeliveryEvidenceStatus Status() => new()
    {
        Enabled = true,
        TransactionCount = _history.OfType<TransactionItem>().Count(),
        LineageEventCount = _history.OfType<LineageItem>().Count(),
        UnlabeledTransactionCount = _unlabeled,
        RetainedStateCount = _stateCount,
        RetainedBytes = _retainedBytes,
        SourceVersion = _source.Version,
        CurrentVersion = _current.Version,
        UnavailableReason = _unavailableReason,
    };

    /// <summary>
    /// The captured history as a delivery operation consumes it. The working snapshot is the
    /// recorder's own current state, so the delivered package is byte-identical to the last
    /// recorded after-state by construction.
    /// </summary>
    public DeliveryEvidenceExport Export(DeliveryReceiptBuildOptions options)
    {
        Reconcile();
        if (_pending is not null)
            Break("evidence was exported while a described batch was in flight");
        var source = new DeliveryDocumentSnapshot("source", _source.Version, _source.Bytes);
        if (_unavailableReason is not null)
        {
            var live = _session.Version;
            var bytes = live == _current.Version ? _current.Bytes : _session.SerializeCleanCheckpoint();
            return new DeliveryEvidenceExport(
                null, source, new DeliveryDocumentSnapshot("working", live, bytes), Status());
        }
        var history = _history.OrderBy(item => item.Sequence).Select(item => item switch
        {
            TransactionItem transaction => DeliveryReceiptHistoryEvent.FromTransaction(
                new DeliveryReceiptTransactionEvidence(
                    transaction.Contribution,
                    new DeliveryDocumentSnapshot(
                        "transaction-before", transaction.Before.Version, transaction.Before.Bytes),
                    new DeliveryDocumentSnapshot(
                        "transaction-after", transaction.After.Version, transaction.After.Bytes))),
            LineageItem lineage => DeliveryReceiptHistoryEvent.FromLineage(lineage.Event),
            _ => throw new InvalidOperationException("unknown history item"),
        }).ToArray();
        var context = new DeliveryReceiptContext(
            history,
            attributionRules: null,
            warnings: _warnings.Distinct(StringComparer.Ordinal),
            privacyProfile: options.PrivacyProfile,
            failOnUnexpectedChanges: options.FailOnUnexpectedChanges);
        return new DeliveryEvidenceExport(
            context,
            source,
            new DeliveryDocumentSnapshot("working", _current.Version, _current.Bytes),
            Status());
    }

    public void Clear()
    {
        _history.Clear();
        _applied.Clear();
        _redo.Clear();
        _warnings.Clear();
        _pending = null;
        _stateCount = 0;
        _retainedBytes = 0;
        _unavailableReason ??= "the session was disposed";
    }

    private DeliveryNormalizedOperation Describe(string tool, string action, string? argumentsJson)
    {
        try
        {
            return DeliveryNormalizedOperation.Create(tool, action, argumentsJson ?? "{}", _limits);
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException)
        {
            _warnings.Add($"Arguments of {tool}/{action} were not recordable ({ex.Message}); recorded without them.");
            return DeliveryNormalizedOperation.Create(tool, action, "{}", _limits);
        }
    }

    private void RecordUnlabeled(State before, State after)
    {
        var step = new MutationBatchStepResult(
            0, UnlabeledTool, UnlabeledAction, new[] { new EditResult { Success = true } }, false);
        var operation = new[] { DeliveryNormalizedOperation.Create(UnlabeledTool, UnlabeledAction, "{}", _limits) };
        RecordTransaction(
            SingleStepResult(MutationBatchMode.Atomic, step, before.Version, after.Version),
            operation, null, before, after);
        _unlabeled++;
        _warnings.Add(
            $"Version {before.Version}->{after.Version} was applied outside a described transaction; " +
            "its packages are exact but its request is unknown.");
    }

    private void RecordTransaction(
        MutationBatchResult result,
        IReadOnlyList<DeliveryNormalizedOperation> operations,
        DeliveryTransactionIdentity? identity,
        State before,
        State after)
    {
        DeliveryTransactionContribution contribution;
        try
        {
            contribution = DeliveryTransactionContribution.FromMutationBatchResult(
                result, before.Manifest, after.Manifest, operations, identity, _limits);
        }
        catch (DeliveryReceiptValidationException ex)
        {
            Break($"a transaction could not be recorded as receipt evidence ({ex.Code}: {ex.Message})");
            return;
        }
        var fingerprint = identity?.RequestFingerprint
            ?? DeliveryReceiptIdentity.RequestFingerprint(result.Mode, operations, _limits);
        var entryId = DeliveryReceiptIdentity.TransactionEntryId(
            fingerprint,
            contribution.BeforeDocument,
            contribution.AfterDocument,
            before.Version,
            after.Version,
            identity?.TransactionId,
            _nextSequence);
        _history.Add(new TransactionItem(_nextSequence++, entryId, contribution, before, after));
        if (!DeliveryReceiptLineageValidator.DocumentEquals(
                contribution.BeforeDocument, contribution.AfterDocument))
        {
            _applied.Add(entryId);
            _redo.Clear();
        }
    }

    private static MutationBatchResult SingleStepResult(
        MutationBatchMode mode,
        MutationBatchStepResult step,
        long baseVersion,
        long resultVersion)
    {
        var rolledBack = mode == MutationBatchMode.Atomic && !step.Success;
        var indexed = step with { Index = 0, RolledBack = rolledBack };
        return new MutationBatchResult
        {
            Mode = mode,
            Success = step.Success,
            RolledBack = rolledBack,
            BaseVersion = baseVersion,
            ResultVersion = resultVersion,
            Steps = new[] { indexed },
            Failure = step.Success ? null : FirstFailure(new[] { indexed }, rolledBack),
        };
    }

    private static MutationBatchFailure? FirstFailure(IReadOnlyList<MutationBatchStepResult> steps, bool rolledBack)
    {
        var failed = steps.FirstOrDefault(step => !step.Success);
        if (failed is null) return null;
        var error = failed.Results.FirstOrDefault(value => !value.Success)?.Error
            ?? new EditError(EditErrorCode.InternalError, "batch step failed without an error");
        return new MutationBatchFailure(failed.Index, failed.Tool, failed.Action, error, rolledBack);
    }

    private bool TryCapture(long version, out State state)
    {
        state = null!;
        try
        {
            var bytes = _session.SerializeCleanCheckpoint();
            state = MakeState(bytes, version);
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or System.IO.IOException)
        {
            Break($"the package at version {version} could not be captured ({ex.Message})");
            return false;
        }
        _stateCount++;
        _retainedBytes += state.Bytes.LongLength;
        if (_stateCount > _maxStates || _retainedBytes > _byteBudget)
        {
            Break(
                $"evidence retention exceeded ({_stateCount} package states, {_retainedBytes} bytes; " +
                $"the limits are {_maxStates} and {_byteBudget})");
            return false;
        }
        return true;
    }

    private State MakeState(byte[] bytes, long version)
    {
        var manifest = PackageManifestGenerator.Generate(bytes, _manifestOptions);
        return new State(version, bytes, manifest, DeliveryDocumentIdentity.FromManifest(manifest, version));
    }

    private void Break(string reason)
    {
        _unavailableReason ??= reason;
        _history.Clear();
        _applied.Clear();
        _redo.Clear();
        _pending = null;
        _stateCount = 0;
        _retainedBytes = 0;
    }
}
