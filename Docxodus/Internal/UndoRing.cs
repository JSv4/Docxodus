#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;

namespace Docxodus.Internal;

/// <summary>
/// Bounded dual-stack ring buffer for undo/redo snapshots. The undo stack holds
/// pre-op snapshots; the redo stack holds post-op snapshots. Recording a new
/// pre-op clears the redo stack (the standard "edit invalidates redo" behavior).
///
/// <para><b>Two bounds, both enforced.</b> A depth alone does not bound memory: each
/// entry is a full deep clone of every snapshot-scoped part, so its cost scales with
/// the DOCUMENT, and a fixed depth on a large document is an unbounded amount of RAM.
/// A 50-deep ring over a long filing is fifty whole-document DOMs held live — fine on a
/// server, fatal in a browser WASM heap. Entries therefore carry a measured cost and the
/// ring evicts the oldest until it is under BOTH the depth cap and the byte budget.</para>
///
/// <para>The most recent entry is never evicted by the budget: a single snapshot larger
/// than the whole budget would otherwise make undo silently unavailable on exactly the
/// documents where a mistake is most expensive. One step of undo is always retained, and
/// <see cref="EvictedForMemory"/> records that the budget bound was the reason so callers
/// can tell "nothing to undo" from "your history was trimmed".</para>
/// </summary>
/// <typeparam name="T">The snapshot type. Cost is read through a caller-supplied selector rather
/// than measured here, so this class stays free of any XML/document dependency — and so the 47
/// existing <c>RecordPreOp(TakeSnapshot())</c> call sites need no change.</typeparam>
internal sealed class UndoRing<T>
{
    private readonly LinkedList<Entry> _undo = new();
    private readonly LinkedList<Entry> _redo = new();
    private readonly int _capacity;
    private readonly long _budgetBytes;
    private readonly Func<T, long>? _costOf;
    private readonly Action<T>? _onRecordPreOp;
    private readonly Action<T>? _onPopUndo;

    private long _undoBytes;
    private long _redoBytes;

    private readonly record struct Entry(T Snapshot, long CostBytes);

    /// <summary>
    /// Opaque copy of both history stacks. Snapshot payloads are intentionally shared rather
    /// than cloned: they are immutable after insertion, and a transaction checkpoint only needs
    /// to retain entries that trimming or redo invalidation might otherwise discard.
    /// </summary>
    internal sealed record State(
        IReadOnlyList<(T Snapshot, long CostBytes)> Undo,
        IReadOnlyList<(T Snapshot, long CostBytes)> Redo,
        long UndoBytes,
        long RedoBytes,
        bool EvictedForMemory);

    /// <param name="capacity">Maximum number of undo entries. Values &lt;= 0 clamp to 1.</param>
    /// <param name="budgetBytes">Approximate byte budget for retained snapshots, counting the
    /// undo and redo sides together. Values &lt;= 0 disable the budget bound (depth only).</param>
    /// <param name="costOf">Approximate retained cost of one snapshot. Null (or a zero budget)
    /// leaves the ring depth-bounded only, exactly as before.</param>
    public UndoRing(
        int capacity,
        long budgetBytes = 0,
        Func<T, long>? costOf = null,
        Action<T>? onRecordPreOp = null,
        Action<T>? onPopUndo = null)
    {
        _capacity = capacity > 0 ? capacity : 1;
        _budgetBytes = budgetBytes > 0 ? budgetBytes : 0;
        _costOf = costOf;
        _onRecordPreOp = onRecordPreOp;
        _onPopUndo = onPopUndo;
    }

    private long CostOf(T snapshot) =>
        _budgetBytes <= 0 || _costOf is null ? 0 : _costOf(snapshot);

    public int UndoCount => _undo.Count;

    public int RedoCount => _redo.Count;

    /// <summary>Approximate bytes currently retained across both stacks.</summary>
    public long RetainedBytes => _undoBytes + _redoBytes;

    /// <summary>True once the byte budget (rather than the depth cap) has discarded at least
    /// one entry. Sticky for the session's lifetime — it answers "was history ever trimmed for
    /// memory?", which is what a caller surfacing a warning needs.</summary>
    public bool EvictedForMemory { get; private set; }

    /// <summary>Record the document state before applying a new mutation.</summary>
    public void RecordPreOp(T preOpSnapshot)
    {
        var costBytes = CostOf(preOpSnapshot);
        _undo.AddLast(new Entry(preOpSnapshot, costBytes));
        _undoBytes += costBytes;
        ClearRedo();
        Trim();
        _onRecordPreOp?.Invoke(preOpSnapshot);
    }

    /// <summary>Pop the most recent pre-op snapshot (for an undo).</summary>
    public (T snapshot, bool ok) PopForUndo()
    {
        if (_undo.Count == 0) return (default!, false);
        var entry = _undo.Last!.Value;
        _undo.RemoveLast();
        _undoBytes -= entry.CostBytes;
        _onPopUndo?.Invoke(entry.Snapshot);
        return (entry.Snapshot, true);
    }

    /// <summary>Record the document state after applying a mutation we just undid.</summary>
    public void RecordForRedo(T postOpSnapshot)
    {
        var costBytes = CostOf(postOpSnapshot);
        _redo.AddLast(new Entry(postOpSnapshot, costBytes));
        _redoBytes += costBytes;
        Trim();
    }

    /// <summary>Pop the most recent post-op snapshot (for a redo).</summary>
    public (T snapshot, bool ok) PopForRedo()
    {
        if (_redo.Count == 0) return (default!, false);
        var entry = _redo.Last!.Value;
        _redo.RemoveLast();
        _redoBytes -= entry.CostBytes;
        return (entry.Snapshot, true);
    }

    /// <summary>Push a snapshot back onto the undo stack (used when applying a redo).</summary>
    public void PushBackForUndo(T snapshot)
    {
        var costBytes = CostOf(snapshot);
        _undo.AddLast(new Entry(snapshot, costBytes));
        _undoBytes += costBytes;
        Trim();
    }

    public void Clear()
    {
        _undo.Clear();
        ClearRedo();
        _undoBytes = 0;
    }

    /// <summary>Capture the exact undo/redo topology for an enclosing transaction.</summary>
    public State CaptureState() => new(
        _undo.Select(e => (e.Snapshot, e.CostBytes)).ToArray(),
        _redo.Select(e => (e.Snapshot, e.CostBytes)).ToArray(),
        _undoBytes,
        _redoBytes,
        EvictedForMemory);

    /// <summary>
    /// Restore a transaction checkpoint without invoking mutation/version callbacks. This also
    /// resurrects entries trimmed while speculative steps were running and restores the sticky
    /// memory-eviction flag, so a rolled-back batch is invisible to history diagnostics.
    /// </summary>
    public void RestoreState(State state)
    {
        ArgumentNullException.ThrowIfNull(state);
        _undo.Clear();
        _redo.Clear();
        foreach (var (snapshot, costBytes) in state.Undo)
            _undo.AddLast(new Entry(snapshot, costBytes));
        foreach (var (snapshot, costBytes) in state.Redo)
            _redo.AddLast(new Entry(snapshot, costBytes));
        _undoBytes = state.UndoBytes;
        _redoBytes = state.RedoBytes;
        EvictedForMemory = state.EvictedForMemory;
    }

    private void ClearRedo()
    {
        _redo.Clear();
        _redoBytes = 0;
    }

    /// <summary>
    /// Enforce both bounds, oldest-first. Redo entries are surrendered before undo entries:
    /// redo is speculative (it only exists between an Undo and the next mutation) whereas an
    /// undo entry is the user's own history.
    /// </summary>
    private void Trim()
    {
        while (_undo.Count > _capacity)
        {
            _undoBytes -= _undo.First!.Value.CostBytes;
            _undo.RemoveFirst();
        }

        if (_budgetBytes <= 0) return;

        while (RetainedBytes > _budgetBytes && _redo.Count > 0)
        {
            _redoBytes -= _redo.First!.Value.CostBytes;
            _redo.RemoveFirst();
            EvictedForMemory = true;
        }

        // Never evict the last undo entry: one step back must stay available even when a single
        // snapshot exceeds the whole budget.
        while (RetainedBytes > _budgetBytes && _undo.Count > 1)
        {
            _undoBytes -= _undo.First!.Value.CostBytes;
            _undo.RemoveFirst();
            EvictedForMemory = true;
        }

        if (_undoBytes < 0) _undoBytes = 0;
        if (_redoBytes < 0) _redoBytes = 0;
    }
}
