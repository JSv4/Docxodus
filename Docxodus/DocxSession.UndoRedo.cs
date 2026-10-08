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
    // ─── Undo / Redo ─────────────────────────────────────────────────────

    /// <summary>
    /// The failure path every mutation's <c>catch</c> shares: keep the exception, roll the op back,
    /// and report the error <see cref="RecordFailedOp"/> chose. One owner, so a change to how a failed
    /// mutation is handled lands once.
    /// </summary>
    private EditResult FailInternal(Exception ex, string? anchorId = null) =>
        EditResult.Fail(RecordFailedOp(ex, anchorId));

    /// <summary>
    /// Keep <paramref name="ex"/> as <see cref="LastInternalError"/>, restore the pre-op snapshot
    /// (<see cref="RollbackFailedOp"/>), and return the error the failing call reports:
    /// <see cref="EditErrorCode.InternalError"/> when the rollback put the document back, or
    /// <see cref="EditErrorCode.SessionCorrupted"/> when the rollback failed too, so the caller learns
    /// from this very call that the session may no longer be edited (issue #963).
    /// </summary>
    private EditError RecordFailedOp(Exception ex, string? anchorId = null)
    {
        LastInternalError = ex;
        RollbackFailedOp();
        return IsCorrupted
            ? new EditError(EditErrorCode.SessionCorrupted,
                $"{ex.Message} (and rolling the edit back failed too, so the document may be half-changed; "
                + "reopen the session from known-good bytes)", anchorId)
            : new EditError(EditErrorCode.InternalError, ex.Message, anchorId);
    }

    /// <summary>
    /// Roll the document back to the pre-op snapshot after a mutation threw partway through.
    /// </summary>
    /// <remarks>
    /// <para>Every mutation calls <see cref="UndoRing{T}.RecordPreOp"/> before entering its
    /// <c>try</c>, so a throw from inside can leave the document HALF-MUTATED — an op that removes
    /// a footnote's body cross-references and then fails before removing the definition has already
    /// stripped the references. Discarding the snapshot without applying it (the historical
    /// <c>_ = _history.PopForUndo()</c>) makes that partial state permanent AND unreachable by
    /// <see cref="Undo"/>, because the record that could have reversed it is the one being thrown
    /// away. Restoring instead is the only outcome that keeps the typed <see cref="EditResult"/>
    /// envelope's implicit promise: a failed op did not change the document.</para>
    ///
    /// <para>Safe to call even when nothing was mutated before the throw — the restore just
    /// re-installs identical XML — so error paths need not reason about how far the op got.
    /// Distinct from the CLEAN-failure paths (<c>else _ = _history.PopForUndo()</c>), which
    /// discard deliberately: those ops detected a problem and returned WITHOUT mutating, so their
    /// snapshot is genuinely spare and must not evict a real edit from the bounded ring.</para>
    ///
    /// <para>A failure of the restore itself is swallowed into <see cref="LastRollbackError"/>
    /// rather than thrown: the caller is already receiving an <see cref="EditErrorCode.InternalError"/>
    /// for the original fault, and replacing that with a secondary one would hide the real cause.
    /// A non-null <see cref="LastRollbackError"/> is the signal that the document may be
    /// inconsistent and the session should be reopened from bytes.</para>
    /// </remarks>
    /// <summary>
    /// Why this session may not mutate, or null when it may: <see cref="EditErrorCode.SessionDisposed"/>
    /// after <see cref="Dispose"/>, <see cref="EditErrorCode.SessionCorrupted"/> once a failed op's
    /// rollback also failed (<see cref="IsCorrupted"/>). Every mutation checks this before touching
    /// the document, so a session whose package may be half-mutated refuses further edits instead of
    /// building on it (issue #963). Reads, <see cref="Save()"/> and <see cref="Dispose"/> are not
    /// refused, so a caller can inspect what happened before reopening from known-good bytes, though a
    /// save can fail on whatever the half-applied op left behind.
    /// </summary>
    private EditError? MutationRefusal()
    {
        if (_disposed) return new EditError(EditErrorCode.SessionDisposed, "session disposed");
        if (LastRollbackError is { } rollbackError)
            return new EditError(EditErrorCode.SessionCorrupted,
                "a failed edit could not be rolled back, so the document may be half-changed; reopen the "
                + $"session from known-good bytes (rollback error: {rollbackError.Message})");
        return null;
    }

    /// <summary><see cref="MutationRefusal"/> for the members with no error channel: throws
    /// <see cref="ObjectDisposedException"/> after <see cref="Dispose"/>, and
    /// <see cref="InvalidOperationException"/> on a corrupted session.</summary>
    private void ThrowIfMutationRefused()
    {
        ThrowIfDisposed();
        if (MutationRefusal() is { } refusal) throw new InvalidOperationException(refusal.Message);
    }

    private void RollbackFailedOp()
    {
        var (preOp, ok) = _history.PopForUndo();
        if (!ok) return;
        try
        {
            RestoreSnapshot(preOp);
        }
        catch (Exception rollbackEx)
        {
            LastRollbackError = rollbackEx;
        }
    }

    public bool Undo()
    {
        if (MutationRefusal() is not null) return false;
        if (_transactions.Count > 0) return false;
        _deliveryEvidence?.Reconcile();
        var nextVersion = NextVersion();
        var (preOp, ok) = _history.PopForUndo();
        if (!ok) return false;
        _history.RecordForRedo(preOp.PackageBytes is null ? TakeSnapshot() : TakePackageSnapshot());
        RestoreSnapshot(preOp);
        _version = nextVersion;
        _deliveryEvidence?.RecordLineage(Verification.DeliveryLineageAction.Undo);
        return true;
    }

    public bool Redo()
    {
        if (MutationRefusal() is not null) return false;
        if (_transactions.Count > 0) return false;
        _deliveryEvidence?.Reconcile();
        var nextVersion = NextVersion();
        var (postOp, ok) = _history.PopForRedo();
        if (!ok) return false;
        _history.PushBackForUndo(postOp.PackageBytes is null ? TakeSnapshot() : TakePackageSnapshot());
        RestoreSnapshot(postOp);
        _version = nextVersion;
        _deliveryEvidence?.RecordLineage(Verification.DeliveryLineageAction.Redo);
        return true;
    }

    private long NextVersion() => checked(_version + 1);

    private void AdvanceVersion() => _version = NextVersion();

    private void OnHistoryRecordPreOp()
    {
        if (_transactions.Count > 0)
        {
            _transactionPendingMutations = checked(_transactionPendingMutations + 1);
            _transactionMutationEpoch = checked(_transactionMutationEpoch + 1);
        }
        else
        {
            // The previous version step is complete and the live package still holds its result:
            // capture it before this op advances the version, so no step goes unrecorded.
            _deliveryEvidence?.Reconcile();
            AdvanceVersion();
        }
    }

    private void OnHistoryPopUndo(DocumentSnapshot snapshot)
    {
        if (_transactions.Count > 0)
        {
            _transactionPendingMutations = Math.Max(0, _transactionPendingMutations - 1);
            return;
        }
        _version = snapshot.Version;
    }
}
