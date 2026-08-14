#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Batch semantics for <see cref="DocxSession"/> (DS43x).
///
/// <para><see cref="DocxSession.Batch"/> exists because a snapshot costs whatever the DOCUMENT
/// costs, not whatever the edit costs: N mutations in a loop paid N whole-document deep clones and
/// consumed N entries of a bounded undo ring, which on a long filing means the sequence a caller
/// just applied is only partially reversible. The contract pinned here is that a batch is <b>one
/// snapshot, one undo step, and all-or-nothing by default</b>.</para>
///
/// <para>The rollback tests reuse the DS42x trigger — a payload XML cannot encode (<c>U+0000</c>)
/// makes an op throw well after it has started mutating — because a batch's hardest promise is the
/// one that only matters when a step fails partway: no per-step snapshot exists to unpick it with,
/// so the batch must reverse from its own.</para>
/// </summary>
public class DocxSessionBatchTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>A payload XML cannot encode — the same realistic trigger DS42x uses.</summary>
    private const string NulPayload = "note\0text";

    private static string[] BodyParagraphs(DocxSession s) =>
        s.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h")
            .Select(t => t.Anchor.Id)
            .ToArray();

    private static string BodyText(DocxSession s)
    {
        using var ms = new MemoryStream(s.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        return string.Concat(
            doc.MainDocumentPart!.GetXDocument().Descendants(W + "t").Select(t => t.Value));
    }

    // ─── One snapshot, one undo step ─────────────────────────────────────

    /// <summary>
    /// The headline property. Three edits through a batch consume ONE undo entry and one
    /// <see cref="DocxSession.Undo"/> reverses all three — where the same three applied in a loop
    /// consume three entries and need three undos.
    /// </summary>
    [Fact]
    public void DS430_Batch_RecordsExactlyOneUndoStepForTheWholeSequence()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);
        var before = BodyText(s);

        var result = s.Batch(new Func<EditResult>[]
        {
            () => s.ReplaceText(paragraphs[0], "first"),
            () => s.ReplaceText(paragraphs[1], "second"),
            () => s.InsertParagraph(paragraphs[1], Position.After, "third"),
        });

        Assert.True(result.Success);
        Assert.Equal(3, result.Applied);
        Assert.Equal(3, result.Steps.Count);
        Assert.False(result.RolledBack);
        Assert.Equal(1, s.UndoCount);

        var after = BodyText(s);
        Assert.Contains("first", after);
        Assert.Contains("third", after);

        Assert.True(s.Undo());
        Assert.Equal(before, BodyText(s));
        Assert.Equal(0, s.UndoCount);
    }

    /// <summary>The loop this replaces, asserted directly — otherwise DS430's "one" has nothing to
    /// be one INSTEAD OF, and a regression that silently restored per-step snapshots would pass.</summary>
    [Fact]
    public void DS431_TheSameEditsWithoutABatchStillCostOneUndoStepEach()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);

        Assert.True(s.ReplaceText(paragraphs[0], "first").Success);
        Assert.True(s.ReplaceText(paragraphs[1], "second").Success);

        Assert.Equal(2, s.UndoCount);
    }

    /// <summary>A batch's single entry must also redo as one step.</summary>
    [Fact]
    public void DS432_Batch_RedoesAsOneStep()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);

        Assert.True(s.Batch(new Func<EditResult>[]
        {
            () => s.ReplaceText(paragraphs[0], "alpha"),
            () => s.ReplaceText(paragraphs[1], "beta"),
        }).Success);
        var applied = BodyText(s);

        Assert.True(s.Undo());
        Assert.DoesNotContain("alpha", BodyText(s));

        Assert.True(s.Redo());
        Assert.Equal(applied, BodyText(s));
    }

    // ─── Failure policy ──────────────────────────────────────────────────

    /// <summary>
    /// Atomic (the default): a clean failure mid-sequence reverses the steps that already
    /// succeeded, so a failed batch never leaves a half-applied document.
    /// </summary>
    [Fact]
    public void DS433_AtomicBatch_ReversesEarlierStepsWhenALaterStepFailsCleanly()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);
        var before = BodyText(s);

        var result = s.Batch(new Func<EditResult>[]
        {
            () => s.ReplaceText(paragraphs[0], "applied"),
            () => s.ReplaceText("p:body:deadbeefdeadbeefdeadbeefdeadbeef", "never"),
        });

        Assert.False(result.Success);
        Assert.True(result.RolledBack);
        Assert.Equal(0, result.Applied);
        Assert.Equal(1, result.FailedStep);
        Assert.Equal(EditErrorCode.AnchorNotFound, result.Error!.Code);

        Assert.Equal(before, BodyText(s));
        Assert.DoesNotContain("applied", BodyText(s));
    }

    /// <summary>A reversed batch leaves the ring as it found it — nothing to undo, and in
    /// particular not the caller's PREVIOUS edit.</summary>
    [Fact]
    public void DS434_RolledBackBatch_DoesNotConsumeOrPolluteTheUndoRing()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);

        Assert.True(s.ReplaceText(paragraphs[0], "the real edit").Success);
        var afterRealEdit = BodyText(s);
        Assert.Equal(1, s.UndoCount);

        var result = s.Batch(new Func<EditResult>[]
        {
            () => s.ReplaceText(paragraphs[1], "doomed"),
            () => s.ReplaceText("p:body:deadbeefdeadbeefdeadbeefdeadbeef", "never"),
        });
        Assert.False(result.Success);

        // The batch is gone from the ring entirely; the next undo reverses the REAL edit.
        Assert.Equal(1, s.UndoCount);
        Assert.Equal(afterRealEdit, BodyText(s));
        Assert.True(s.Undo());
        Assert.DoesNotContain("the real edit", BodyText(s));
    }

    /// <summary>
    /// Best-effort tolerates the failures that provably did not touch the document, keeps the
    /// successes, and still costs one undo step.
    /// </summary>
    [Fact]
    public void DS435_BestEffortBatch_KeepsSuccessesPastACleanFailure()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);
        var before = BodyText(s);

        var result = s.Batch(
            new Func<EditResult>[]
            {
                () => s.ReplaceText(paragraphs[0], "kept one"),
                () => s.ReplaceText("p:body:deadbeefdeadbeefdeadbeefdeadbeef", "never"),
                () => s.ReplaceText(paragraphs[1], "kept two"),
            },
            new BatchOptions { Atomic = false, StopOnError = false });

        Assert.False(result.Success);           // a step failed…
        Assert.False(result.RolledBack);        // …but nothing was reversed
        Assert.Equal(2, result.Applied);
        Assert.Equal(3, result.Steps.Count);
        Assert.Equal(1, result.FailedStep);

        var after = BodyText(s);
        Assert.Contains("kept one", after);
        Assert.Contains("kept two", after);

        // Still ONE undo step, and it reverses both surviving edits together.
        Assert.Equal(1, s.UndoCount);
        Assert.True(s.Undo());
        Assert.Equal(before, BodyText(s));
    }

    /// <summary>Best-effort with stop-on-error attempts nothing after the first failure.</summary>
    [Fact]
    public void DS436_BestEffortBatch_StopOnError_AttemptsNothingAfterTheFailure()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);

        var result = s.Batch(
            new Func<EditResult>[]
            {
                () => s.ReplaceText(paragraphs[0], "kept"),
                () => s.ReplaceText("p:body:deadbeefdeadbeefdeadbeefdeadbeef", "never"),
                () => s.ReplaceText(paragraphs[1], "not attempted"),
            },
            new BatchOptions { Atomic = false, StopOnError = true });

        Assert.False(result.Success);
        Assert.Equal(2, result.Steps.Count);          // the third was never run
        Assert.Equal(1, result.Applied);
        Assert.DoesNotContain("not attempted", BodyText(s));
    }

    // ─── Damage overrides policy ─────────────────────────────────────────

    /// <summary>
    /// The sharp case. A step that throws partway has already mutated, and inside a batch there is
    /// no per-step snapshot left to unpick it with — so the batch reverses EVEN under best-effort,
    /// where a clean failure would have been tolerated. Without this, batching would silently
    /// reintroduce exactly the half-applied-mutation bug the per-op rollback fixed.
    /// </summary>
    [Fact]
    public void DS437_BestEffortBatch_StillReversesWhenAStepThrewPartway()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);
        var before = BodyText(s);

        var result = s.Batch(
            new Func<EditResult>[]
            {
                () => s.ReplaceText(paragraphs[0], "applied first"),
                () => s.InsertFootnote(paragraphs[1], 0, NulPayload),   // throws mid-op
                () => s.ReplaceText(paragraphs[1], "never reached"),
            },
            new BatchOptions { Atomic = false, StopOnError = false });

        Assert.False(result.Success);
        Assert.True(result.RolledBack);
        Assert.Equal(0, result.Applied);
        Assert.Equal(EditErrorCode.InternalError, result.Error!.Code);

        Assert.Equal(before, BodyText(s));
        Assert.DoesNotContain("applied first", BodyText(s));
        Assert.Equal(0, s.UndoCount);
    }

    /// <summary>The scaffold a thrown step built must not survive the batch either — the DS421
    /// property, one level up.</summary>
    [Fact]
    public void DS438_ThrownStepInABatch_LeavesNoOrphanedScaffold()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);

        Assert.False(s.Batch(new Func<EditResult>[]
        {
            () => s.ReplaceText(paragraphs[0], "applied first"),
            () => s.InsertFootnote(paragraphs[1], 0, NulPayload),
        }).Success);

        using var ms = new MemoryStream(s.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;

        Assert.Null(main.FootnotesPart);
        Assert.Null(main.DocumentSettingsPart?.GetXDocument().Root?.Element(W + "footnotePr"));
        Assert.Empty(main.GetXDocument().Descendants(W + "footnoteReference"));
    }

    /// <summary>A step that throws OUT of the batch — rather than one whose op caught and reported —
    /// is contained the same way: the exception does not escape, and the document is restored.</summary>
    [Fact]
    public void DS439_StepThrowingOutOfTheBatch_IsContainedAndReversed()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);
        var before = BodyText(s);

        var result = s.Batch(new Func<EditResult>[]
        {
            () => s.ReplaceText(paragraphs[0], "applied first"),
            () => throw new InvalidOperationException("caller bug"),
        });

        Assert.False(result.Success);
        Assert.True(result.RolledBack);
        Assert.Equal(EditErrorCode.InternalError, result.Error!.Code);
        Assert.Equal("caller bug", result.Error.Message);
        Assert.Equal(before, BodyText(s));
    }

    // ─── Shape ───────────────────────────────────────────────────────────

    /// <summary>An empty batch is a success that touches nothing — no snapshot, no undo entry.</summary>
    [Fact]
    public void DS43A_EmptyBatch_IsANoOp()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());

        var result = s.Batch(Array.Empty<Func<EditResult>>());

        Assert.True(result.Success);
        Assert.Empty(result.Steps);
        Assert.Equal(0, s.UndoCount);
    }

    /// <summary>The aggregate anchor lists are the union across applied steps, de-duplicated — a
    /// host repaints from these, so a paragraph edited twice must appear once.</summary>
    [Fact]
    public void DS43B_Batch_AggregatesAndDedupesTouchedAnchors()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);

        var result = s.Batch(new Func<EditResult>[]
        {
            () => s.ReplaceText(paragraphs[0], "once"),
            () => s.ReplaceText(paragraphs[0], "twice"),
            () => s.InsertParagraph(paragraphs[1], Position.After, "new one"),
        });

        Assert.True(result.Success);
        Assert.Single(result.Modified);
        Assert.Equal(paragraphs[0], result.Modified[0].Id);
        Assert.Single(result.Created);
        Assert.Equal(result.Modified.Select(a => a.Id).Distinct().Count(), result.Modified.Count);
    }

    /// <summary>A batch nested inside a step joins the outer one rather than opening a second
    /// snapshot: the whole tree is still one undo step, and the outer batch owns the rollback.</summary>
    [Fact]
    public void DS43C_NestedBatch_JoinsTheOuterOne()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);
        var before = BodyText(s);

        var result = s.Batch(new Func<EditResult>[]
        {
            () => s.ReplaceText(paragraphs[0], "outer"),
            () => s.Batch(new Func<EditResult>[]
            {
                () => s.ReplaceText(paragraphs[1], "inner one"),
                () => s.InsertParagraph(paragraphs[1], Position.After, "inner two"),
            }).Success
                ? new EditResult { Success = true }
                : EditResult.Fail(EditErrorCode.InternalError, "nested batch failed"),
        });

        Assert.True(result.Success);
        Assert.Equal(1, s.UndoCount);

        var after = BodyText(s);
        Assert.Contains("outer", after);
        Assert.Contains("inner two", after);

        Assert.True(s.Undo());
        Assert.Equal(before, BodyText(s));
    }

    /// <summary>A failure inside a nested batch reverses the whole tree, not just the inner half.</summary>
    [Fact]
    public void DS43D_NestedBatchFailure_ReversesTheOuterBatchToo()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);
        var before = BodyText(s);

        var result = s.Batch(new Func<EditResult>[]
        {
            () => s.ReplaceText(paragraphs[0], "outer edit"),
            () =>
            {
                var inner = s.Batch(new Func<EditResult>[]
                {
                    () => s.ReplaceText(paragraphs[1], "inner edit"),
                    () => s.InsertFootnote(paragraphs[1], 0, NulPayload),   // throws mid-op
                });
                return inner.Success
                    ? new EditResult { Success = true }
                    : EditResult.Fail(inner.Error!.Code, inner.Error.Message);
            },
        });

        Assert.False(result.Success);
        Assert.True(result.RolledBack);
        Assert.Equal(before, BodyText(s));
        Assert.DoesNotContain("outer edit", BodyText(s));
        Assert.Equal(0, s.UndoCount);
    }

    /// <summary>A disposed session reports rather than throws, matching every other op.</summary>
    [Fact]
    public void DS43E_DisposedSession_ReturnsSessionDisposed()
    {
        var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        s.Dispose();

        var result = s.Batch(new Func<EditResult>[] { () => new EditResult { Success = true } });

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.SessionDisposed, result.Error!.Code);
    }

    // ─── The explicit scope form ─────────────────────────────────────────

    /// <summary>
    /// BeginBatch/EndBatch is the form a dispatcher uses when its steps cannot be expressed as
    /// <see cref="EditResult"/>-returning delegates — it runs its own switch and reports its own
    /// JSON. Committing keeps the edits as one undo step, exactly as <see cref="DocxSession.Batch"/> does.
    /// </summary>
    [Fact]
    public void DS440_BeginEndBatch_CommitsAsOneUndoStep()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);
        var before = BodyText(s);

        Assert.True(s.BeginBatch());
        Assert.True(s.ReplaceText(paragraphs[0], "scoped one").Success);
        Assert.True(s.ReplaceText(paragraphs[1], "scoped two").Success);
        var close = s.EndBatch(commit: true);

        Assert.True(close.Success);
        Assert.False(close.RolledBack);
        Assert.Equal(1, s.UndoCount);
        Assert.Contains("scoped two", BodyText(s));

        Assert.True(s.Undo());
        Assert.Equal(before, BodyText(s));
    }

    /// <summary>Closing without committing reverses everything the scope applied.</summary>
    [Fact]
    public void DS441_BeginEndBatch_RollsBackWhenNotCommitted()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);
        var before = BodyText(s);

        Assert.True(s.BeginBatch());
        Assert.True(s.ReplaceText(paragraphs[0], "discarded").Success);
        var close = s.EndBatch(commit: false);

        Assert.False(close.Success);
        Assert.True(close.RolledBack);
        Assert.Equal(before, BodyText(s));
        Assert.Equal(0, s.UndoCount);
    }

    /// <summary>
    /// A commit is not unconditional: a step that mutated before failing reverses the scope even
    /// when the caller asked to keep it. This is the property that stops the scope form from being
    /// a way to opt out of the half-applied-mutation guarantee.
    /// </summary>
    [Fact]
    public void DS442_BeginEndBatch_CommitIsRefusedWhenAStepMutatedBeforeFailing()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);
        var before = BodyText(s);

        Assert.True(s.BeginBatch());
        Assert.True(s.ReplaceText(paragraphs[0], "applied").Success);
        Assert.False(s.InsertFootnote(paragraphs[1], 0, NulPayload).Success);   // throws mid-op
        var close = s.EndBatch(commit: true);                                   // asks to keep it anyway

        Assert.False(close.Success);
        Assert.True(close.RolledBack);
        Assert.Equal(before, BodyText(s));
        Assert.Equal(0, s.UndoCount);
    }

    /// <summary>An unpaired close reports rather than corrupting the ring — it must never pop an
    /// entry that belongs to an ordinary edit.</summary>
    [Fact]
    public void DS443_EndBatch_WithNoBatchOpen_ReportsAndLeavesTheRingAlone()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(s);

        Assert.True(s.ReplaceText(paragraphs[0], "a real edit").Success);
        Assert.Equal(1, s.UndoCount);

        var close = s.EndBatch(commit: false);

        Assert.False(close.Success);
        Assert.False(close.RolledBack);
        Assert.Equal(1, s.UndoCount);
        Assert.Contains("a real edit", BodyText(s));
    }

    // ─── The cost this exists to remove ──────────────────────────────────

    /// <summary>
    /// The reason batching is a feature and not a convenience: a batch takes ONE snapshot, so its
    /// retained undo memory is flat in the number of steps, where the loop's grows with it. Asserted
    /// as a ratio rather than an absolute so it pins the shape of the cost, not a machine's numbers.
    /// </summary>
    [Fact]
    public void DS43F_Batch_RetainedUndoMemoryIsFlatInTheNumberOfSteps()
    {
        var paragraphTexts = Enumerable.Range(0, 12).Select(i => $"edit {i}").ToArray();

        long LoopBytes()
        {
            using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
            var anchor = BodyParagraphs(s)[0];
            foreach (var text in paragraphTexts) Assert.True(s.ReplaceText(anchor, text).Success);
            return s.UndoMemoryBytes;
        }

        long BatchBytes()
        {
            using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
            var anchor = BodyParagraphs(s)[0];
            var steps = paragraphTexts.Select<string, Func<EditResult>>(
                text => () => s.ReplaceText(anchor, text)).ToArray();
            Assert.True(s.Batch(steps).Success);
            return s.UndoMemoryBytes;
        }

        var loop = LoopBytes();
        var batch = BatchBytes();

        // Twelve steps, one snapshot instead of twelve: comfortably under a quarter of the loop's
        // retained cost even allowing for per-entry overhead.
        Assert.True(batch * 4 < loop, $"batch retained {batch} bytes vs loop {loop} for 12 edits");
    }
}
