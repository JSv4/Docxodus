#nullable enable

using System;
using System.Collections.Generic;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.3 Task 4 — the test-side parity adapter. Exposes the IR diff pipeline through a
/// <see cref="WmlComparer"/>-shaped <c>GetRevisions</c> surface so the RUNNABLE_NOW rows of the
/// WmlComparer test suite (those that assert on <c>GetRevisions</c> counts / types / texts / move
/// semantics, NOT on produced OOXML markup) can be re-expressed against the new engine and scored.
///
/// <para><b>Pipeline.</b> <see cref="IrReader.Read"/> ×2 (the same <see cref="WcCorpus.ReadOpts"/> the
/// differential harness uses: <c>RetainSources=false</c>, <c>RevisionView=Accept</c>) →
/// <see cref="IrEditScriptBuilder.Build"/> → <see cref="IrRevisionRenderer.Render"/>. The result is a
/// flat <c>List&lt;IrRevision&gt;</c> — the IR analogue of
/// <c>List&lt;WmlComparer.WmlComparerRevision&gt;</c>.</para>
///
/// <para><b>Why an adapter and not the comparer.</b> The original tests call
/// <c>WmlComparer.Compare(left,right,settings)</c> to PRODUCE a tracked-revisions document and then
/// <c>WmlComparer.GetRevisions(compared,settings)</c> to read it back. The IR engine produces no OOXML
/// document yet (that is M2.4); its revisions surface comes straight off the edit script. So the adapter
/// skips the produce-then-reparse round-trip and renders revisions directly — semantically the same
/// <c>GetRevisions</c> contract, structurally a shortcut. Tests whose assertions ride on the produced
/// document (validation, accept/reject, native markup elements) are NOT adaptable here — they are the
/// MARKUP_BLOCKED rows of the scoreboard.</para>
/// </summary>
internal static class IrWmlComparerAdapter
{
    /// <summary>
    /// The <see cref="WmlComparer"/>-shaped entry point: run the IR pipeline over two in-memory documents
    /// under settings mapped from <see cref="WmlComparerSettings"/>, returning the rendered revisions.
    /// </summary>
    public static List<IrRevision> GetRevisions(WmlDocument left, WmlDocument right, WmlComparerSettings settings)
    {
        var diff = MapSettings(settings);
        var irLeft = IrReader.Read(left, WcCorpus.ReadOpts);
        var irRight = IrReader.Read(right, WcCorpus.ReadOpts);
        var script = IrEditScriptBuilder.Build(irLeft, irRight, diff);
        var revisions = IrRevisionRenderer.Render(script, irLeft, irRight, diff);
        return new List<IrRevision>(revisions);
    }

    /// <summary>
    /// Map the consumer-relevant subset of <see cref="WmlComparerSettings"/> onto <see cref="IrDiffSettings"/>.
    /// Every field that has a faithful IR analogue is carried; the rest are documented unmappable.
    ///
    /// <para><b>Mapped 1:1.</b></para>
    /// <list type="bullet">
    /// <item><c>AuthorForRevisions</c> → <see cref="IrDiffSettings.AuthorForRevisions"/> — same default
    /// (<c>"Open-Xml-PowerTools"</c>), so the adapter is author-comparable out of the box.</item>
    /// <item><c>CaseInsensitive</c> → <see cref="IrDiffSettings.CaseInsensitive"/>; <c>CultureInfo</c> →
    /// <see cref="IrDiffSettings.Culture"/> (null ⇒ ordinal/invariant folding).</item>
    /// <item><c>ConflateBreakingAndNonbreakingSpaces</c> →
    /// <see cref="IrDiffSettings.ConflateBreakingAndNonbreakingSpaces"/>.</item>
    /// <item><c>MoveSimilarityThreshold</c> → <see cref="IrDiffSettings.MoveSimilarityThreshold"/>;
    /// <c>MoveMinimumWordCount</c> → <see cref="IrDiffSettings.MoveMinimumTokenCount"/> (the IR engine
    /// counts Word-kind tokens; the comparer counts words — the same quantity for these fixtures).</item>
    /// </list>
    ///
    /// <para><b>Mapped with a caveat.</b></para>
    /// <list type="bullet">
    /// <item><c>DetectMoves</c> → there is no IR "enable moves" boolean: cross-gap fuzzy moves are gated by
    /// <see cref="IrDiffSettings.MoveSimilarityThreshold"/> / <see cref="IrDiffSettings.MoveMinimumTokenCount"/>.
    /// When <c>DetectMoves=false</c> we push the threshold ABOVE 1.0 (<see cref="DisableMovesThreshold"/>),
    /// which no Jaccard score can meet, switching off the FUZZY move pass. <b>It does NOT switch off
    /// exact-content relocations caught by the aligner's off-spine anchoring</b> (a structural property of
    /// the IR aligner, not a tunable), so an exact paragraph swap can still render as <c>Moved</c> under
    /// <c>DetectMoves=false</c>. This is a documented engine difference the scoreboard measures, not a
    /// mapping bug — see <c>MoveDetection_Disabled_ShouldNotDetectMoves</c>.</item>
    /// </list>
    ///
    /// <para><b>Unmappable (no IR analogue — left at IR defaults).</b></para>
    /// <list type="bullet">
    /// <item><c>DetailThreshold</c> (default 0.15) — the comparer's whole-document LCS detail knob. The IR
    /// engine has no global LCS detail parameter; granularity is governed by per-block tokenization and the
    /// <see cref="IrDiffSettings.BlockSimilarityThreshold"/> in-gap pairing floor. No faithful mapping;
    /// ignored. (A divergence source where the comparer's atomization is detail-tuned.)</item>
    /// <item><c>DetectFormatChanges</c> (default true) — the IR engine ALWAYS computes modeled format
    /// deltas (FormatChanged token spans / FormatOnly blocks) under
    /// <see cref="IrDiffSettings.FormatComparison"/>; there is no off switch. For <c>DetectFormatChanges=true</c>
    /// (the suite's format tests) this matches; a hypothetical <c>false</c> case has no IR analogue and is
    /// not exercised by the runnable rows.</item>
    /// <item><c>SimplifyMoveMarkup</c> — a PRODUCED-MARKUP transform (rewrite moveFrom/moveTo as del/ins in
    /// the output document). The revisions surface is pre-markup, so this is inherently MARKUP_BLOCKED; no
    /// IR analogue and never reached by an adapter row.</item>
    /// <item><c>DateTimeForRevisions</c> — the IR engine pins a deterministic epoch by default
    /// (<see cref="IrDiffSettings.DeterministicEpoch"/>); no runnable row asserts on revision dates, so the
    /// wall-clock default is deliberately NOT propagated (keeping adapter output reproducible).</item>
    /// </list>
    /// </summary>
    public static IrDiffSettings MapSettings(WmlComparerSettings settings)
    {
        double moveThreshold = settings.DetectMoves
            ? settings.MoveSimilarityThreshold
            : DisableMovesThreshold;

        return new IrDiffSettings
        {
            AuthorForRevisions = settings.AuthorForRevisions,
            CaseInsensitive = settings.CaseInsensitive,
            Culture = settings.CultureInfo,
            ConflateBreakingAndNonbreakingSpaces = settings.ConflateBreakingAndNonbreakingSpaces,
            MoveSimilarityThreshold = moveThreshold,
            MoveMinimumTokenCount = settings.MoveMinimumWordCount,
        };
    }

    /// <summary>
    /// The fuzzy-move threshold used to switch the cross-gap move pass OFF for <c>DetectMoves=false</c>:
    /// a value strictly greater than 1.0, which no Jaccard similarity (∈ [0,1]) can satisfy. Exact-content
    /// relocations caught by aligner anchoring are unaffected (see <see cref="MapSettings"/>).
    /// </summary>
    public const double DisableMovesThreshold = 2.0;
}
