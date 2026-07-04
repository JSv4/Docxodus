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
    /// <para><b>Mapped at render time (M2.4 Task 2).</b></para>
    /// <list type="bullet">
    /// <item><c>DetectMoves</c> → <see cref="IrDiffSettings.RenderMoves"/>. The engine ALWAYS aligns a
    /// relocated block as a move; <see cref="IrDiffSettings.RenderMoves"/> only controls whether the renderer
    /// PROJECTS it as a <c>Moved</c> pair or as a plain Inserted+Deleted pair. When <c>DetectMoves=false</c>
    /// we set <see cref="IrDiffSettings.RenderMoves"/> false, so EVERY aligned move — whether caught by the
    /// aligner's off-spine anchoring OR the fuzzy similarity pass — renders as Inserted+Deleted. This is the
    /// faithful render-time analogue of the comparer's switch (the earlier threshold-pushing approach could
    /// only gate the fuzzy pass and left exact relocations rendering as Moved; that gap is now closed).</item>
    /// <item><c>RevisionGranularity</c> → <see cref="Docxodus.Ir.Diff.RevisionGranularity.WmlComparerCompatible"/>,
    /// unconditionally — the adapter exists to be count/text-comparable to <c>GetRevisions</c>, so it always
    /// renders in the comparer's coarser contiguous-region grain.</item>
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
        return new IrDiffSettings
        {
            AuthorForRevisions = settings.AuthorForRevisions,
            CaseInsensitive = settings.CaseInsensitive,
            Culture = settings.CultureInfo,
            ConflateBreakingAndNonbreakingSpaces = settings.ConflateBreakingAndNonbreakingSpaces,
            MoveSimilarityThreshold = settings.MoveSimilarityThreshold,
            MoveMinimumTokenCount = settings.MoveMinimumWordCount,

            // M2.4 Task 2 — render-time WmlComparer parity. The adapter targets the shipped comparer's
            // GetRevisions surface, so it renders in WmlComparer-compatible granularity (contiguous-region
            // coalescing + common-affix trim + zero-width prune) rather than the engine's native fine grain.
            RevisionGranularity = RevisionGranularity.WmlComparerCompatible,

            // DetectMoves is a RENDER-TIME relabel here (M2.4 Task 2): when off, the renderer projects an
            // aligned move's two halves as a plain Inserted+Deleted pair instead of a Moved pair. The engine
            // alignment is unchanged — this is purely how the move is reported — so it switches off move
            // SEMANTICS regardless of how the move arose (aligner off-spine anchoring OR fuzzy similarity),
            // which the threshold-pushing approach could not do (it only gated the fuzzy pass). The move
            // detection thresholds still map 1:1 above so a move that IS rendered respects the caller's
            // similarity/min-word tuning.
            RenderMoves = settings.DetectMoves,
        };
    }
}
