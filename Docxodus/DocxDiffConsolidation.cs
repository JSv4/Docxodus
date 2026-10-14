#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Docxodus.Internal;
using Docxodus.Ir;
using Docxodus.Ir.Diff;

namespace Docxodus;

/// <summary>
/// A memoized N-way consolidation of one base document and its reviewers — the multi-reviewer twin
/// of <see cref="DocxDiffComparison"/>. The expensive work (pre-accept, the reads, the merge) runs
/// at most once and every product is served from it.
/// </summary>
/// <remarks>
/// <para><b>Why it exists (issue #617).</b> The four N-way statics each read the full reviewer set
/// independently, so the natural workflow — inspect the conflicts, choose a resolution policy, then
/// consolidate — read every document twice. An <c>N</c>-reviewer set is <c>N+1</c> documents, so
/// that is <c>2(N+1)</c> reads to produce an answer that needs <c>N+1</c>. Create one of these and
/// the set is read once however many products are asked for.</para>
/// <example>
/// <code>
/// var consolidation = DocxDiff.CreateConsolidation(baseDocument, reviewers);
/// foreach (var conflict in consolidation.GetConflicts()) { /* review */ }
/// var merged = consolidation.Consolidate();      // no second read
/// </code>
/// </example>
/// <para><b>Products are the statics' answers.</b> Each product is exactly what the corresponding
/// <see cref="DocxDiff"/> static returns for the same inputs and settings — the statics delegate to
/// a single-use instance of this class, so there is one implementation rather than two that have to
/// be kept in step.</para>
/// <para><b>Compatibility pre-flight.</b> Fires on the first product call, like
/// <see cref="DocxDiffComparison"/>. Previously only <see cref="Consolidate"/> ran it, so a caller
/// who set <c>ThrowOnCompatibilityWarning</c> and asked for conflicts, consolidated revisions or the
/// consolidated edit script was silently never told — the same shape of gap #622 closed on the
/// pairwise side.</para>
/// <para><b>Memory.</b> Reads retain provenance, because the markup renderer clones source elements
/// out of them rather than re-reading; that pins the parsed XML of every document in the set for the
/// instance's lifetime. Scope it to the run that needs the products.</para>
/// <para><b>Thread-safety.</b> All memoization is
/// <see cref="LazyThreadSafetyMode.ExecutionAndPublication"/>.</para>
/// </remarks>
public sealed class DocxDiffConsolidation
{
    private readonly DocxDiffConsolidateSettings _settings;
    private readonly IrDiffSettings _diff;
    private readonly bool _empty;

    private readonly Lazy<(WmlDocument Base, IReadOnlyList<DocxDiffReviewer> Reviewers)> _preflighted;
    private readonly Lazy<(IrDocument BaseIr, List<(string Author, IrDocument Ir)> ReviewerIrs)> _ir;
    private readonly Lazy<IrCompositeScript> _script;
    private readonly Lazy<WmlDocument> _redline;
    private readonly Lazy<IReadOnlyList<DocxDiffConflict>> _conflicts;
    private readonly Lazy<IReadOnlyList<DocxDiffConsolidatedRevision>> _revisions;
    private readonly Lazy<string> _editScriptJson;

    internal DocxDiffConsolidation(
        WmlDocument baseDocument,
        IReadOnlyList<DocxDiffReviewer> reviewers,
        DocxDiffConsolidateSettings settings)
    {
        _settings = settings;
        _diff = settings.Diff.ToIrDiffSettings();
        _empty = reviewers.Count == 0;
        var originalBase = baseDocument;

        _preflighted = new(() =>
        {
            // The opt-in accept-all pre-flatten (default off → no-op), then the pre-flight over the
            // whole set, exactly as Consolidate performed them.
            var preBase = DocxDiff.PreAccept(_settings.Diff, originalBase);
            var preReviewers = reviewers
                .Select(r => new DocxDiffReviewer
                {
                    Author = r.Author,
                    Document = DocxDiff.PreAccept(_settings.Diff, r.Document),
                })
                .ToList();
            DocxDiff.PreflightCompatibility(
                _settings.Diff,
                new[] { preBase }.Concat(preReviewers.Select(r => r.Document)).ToArray());
            return (preBase, (IReadOnlyList<DocxDiffReviewer>)preReviewers);
        }, LazyThreadSafetyMode.ExecutionAndPublication);

        _ir = new(() =>
        {
            var (preBase, preReviewers) = _preflighted.Value;
            return DocxDiff.ReadReviewerSet(preBase, preReviewers, DocxDiff.RenderReadOpts);
        }, LazyThreadSafetyMode.ExecutionAndPublication);

        _script = new(() =>
        {
            var (baseIr, revIr) = _ir.Value;
            return IrCompositeMerger.Merge(baseIr, revIr, _settings.ConflictResolution, _diff);
        }, LazyThreadSafetyMode.ExecutionAndPublication);

        _redline = new(() =>
        {
            if (_empty) return originalBase;
            var (preBase, preReviewers) = _preflighted.Value;
            var (baseIr, revIr) = _ir.Value;
            return IrCompositeMarkupRenderer.Render(
                _script.Value, preBase, preReviewers.Select(r => (r.Author, r.Document)).ToList(), _diff,
                baseIr, revIr.Select(x => x.Ir).ToList());
        }, LazyThreadSafetyMode.ExecutionAndPublication);

        _conflicts = new(() => _empty
            ? Array.Empty<DocxDiffConflict>()
            : _script.Value.Conflicts.Select(DocxDiffConflict.FromIr).ToList(),
            LazyThreadSafetyMode.ExecutionAndPublication);

        _revisions = new(() =>
        {
            if (_empty) return Array.Empty<DocxDiffConsolidatedRevision>();
            var (baseIr, revIr) = _ir.Value;
            return DocxDiff.ProjectConsolidatedRevisions(
                IrCompositeRevisionRenderer.Render(_script.Value, baseIr, revIr, _diff));
        }, LazyThreadSafetyMode.ExecutionAndPublication);

        _editScriptJson = new(() => IrCompositeScriptJson.Write(_empty
            ? new IrCompositeScript(
                IrNodeList.From(Array.Empty<IrCompositeOp>()),
                IrNodeList.From(Array.Empty<IrConflict>()))
            : _script.Value),
            LazyThreadSafetyMode.ExecutionAndPublication);
    }

    /// <summary>The consolidated document — what <see cref="DocxDiff.Consolidate"/> returns for the
    /// same inputs and settings, memoized.</summary>
    public WmlDocument Consolidate() => _redline.Value;

    /// <summary>The conflicts — what <see cref="DocxDiff.GetConflicts"/> returns for the same inputs
    /// and settings, memoized.</summary>
    public IReadOnlyList<DocxDiffConflict> GetConflicts() => _conflicts.Value;

    /// <summary>The attributed revision list — what
    /// <see cref="DocxDiff.GetConsolidatedRevisions"/> returns, memoized.</summary>
    public IReadOnlyList<DocxDiffConsolidatedRevision> GetConsolidatedRevisions() => _revisions.Value;

    /// <summary>The composite edit script as JSON — what
    /// <see cref="DocxDiff.GetConsolidatedEditScriptJson"/> returns, memoized.</summary>
    public string GetConsolidatedEditScriptJson() => _editScriptJson.Value;
}
