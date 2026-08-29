#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Docxodus.Verification;

namespace Docxodus;

/// <summary>
/// A memoized comparison of one document pair (issue #594): the expensive alignment work —
/// input normalization, the IR reads, and the edit-script build — runs at most once, and every
/// data product (<see cref="ToRedline"/>, <see cref="GetRevisions"/>,
/// <see cref="GetEditScriptJson"/>, <see cref="GetSemanticChanges"/>) is served from that single
/// pass. Created by <see cref="DocxDiff.CreateComparison"/>; the stateless
/// <see cref="DocxDiff"/> statics delegate to a single-use instance, so each product here is
/// exactly what the corresponding static returns for the same inputs and settings.
/// </summary>
/// <remarks>
/// <para><b>Snapshot semantics.</b> The comparison captures the two <see cref="WmlDocument"/>
/// instances at creation; every product observes that one consistent snapshot, unlike three
/// sequential stateless calls, which could in principle observe different bytes if the caller
/// swaps inputs in between.</para>
/// <para><b>Laziness.</b> Nothing is computed at creation. Each product materializes on first
/// request and is cached — repeated calls return the same instance. Compatibility preflight
/// (<see cref="DocxDiffSettings.OnCompatibilityWarning"/>/
/// <see cref="DocxDiffSettings.ThrowOnCompatibilityWarning"/>) therefore fires on the first
/// product call, exactly as it fires inside each stateless static.</para>
/// <para><b>Semantic products.</b> With no explicit options, the comparison's own
/// <see cref="DocxDiffSettings"/> flow into the semantic pass (as
/// <see cref="SemanticDiffOptions.DiffSettings"/>) and the result is memoized. Passing explicit
/// <see cref="SemanticDiffOptions"/> bypasses the memo and runs the semantic pipeline with those
/// options verbatim. The semantic pipeline keeps its own package-preflighted read (its reader
/// retains sources, the diff read does not), so it shares the snapshot but not the IR pass.</para>
/// <para><b>Memory.</b> The instance retains the input packages and every materialized product
/// for its lifetime; scope it to the pipeline run that needs the products.</para>
/// <para><b>Thread-safety.</b> All memoization is <see cref="LazyThreadSafetyMode.ExecutionAndPublication"/>;
/// concurrent product calls are safe.</para>
/// </remarks>
public sealed class DocxDiffComparison
{
    private readonly WmlDocument _originalLeft;
    private readonly WmlDocument _originalRight;
    private readonly DocxDiffSettings _settings;

    private readonly Lazy<(WmlDocument Left, WmlDocument Right)> _preflighted;
    private readonly Lazy<(IrDocument Left, IrDocument Right)> _ir;
    private readonly Lazy<IrEditScript> _dataScript;
    private readonly Lazy<WmlDocument> _redline;
    private readonly Lazy<IReadOnlyList<DocxDiffRevision>> _revisions;
    private readonly Lazy<string> _editScriptJson;
    private readonly Lazy<SemanticChangeSet> _semanticChanges;
    private readonly Lazy<string> _semanticChangesJson;

    internal DocxDiffComparison(WmlDocument left, WmlDocument right, DocxDiffSettings? settings)
    {
        _originalLeft = left;
        _originalRight = right;
        _settings = settings ?? new DocxDiffSettings();

        // PreAccept re-reads and can rewrite the whole package (strict-namespace normalization,
        // mc:AlternateContent resolution, optional accept-flatten). The two sides never look at each
        // other, so they run concurrently; the preflight that DOES span both runs after the join.
        _preflighted = new(() =>
        {
            var leftTask = Task.Run(() => DocxDiff.PreAccept(_settings, _originalLeft));
            var preRight = DocxDiff.PreAccept(_settings, _originalRight);
            var preLeft = leftTask.GetAwaiter().GetResult();
            DocxDiff.PreflightCompatibility(_settings, preLeft, preRight);
            return (preLeft, preRight);
        }, LazyThreadSafetyMode.ExecutionAndPublication);

        // ONE read per side serves every product, including the markup renderer's clone-from-provenance
        // pass — hence DocxDiff.RenderReadOpts (RetainSources ON) rather than DocxDiff.ReadOpts. Provenance
        // is equality-neutral (see IrProvenance), so this snapshot is node-for-node value-equal to the
        // retention-off one and the edit script built over it is unchanged; what it buys is the renderer's
        // two full re-reads of these same documents, which on a heavyweight document dominate the compare.
        // The two sides are independent pure reads, so they run concurrently.
        _ir = new(() =>
        {
            var (preLeft, preRight) = _preflighted.Value;
            var leftTask = Task.Run(() => IrReader.Read(preLeft, DocxDiff.RenderReadOpts));
            var irRight = IrReader.Read(preRight, DocxDiff.RenderReadOpts);
            return (leftTask.GetAwaiter().GetResult(), irRight);
        }, LazyThreadSafetyMode.ExecutionAndPublication);

        // The DATA script: CrossParagraphTokenDiff forced off, exactly as GetRevisions and
        // GetEditScriptJson force it (it is a markup-only refinement; see those statics).
        _dataScript = new(() =>
        {
            var diff = _settings.ToIrDiffSettings() with { CrossParagraphTokenDiff = false };
            var (irLeft, irRight) = _ir.Value;
            return IrEditScriptBuilder.Build(irLeft, irRight, diff);
        }, LazyThreadSafetyMode.ExecutionAndPublication);

        _redline = new(BuildRedline, LazyThreadSafetyMode.ExecutionAndPublication);

        _revisions = new(() =>
        {
            // Byte-identical packages have nothing to report, and proving that by running the whole
            // pipeline costs as much as a real comparison. This is the same guard ToRedline uses, for
            // the same reason and with the same exception: an explicit accept-all request rewrites the
            // input even when the two sides match, so it is not a no-op. Note the edit script has no
            // equivalent shortcut — its all-Equal operations are the answer callers asked for.
            if (DocxCompare.HasIdenticalPackageBytes(_originalLeft, _originalRight) &&
                !(_settings.PreAcceptInputRevisions && !_settings.PreserveInputRevisions))
                return Array.Empty<DocxDiffRevision>();

            var diff = _settings.ToIrDiffSettings() with { CrossParagraphTokenDiff = false };
            var (irLeft, irRight) = _ir.Value;
            return IrRevisionRenderer.Render(_dataScript.Value, irLeft, irRight, diff)
                .Select(DocxDiffRevision.FromIr).ToList();
        }, LazyThreadSafetyMode.ExecutionAndPublication);

        _editScriptJson = new(
            () => IrEditScriptJson.Write(_dataScript.Value),
            LazyThreadSafetyMode.ExecutionAndPublication);

        _semanticChanges = new(
            () => SemanticDiff.Compare(_originalLeft, _originalRight, DefaultSemanticOptions()),
            LazyThreadSafetyMode.ExecutionAndPublication);

        _semanticChangesJson = new(
            () => _semanticChanges.Value.ToJson(indented: true),
            LazyThreadSafetyMode.ExecutionAndPublication);
    }

    /// <summary>
    /// The native tracked-changes redline — what <see cref="DocxDiff.Compare"/> returns for the
    /// same inputs and settings, memoized. Repeated calls return the same instance.
    /// </summary>
    public WmlDocument ToRedline() => _redline.Value;

    /// <summary>
    /// The consumer revision list — what <see cref="DocxDiff.GetRevisions"/> returns for the same
    /// inputs and settings, memoized. Repeated calls return the same list instance.
    /// </summary>
    public IReadOnlyList<DocxDiffRevision> GetRevisions() => _revisions.Value;

    /// <summary>
    /// The edit script as JSON — what <see cref="DocxDiff.GetEditScriptJson"/> returns for the
    /// same inputs and settings, memoized.
    /// </summary>
    public string GetEditScriptJson() => _editScriptJson.Value;

    /// <summary>
    /// The semantic change set. With <paramref name="options"/> null, the comparison's own
    /// settings flow into the pass and the result is memoized; explicit options run the semantic
    /// pipeline with those options verbatim (unmemoized).
    /// </summary>
    public SemanticChangeSet GetSemanticChanges(SemanticDiffOptions? options = null) =>
        options is null
            ? _semanticChanges.Value
            : SemanticDiff.Compare(_originalLeft, _originalRight, options);

    /// <summary>
    /// JSON counterpart of <see cref="GetSemanticChanges"/>. The default (no options, indented)
    /// is memoized; any other shape serializes from the corresponding change set.
    /// </summary>
    public string GetSemanticChangesJson(SemanticDiffOptions? options = null, bool indented = true) =>
        options is null
            ? indented ? _semanticChangesJson.Value : _semanticChanges.Value.ToJson(indented: false)
            : SemanticDiff.Compare(_originalLeft, _originalRight, options).ToJson(indented);

    private SemanticDiffOptions DefaultSemanticOptions() => new() { DiffSettings = _settings };

    /// <summary>Replicates <see cref="DocxDiff.Compare"/>'s exact structure, product-shared:
    /// both identical-bytes fast paths (including the gated preflight on the first), then the
    /// markup render — reusing the memoized data script when the fused and data settings agree.</summary>
    private WmlDocument BuildRedline()
    {
        var s = _settings;
        // Exact identity is a no-op even for Strict OOXML or revision-bearing packages. An explicit
        // accept-all request is an exception — see DocxDiff.Compare, which this mirrors verbatim.
        if (DocxCompare.HasIdenticalPackageBytes(_originalLeft, _originalRight) &&
            !(s.PreAcceptInputRevisions && !s.PreserveInputRevisions))
        {
            if (s.OnCompatibilityWarning != null || s.ThrowOnCompatibilityWarning)
            {
                var preflightLeft = DocxDiff.PreAccept(s, _originalLeft);
                var preflightRight = DocxDiff.PreAccept(s, _originalRight);
                DocxDiff.PreflightCompatibility(s, preflightLeft, preflightRight);
            }

            return new WmlDocument(_originalLeft);
        }

        var (left, right) = _preflighted.Value;
        if (DocxCompare.HasIdenticalPackageBytes(left, right))
            return new WmlDocument(left);

        var diff = s.ToIrDiffSettings();
        var script = diff.CrossParagraphTokenDiff
            ? BuildFusedScript(diff)
            : _dataScript.Value;
        var (irLeft, irRight) = _ir.Value;
        return IrMarkupRenderer.Render(script, left, right, diff, irLeft, irRight);
    }

    private IrEditScript BuildFusedScript(IrDiffSettings diff)
    {
        var (irLeft, irRight) = _ir.Value;
        return IrEditScriptBuilder.Build(irLeft, irRight, diff);
    }
}
