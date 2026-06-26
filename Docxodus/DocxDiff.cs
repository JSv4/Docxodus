#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;

namespace Docxodus;

/// <summary>
/// The IR diff engine's public facade — a structure-aware DOCX comparison engine that produces native
/// tracked-changes markup, a consumer revision list, AND the edit script as data (JSON). This is the
/// public surface over the internal <c>Docxodus.Ir.Diff</c> pipeline
/// (<c>IrReader → IrEditScriptBuilder → IrMarkupRenderer / IrRevisionRenderer / IrEditScriptJson</c>).
/// </summary>
/// <remarks>
/// <para><b>Relationship to <see cref="WmlComparer"/>.</b> <see cref="WmlComparer"/> remains the
/// default, blessed comparison API. <see cref="DocxDiff"/> is the NEW engine: it is built on an
/// intermediate document representation (IR) with anchor-addressed blocks, so its revisions carry
/// stable anchors (<see cref="DocxDiffRevision.LeftAnchor"/>/<see cref="DocxDiffRevision.RightAnchor"/>)
/// and it can emit its edit script as data (<see cref="GetEditScriptJson"/>) — neither of which
/// <see cref="WmlComparer"/> offers. It is shipped as a <b>production-candidate</b> pending the Word
/// manual-verification checklist and burn-in; prefer <see cref="WmlComparer"/> for production redlines
/// until that swap is ratified (decision D4). See <c>docs/architecture/ir_diff_engine.md</c>.</para>
///
/// <para><b>Multi-author / consolidate-forward stance.</b> Nothing here is static or process-global:
/// the author stamped on revisions flows per call via <see cref="DocxDiffSettings.AuthorForRevisions"/>,
/// so a multi-author pipeline simply passes a different author per comparison. The facade compares two
/// documents per call; an N-way consolidate is a forward composition of pairwise diffs (each carrying
/// its own author) rather than a baked-in "exactly two documents" assumption — the types here make no
/// such assumption and are consolidate-compatible.</para>
///
/// <para><b>Determinism.</b> By default (<see cref="DocxDiffSettings.Deterministic"/> true) revision
/// dates are pinned to a fixed epoch, so two diffs of the same inputs are byte-identical. Set
/// <see cref="DocxDiffSettings.Deterministic"/> false (optionally with an explicit
/// <see cref="DocxDiffSettings.DateTimeForRevisions"/>) for wall-clock dates.</para>
///
/// <para><b>Thread-safety.</b> All methods are pure functions of their arguments with no shared mutable
/// state; concurrent calls on distinct inputs are safe.</para>
/// </remarks>
public static class DocxDiff
{
    // The diff engine reads with provenance OFF (it never needs element-level provenance and the lower
    // footprint matters for bulk pipelines) and revisions ACCEPTED (the IR the script is built over is
    // the accepted view of each side). The markup renderer re-reads internally with its own options.
    private static readonly IrReaderOptions ReadOpts =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };

    /// <summary>
    /// Compare <paramref name="left"/> and <paramref name="right"/> and produce a tracked-changes
    /// <see cref="WmlDocument"/> carrying native Word revision markup
    /// (<c>w:ins</c>/<c>w:del</c>/<c>w:moveFrom</c>/<c>w:moveTo</c>/<c>w:rPrChange</c>). The result
    /// satisfies the WmlComparer contract: <c>RevisionAccepter.AcceptRevisions(result)</c> content-equals
    /// <paramref name="right"/> and <c>RevisionProcessor.RejectRevisions(result)</c> content-equals
    /// <paramref name="left"/> at the per-block text level.
    /// </summary>
    /// <param name="left">The earlier/original document.</param>
    /// <param name="right">The later/revised document.</param>
    /// <param name="settings">Diff settings; <c>null</c> uses the defaults.</param>
    /// <returns>A new <see cref="WmlDocument"/> with tracked-changes markup; the inputs are unchanged.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="left"/> or <paramref name="right"/> is null.</exception>
    public static WmlDocument Compare(WmlDocument left, WmlDocument right, DocxDiffSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(right);
        var s = settings ?? new DocxDiffSettings();
        // Opt-in accept-all pre-flatten (default off → no-op): diff (and clone the output from) the accepted
        // view of both inputs so no pre-existing input revision survives into the result. See the flag's docs.
        left = PreAccept(s, left);
        right = PreAccept(s, right);
        PreflightCompatibility(s, left, right);
        var diff = s.ToIrDiffSettings();
        var irLeft = IrReader.Read(left, ReadOpts);
        var irRight = IrReader.Read(right, ReadOpts);
        var script = IrEditScriptBuilder.Build(irLeft, irRight, diff);
        // IrMarkupRenderer re-reads both packages with provenance (RetainSources=true) to clone source
        // block elements; it takes the WmlDocuments, not the IR snapshots above.
        return IrMarkupRenderer.Render(script, left, right, diff);
    }

    /// <summary>
    /// Compare <paramref name="left"/> and <paramref name="right"/> and return the consumer revision list
    /// (insertions, deletions, moves, format changes) rendered from the edit script — the diff-as-revisions
    /// view, analogous to <see cref="WmlComparer"/>'s <c>GetRevisions</c> but anchor-addressed and produced
    /// directly off the IR script (no produce-then-reparse round-trip).
    /// </summary>
    /// <param name="left">The earlier/original document.</param>
    /// <param name="right">The later/revised document.</param>
    /// <param name="settings">Diff settings; <c>null</c> uses the defaults.</param>
    /// <returns>The revisions in document order.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="left"/> or <paramref name="right"/> is null.</exception>
    public static IReadOnlyList<DocxDiffRevision> GetRevisions(
        WmlDocument left, WmlDocument right, DocxDiffSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(right);
        var s = settings ?? new DocxDiffSettings();
        left = PreAccept(s, left);
        right = PreAccept(s, right);
        PreflightCompatibility(s, left, right);
        var diff = s.ToIrDiffSettings();
        var irLeft = IrReader.Read(left, ReadOpts);
        var irRight = IrReader.Read(right, ReadOpts);
        var script = IrEditScriptBuilder.Build(irLeft, irRight, diff);
        var rendered = IrRevisionRenderer.Render(script, irLeft, irRight, diff);
        return rendered.Select(DocxDiffRevision.FromIr).ToList();
    }

    /// <summary>
    /// Compare <paramref name="left"/> and <paramref name="right"/> and return the engine's edit script as
    /// a JSON string — the <b>diff-as-data</b> differentiator. The script is the anchor-addressed list of
    /// block operations (equal/insert/delete/modify/move/format-only, with nested token diffs and
    /// row/cell-precise table diffs) that the markup and revision renderers both consume. Stable and
    /// machine-readable: suitable for storage, transport to non-.NET consumers, review tooling, and audit.
    /// </summary>
    /// <param name="left">The earlier/original document.</param>
    /// <param name="right">The later/revised document.</param>
    /// <param name="settings">Diff settings; <c>null</c> uses the defaults.</param>
    /// <returns>The edit script serialized as indented JSON.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="left"/> or <paramref name="right"/> is null.</exception>
    public static string GetEditScriptJson(
        WmlDocument left, WmlDocument right, DocxDiffSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(right);
        var s = settings ?? new DocxDiffSettings();
        left = PreAccept(s, left);
        right = PreAccept(s, right);
        PreflightCompatibility(s, left, right);
        var diff = s.ToIrDiffSettings();
        var irLeft = IrReader.Read(left, ReadOpts);
        var irRight = IrReader.Read(right, ReadOpts);
        var script = IrEditScriptBuilder.Build(irLeft, irRight, diff);
        return IrEditScriptJson.Write(script);
    }

    /// <summary>
    /// Consolidate the edits of N reviewers — each an independently revised copy of the SAME
    /// <paramref name="baseDocument"/> — into a single tracked-changes <see cref="WmlDocument"/> carrying
    /// native Word revision markup (<c>w:ins</c>/<c>w:del</c>/<c>w:moveFrom</c>/<c>w:moveTo</c>/<c>w:rPrChange</c>),
    /// with each reviewer's contribution attributed to that reviewer's own author name. This is the N-way,
    /// shared-base composition counterpart to the pairwise <see cref="Compare"/>: every reviewer is diffed
    /// against the common base and the resulting per-reviewer edit scripts are merged at the block- and
    /// token-span level into one document.
    /// </summary>
    /// <remarks>
    /// <para><b>Shared-base contract.</b> All reviewers MUST be derived from the same
    /// <paramref name="baseDocument"/>; the merge is anchored on the base's deterministic block anchors, and a
    /// reviewer document with no common ancestry will simply read as a wholesale rewrite. Reviewers are
    /// processed in list order, which is significant for tie-breaking under
    /// <see cref="ConflictResolution.FirstReviewerWins"/> and for the emission order under
    /// <see cref="ConflictResolution.StackAll"/>.</para>
    ///
    /// <para><b>Multi-author attribution.</b> Each reviewer's revisions are stamped with that reviewer's
    /// <see cref="DocxDiffReviewer.Author"/> (not the single
    /// <see cref="DocxDiffSettings.AuthorForRevisions"/>), so the consolidated document distinguishes who made
    /// each edit — directly consumable by Word's reviewing pane and by consolidate-forward pipelines.</para>
    ///
    /// <para><b>Round-trip contract.</b> The result satisfies the WmlComparer round-trip relative to the
    /// policy-resolved outcome: <c>RevisionProcessor.RejectRevisions(result)</c> content-equals
    /// <paramref name="baseDocument"/> (rejecting every reviewer's edits restores the base) and
    /// <c>RevisionAccepter.AcceptRevisions(result)</c> content-equals the document the chosen
    /// <see cref="DocxDiffConsolidateSettings.ConflictResolution"/> policy resolves to (e.g. base text retained
    /// at conflicted spans under <see cref="ConflictResolution.BaseWins"/>, the first reviewer's text under
    /// <see cref="ConflictResolution.FirstReviewerWins"/>).</para>
    ///
    /// <para><b>Conflicts.</b> A base span edited DIFFERENTLY by two or more reviewers is a conflict: it is
    /// resolved per the configured policy in the OUTPUT, and every competing edit is recorded — call
    /// <see cref="GetConflicts"/> on the same inputs/settings to retrieve the conflict list (id, base anchor,
    /// token span, applied policy, and the competing edits per reviewer).</para>
    ///
    /// <para><b>Determinism &amp; thread-safety.</b> With
    /// <see cref="DocxDiffConsolidateSettings.Diff"/>'s <see cref="DocxDiffSettings.Deterministic"/> true (the
    /// default), the consolidated output is byte-identical across runs of the same inputs. The method is a pure
    /// function of its arguments with no shared mutable state; concurrent calls on distinct inputs are safe.</para>
    ///
    /// <para><b>Zero reviewers.</b> An empty <paramref name="reviewers"/> list returns
    /// <paramref name="baseDocument"/> unchanged (no edits to consolidate).</para>
    /// </remarks>
    /// <param name="baseDocument">The shared base document every reviewer revised.</param>
    /// <param name="reviewers">The reviewers' revised copies and their author names, in priority order.</param>
    /// <param name="settings">Consolidate settings (diff settings + conflict policy); <c>null</c> uses the defaults.</param>
    /// <returns>
    /// A new <see cref="WmlDocument"/> with multi-author tracked-changes markup; the inputs are unchanged.
    /// When <paramref name="reviewers"/> is empty, <paramref name="baseDocument"/> is returned as-is.
    /// </returns>
    /// <exception cref="ArgumentNullException"><paramref name="baseDocument"/> or <paramref name="reviewers"/> is null.</exception>
    public static WmlDocument Consolidate(
        WmlDocument baseDocument, IReadOnlyList<DocxDiffReviewer> reviewers,
        DocxDiffConsolidateSettings? settings = null)
    {
        var s = ValidateConsolidateArgs(baseDocument, reviewers, settings);
        var diff = s.Diff.ToIrDiffSettings();
        if (reviewers.Count == 0) return baseDocument;
        // Opt-in accept-all pre-flatten (default off → no-op): consolidate the accepted view of the base and
        // every reviewer so no pre-existing input revision survives into the consolidated result.
        baseDocument = PreAccept(s.Diff, baseDocument);
        reviewers = PreAccept(s.Diff, reviewers);
        PreflightCompatibility(s.Diff, new[] { baseDocument }.Concat(reviewers.Select(r => r.Document)).ToArray());
        var baseIr = IrReader.Read(baseDocument, ReadOpts);
        var revIr = reviewers.Select(r => (r.Author, IrReader.Read(r.Document, ReadOpts))).ToList();
        var script = IrCompositeMerger.Merge(baseIr, revIr, s.ConflictResolution, diff);
        return IrCompositeMarkupRenderer.Render(
            script, baseDocument, reviewers.Select(r => (r.Author, r.Document)).ToList(), diff);
    }

    /// <summary>
    /// Return the conflicts that an N-way <see cref="Consolidate"/> of the same inputs would record — each a
    /// base span edited DIFFERENTLY by two or more reviewers — WITHOUT producing the consolidated document.
    /// This is the inspect-before-merge view: it runs the same shared-base merge as <see cref="Consolidate"/>
    /// and surfaces only the conflict list, so callers can review disagreements (and choose a
    /// <see cref="DocxDiffConsolidateSettings.ConflictResolution"/> policy) before committing to an output.
    /// </summary>
    /// <remarks>
    /// <para><b>Shared-base contract.</b> As with <see cref="Consolidate"/>, all reviewers MUST derive from the
    /// same <paramref name="baseDocument"/>, and reviewer LIST ORDER is significant (it determines competitor
    /// order within each conflict and policy tie-breaking).</para>
    ///
    /// <para><b>What a conflict carries.</b> Each <see cref="DocxDiffConflict"/> records the base anchor and
    /// token span where reviewers disagree, the <see cref="ConflictResolution"/> that WOULD be applied, and the
    /// per-reviewer competing edits (<see cref="DocxDiffConflict.Competitors"/>, each with the reviewer's
    /// <see cref="DocxDiffConflictCompetitor.Author"/> and the text that reviewer's edit would produce). A span
    /// edited by only one reviewer, or identically by several, is NOT a conflict and does not appear here.</para>
    ///
    /// <para><b>Consistency with <see cref="Consolidate"/>.</b> Given identical
    /// <paramref name="baseDocument"/>, <paramref name="reviewers"/>, and <paramref name="settings"/>, the
    /// conflicts returned here are exactly those recorded by <see cref="Consolidate"/>, with matching
    /// <see cref="DocxDiffConflict.Id"/> values — so the two calls can be correlated.</para>
    ///
    /// <para><b>Determinism &amp; thread-safety.</b> Deterministic for the same inputs and a pure function of
    /// its arguments with no shared mutable state; concurrent calls on distinct inputs are safe.</para>
    ///
    /// <para><b>Zero reviewers.</b> An empty <paramref name="reviewers"/> list yields an empty conflict list.</para>
    /// </remarks>
    /// <param name="baseDocument">The shared base document every reviewer revised.</param>
    /// <param name="reviewers">The reviewers' revised copies and their author names, in priority order.</param>
    /// <param name="settings">Consolidate settings (diff settings + conflict policy); <c>null</c> uses the defaults.</param>
    /// <returns>
    /// The conflicts in document order; empty when there are no conflicting spans (or no reviewers).
    /// </returns>
    /// <exception cref="ArgumentNullException"><paramref name="baseDocument"/> or <paramref name="reviewers"/> is null.</exception>
    public static IReadOnlyList<DocxDiffConflict> GetConflicts(
        WmlDocument baseDocument, IReadOnlyList<DocxDiffReviewer> reviewers,
        DocxDiffConsolidateSettings? settings = null)
    {
        var s = ValidateConsolidateArgs(baseDocument, reviewers, settings);
        var diff = s.Diff.ToIrDiffSettings();
        if (reviewers.Count == 0) return System.Array.Empty<DocxDiffConflict>();
        baseDocument = PreAccept(s.Diff, baseDocument);
        reviewers = PreAccept(s.Diff, reviewers);
        var baseIr = IrReader.Read(baseDocument, ReadOpts);
        var revIr = reviewers.Select(r => (r.Author, IrReader.Read(r.Document, ReadOpts))).ToList();
        var script = IrCompositeMerger.Merge(baseIr, revIr, s.ConflictResolution, diff);
        return script.Conflicts.Select(DocxDiffConflict.FromIr).ToList();
    }

    /// <summary>
    /// Consolidate N reviewers' edits of the same <paramref name="baseDocument"/> and return the ATTRIBUTED
    /// consumer revision list — the consolidate counterpart to <see cref="GetRevisions"/>. Each revision carries
    /// the contributing reviewer's <see cref="DocxDiffConsolidatedRevision.Author"/> (not a single document
    /// author) and, when it participates in a conflicted span, the
    /// <see cref="DocxDiffConsolidatedRevision.ConflictId"/> linking it to the matching
    /// <see cref="DocxDiffConflict"/> from <see cref="GetConflicts"/>.
    /// </summary>
    /// <remarks>
    /// <para><b>Attribution.</b> An insertion is attributed to the reviewer who inserted it; a deletion to the
    /// reviewer who deleted that base content; a format change to the reviewer who changed the formatting. When a
    /// single base paragraph was edited by several reviewers, each reviewer's word-level edits are surfaced as
    /// separate revisions under that reviewer's name (the multi-author generalization of
    /// <see cref="GetRevisions"/>).</para>
    ///
    /// <para><b>Consistency.</b> Given identical <paramref name="baseDocument"/>, <paramref name="reviewers"/>,
    /// and <paramref name="settings"/>, the revisions returned here describe exactly the edits that
    /// <see cref="Consolidate"/> places in the consolidated document, and conflict ids match those from
    /// <see cref="GetConflicts"/>. Deterministic for the same inputs; a pure function of its arguments.</para>
    ///
    /// <para><b>Zero reviewers.</b> An empty <paramref name="reviewers"/> list yields an empty revision list.</para>
    /// </remarks>
    /// <param name="baseDocument">The shared base document every reviewer revised.</param>
    /// <param name="reviewers">The reviewers' revised copies and their author names, in priority order.</param>
    /// <param name="settings">Consolidate settings (diff settings + conflict policy); <c>null</c> uses the defaults.</param>
    /// <returns>The attributed revisions in document order; empty when there are no reviewers.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="baseDocument"/> or <paramref name="reviewers"/> is null.</exception>
    public static IReadOnlyList<DocxDiffConsolidatedRevision> GetConsolidatedRevisions(
        WmlDocument baseDocument, IReadOnlyList<DocxDiffReviewer> reviewers,
        DocxDiffConsolidateSettings? settings = null)
    {
        var s = ValidateConsolidateArgs(baseDocument, reviewers, settings);
        var diff = s.Diff.ToIrDiffSettings();
        if (reviewers.Count == 0) return System.Array.Empty<DocxDiffConsolidatedRevision>();
        baseDocument = PreAccept(s.Diff, baseDocument);
        reviewers = PreAccept(s.Diff, reviewers);
        var baseIr = IrReader.Read(baseDocument, ReadOpts);
        var revIr = reviewers.Select(r => (r.Author, IrReader.Read(r.Document, ReadOpts))).ToList();
        var script = IrCompositeMerger.Merge(baseIr, revIr, s.ConflictResolution, diff);
        var rendered = IrCompositeRevisionRenderer.Render(script, baseIr, revIr, diff);
        return rendered.Select(x => new DocxDiffConsolidatedRevision(
            type: x.Rev.Type switch
            {
                IrRevisionType.Inserted => DocxDiffRevisionType.Inserted,
                IrRevisionType.Deleted => DocxDiffRevisionType.Deleted,
                IrRevisionType.Moved => DocxDiffRevisionType.Moved,
                IrRevisionType.FormatChanged => DocxDiffRevisionType.FormatChanged,
                _ => throw new ArgumentOutOfRangeException(nameof(reviewers), x.Rev.Type, "Unknown IrRevisionType."),
            },
            text: x.Rev.Text,
            author: x.Author,
            date: x.Rev.Date,
            moveGroupId: x.Rev.MoveGroupId,
            isMoveSource: x.Rev.IsMoveSource,
            formatChange: x.Rev.FormatChange is { } fc ? new DocxDiffFormatChange(fc) : null,
            leftAnchor: x.Rev.LeftAnchor,
            rightAnchor: x.Rev.RightAnchor,
            conflictId: x.ConflictId)).ToList();
    }

    /// <summary>
    /// Consolidate N reviewers' edits of the same <paramref name="baseDocument"/> and return the COMPOSITE edit
    /// script as a JSON string — the consolidate counterpart to <see cref="GetEditScriptJson"/>. It additively
    /// extends the two-way edit-script JSON: every operation carries an <c>author</c> and <c>sourceReviewer</c>,
    /// a <c>conflictId</c> when the op is a conflict winner, and (for a composed multi-reviewer paragraph)
    /// <c>authoredTokens</c> + <c>sourceRightAnchors</c>; the document additionally carries a top-level
    /// <c>conflicts</c> array (id, base anchor, token span, applied policy, and the per-reviewer competing edits).
    /// </summary>
    /// <remarks>
    /// <para><b>Diff-as-data.</b> This is the machine-readable, transport-friendly form of a consolidate:
    /// suitable for storage, audit, review tooling, and non-.NET consumers. It is stable and deterministic for
    /// the same inputs.</para>
    ///
    /// <para><b>Zero reviewers.</b> An empty <paramref name="reviewers"/> list yields a JSON document with empty
    /// <c>operations</c> and <c>conflicts</c> arrays.</para>
    /// </remarks>
    /// <param name="baseDocument">The shared base document every reviewer revised.</param>
    /// <param name="reviewers">The reviewers' revised copies and their author names, in priority order.</param>
    /// <param name="settings">Consolidate settings (diff settings + conflict policy); <c>null</c> uses the defaults.</param>
    /// <returns>The composite edit script serialized as indented JSON.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="baseDocument"/> or <paramref name="reviewers"/> is null.</exception>
    public static string GetConsolidatedEditScriptJson(
        WmlDocument baseDocument, IReadOnlyList<DocxDiffReviewer> reviewers,
        DocxDiffConsolidateSettings? settings = null)
    {
        var s = ValidateConsolidateArgs(baseDocument, reviewers, settings);
        var diff = s.Diff.ToIrDiffSettings();
        if (reviewers.Count == 0)
            return IrCompositeScriptJson.Write(new IrCompositeScript(
                IrNodeList.From(System.Array.Empty<IrCompositeOp>()),
                IrNodeList.From(System.Array.Empty<IrConflict>())));
        baseDocument = PreAccept(s.Diff, baseDocument);
        reviewers = PreAccept(s.Diff, reviewers);
        var baseIr = IrReader.Read(baseDocument, ReadOpts);
        var revIr = reviewers.Select(r => (r.Author, IrReader.Read(r.Document, ReadOpts))).ToList();
        var script = IrCompositeMerger.Merge(baseIr, revIr, s.ConflictResolution, diff);
        return IrCompositeScriptJson.Write(script);
    }

    /// <summary>
    /// Pre-flight a document against the DocxDiff feature catalog, returning a warning for each construct
    /// DocxDiff has not had a fidelity campaign for. Does NOT diff anything and does not change
    /// <see cref="Compare(WmlDocument, WmlDocument, DocxDiffSettings?)"/>; the caller decides whether to log,
    /// gate, or throw on <see cref="DocxDiffCompatibilityReport.HasWarnings"/>. See
    /// <see cref="DocxDiffCompatibility"/> for the catalog (the campaign roadmap).
    /// </summary>
    public static DocxDiffCompatibilityReport InspectCompatibility(WmlDocument doc) =>
        DocxDiffCompatibility.Inspect(doc);

    /// <inheritdoc cref="InspectCompatibility(WmlDocument)"/>
    public static DocxDiffCompatibilityReport InspectCompatibility(byte[] docxBytes) =>
        DocxDiffCompatibility.Inspect(docxBytes);

    // ----------------------------------------------------------------- compatibility pre-flight

    /// <summary>
    /// Opt-in pre-flight a diff's inputs against the compatibility catalog. A no-op (NO scan, zero cost) unless
    /// <see cref="DocxDiffSettings.OnCompatibilityWarning"/> or <see cref="DocxDiffSettings.ThrowOnCompatibilityWarning"/>
    /// is set. When engaged and a warning is found, invokes the callback then optionally throws
    /// <see cref="DocxDiffCompatibilityException"/>.
    /// </summary>
    private static void PreflightCompatibility(DocxDiffSettings settings, params WmlDocument[] inputs)
    {
        if (settings.OnCompatibilityWarning == null && !settings.ThrowOnCompatibilityWarning)
            return;
        var report = CombineCompatibility(inputs);
        if (!report.HasWarnings)
            return;
        settings.OnCompatibilityWarning?.Invoke(report);
        if (settings.ThrowOnCompatibilityWarning)
            throw new DocxDiffCompatibilityException(report);
    }

    /// <summary>Union the compatibility reports of several inputs: warnings keyed by feature id with occurrences
    /// summed across inputs, CoveredPresent unioned.</summary>
    private static DocxDiffCompatibilityReport CombineCompatibility(IReadOnlyList<WmlDocument> inputs)
    {
        var warnings = new Dictionary<string, DocxDiffCompatibilityFinding>();
        var covered = new Dictionary<string, DocxDiffFeature>();
        foreach (var input in inputs)
        {
            var r = DocxDiffCompatibility.Inspect(input);
            foreach (var w in r.Warnings)
                warnings[w.Feature.Id] = warnings.TryGetValue(w.Feature.Id, out var existing)
                    ? existing with { Occurrences = existing.Occurrences + w.Occurrences }
                    : w;
            foreach (var c in r.CoveredPresent)
                covered[c.Id] = c;
        }
        return new DocxDiffCompatibilityReport(warnings.Values.ToList(), covered.Values.ToList());
    }

    /// <summary>
    /// Validate the shared consolidate arguments once for all four N-way entry points and resolve the
    /// settings: rejects a null base/reviewers, a null <see cref="DocxDiffConsolidateSettings.Diff"/>, and any
    /// null reviewer element or reviewer <see cref="DocxDiffReviewer.Document"/> — surfacing a clear,
    /// attributed <see cref="ArgumentException"/> at the public boundary instead of a downstream
    /// <see cref="NullReferenceException"/>.
    /// </summary>
    private static DocxDiffConsolidateSettings ValidateConsolidateArgs(
        WmlDocument baseDocument, IReadOnlyList<DocxDiffReviewer> reviewers,
        DocxDiffConsolidateSettings? settings)
    {
        ArgumentNullException.ThrowIfNull(baseDocument);
        ArgumentNullException.ThrowIfNull(reviewers);
        var s = settings ?? new DocxDiffConsolidateSettings();
        if (s.Diff is null)
            throw new ArgumentException(
                $"{nameof(DocxDiffConsolidateSettings)}.{nameof(DocxDiffConsolidateSettings.Diff)} must not be null.",
                nameof(settings));
        for (int i = 0; i < reviewers.Count; i++)
        {
            if (reviewers[i] is null)
                throw new ArgumentException($"reviewers[{i}] is null.", nameof(reviewers));
            if (reviewers[i].Document is null)
                throw new ArgumentException(
                    $"reviewers[{i}].Document is null (author '{reviewers[i].Author}').", nameof(reviewers));
        }
        return s;
    }

    /// <summary>
    /// Apply <see cref="DocxDiffSettings.PreAcceptInputRevisions"/>: when set, return the fully accepted view of
    /// <paramref name="doc"/> (a new <see cref="WmlDocument"/>; the input is untouched) so both the IR read AND
    /// the output-package clone are revision-free; when not set, return <paramref name="doc"/> unchanged so the
    /// default path is byte-for-byte what it was before the flag existed.
    /// </summary>
    private static WmlDocument PreAccept(DocxDiffSettings settings, WmlDocument doc) =>
        settings.PreAcceptInputRevisions ? RevisionProcessor.AcceptRevisions(doc) : doc;

    /// <summary>Per-reviewer <see cref="PreAccept(DocxDiffSettings, WmlDocument)"/> for the N-way entry points;
    /// returns the original list unchanged when the flag is off.</summary>
    private static IReadOnlyList<DocxDiffReviewer> PreAccept(
        DocxDiffSettings settings, IReadOnlyList<DocxDiffReviewer> reviewers)
    {
        if (!settings.PreAcceptInputRevisions)
            return reviewers;
        return reviewers
            .Select(r => new DocxDiffReviewer { Author = r.Author, Document = RevisionProcessor.AcceptRevisions(r.Document) })
            .ToList();
    }
}

/// <summary>
/// How <see cref="DocxDiff"/> projects the edit script to consumer revisions. The edit script's grain is
/// the engine's truth and is untouched by this setting — it only governs coalescing/trimming on the way
/// out of <see cref="DocxDiff.GetRevisions"/>.
/// </summary>
public enum DocxDiffRevisionGranularity
{
    /// <summary>
    /// The engine's native, finest-grain projection (the default): one revision per token-op span. A
    /// paragraph whose every word changed yields one inserted + one deleted revision per word. This is the
    /// faithful, byte-stable mirror of the edit script — the right grain for review UIs that map a revision
    /// to a token span, blame, and structured indexers.
    /// </summary>
    Fine,

    /// <summary>
    /// A render-time projection that reproduces <see cref="WmlComparer"/>'s coarser <c>GetRevisions</c>
    /// atomization, so a <see cref="DocxDiff"/> revision set is count/text-comparable to the shipped
    /// comparer's. Per modified block it coalesces adjacent same-kind revisions into one maximal contiguous
    /// changed region (bridging purely-separator gaps), trims the common character prefix/suffix shared by a
    /// region's deleted and inserted text, and prunes zero-width (content-less) revisions. Move and
    /// format-change revisions pass through untouched. Choose this only when matching the legacy comparer's
    /// revision counts/texts matters; otherwise <see cref="Fine"/> is the more faithful surface.
    /// </summary>
    WmlComparerCompatible,
}

/// <summary>
/// How <see cref="DocxDiff"/> compares run formatting. A purely diff-time policy: it changes which format
/// facts a comparison treats as significant, never the documents themselves.
/// </summary>
public enum DocxDiffFormatComparison
{
    /// <summary>
    /// Compare only the MODELED run-format fields (bold, italic, underline, size, color, sub/superscript,
    /// strike, caps, highlight, …) — the default. A format change is reported only when a modeled field
    /// differs. <b>Trade-off:</b> a visible but UNMODELED rPr difference (e.g. <c>w:shd</c> run shading,
    /// complex-script toggles, secondary font faces, <c>w:lang</c>) reads as Unchanged — a false negative.
    /// This default exists because a <c>w:rPrChange</c>-grade report can only ever DESCRIBE modeled fields,
    /// so reporting an undescribable unmodeled-only flip is noise; comparing modeled fields collapses that
    /// noise without losing any delta a format-change report could express.
    /// </summary>
    ModeledOnly,

    /// <summary>
    /// Compare the FULL run format including the unmodeled rPr digest — every rPr difference (lang,
    /// complex-script toggles, secondary font faces, shading) is significant. Choose this for byte-fidelity
    /// consumers that must DETECT (even if they cannot fully describe) every formatting difference. The
    /// trade-off is the inverse of <see cref="ModeledOnly"/>: format-change noise from cosmetic rPr churn
    /// (e.g. editing tools rewriting <c>w:lang</c>/<c>w:bCs</c>) is reported.
    /// </summary>
    Full,
}

/// <summary>
/// Settings for <see cref="DocxDiff"/>. Defaults mirror <see cref="WmlComparerSettings"/> so the engine
/// reproduces the shipped comparer's word granularity and normalization out of the box, with two honest
/// deviations called out below (<see cref="Deterministic"/> revision dates and
/// <see cref="FormatComparison"/> defaulting to modeled-only). Its properties are settable; construct a
/// fresh instance per call rather than mutating one shared across concurrent diffs.
/// </summary>
/// <remarks>
/// This is the public mirror of the internal <c>IrDiffSettings</c>; it exposes the consumer-relevant
/// subset and maps onto it. Settings carry no per-document or process-global state — pass a different
/// <see cref="AuthorForRevisions"/> per call for multi-author pipelines.
/// </remarks>
public sealed class DocxDiffSettings
{
    /// <summary>
    /// Author name stamped on every revision's <see cref="DocxDiffRevision.Author"/> and on the produced
    /// markup's <c>w:author</c> attributes. Default <c>"Open-Xml-PowerTools"</c>, matching
    /// <see cref="WmlComparerSettings.AuthorForRevisions"/>. Flows per call — set a different author per
    /// comparison in a multi-author/consolidate pipeline.
    /// </summary>
    public string AuthorForRevisions { get; set; } = "Open-Xml-PowerTools";

    /// <summary>
    /// When true (the DEFAULT), revision dates are pinned to a fixed epoch
    /// (<c>2000-01-01T00:00:00Z</c>) so two diffs of the same inputs are byte-identical. Set false for
    /// wall-clock dates (matching <see cref="WmlComparerSettings"/>'s <c>DateTime.Now</c> default). This is
    /// an intentional deviation from <see cref="WmlComparerSettings"/>, which is nondeterministic by default.
    /// </summary>
    public bool Deterministic { get; set; } = true;

    /// <summary>
    /// The ISO-8601 date string stamped on every revision and on the produced markup's <c>w:date</c>
    /// attributes. When null (the default), the date is derived from <see cref="Deterministic"/>: the fixed
    /// epoch when deterministic, else <c>DateTime.Now</c> in round-trip ("o") format captured once per diff.
    /// An explicit non-null value always wins over both, and must be a parseable ISO-8601 date/time (it is
    /// stamped verbatim into the markup's <c>w:date</c> attributes) — a non-parseable value throws
    /// <see cref="ArgumentException"/>.
    /// </summary>
    public string? DateTimeForRevisions { get; set; }

    /// <summary>
    /// Optional pre-flight: a callback invoked once (BEFORE diffing) with a combined
    /// <see cref="DocxDiffCompatibilityReport"/> when an input document contains a feature DocxDiff has not had a
    /// fidelity campaign for (see <see cref="DocxDiffCompatibility"/>). Both inputs are inspected and their
    /// warnings unioned (occurrences summed); only <see cref="DocxDiffCoverage.Untested"/>/
    /// <see cref="DocxDiffCoverage.Partial"/> features trigger it. <b>Null (the default) skips the scan
    /// entirely</b> — zero behavior and zero performance change. The diff then proceeds normally (non-fatal);
    /// pair with <see cref="ThrowOnCompatibilityWarning"/> for a hard gate.
    /// </summary>
    public Action<DocxDiffCompatibilityReport>? OnCompatibilityWarning { get; set; }

    /// <summary>
    /// When true, a diff method THROWS <see cref="DocxDiffCompatibilityException"/> (after invoking
    /// <see cref="OnCompatibilityWarning"/>, if set) when an input contains an under-tested feature — a hard
    /// gate that produces no output. Default false. Orthogonal to <see cref="OnCompatibilityWarning"/>: the
    /// pre-flight scan runs only when at least one of the two is engaged, so the default is still zero-cost.
    /// </summary>
    public bool ThrowOnCompatibilityWarning { get; set; }

    /// <summary>
    /// When true, word match keys are case-folded (per <see cref="Culture"/>, or ordinal/invariant when
    /// <see cref="Culture"/> is null) so "Foo" matches "foo". Default false, matching
    /// <see cref="WmlComparerSettings.CaseInsensitive"/>.
    /// </summary>
    public bool CaseInsensitive { get; set; }

    /// <summary>
    /// Culture used for case folding when <see cref="CaseInsensitive"/> is true. Null (the default) means
    /// ordinal/invariant folding — no culture-specific casing. Mirrors
    /// <see cref="WmlComparerSettings.CultureInfo"/>.
    /// </summary>
    public CultureInfo? Culture { get; set; }

    /// <summary>
    /// When true (the DEFAULT), a non-breaking space (U+00A0) folds to an ordinary space (U+0020) in match
    /// keys, so NBSP-separated text matches space-separated text. The non-breaking hyphen (U+2011) is not
    /// folded. Mirrors <see cref="WmlComparerSettings.ConflateBreakingAndNonbreakingSpaces"/>.
    /// </summary>
    public bool ConflateBreakingAndNonbreakingSpaces { get; set; } = true;

    /// <summary>
    /// Characters that split a run's text into word vs. separator tokens; each separator character becomes
    /// its own token. <b>Null (the default)</b> uses the same default set as
    /// <see cref="WmlComparerSettings.WordSeparators"/>
    /// (<c>{ ' ', '-', ')', '(', ';', ',', and CJK punctuation }</c>). An <b>explicit set is honored
    /// verbatim — including an empty array</b>, which means nothing splits (the run becomes a single token,
    /// modulo NBSP conflation); only null falls back to the default.
    /// </summary>
    public char[]? WordSeparators { get; set; }

    /// <summary>
    /// When true (the DEFAULT), relocated content is reported as a move pair (a source revision sharing a
    /// <see cref="DocxDiffRevision.MoveGroupId"/> with its destination) and the produced markup uses native
    /// <c>w:moveFrom</c>/<c>w:moveTo</c>. When false, an aligned move is projected as an ordinary
    /// inserted+deleted pair and the markup uses plain <c>w:del</c>/<c>w:ins</c>. Mirrors
    /// <see cref="WmlComparerSettings.DetectMoves"/>. The engine always ALIGNS a relocation as a move; this
    /// only controls how it is REPORTED, so it works regardless of how the move was detected.
    /// </summary>
    public bool DetectMoves { get; set; } = true;

    /// <summary>
    /// Minimum block similarity (Jaccard over token match-key multisets, 0.0–1.0) for two leftover blocks to
    /// be re-paired as a cross-gap fuzzy move. Default 0.8, matching
    /// <see cref="WmlComparerSettings.MoveSimilarityThreshold"/>.
    /// </summary>
    public double MoveSimilarityThreshold { get; set; } = 0.8;

    /// <summary>
    /// Minimum number of word tokens both sides of a candidate fuzzy move must carry for it to be considered
    /// a move (short fragments are excluded to avoid false positives). Default 3, matching
    /// <see cref="WmlComparerSettings.MoveMinimumWordCount"/>.
    /// </summary>
    public int MoveMinimumWordCount { get; set; } = 3;

    /// <summary>
    /// How <see cref="DocxDiff.GetRevisions"/> projects the edit script to revisions. Default
    /// <see cref="DocxDiffRevisionGranularity.Fine"/> (the engine's native one-revision-per-token-span
    /// grain). Set <see cref="DocxDiffRevisionGranularity.WmlComparerCompatible"/> for revision counts/texts
    /// comparable to the shipped comparer's. Does not affect <see cref="DocxDiff.Compare"/> or
    /// <see cref="DocxDiff.GetEditScriptJson"/>.
    /// </summary>
    public DocxDiffRevisionGranularity RevisionGranularity { get; set; } = DocxDiffRevisionGranularity.Fine;

    /// <summary>
    /// How run formatting is compared. Default <see cref="DocxDiffFormatComparison.ModeledOnly"/> — see that
    /// member for the false-negative trade-off and the rationale, and <see cref="DocxDiffFormatComparison.Full"/>
    /// for byte-fidelity comparison.
    /// </summary>
    public DocxDiffFormatComparison FormatComparison { get; set; } = DocxDiffFormatComparison.ModeledOnly;

    /// <summary>
    /// When true, every input document is run through <see cref="RevisionProcessor.AcceptRevisions(WmlDocument)"/>
    /// BEFORE it is diffed — so the comparison, AND the output package the markup renderer clones from those
    /// inputs, are built over the fully <b>accepted view</b> of each side. This is the first-class, opt-in form
    /// of the "accept-all-both-sides, then compare" wrapper. <b>Default false</b>, which is a zero behavior change:
    /// the engine already diffs the accepted view at the edit-script/body level, but carries pre-existing
    /// revision markup in carried-over parts (headers/footers, unchanged footnotes/endnotes, styles, comments)
    /// straight through into the result. Setting this true runs the accept-all, which flattens the body,
    /// headers/footers, footnotes/endnotes, and styles, so every revision the markup renderer emits in those
    /// scopes is attributable to THIS diff (none is a stale input revision passed through) and the round-trip
    /// holds against the accepted view of each side there. <b>Coverage is exactly the parts
    /// <see cref="RevisionProcessor.AcceptRevisions(WmlDocument)"/> processes</b> — the main body,
    /// headers/footers, footnotes/endnotes, and styles. Carried-over parts it does NOT process — notably the
    /// comments part and the glossary (building-blocks / AutoText) document part — are left intact, so a
    /// pre-existing revision inside a comment definition or a glossary entry survives regardless of this flag.
    ///
    /// <para><b>Two honest costs of accept-all (read before enabling).</b></para>
    /// <para>1. <b>It flattens pre-existing authorship and change boundaries.</b> Accepting collapses each input's
    /// own tracked changes into final text, so WHO made a prior edit and WHERE the prior change boundaries were
    /// are lost — the result's authorship reflects only this diff, not the inputs' edit history.</para>
    /// <para>2. <b>"Accept all" is itself a policy.</b> It overrides any change a prior reviewer had left
    /// unaccepted — including one they had effectively rejected by leaving in tracked form — by materializing
    /// every insertion and dropping every deletion. If you need to preserve or re-adjudicate the inputs'
    /// in-flight revisions, do not enable this; resolve them first by your own policy, then diff.</para>
    ///
    /// <para>.NET-only in v1: not yet surfaced on the WASM/npm/python bridges.</para>
    /// </summary>
    public bool PreAcceptInputRevisions { get; set; }

    /// <summary>Map this public settings object onto the internal <c>IrDiffSettings</c>.</summary>
    internal IrDiffSettings ToIrDiffSettings()
    {
        // Resolve the revision date: an explicit value always wins; otherwise derive from Deterministic.
        // When non-deterministic with no explicit date, capture DateTime.Now once here (not per revision)
        // so a single diff is internally consistent, mirroring a single WmlComparerSettings instance.
        var (deterministic, date) = (Deterministic, DateTimeForRevisions) switch
        {
            (_, { } explicitDate) => (Deterministic, explicitDate),
            (true, null) => (true, IrDiffSettings.DeterministicEpoch),
            (false, null) => (false, DateTime.Now.ToString("o", CultureInfo.InvariantCulture)),
        };

        // An explicit date is stamped verbatim into every w:date attribute (xsd:dateTime); validate it parses
        // so garbage fails at the boundary instead of producing schema-questionable markup. The deterministic
        // epoch and the captured DateTime.Now are always valid, so only the caller-supplied value is checked.
        if (DateTimeForRevisions is { } supplied &&
            !DateTimeOffset.TryParse(supplied, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out _))
        {
            throw new ArgumentException(
                $"DateTimeForRevisions '{supplied}' is not a valid ISO-8601 date/time.",
                nameof(DateTimeForRevisions));
        }

        return new IrDiffSettings
        {
            AuthorForRevisions = AuthorForRevisions,
            Deterministic = deterministic,
            DateTimeForRevisions = date,
            CaseInsensitive = CaseInsensitive,
            Culture = Culture,
            ConflateBreakingAndNonbreakingSpaces = ConflateBreakingAndNonbreakingSpaces,
            // Only NULL falls back to the default set; an explicit set — even empty — is honored verbatim.
            WordSeparators = WordSeparators is { } seps
                ? System.Collections.Immutable.ImmutableHashSet.CreateRange(seps)
                : IrDiffSettings.DefaultWordSeparators,
            RenderMoves = DetectMoves,
            MoveSimilarityThreshold = MoveSimilarityThreshold,
            MoveMinimumTokenCount = MoveMinimumWordCount,
            RevisionGranularity = RevisionGranularity == DocxDiffRevisionGranularity.WmlComparerCompatible
                ? Docxodus.Ir.Diff.RevisionGranularity.WmlComparerCompatible
                : Docxodus.Ir.Diff.RevisionGranularity.Fine,
            FormatComparison = FormatComparison == DocxDiffFormatComparison.Full
                ? IrFormatComparison.Full
                : IrFormatComparison.ModeledOnly,
        };
    }
}

/// <summary>
/// The kind of a <see cref="DocxDiffRevision"/>. One-for-one with
/// <see cref="WmlComparer.WmlComparerRevisionType"/> plus the <see cref="Moved"/> and
/// <see cref="FormatChanged"/> kinds the IR engine surfaces natively.
/// </summary>
public enum DocxDiffRevisionType
{
    /// <summary>Right-only content (a <c>w:ins</c>-grade insertion).</summary>
    Inserted,

    /// <summary>Left-only content (a <c>w:del</c>-grade deletion).</summary>
    Deleted,

    /// <summary>
    /// Relocated content. Source and destination revisions share a <see cref="DocxDiffRevision.MoveGroupId"/>;
    /// <see cref="DocxDiffRevision.IsMoveSource"/> distinguishes the two. Only produced when
    /// <see cref="DocxDiffSettings.DetectMoves"/> is true.
    /// </summary>
    Moved,

    /// <summary>
    /// Content unchanged in text but changed in modeled run formatting (a <c>w:rPrChange</c>-grade change).
    /// <see cref="DocxDiffRevision.FormatChange"/> carries the old/new properties and the changed names.
    /// </summary>
    FormatChanged,
}

/// <summary>
/// Details of a <see cref="DocxDiffRevisionType.FormatChanged"/> revision: the modeled run-format fields
/// before and after, plus the names of the fields that differ. Mirrors
/// <see cref="WmlComparer.FormatChangeDetails"/>.
/// </summary>
/// <remarks>
/// The dictionaries enumerate only MODELED format fields (bold, italic, underline, fontSize, color, …),
/// keyed by friendly property names. A field present on one side and absent on the other (e.g. bold added)
/// appears in only that side's dictionary and in <see cref="ChangedPropertyNames"/>.
/// </remarks>
public sealed class DocxDiffFormatChange
{
    /// <summary>The modeled format properties on the LEFT (old) side, keyed by friendly name.</summary>
    public IReadOnlyDictionary<string, string> OldProperties { get; }

    /// <summary>The modeled format properties on the RIGHT (new) side, keyed by friendly name.</summary>
    public IReadOnlyDictionary<string, string> NewProperties { get; }

    /// <summary>The friendly names of the properties that differ between the two sides.</summary>
    public IReadOnlyList<string> ChangedPropertyNames { get; }

    internal DocxDiffFormatChange(IrFormatChangeDetails details)
    {
        OldProperties = details.OldProperties;
        NewProperties = details.NewProperties;
        ChangedPropertyNames = details.ChangedPropertyNames;
    }
}

/// <summary>
/// One consumer-facing revision from <see cref="DocxDiff.GetRevisions"/>. Mirrors the consumer-relevant
/// shape of <see cref="WmlComparer.WmlComparerRevision"/>
/// (<see cref="Type"/>/<see cref="Text"/>/<see cref="Author"/>/<see cref="Date"/>/<see cref="MoveGroupId"/>/
/// <see cref="IsMoveSource"/>/<see cref="FormatChange"/>) and ADDS the block anchors the revision derives
/// from — the IR engine's differentiator over <see cref="WmlComparer.WmlComparerRevision"/>.
/// </summary>
/// <remarks>
/// <para><b>Anchor grammar.</b> <see cref="LeftAnchor"/> and <see cref="RightAnchor"/> are stable block
/// anchors of the form <c>kind:scope:unid</c> — e.g. <c>p:body:a1b2c3d4</c> (a body paragraph),
/// <c>li:body:…</c> (a list item), <c>tbl:body:…</c> (a table), <c>p:fn3:…</c> (a paragraph in footnote 3).
/// The <c>kind</c> and <c>scope</c> match the anchor grammar of <see cref="WmlToMarkdownConverter"/>'s
/// markdown projection and of <see cref="DocxSession"/>'s anchor-addressed editing API, so a revision can be
/// located in the markdown projection or fed straight to a <see cref="DocxSession"/> mutation
/// (<c>ReplaceText</c>, <c>GetBlockMetadata</c>, <c>AddAnnotation</c>, …) on the corresponding document.
/// The <c>unid</c> is the IR's deterministic per-element id.</para>
///
/// <para><b>Anchor presence by <see cref="Type"/>.</b> Each type's PRIMARY anchor is ALWAYS present; the
/// opposite anchor MAY also be present for a TOKEN-LEVEL revision (see below). The full rule:
/// <list type="bullet">
/// <item><see cref="DocxDiffRevisionType.Inserted"/> — <see cref="RightAnchor"/> always; <see cref="LeftAnchor"/>
/// null for a whole-block insertion, but PRESENT for a token-level insert inside a modified block.</item>
/// <item><see cref="DocxDiffRevisionType.Deleted"/> — <see cref="LeftAnchor"/> always; <see cref="RightAnchor"/>
/// null for a whole-block deletion, but PRESENT for a token-level delete inside a modified block.</item>
/// <item><see cref="DocxDiffRevisionType.FormatChanged"/> — BOTH (it is always a content-equal block pair).</item>
/// <item><see cref="DocxDiffRevisionType.Moved"/> — EXCLUSIVE: source carries <see cref="LeftAnchor"/> only
/// (<see cref="RightAnchor"/> null), destination carries <see cref="RightAnchor"/> only.</item>
/// </list>
/// A TOKEN-LEVEL revision is one describing an insert/delete WITHIN a modified (or moved-and-edited) paragraph
/// — that paragraph exists on both sides, so the revision carries BOTH the left and right anchors of its
/// enclosing block (useful for locating the edit in either document). A BLOCK-LEVEL insert/delete (a whole
/// new/removed paragraph) carries only its primary anchor. Each anchor resolves against the IR of the document
/// it came FROM: left anchors against <see cref="DocxDiff.GetRevisions"/>'s <c>left</c> argument, right anchors
/// against <c>right</c>.</para>
/// </remarks>
public sealed class DocxDiffRevision
{
    /// <summary>The kind of change.</summary>
    public DocxDiffRevisionType Type { get; }

    /// <summary>The affected text.</summary>
    public string Text { get; }

    /// <summary>The author stamped on the revision (from <see cref="DocxDiffSettings.AuthorForRevisions"/>).</summary>
    public string Author { get; }

    /// <summary>The ISO-8601 revision date (deterministic epoch by default).</summary>
    public string Date { get; }

    /// <summary>
    /// For <see cref="DocxDiffRevisionType.Moved"/> revisions, the id linking a move's source and
    /// destination; null otherwise.
    /// </summary>
    public int? MoveGroupId { get; }

    /// <summary>
    /// For <see cref="DocxDiffRevisionType.Moved"/> revisions: true at the source (moved FROM here), false at
    /// the destination (moved TO here); null for non-move revisions.
    /// </summary>
    public bool? IsMoveSource { get; }

    /// <summary>
    /// For <see cref="DocxDiffRevisionType.FormatChanged"/> revisions, the old/new modeled properties and the
    /// changed names; null otherwise.
    /// </summary>
    public DocxDiffFormatChange? FormatChange { get; }

    /// <summary>
    /// The LEFT-document block anchor (<c>kind:scope:unid</c>) the revision derives from, or null when it has
    /// no left side (a pure insertion). See the type remarks for the anchor grammar and presence rules.
    /// </summary>
    public string? LeftAnchor { get; }

    /// <summary>
    /// The RIGHT-document block anchor (<c>kind:scope:unid</c>) the revision derives from, or null when it has
    /// no right side (a pure deletion). See the type remarks for the anchor grammar and presence rules.
    /// </summary>
    public string? RightAnchor { get; }

    private DocxDiffRevision(
        DocxDiffRevisionType type, string text, string author, string date,
        int? moveGroupId, bool? isMoveSource, DocxDiffFormatChange? formatChange,
        string? leftAnchor, string? rightAnchor)
    {
        Type = type;
        Text = text;
        Author = author;
        Date = date;
        MoveGroupId = moveGroupId;
        IsMoveSource = isMoveSource;
        FormatChange = formatChange;
        LeftAnchor = leftAnchor;
        RightAnchor = rightAnchor;
    }

    internal static DocxDiffRevision FromIr(IrRevision r) => new(
        type: r.Type switch
        {
            IrRevisionType.Inserted => DocxDiffRevisionType.Inserted,
            IrRevisionType.Deleted => DocxDiffRevisionType.Deleted,
            IrRevisionType.Moved => DocxDiffRevisionType.Moved,
            IrRevisionType.FormatChanged => DocxDiffRevisionType.FormatChanged,
            _ => throw new ArgumentOutOfRangeException(nameof(r), r.Type, "Unknown IrRevisionType."),
        },
        text: r.Text,
        author: r.Author,
        date: r.Date,
        moveGroupId: r.MoveGroupId,
        isMoveSource: r.IsMoveSource,
        formatChange: r.FormatChange is { } fc ? new DocxDiffFormatChange(fc) : null,
        leftAnchor: r.LeftAnchor,
        rightAnchor: r.RightAnchor);
}
