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
        var diff = (settings ?? new DocxDiffSettings()).ToIrDiffSettings();
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
        var diff = (settings ?? new DocxDiffSettings()).ToIrDiffSettings();
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
        var diff = (settings ?? new DocxDiffSettings()).ToIrDiffSettings();
        var irLeft = IrReader.Read(left, ReadOpts);
        var irRight = IrReader.Read(right, ReadOpts);
        var script = IrEditScriptBuilder.Build(irLeft, irRight, diff);
        return IrEditScriptJson.Write(script);
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
/// <see cref="FormatComparison"/> defaulting to modeled-only). Immutable; construct fresh per call.
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
    /// An explicit non-null value always wins over both.
    /// </summary>
    public string? DateTimeForRevisions { get; set; }

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
    /// its own token. Null (the default) uses the same default set as
    /// <see cref="WmlComparerSettings.WordSeparators"/>
    /// (<c>{ ' ', '-', ')', '(', ';', ',', and CJK punctuation }</c>).
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

        return new IrDiffSettings
        {
            AuthorForRevisions = AuthorForRevisions,
            Deterministic = deterministic,
            DateTimeForRevisions = date,
            CaseInsensitive = CaseInsensitive,
            Culture = Culture,
            ConflateBreakingAndNonbreakingSpaces = ConflateBreakingAndNonbreakingSpaces,
            WordSeparators = WordSeparators is { Length: > 0 }
                ? System.Collections.Immutable.ImmutableHashSet.CreateRange(WordSeparators)
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
