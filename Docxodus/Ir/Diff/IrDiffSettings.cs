#nullable enable

using System.Collections.Immutable;
using System.Globalization;

namespace Docxodus.Ir.Diff;

/// <summary>
/// How the diff engine compares formatting (M2.2 Task 4). A purely DIFF-TIME policy: it changes which
/// format facts a comparison treats as significant, NOT the IR's stored hashes (no snapshot churn).
/// </summary>
internal enum IrFormatComparison
{
    /// <summary>
    /// Compare only the MODELED run-format fields (Bold/Italic/Underline/Size/Color/… — the
    /// <see cref="IrRunFormat"/> record EXCLUDING its <see cref="IrRunFormat.UnmodeledDigest"/>). The
    /// DEFAULT.
    /// <para><b>Why default.</b> The WC-BodyBookmarks diagnosis (M2.2 Task 4, sub-task B) showed the
    /// corpus' entire FormatOnly population (1,714 entries) comes from content-equal paragraphs whose
    /// ONLY format difference is unmodeled rPr leftovers — <c>w:lang</c> (4597), <c>w:iCs</c> (1328),
    /// <c>w:bCs</c> (550), <c>w:rFonts</c> hAnsi/cs faces (33), <c>w:szCs</c>/<c>w:rtl</c> — with every
    /// MODELED field byte-identical. Those are legitimate IR facts but pure noise for diff purposes:
    /// a <c>w:rPrChange</c>-grade format-change report can only ever DESCRIBE modeled fields anyway, so
    /// reporting a format change driven by an undescribable unmodeled-digest flip is a false positive.
    /// Comparing modeled fields only collapses that noise (FormatOnly → Unchanged) without losing any
    /// format delta a <c>w:rPrChange</c>-grade report could DESCRIBE. The honest trade-off: a visible
    /// but UNMODELED format change (e.g. <c>w:shd</c> run shading) is a false NEGATIVE under this
    /// default — it reads as Unchanged. Consumers needing to detect (if not describe) such changes use
    /// <see cref="Full"/>.</para>
    /// </summary>
    ModeledOnly,

    /// <summary>
    /// Compare the FULL run format including <see cref="IrRunFormat.UnmodeledDigest"/> — i.e. trust the
    /// reader-computed <c>FormatFingerprint</c> verbatim. Available for byte-fidelity consumers that
    /// must see every rPr difference (lang, complex-script toggles, secondary font faces). This is the
    /// M2.1 behavior.
    /// </summary>
    Full,
}

/// <summary>
/// RENDER-TIME revision granularity policy (M2.4 Task 2). A purely <see cref="IrRevisionRenderer"/>-level
/// transform: it changes HOW the (unchanged) edit script is PROJECTED to consumer revisions, never the
/// script itself, the aligner, or the token diff. The edit script's grain is the engine's truth and is
/// untouchable; this only governs coalescing/trimming on the way out.
/// </summary>
internal enum RevisionGranularity
{
    /// <summary>
    /// The engine's native, finest-grain projection (the DEFAULT and the M2.3 behavior): one revision per
    /// token-op span. A paragraph whose every word changed yields one Inserted + one Deleted per word. This
    /// is the faithful mirror of the edit script — byte-stable across releases — and the right grain for
    /// consumers that want the engine's actual atomization (review UIs that map a revision to a token span,
    /// blame, structured indexers).
    /// </summary>
    Fine,

    /// <summary>
    /// A render-time projection that reproduces <c>WmlComparer.GetRevisions</c>'s coarser atomization, so an
    /// IR-rendered revision SET is count/text-comparable to the shipped comparer's. WmlComparer's revisions
    /// come from contiguous <c>w:ins</c>/<c>w:del</c> regions of the produced document — one revision per
    /// maximal contiguous changed region — so this mode, per Modified block:
    /// <list type="number">
    /// <item><b>Coalesces</b> adjacent same-kind token-op revisions into one, INCLUDING across an Equal
    /// op that is PURELY separators (whitespace/punctuation) sitting between two changed words — those
    /// separators are part of WmlComparer's contiguous region. An Equal op containing any Word token is a
    /// true region boundary and is NOT bridged.</item>
    /// <item><b>Trims</b> the common character prefix and suffix shared by a coalesced region's deleted and
    /// inserted text (WmlComparer keeps the common edges unchanged and only marks the differing middle),
    /// dropping a side that becomes empty.</item>
    /// <item><b>Prunes</b> zero-width revisions (empty <see cref="IrRevision.Text"/> Inserted/Deleted) that
    /// arise from non-text placeholder tokens (masked textbox placeholders, section breaks) — WmlComparer
    /// reports no revision for a content-less change at this surface.</item>
    /// </list>
    /// Move and FormatChanged revisions are passed through untouched; only Inserted/Deleted are coalesced.
    /// </summary>
    WmlComparerCompatible,
}

/// <summary>
/// Diff-time settings for the IR diff engine (Phase 2). These govern how IR paragraphs are
/// tokenized and compared; they are <b>not</b> document facts. Per the IR spec (§1 non-goals,
/// "Not the diff's tokenization"), word splitting, case folding, and separator policy are
/// comparison settings that live here — the IR itself stores raw runs and never applies them.
/// </summary>
/// <remarks>
/// The defaults mirror <see cref="WmlComparerSettings"/> so the IR diff path reproduces the
/// shipped comparer's word granularity and normalization out of the box.
/// </remarks>
internal sealed record IrDiffSettings
{
    /// <summary>
    /// DIFF-TIME setting. How formatting is compared at both the token level (the differ's
    /// FormatChanged post-pass) and the block level (the aligner's FormatOnly classification). Default
    /// <see cref="IrFormatComparison.ModeledOnly"/> — see that member for the evidence.
    /// </summary>
    /// <remarks>
    /// <b>Layering (purely diff-time; the IR's stored hashes never change).</b>
    /// <list type="bullet">
    /// <item><b>Token level.</b> <see cref="IrFormatComparison.ModeledOnly"/> compares
    /// <see cref="IrRunFormat"/> records EXCLUDING <see cref="IrRunFormat.UnmodeledDigest"/>, so a
    /// lang/iCs/bCs-only difference does not raise a FormatChanged token span.</item>
    /// <item><b>Block level.</b> The aligner cannot trust the reader's block <c>FormatFingerprint</c>
    /// for FormatOnly under ModeledOnly (that fingerprint folds in the UnmodeledDigest AND is
    /// run-boundary-sensitive). Instead it recomputes a BOUNDARY-NORMALIZED modeled-only signature at
    /// diff time: the sequence of <c>(token MatchKey, modeled-format key)</c> over the paragraph's
    /// tokens. Because it keys on the boundary-independent token stream rather than the raw run
    /// segmentation, editing churn that re-segments runs (the M2.1 finding) no longer flips it. A pair
    /// that is ContentHash-equal but whose modeled-only signatures differ is FormatOnly; equal
    /// signatures are Unchanged.</item>
    /// </list>
    /// Under <see cref="IrFormatComparison.Full"/> both levels fall back to the M2.1 behavior (full
    /// record equality at the token level; the stored block FormatFingerprint at the block level).
    /// </remarks>
    public IrFormatComparison FormatComparison { get; init; } = IrFormatComparison.ModeledOnly;

    /// <summary>
    /// RENDER-TIME setting (M2.4 Task 2). How <see cref="IrRevisionRenderer"/> projects the edit script to
    /// consumer revisions. Default <see cref="Docxodus.Ir.Diff.RevisionGranularity.Fine"/> (the engine's
    /// native one-revision-per-token-span grain — byte-stable). The adapter that targets
    /// <c>WmlComparer.GetRevisions</c> parity sets
    /// <see cref="Docxodus.Ir.Diff.RevisionGranularity.WmlComparerCompatible"/>. This NEVER affects the edit
    /// script, the aligner, or the token diff — only the rendered revision list.
    /// </summary>
    public RevisionGranularity RevisionGranularity { get; init; } = RevisionGranularity.Fine;

    /// <summary>
    /// RENDER-TIME setting (M2.4 Task 2). When false, <see cref="IrRevisionRenderer"/> renders
    /// <c>Moved</c>/<c>MoveModify</c> ops as an Inserted (at the destination) + Deleted (at the source) PAIR
    /// instead of a <c>Moved</c> pair — the projection a consumer that has move detection turned off expects.
    /// Default true. This is the render-time analogue of <c>WmlComparerSettings.DetectMoves</c>; the engine
    /// alignment is unchanged (a relocated block is still ALIGNED as a move — we only relabel its projection),
    /// so it works regardless of how the move arose (aligner off-spine anchoring OR fuzzy similarity).
    /// </summary>
    public bool RenderMoves { get; init; } = true;

    /// <summary>
    /// MARKUP-RENDER setting (M2.4 Task 4). When true, <see cref="IrMarkupRenderer"/> rewrites the native
    /// <c>w:moveFrom</c>/<c>w:moveTo</c> markup it produces into plain <c>w:del</c>/<c>w:ins</c> and strips the
    /// move range markers, as a post-pass over the assembled document. The render-time analogue of
    /// <c>WmlComparerSettings.SimplifyMoveMarkup</c> (a Word-compatibility workaround for renderers that do not
    /// honor native move markup). Default false — native move markup is emitted. Only meaningful when
    /// <see cref="RenderMoves"/> is true (otherwise no move markup is produced to simplify).
    /// </summary>
    public bool SimplifyMoveMarkup { get; init; }

    /// <summary>
    /// DIFF-TIME setting. Characters that split an <c>IrTextRun</c>'s text into word vs. separator
    /// tokens. Each separator character becomes its own <see cref="IrDiffTokenKind.Separator"/> token
    /// (matching <c>WmlComparer</c>'s atom granularity — one atom per separator char).
    /// </summary>
    /// <remarks>
    /// Default copied verbatim from <c>WmlComparerSettings.WordSeparators</c> (Docxodus/WmlComparer.cs
    /// ~line 123): <c>{ ' ', '-', ')', '(', ';', ',', '（', '）', '，', '、', '、', '，', '；', '。',
    /// '：', '的' }</c>. Held as an <see cref="ImmutableHashSet{T}"/> for O(1) membership during the
    /// per-character tokenizer walk (the comparer's source carries duplicate CJK entries, which the
    /// set folds away harmlessly).
    /// </remarks>
    public ImmutableHashSet<char> WordSeparators { get; init; } = DefaultWordSeparators;

    /// <summary>
    /// DIFF-TIME setting. When true, word match keys are case-folded (per <see cref="Culture"/>, or
    /// ordinal/invariant when <see cref="Culture"/> is null) so "Foo" matches "foo". Default false,
    /// matching <c>WmlComparerSettings.CaseInsensitive</c>.
    /// </summary>
    public bool CaseInsensitive { get; init; }

    /// <summary>
    /// DIFF-TIME setting. When true, a non-breaking space (U+00A0) folds to an ordinary space
    /// (U+0020) in match keys, so NBSP-separated text matches space-separated text. The non-breaking
    /// hyphen (U+2011) is deliberately <b>not</b> folded — it is not a space. Default true, matching
    /// <c>WmlComparerSettings.ConflateBreakingAndNonbreakingSpaces</c>.
    /// </summary>
    public bool ConflateBreakingAndNonbreakingSpaces { get; init; } = true;

    /// <summary>
    /// DIFF-TIME setting. Culture used for case folding when <see cref="CaseInsensitive"/> is true.
    /// Null (the default) means ordinal/invariant folding (<c>ToLowerInvariant</c>) — no
    /// culture-specific casing.
    /// </summary>
    public CultureInfo? Culture { get; init; }

    /// <summary>
    /// DIFF-TIME setting. Minimum block similarity (Jaccard over token <c>MatchKey</c> multisets,
    /// 0.0–1.0) for two blocks left UNPAIRED after a gap's exact refinement to be paired as
    /// <c>Modified</c> (a "same block, edited" pairing) rather than falling out as separate
    /// <c>Deleted</c>+<c>Inserted</c>. Default 0.35.
    /// </summary>
    /// <remarks>
    /// <b>Why 0.35.</b> Word's own compare pairs paragraphs well below half token-overlap — e.g. a
    /// sentence rewrite keeping only a trailing clause (Jaccard ≈ 0.375) renders in Word's redline as ONE
    /// paragraph with interleaved ins/del runs, not as a deleted + an inserted paragraph. 0.35 was
    /// calibrated against a 200-document Word-compare oracle corpus: it captures those rewrites (raising
    /// pixel fidelity vs Word's arrangement measurably) while still refusing junk pairings — at even
    /// lower floors (≈0.15) the pairing starts claiming unrelated paragraphs are revisions of each other,
    /// which misleads review UIs and measurably WORSENS fidelity on rewritten documents. This is the
    /// in-gap floor; cross-gap MOVES demand the stricter <see cref="MoveSimilarityThreshold"/> because
    /// relocating-and-editing is a stronger claim than editing in place.
    /// </remarks>
    public double BlockSimilarityThreshold { get; init; } = 0.35;

    /// <summary>
    /// DIFF-TIME setting. Minimum block similarity (Jaccard over token <c>MatchKey</c> multisets,
    /// 0.0–1.0) for two GLOBALLY-leftover blocks (one deleted, one inserted, in different gaps) to be
    /// re-paired as a cross-gap fuzzy move (<c>MovedModified</c>). Default 0.8.
    /// </summary>
    /// <remarks>
    /// Default 0.8 mirrors <c>WmlComparerSettings.MoveSimilarityThreshold</c> (Docxodus/WmlComparer.cs
    /// ~line 85, "80% word overlap required") so the IR diff's fuzzy-move bar matches the shipped
    /// comparer's. Strictly higher than <see cref="BlockSimilarityThreshold"/>: a move asserts the block
    /// relocated AND was edited, a stronger claim than an in-place edit, so it needs stronger evidence.
    /// </remarks>
    public double MoveSimilarityThreshold { get; init; } = 0.8;

    /// <summary>
    /// DIFF-TIME setting. Minimum number of <see cref="IrDiffTokenKind.Word"/> tokens that BOTH sides of
    /// a candidate cross-gap fuzzy move must carry for it to be considered a <c>MovedModified</c> pair.
    /// Counts Word-kind tokens only (separators, tabs, breaks, refs, images do not count). Default 3.
    /// </summary>
    /// <remarks>
    /// Default 3 mirrors <c>WmlComparerSettings.MoveMinimumWordCount</c> (Docxodus/WmlComparer.cs
    /// ~line 92, "very short text is excluded to avoid false positives"). Short fragments (a heading word,
    /// a list bullet) are similar to too many candidates by coincidence, so excluding them is the
    /// dominant false-positive guard for move detection.
    /// </remarks>
    public int MoveMinimumTokenCount { get; init; } = 3;

    /// <summary>
    /// DIFF-TIME setting (M2.6). When true (the DEFAULT), the aligner's gap fill runs the 1:N paragraph
    /// split / N:1 merge containment scan (after similarity pairing, before the 1×1-residue rule),
    /// emitting <see cref="IrEditOpKind.SplitBlock"/>/<see cref="IrEditOpKind.MergeBlock"/> ops — so a
    /// paragraph split mid-text (Enter pressed) or two paragraphs fused report as a paragraph-mark
    /// revision plus per-segment edits instead of an inflated whole-paragraph delete+insert pair
    /// (WC-1450/WC-1830). Set false for strict 1:1 op semantics (every op carries at most one anchor
    /// per side; splits fall back to the pre-M2.6 Modify+Insert account).
    /// </summary>
    public bool DetectSplitMerge { get; init; } = true;

    /// <summary>
    /// DIFF-TIME setting (M2.6). Minimum in-order LCS coverage of the singular-side paragraph's content
    /// tokens by the candidate run for a split/merge to fire. STARTING HYPOTHESIS 0.90 (spec §2.2) —
    /// the corpus sweep (IrSplitThresholdSweepTests) is the gate that pins the shipped value (F4.1).
    /// </summary>
    public double SplitCoverageThreshold { get; init; } = 0.90;

    /// <summary>
    /// DIFF-TIME setting (M2.6). Maximum fraction of the candidate run's content tokens NOT matched by
    /// the LCS (net-new content, e.g. WC-1830's inserted math paragraph). Starting hypothesis 0.34; swept.
    /// </summary>
    public double SplitForeignSlack { get; init; } = 0.34;

    /// <summary>DIFF-TIME setting (M2.6). Hard cap on a split/merge candidate run's block count
    /// (bounds the per-gap O(G²) candidate scan on pathological gaps).</summary>
    public int SplitMaxRunLength { get; init; } = 8;

    /// <summary>
    /// DIFF-TIME setting (header/footer campaign, 2026-07-03). When true (the DEFAULT — Word's own
    /// Compare "Headers and footers" granularity default), <see cref="IrEditScriptBuilder"/> diffs the
    /// header/footer stories (paired per section ordinal × occurrence kind, with Word's
    /// previous-section inheritance rule) into <see cref="IrEditScript.HeaderFooterOps"/>, the markup
    /// renderer rebuilds changed stories with native tracked-changes markup, and Fine-granularity
    /// revisions include hdr/ftr-anchored entries. When false, header/footer scopes are ignored
    /// entirely — the pre-campaign behavior: the output carries the LEFT package's header/footer parts
    /// verbatim and no header change is reported anywhere.
    /// </summary>
    public bool CompareHeadersFooters { get; init; } = true;

    /// <summary>
    /// RENDER-TIME setting (Word-parity input-revision preservation). When true, the markup renderer
    /// carries the RIGHT input's pre-existing tracked-revision markup through into the output for
    /// content-EQUAL blocks (verbatim from the ORIGINAL right element rather than the accepted working
    /// copy) and for whole-block INSERTED content (the diff's <c>w:ins</c> wraps only plain runs; a
    /// foreign <c>w:ins</c>/<c>w:del</c> child stays as-is, never nested same-kind) — in the body and
    /// in footnote/endnote bodies (note definitions pair by id). All char-precise paths
    /// (modify/format-only/split/merge/move, changed header/footer stories) keep rendering
    /// over the accepted view. Default false — the accepted-view-only behavior. Public mirror:
    /// <see cref="DocxDiffSettings.PreserveInputRevisions"/> (see it for the one-sided round-trip
    /// contract: accept ≡ accept(right) holds; reject ≠ left where foreign markup exists).
    /// </summary>
    public bool PreserveInputRevisions { get; init; }

    /// <summary>
    /// DIFF-TIME setting (block-format-change family, 2026-07-03). When true (the DEFAULT), paragraph-and-above
    /// property changes are DETECTED and TRACKED: the aligner's modeled-only block signature includes the
    /// paragraph's modeled <see cref="IrParaFormat"/> key (so a pPr-only change classifies FormatOnly instead of
    /// Unchanged), and the markup renderer emits native property-revision markup (<c>w:pPrChange</c>, and — as
    /// later phases land — the table-shell and section variants) at the sites that clone right-side properties.
    /// When false, the pre-campaign behavior is restored exactly: block-property deltas are invisible to
    /// classification and applied untracked. Forced off by <see cref="IrCompositeMerger"/> for its per-reviewer
    /// diffs — the Consolidate v1 ceiling, pinned by
    /// <c>BlockFormatChangeTests.Consolidate_ignores_block_format_changes_v1_ceiling</c> (the
    /// <see cref="CompareHeadersFooters"/> precedent).
    /// </summary>
    public bool TrackBlockFormatChanges { get; init; } = true;

    /// <summary>
    /// DIFF-TIME setting (Consolidate sub-project B). The PARAGRAPH-scope slice of
    /// <see cref="TrackBlockFormatChanges"/>: gates ONLY the <c>w:pPrChange</c> (paragraph properties + mark
    /// rPr) detection/emission, NOT the table-shell or section variants. Defaults equal to
    /// <see cref="TrackBlockFormatChanges"/> (so every two-way call behaves byte-identically — the split is
    /// invisible outside the composite). The composite merger sets this TRUE while forcing
    /// <see cref="TrackBlockFormatChanges"/> FALSE, so <c>Consolidate</c> merges reviewers' pPr changes
    /// (B1) while table-shell/section composite merge stays a v1 ceiling (B2).
    /// </summary>
    public bool TrackParagraphFormatChanges { get; init; } = true;

    /// <summary>
    /// DIFF-TIME setting (Consolidate sub-project B2). The TABLE-SHELL slice of
    /// <see cref="TrackBlockFormatChanges"/>: gates ONLY the table-shell property-revision markup
    /// (<c>w:tcPrChange</c>/<c>w:trPrChange</c>/<c>w:tblPrChange</c>/<c>w:tblGridChange</c>/<c>w:tblPrExChange</c>),
    /// NOT the paragraph or section variants. Defaults equal to <see cref="TrackBlockFormatChanges"/> (so every
    /// two-way call behaves byte-identically — the split is invisible outside the composite). The composite
    /// merger + renderers set this TRUE while forcing <see cref="TrackBlockFormatChanges"/> FALSE, so
    /// <c>Consolidate</c> merges reviewers' table-shell changes (B2) with per-element attribution.
    /// </summary>
    public bool TrackTableFormatChanges { get; init; } = true;

    /// <summary>
    /// DIFF-TIME setting (Consolidate sub-project B2). The SECTION slice of
    /// <see cref="TrackBlockFormatChanges"/>: gates ONLY the section property-revision markup
    /// (<c>w:sectPrChange</c> on the trailing body section AND on an inline in-<c>w:pPr</c> section), NOT the
    /// paragraph or table-shell variants. Defaults equal to <see cref="TrackBlockFormatChanges"/> (two-way is
    /// byte-identical). The composite merger + renderers set this TRUE while forcing
    /// <see cref="TrackBlockFormatChanges"/> FALSE, so <c>Consolidate</c> merges reviewers' section changes (B2).
    /// </summary>
    public bool TrackSectionFormatChanges { get; init; } = true;

    /// <summary>
    /// DIFF-TIME setting (Consolidate B2). Whether the two-way revision renderer appends the DOCUMENT-LEVEL
    /// trailing-<c>w:sectPr</c> FormatChanged revision (it compares the two documents' final section blocks, so
    /// it is not a per-op revision). Default true (every two-way call reports it). The composite revision
    /// renderer sets this FALSE for its PER-OP mini-scripts — otherwise a section-changing reviewer with N ops
    /// would emit the same trailing-section revision N times — and instead emits it exactly once from
    /// <c>IrCompositeScript.TrailingSectPr</c>, attributed to the section winner.
    /// </summary>
    public bool EmitTrailingSectionRevision { get; init; } = true;

    /// <summary>
    /// REVISIONS-SURFACE setting (M2.3 Task 1). Author name stamped on every <see cref="IrRevision"/>'s
    /// <see cref="IrRevision.Author"/>. Default <c>"Open-Xml-PowerTools"</c> — copied verbatim from
    /// <c>WmlComparerSettings.AuthorForRevisions</c> (Docxodus/WmlComparer.cs ~line 54) so an IR-rendered
    /// revision set is author-comparable to the shipped comparer's out of the box.
    /// </summary>
    public string AuthorForRevisions { get; init; } = "Open-Xml-PowerTools";

    /// <summary>
    /// REVISIONS-SURFACE setting (M2.3 Task 1). When true (the DEFAULT), <see cref="DateTimeForRevisions"/>
    /// is pinned to a fixed epoch (<see cref="DeterministicEpoch"/>) so two renders of the same inputs
    /// produce byte-identical revision dates.
    /// </summary>
    /// <remarks>
    /// <b>Why deterministic by default.</b> Reproducible output is a program principle (the IR/diff layer
    /// is value-equal and JSON-round-trippable end to end). <c>WmlComparerSettings.DateTimeForRevisions</c>
    /// defaults to <c>DateTime.Now.ToString("o")</c> (Docxodus/WmlComparer.cs ~line 55), a documented
    /// nondeterminism wart: the same compare run twice yields different revision dates, which breaks
    /// golden-output and round-trip determinism tests. We invert that default — pinned epoch unless the
    /// caller opts into wall-clock behavior by setting this false (and optionally supplying
    /// <see cref="DateTimeForRevisions"/> explicitly).
    /// </remarks>
    public bool Deterministic { get; init; } = true;

    /// <summary>
    /// REVISIONS-SURFACE setting (M2.3 Task 1). The ISO-8601 date string stamped on every
    /// <see cref="IrRevision.Date"/>. Defaults to:
    /// <list type="bullet">
    /// <item>the fixed <see cref="DeterministicEpoch"/> when <see cref="Deterministic"/> is true (the
    /// default), so output is reproducible;</item>
    /// <item><c>DateTime.Now.ToString("o")</c> — captured once at settings construction — when
    /// <see cref="Deterministic"/> is false, matching <c>WmlComparerSettings.DateTimeForRevisions</c>'s
    /// "o" round-trip format (Docxodus/WmlComparer.cs ~line 55).</item>
    /// </list>
    /// A caller may set this explicitly to any string; an explicit value always wins over both defaults.
    /// </summary>
    public string DateTimeForRevisions { get; init; } = DeterministicEpoch;

    /// <summary>
    /// The fixed epoch stamped on revision dates when <see cref="Deterministic"/> is true:
    /// <c>2000-01-01T00:00:00Z</c>. Chosen as a recognizable, timezone-explicit ISO-8601 instant well
    /// before any plausible document date, so a pinned revision date is obviously synthetic.
    /// </summary>
    public const string DeterministicEpoch = "2000-01-01T00:00:00Z";

    /// <summary>
    /// The default separator set, copied verbatim from <c>WmlComparerSettings.WordSeparators</c>
    /// (Docxodus/WmlComparer.cs ~line 123). The comparer's literal includes duplicate CJK entries;
    /// the set folds them.
    /// </summary>
    public static readonly ImmutableHashSet<char> DefaultWordSeparators = ImmutableHashSet.Create(
        ' ', '-', ')', '(', ';', ',', '（', '）', '，', '、', '、', '，', '；', '。', '：', '的');

    /// <summary>
    /// Build a settings instance for NONDETERMINISTIC revision rendering: <see cref="Deterministic"/>
    /// false with <see cref="DateTimeForRevisions"/> captured once from <c>DateTime.Now</c> in the "o"
    /// format (so all revisions in one render share a single timestamp, mirroring a single
    /// <c>WmlComparerSettings</c> instance). Other settings keep their defaults; chain <c>with</c> to
    /// override. The wall-clock read happens here, NOT lazily per revision, so a render is internally
    /// consistent.
    /// </summary>
    public static IrDiffSettings WithWallClockRevisionDate() => new()
    {
        Deterministic = false,
        DateTimeForRevisions = System.DateTime.Now.ToString("o", CultureInfo.InvariantCulture),
    };
}
