#nullable enable

using System.Collections.Immutable;
using System.Globalization;

namespace Docxodus.Ir.Diff;

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
    /// <c>Deleted</c>+<c>Inserted</c>. Default 0.5.
    /// </summary>
    /// <remarks>
    /// <b>Why 0.5.</b> Below half token-overlap, treating two blocks as "the same block edited" produces
    /// a WORSE edit script than a clean Insert+Delete: a Modified pairing forces a token diff whose
    /// shared run is a minority of the content, so the diff is mostly Delete-then-Insert anyway but now
    /// carries the false claim that the destination paragraph is a revision of that particular source
    /// paragraph (misleading review UIs, bad blame). At ≥0.5 the majority of tokens are shared, so the
    /// "edited in place" framing is the faithful one. 0.5 is the in-gap floor; cross-gap MOVES demand the
    /// stricter <see cref="MoveSimilarityThreshold"/> because relocating-and-editing is a stronger claim
    /// than editing in place.
    /// </remarks>
    public double BlockSimilarityThreshold { get; init; } = 0.5;

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
    /// The default separator set, copied verbatim from <c>WmlComparerSettings.WordSeparators</c>
    /// (Docxodus/WmlComparer.cs ~line 123). The comparer's literal includes duplicate CJK entries;
    /// the set folds them.
    /// </summary>
    public static readonly ImmutableHashSet<char> DefaultWordSeparators = ImmutableHashSet.Create(
        ' ', '-', ')', '(', ';', ',', '（', '）', '，', '、', '、', '，', '；', '。', '：', '的');
}
