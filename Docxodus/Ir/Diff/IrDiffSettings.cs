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
    /// The default separator set, copied verbatim from <c>WmlComparerSettings.WordSeparators</c>
    /// (Docxodus/WmlComparer.cs ~line 123). The comparer's literal includes duplicate CJK entries;
    /// the set folds them.
    /// </summary>
    public static readonly ImmutableHashSet<char> DefaultWordSeparators = ImmutableHashSet.Create(
        ' ', '-', ')', '(', ';', ',', '（', '）', '，', '、', '、', '，', '；', '。', '：', '的');
}
