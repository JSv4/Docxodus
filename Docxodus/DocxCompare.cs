#nullable enable

namespace Docxodus;

/// <summary>
/// The single, shared front door for two-document DOCX comparison, owning the ONE
/// <see cref="WmlComparer"/>-vs-<see cref="DocxDiff"/> branch in the codebase. The CLI
/// (<c>tools/redline</c>), the WASM bridge (<c>DocumentComparer</c>), and — transitively — the npm
/// wrappers all route their primary "compare these two documents → redlined DOCX" call through
/// <see cref="Compare"/>, so the engine choice lives in exactly one place (mirroring the single-owner
/// facade pattern used by <see cref="Internal.DocxDiffOps"/> / <c>HtmlConversionOps</c>).
///
/// <para>Both engines emit native tracked-changes markup (<c>w:ins</c>/<c>w:del</c>/<c>w:moveFrom</c>/…),
/// so callers can count revisions on the returned document via
/// <see cref="WmlComparer.GetRevisions(WmlDocument, WmlComparerSettings)"/> regardless of which engine
/// produced it.</para>
///
/// <para>The parameter is a <see cref="WmlComparerSettings"/> — the settings shape all three surfaces
/// already build — so wiring the selector in only adds an <see cref="ComparisonEngine"/> argument; the
/// surfaces' settings-construction code is untouched. On the <see cref="ComparisonEngine.DocxDiff"/>
/// branch the common option set is mapped to <see cref="DocxDiffSettings"/> by
/// <see cref="ToDocxDiffSettings"/> (WmlComparer-only knobs are dropped — see that method).</para>
/// </summary>
public static class DocxCompare
{
    /// <summary>
    /// Compare <paramref name="left"/> against <paramref name="right"/> with the selected
    /// <paramref name="engine"/> and return the redlined document. <see cref="ComparisonEngine.WmlComparer"/>
    /// (the default) is a straight delegate to <see cref="WmlComparer.Compare"/> — byte-for-byte today's
    /// behavior; <see cref="ComparisonEngine.DocxDiff"/> routes to <see cref="DocxDiff.Compare"/> with the
    /// mapped settings.
    /// </summary>
    /// <param name="left">The earlier / original document.</param>
    /// <param name="right">The later / revised document.</param>
    /// <param name="engine">Which comparison engine to use.</param>
    /// <param name="settings">Comparison settings (the same <see cref="WmlComparerSettings"/> shape both engines accept via mapping).</param>
    public static WmlDocument Compare(
        WmlDocument left,
        WmlDocument right,
        ComparisonEngine engine,
        WmlComparerSettings settings)
        => engine == ComparisonEngine.DocxDiff
            ? DocxDiff.Compare(left, right, ToDocxDiffSettings(settings))
            : WmlComparer.Compare(left, right, settings);

    /// <summary>
    /// Map the option set shared by both engines from <see cref="WmlComparerSettings"/> onto a fresh
    /// <see cref="DocxDiffSettings"/>. WmlComparer-only knobs are intentionally dropped: <c>DetailThreshold</c>
    /// (an LCS-granularity knob with no IR equivalent), <c>SimplifyMoveMarkup</c> (DocxDiff renders moves
    /// natively), and <c>DetectFormatChanges</c> (DocxDiff uses <see cref="DocxDiffSettings.FormatComparison"/>,
    /// left at its default). DocxDiff-specific settings keep their defaults. An explicit
    /// <c>DateTimeForRevisions</c> is carried through and wins over <see cref="DocxDiffSettings.Deterministic"/>.
    /// </summary>
    internal static DocxDiffSettings ToDocxDiffSettings(WmlComparerSettings settings) => new()
    {
        AuthorForRevisions = settings.AuthorForRevisions,
        DateTimeForRevisions = settings.DateTimeForRevisions,
        CaseInsensitive = settings.CaseInsensitive,
        ConflateBreakingAndNonbreakingSpaces = settings.ConflateBreakingAndNonbreakingSpaces,
        DetectMoves = settings.DetectMoves,
        MoveSimilarityThreshold = settings.MoveSimilarityThreshold,
        MoveMinimumWordCount = settings.MoveMinimumWordCount,
        // Engine equivalence: WmlComparer.Compare accepts BOTH inputs' pre-existing tracked
        // revisions before comparing (Microsoft Word's compare does the same), so the DocxDiff
        // branch must too — otherwise inputs that already carry revisions diff their raw
        // revision-bearing surface and emit whole-document churn. The raw DocxDiff API keeps
        // its opt-in default; this is the WmlComparer-parity surface.
        PreAcceptInputRevisions = true,
    };

    /// <summary>
    /// Parse a case-insensitive engine name — <c>wmlcomparer</c> or <c>docxdiff</c> — as accepted by the
    /// redline CLI's <c>--engine=</c> flag. Surrounding whitespace is trimmed. Returns <c>false</c> for an
    /// unrecognized value, in which case <paramref name="engine"/> is set to the default
    /// <see cref="ComparisonEngine.WmlComparer"/>.
    /// </summary>
    public static bool TryParseEngine(string? value, out ComparisonEngine engine)
    {
        switch (value?.Trim().ToLowerInvariant())
        {
            case "wmlcomparer":
                engine = ComparisonEngine.WmlComparer;
                return true;
            case "docxdiff":
                engine = ComparisonEngine.DocxDiff;
                return true;
            default:
                engine = ComparisonEngine.WmlComparer;
                return false;
        }
    }
}
