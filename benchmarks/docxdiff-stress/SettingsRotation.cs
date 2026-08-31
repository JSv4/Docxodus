// Round-robin over the DocxDiffSettings surface (issue #624, item 3).
//
// The matrix used to sample four settings combinations out of a surface of twenty properties, and
// the #616 regression sat behind one of the sixteen that were never set: the harness DOES digest a
// thrown exception, so had ThrowOnCompatibilityWarning=true appeared anywhere in the matrix, the
// existing code would have caught it. That is the more dangerous of the two failure modes -- the
// channel IS observed, but no input ever puts the code in a state where it differs, so the harness
// looks like it has you covered.
//
// A cross-product of twenty properties explodes. A rotation does not: document i also runs with
// setting (i mod N) varied from its default, so every setting is exercised on roughly
// documents/N of the corpus for ONE extra comparison per document rather than twenty.

using System.Globalization;
using Docxodus;

namespace Docxodus.Stress;

internal static class SettingsRotation
{
    /// <summary>
    /// Each entry names one non-default setting and builds a settings object carrying only that
    /// deviation, so a mismatch on a rotated observation names the setting responsible.
    /// <para>
    /// <c>Deterministic = false</c> is deliberately NOT here. It is a real setting and it is not
    /// covered, but by construction it makes two comparisons of the same inputs emit different
    /// bytes — which is what the differential exists to detect, so including it would make the
    /// harness disagree with itself on every run and drown genuine regressions. A setting whose
    /// whole purpose is to defeat reproducibility cannot be covered by a reproducibility check.
    /// </para>
    /// </summary>
    private static readonly (string Name, Func<DocxDiffSettings> Build)[] Variations =
    {
        // The one that would have caught #616: it turns a compatibility report into a thrown
        // exception, which the harness records as an outcome like any other.
        ("throw-on-compat", () => new DocxDiffSettings { ThrowOnCompatibilityWarning = true }),
        ("case-insensitive", () => new DocxDiffSettings { CaseInsensitive = true }),
        ("no-space-conflation", () => new DocxDiffSettings { ConflateBreakingAndNonbreakingSpaces = false }),
        ("no-moves", () => new DocxDiffSettings { DetectMoves = false }),
        ("loose-moves", () => new DocxDiffSettings { MoveSimilarityThreshold = 0.5, MoveMinimumWordCount = 1 }),
        ("coarse-revisions", () => new DocxDiffSettings
        {
            RevisionGranularity = DocxDiffRevisionGranularity.WmlComparerCompatible,
        }),
        ("full-format", () => new DocxDiffSettings { FormatComparison = DocxDiffFormatComparison.Full }),
        ("normalize-authors", () => new DocxDiffSettings { NormalizeRevisionAuthors = true }),
        ("no-header-footer", () => new DocxDiffSettings { CompareHeadersFooters = false }),
        ("no-block-format", () => new DocxDiffSettings { TrackBlockFormatChanges = false }),
        ("no-cross-paragraph", () => new DocxDiffSettings { CrossParagraphTokenDiff = false }),
        ("author-and-date", () => new DocxDiffSettings
        {
            AuthorForRevisions = "Rotation",
            DateTimeForRevisions = "2020-01-01T00:00:00Z",
        }),
        ("invariant-culture", () => new DocxDiffSettings { Culture = CultureInfo.InvariantCulture }),
    };

    public static int Count => Variations.Length;

    public static (string Name, DocxDiffSettings Settings) For(int documentIndex)
    {
        var (name, build) = Variations[((documentIndex % Count) + Count) % Count];
        return (name, build());
    }
}
