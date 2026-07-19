#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

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
    /// (the default) delegates directly for transitional documents and normalizes Word Strict inputs first,
    /// matching Word's open behavior; byte-identical inputs instead return a detached exact clone without
    /// normalization or reserialization when the legacy comparer need not repair malformed math revision
    /// markup; <see cref="ComparisonEngine.DocxDiff"/> routes to
    /// <see cref="DocxDiff.Compare"/> with the mapped settings.
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
    {
        // An exact same-package comparison has no revisions to produce. More importantly, a no-op
        // must not silently rewrite a valid Strict package or discard unrelated existing revision
        // markup merely because it passed through the comparison API. Return a detached clone so the
        // result remains safe for callers to mutate/save independently of the input.
        if (CanReturnExactNoOp(left, right))
            return new WmlDocument(left);

        return engine == ComparisonEngine.DocxDiff
            ? DocxDiff.Compare(left, right, ToDocxDiffSettings(settings))
            : WmlComparer.Compare(
                StrictOoxmlNormalizer.NormalizeToTransitional(left),
                StrictOoxmlNormalizer.NormalizeToTransitional(right),
                settings);
    }

    /// <summary>Whether two documents are the exact same package bytes, not merely semantically equal.</summary>
    internal static bool HasIdenticalPackageBytes(WmlDocument left, WmlDocument right) =>
        left.DocumentByteArray.AsSpan().SequenceEqual(right.DocumentByteArray);

    /// <summary>
    /// Whether an exact-package comparison can skip the legacy comparer safely. A small set of old Word
    /// documents place tracked-revision wrappers directly inside an Office Math run, which is schema-invalid.
    /// The legacy preprocessing path repairs that shape; returning an exact clone would retain the invalid
    /// markup and violate the comparer’s long-standing valid-output contract.
    /// </summary>
    internal static bool CanReturnExactNoOp(WmlDocument left, WmlDocument right) =>
        HasIdenticalPackageBytes(left, right) && !HasRevisionMarkupInsideMathRun(left);

    private static readonly HashSet<XName> TrackedRevisionNames = new(RevisionProcessor.TrackedRevisionsElements);

    private static bool HasRevisionMarkupInsideMathRun(WmlDocument document)
    {
        try
        {
            using var streamDoc = new OpenXmlMemoryStreamDocument(document);
            using var wordDoc = streamDoc.GetWordprocessingDocument();
            var mainXDoc = wordDoc.MainDocumentPart?.GetXDocument();
            return mainXDoc?.Descendants(M.r).Any(mathRun =>
                mathRun.Descendants().Any(element => TrackedRevisionNames.Contains(element.Name))) ?? false;
        }
        catch (Exception)
        {
            // A damaged package cannot safely take an exact-package shortcut. Let the established comparer
            // path report or repair it with its normal diagnostics instead.
            return true;
        }
    }

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
        // Input-revision policy on the selector path: the DIFF must run over the accepted view (as
        // WmlComparer does internally — otherwise revision-bearing inputs diff their raw surface and
        // emit whole-document churn), and Word's Compare additionally PRESERVES the inputs' own markup
        // in its output (original author/date rides through, verified against Word-oracle outputs).
        // Preserve WINS over the pre-accept by precedence: matching still happens on the accepted view
        // (the IR read accepts regardless), the byte-level flatten is skipped, and equal/inserted blocks
        // carry the input's markup through. See DocxDiffSettings.PreserveInputRevisions for the
        // one-sided round-trip contract this implies (accept ≡ right holds; reject ≠ left where foreign
        // markup exists — exactly Word). The raw DocxDiff API keeps both flags' opt-in defaults.
        PreAcceptInputRevisions = true,
        PreserveInputRevisions = true,
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
