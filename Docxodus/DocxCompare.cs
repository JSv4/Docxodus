#nullable enable

using System;
using System.Linq;

namespace Docxodus;

/// <summary>
/// The shared front door for two-document DOCX comparison. The CLI (<c>tools/redline</c>), the WASM
/// bridge (<c>DocumentComparer</c>), and — transitively — the npm wrappers all route their
/// "compare these two documents → redlined DOCX" call through <see cref="Compare"/>, so the
/// comparison POLICY lives in exactly one place (mirroring the single-owner facade pattern used by
/// <see cref="Internal.DocxDiffOps"/> / <c>HtmlConversionOps</c>).
///
/// <para><b>Front door vs. raw engine.</b> <see cref="DocxDiff.Compare"/> is the engine; this is the
/// product. The difference is the input-revision policy: the front door always compares on the
/// ACCEPTED view, because that is what Word's Compare does. The raw <see cref="DocxDiff"/> API keeps
/// that flag at its opt-in default for callers who want the engine's unopinionated behavior. Calling
/// <see cref="DocxDiff.Compare"/> directly with a fresh <see cref="DocxDiffSettings"/> is therefore
/// NOT equivalent to calling this — see <see cref="ApplyFrontDoorRevisionPolicy"/>.</para>
///
/// <para>Before v11.0.0 this type also owned the one <c>WmlComparer</c>-vs-<c>DocxDiff</c> engine
/// branch in the codebase. The legacy engine is gone; what remains is the policy it used to share.</para>
/// </summary>
public static class DocxCompare
{
    /// <summary>
    /// Compare <paramref name="left"/> against <paramref name="right"/> and return the redlined
    /// document, with the front-door input-revision policy applied on top of
    /// <paramref name="settings"/>. Byte-identical TRANSITIONAL inputs return a detached exact clone
    /// without reserialization — a no-op must not rewrite a valid package merely because it passed
    /// through the comparison API; byte-identical STRICT inputs are normalized to transitional.
    /// </summary>
    /// <param name="left">The earlier / original document.</param>
    /// <param name="right">The later / revised document.</param>
    /// <param name="settings">Comparison settings; <c>null</c> takes <see cref="DocxDiffSettings"/> defaults.</param>
    public static WmlDocument Compare(
        WmlDocument left,
        WmlDocument right,
        DocxDiffSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(right);

        // An exact same-package comparison has no revisions to produce; return a detached clone so
        // the result remains safe for callers to mutate/save independently of the input. A STRICT
        // package is still normalized to transitional on the way out — Word converts on open no
        // matter what the compare finds, and strict bytes break downstream consumers (LibreOffice
        // renders them poorly, python-docx rejects them). Transitional inputs stay byte-identical.
        if (CanReturnExactNoOp(left, right))
        {
            var normalized = StrictOoxmlNormalizer.NormalizeToTransitional(left);
            return ReferenceEquals(normalized, left) ? new WmlDocument(left) : normalized;
        }

        return DocxDiff.Compare(left, right, ApplyFrontDoorRevisionPolicy(settings));
    }

    /// <summary>Whether two documents are the exact same package bytes, not merely semantically equal.</summary>
    internal static bool HasIdenticalPackageBytes(WmlDocument left, WmlDocument right) =>
        left.DocumentByteArray.AsSpan().SequenceEqual(right.DocumentByteArray);

    /// <summary>
    /// Whether an exact-package comparison can skip the engine. Byte equality is the whole test.
    ///
    /// <para>Through v10 this additionally refused the shortcut for documents carrying tracked-revision
    /// wrappers inside an Office Math run — schema-invalid markup that <c>WmlComparer</c>'s
    /// preprocessing repaired as a side effect, so routing through the engine was strictly better than
    /// cloning. <see cref="DocxDiff"/> performs no such repair: on that input it returns the source
    /// bytes unchanged, with the same validation error. The guard therefore bought nothing but a full
    /// comparison, and was removed with the engine that motivated it. Preserving that invalid input
    /// rather than silently rewriting it is a decided position, not an oversight — see "Office Math:
    /// revision wrappers nested inside <c>m:r</c>" in <c>docs/ooxml_corner_cases.md</c> (issue #642).</para>
    /// </summary>
    internal static bool CanReturnExactNoOp(WmlDocument left, WmlDocument right) =>
        HasIdenticalPackageBytes(left, right);

    /// <summary>
    /// Layer the front-door input-revision policy onto the caller's settings.
    ///
    /// <para>The DIFF runs over the ACCEPTED view. Word's Compare dialog says it outright — "Word will
    /// treat them as accepted" — and its outputs confirm it: text an input had struck through is absent
    /// from the compare result entirely, the surviving text is re-detected as the compare author's own
    /// insertions, and the output collapses to a single revision author. Without the pre-accept,
    /// revision-bearing inputs diff their raw surface and emit whole-document churn.</para>
    ///
    /// <para>Earlier releases ALSO preserved the inputs' own markup here — Word's COMBINE behavior,
    /// decoded from a batch of oracle documents that turned out to be Combine-shaped. Compare is the
    /// operation this surface models, so only the pre-accept flatten runs. Callers who want the inputs'
    /// revisions carried through use the raw <see cref="DocxDiff"/> API with
    /// <see cref="DocxDiffSettings.PreserveInputRevisions"/>.</para>
    ///
    /// <para>This is unconditional, exactly as the pre-v11 settings mapping was: it is the front door's
    /// defining behavior, not a default a caller is expected to opt into. A caller who wants the
    /// engine's opt-in defaults should call <see cref="DocxDiff.Compare"/> directly.</para>
    /// </summary>
    internal static DocxDiffSettings ApplyFrontDoorRevisionPolicy(DocxDiffSettings? settings)
    {
        var result = settings is null ? new DocxDiffSettings() : settings.Clone();
        result.PreAcceptInputRevisions = true;
        return result;
    }
}
