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
/// accepted view AND preserves the inputs' own revision markup, because that is what Word's Compare
/// does and what every shipping surface wants. The raw <see cref="DocxDiff"/> API keeps both flags at
/// their opt-in defaults for callers who want the engine's unopinionated behavior. Calling
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
    /// <paramref name="settings"/>. Byte-identical inputs return a detached exact clone without
    /// reserialization — a no-op must not rewrite a valid package merely because it passed through
    /// the comparison API.
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

        // An exact same-package comparison has no revisions to produce. More importantly, a no-op must
        // not silently rewrite a valid Strict package or discard unrelated existing revision markup.
        // Return a detached clone so the result stays safe for callers to mutate/save independently.
        if (CanReturnExactNoOp(left, right))
            return new WmlDocument(left);

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
    /// comparison, and was removed with the engine that motivated it. Tracked as issue #642.</para>
    /// </summary>
    internal static bool CanReturnExactNoOp(WmlDocument left, WmlDocument right) =>
        HasIdenticalPackageBytes(left, right);

    /// <summary>
    /// Layer the front-door input-revision policy onto the caller's settings.
    ///
    /// <para>The DIFF must run over the accepted view — otherwise revision-bearing inputs diff their
    /// raw surface and emit whole-document churn — and Word's Compare additionally PRESERVES the
    /// inputs' own markup in its output (original author/date rides through, verified against
    /// Word-oracle outputs). Preserve WINS over the pre-accept by precedence: matching still happens on
    /// the accepted view (the IR read accepts regardless), the byte-level flatten is skipped, and
    /// equal/inserted blocks carry the input's markup through. See
    /// <see cref="DocxDiffSettings.PreserveInputRevisions"/> for the one-sided round-trip contract this
    /// implies (accept ≡ right holds; reject ≠ left where foreign markup exists — exactly Word).</para>
    ///
    /// <para>This is unconditional, exactly as the pre-v11 settings mapping was: it is the front door's
    /// defining behavior, not a default a caller is expected to opt into. A caller who wants the
    /// engine's opt-in defaults should call <see cref="DocxDiff.Compare"/> directly.</para>
    /// </summary>
    internal static DocxDiffSettings ApplyFrontDoorRevisionPolicy(DocxDiffSettings? settings)
    {
        var result = settings is null ? new DocxDiffSettings() : settings.Clone();
        result.PreAcceptInputRevisions = true;
        result.PreserveInputRevisions = true;
        return result;
    }
}
