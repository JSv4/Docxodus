#nullable enable

using System;
using System.Threading;
using Docxodus.Ir;

namespace Docxodus;

/// <summary>
/// One document, read once, reusable across comparisons (issue #617). A comparison's single
/// largest stage is the IR read, and nothing let a caller say <em>"I already read this document."</em>
/// A snapshot is that statement: hand the same one to every comparison a baseline participates in
/// and the baseline is read once instead of once per counterparty.
/// </summary>
/// <remarks>
/// <para><b>The workloads this exists for.</b> One baseline against many counterparties' markups
/// (the baseline is otherwise read N times); a version chain A→B→C→D (every interior version is
/// otherwise read twice); inspecting conflicts before consolidating (everything is otherwise read
/// twice).</para>
/// <example>
/// <code>
/// var baseline = DocxDiff.CreateSnapshot(original);
/// foreach (var candidate in candidates)
///     yield return DocxDiff
///         .CreateComparison(baseline, DocxDiff.CreateSnapshot(candidate))
///         .GetRevisions();
/// </code>
/// </example>
/// <para><b>Laziness.</b> Nothing is read at creation. The pre-accept normalization and the IR read
/// materialize on the first comparison that needs them and are then cached, so creating snapshots
/// for a batch costs nothing until the batch runs.</para>
/// <para><b>What it is valid for.</b> A snapshot is read under one input-revision policy and is
/// valid only for comparisons that share it — see <see cref="InputRevisionsAccepted"/>. Everything
/// else in <see cref="DocxDiffSettings"/> is a diff-time or render-time policy that does not reach
/// the read, so a snapshot is reusable across every other settings difference.
/// <see cref="DocxDiff.CreateComparison(DocxDiffSnapshot, DocxDiffSnapshot, DocxDiffSettings?)"/>
/// <b>rejects</b> a mismatch rather than silently reusing a snapshot that was read differently.</para>
/// <para><b>Memory.</b> This is the trade the reuse is made of: a materialized snapshot roots the
/// parsed <c>XDocument</c> of every story in the document, not merely the IR values, because the
/// markup renderer clones source elements out of it. Reckon on several times the package's XML size
/// resident per snapshot, and hold only the ones a run is actually reusing — a hundred retained
/// snapshots is a hundred parsed documents.</para>
/// <para><b>Thread-safety.</b> Memoization is
/// <see cref="LazyThreadSafetyMode.ExecutionAndPublication"/>, so one snapshot can serve several
/// comparisons concurrently; that is the point of the fan-out case.</para>
/// </remarks>
public sealed class DocxDiffSnapshot
{
    private readonly Lazy<WmlDocument> _preAccepted;
    private readonly Lazy<IrDocument> _ir;

    internal DocxDiffSnapshot(WmlDocument document, bool inputRevisionsAccepted)
    {
        Document = document;
        InputRevisionsAccepted = inputRevisionsAccepted;

        // The pre-accept is the same one DocxDiff.PreAccept performs for an unsnapshotted
        // comparison: strict-namespace normalization and mc:AlternateContent resolution always, the
        // byte-level accept-flatten only under the policy this snapshot was created for.
        _preAccepted = new(
            () => DocxDiff.PreAccept(SettingsForRead(inputRevisionsAccepted), document),
            LazyThreadSafetyMode.ExecutionAndPublication);

        // Read WITH provenance retained, exactly as DocxDiffComparison reads: the markup renderer
        // clones source elements out of this snapshot rather than re-reading the document, which is
        // what makes one read serve every product. Provenance is equality-neutral, so an edit script
        // built over this snapshot is identical to one built over a retention-off read.
        _ir = new(
            () => IrReader.Read(_preAccepted.Value, DocxDiff.RenderReadOpts),
            LazyThreadSafetyMode.ExecutionAndPublication);
    }

    /// <summary>The document this snapshot was created from, unmodified. Comparisons that need the
    /// caller's original bytes — the byte-identical shortcut, the semantic change set — read it
    /// from here.</summary>
    public WmlDocument Document { get; }

    /// <summary>
    /// Whether the read flattened the document's own tracked revisions away first, i.e. whether
    /// <c>PreAcceptInputRevisions</c> was set and <c>PreserveInputRevisions</c> was not. It is the
    /// only <see cref="DocxDiffSettings"/> value that reaches the read, and therefore the only one a
    /// snapshot has to agree with the comparison about.
    /// </summary>
    public bool InputRevisionsAccepted { get; }

    /// <summary>Whether the document has already been read. Diagnostic only: a comparison forces it
    /// on the first product that needs it.</summary>
    public bool IsMaterialized => _ir.IsValueCreated;

    internal WmlDocument PreAccepted => _preAccepted.Value;

    internal IrDocument Ir => _ir.Value;

    /// <summary>The input-revision policy a settings object asks the read for.</summary>
    internal static bool AcceptsInputRevisions(DocxDiffSettings settings) =>
        settings.PreAcceptInputRevisions && !settings.PreserveInputRevisions;

    /// <summary>A minimal settings object carrying only what <see cref="DocxDiff.PreAccept"/> reads,
    /// so a snapshot's read cannot accidentally depend on a diff-time policy.</summary>
    private static DocxDiffSettings SettingsForRead(bool inputRevisionsAccepted) =>
        new() { PreAcceptInputRevisions = inputRevisionsAccepted };
}
