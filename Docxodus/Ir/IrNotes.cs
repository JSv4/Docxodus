#nullable enable

using System.Collections.Generic;

namespace Docxodus.Ir;

/// <summary>
/// Footnote or endnote store: a map from note id (`w:id`) to the <see cref="IrScope"/> holding
/// that note's blocks.
/// </summary>
/// <remarks>
/// The backing dictionary keeps reference equality (it is a derived index, not modeled content);
/// node-for-node value equality of an <see cref="IrDocument"/> is defined over the scopes it
/// contains, not over this dictionary.
/// </remarks>
internal sealed record IrNoteStore(IReadOnlyDictionary<string, IrScope> Notes)
{
    public static readonly IrNoteStore Empty = new(new Dictionary<string, IrScope>());
}

/// <summary>The set of document comments, each modeled as an <see cref="IrComment"/>.</summary>
internal sealed record IrCommentStore(IrNodeList<IrComment> Comments)
{
    public static readonly IrCommentStore Empty = new(IrNodeList.Empty<IrComment>());
}

/// <summary>
/// A single comment: its identity anchor, authorship metadata, block content, and the spans of
/// document text it targets.
/// </summary>
internal sealed record IrComment(IrAnchor Anchor, string Author, string? Initials, string? Date,
                                 IrNodeList<IrBlock> Blocks, IrNodeList<IrCommentTarget> Targets);

/// <summary>
/// A character range a comment targets within a given block (rule N15).
/// </summary>
/// <remarks>
/// <para><b>Char-offset rule.</b> <see cref="StartChar"/>/<see cref="EndChar"/> count
/// <em>visible text characters</em> within the block — the summed lengths of the block's
/// <c>IrTextRun</c>s. Tabs, breaks, images, note references, fields, and opaque inlines all count as
/// 0. This is the simplest rule that is stable under the N5 run-coalescing pass (coalescing never
/// changes a block's total text length).</para>
/// <para><b>Cross-block ranges.</b> A comment range that spans multiple blocks produces one
/// <see cref="IrCommentTarget"/> per touched block: the first runs from its start offset to that
/// block's end, intermediate blocks run from 0 to their end, and the last runs from 0 to the close
/// offset (spec §12 open-question #2).</para>
/// <para>A <c>commentReference</c> for a comment that has no ranges records a single zero-length
/// target (<see cref="StartChar"/> == <see cref="EndChar"/>) at the reference's offset.</para>
/// </remarks>
internal sealed record IrCommentTarget(IrAnchor BlockAnchor, int StartChar, int EndChar);
