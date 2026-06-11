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

/// <summary>A character range a comment targets within a given block.</summary>
internal sealed record IrCommentTarget(IrAnchor BlockAnchor, int StartChar, int EndChar);
