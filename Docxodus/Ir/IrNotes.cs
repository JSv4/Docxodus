#nullable enable

using System.Collections.Generic;

namespace Docxodus.Ir;

/// <summary>
/// Footnote or endnote store: a map from note id (`w:id`) to the <see cref="IrScope"/> holding
/// that note's blocks, plus a parallel map from note id to that note element's own
/// <c>pt:Unid</c> (<see cref="NoteUnids"/>).
/// </summary>
/// <remarks>
/// <para>The backing dictionaries keep reference equality (they are derived indexes, not modeled
/// content); node-for-node value equality of an <see cref="IrDocument"/> is defined over the scopes
/// it contains, not over these dictionaries.</para>
/// <para><b>Why <see cref="NoteUnids"/>.</b> The markdown projection's note LABEL
/// (<c>[^fn-…]</c>/<c>[^en-…]</c>) is derived from the <c>w:footnote</c>/<c>w:endnote</c> element's
/// own <c>pt:Unid</c> — NOT the first block's anchor unid — and a body <c>IrNoteRef</c> carries only
/// the <c>w:id</c>. The emitter resolves <c>w:id → note Unid</c> through this map so it can render
/// the label the oracle does. The note element's Unid is otherwise not addressable (a note is a
/// scope, not a block), so this is the additive fact the projection needs.</para>
/// </remarks>
internal sealed record IrNoteStore(
    IReadOnlyDictionary<string, IrScope> Notes,
    IReadOnlyDictionary<string, string> NoteUnids)
{
    public static readonly IrNoteStore Empty =
        new(new Dictionary<string, IrScope>(), new Dictionary<string, string>());
}

/// <summary>The set of document comments, each modeled as an <see cref="IrComment"/>.</summary>
/// <remarks><see cref="PartUri"/> is the comments part's URI, populated in BOTH retention modes
/// (the comment scope's analogue of <see cref="IrScope.PartUri"/>) so the markdown emitter can label
/// <c>{#cmt:…}</c> anchors without depending on <see cref="IrDocument.Sources"/>, which is empty when
/// <see cref="IrReaderOptions.RetainSources"/> is off.</remarks>
internal sealed record IrCommentStore(IrNodeList<IrComment> Comments, System.Uri? PartUri = null)
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
/// <para><b>Char-offset rule.</b> <see cref="StartChar"/>/<see cref="EndChar"/> count the characters
/// of every <c>IrTextRun</c> the block emits, measured at emission time. This <em>includes</em> text
/// that lives inside a field's cached result (which is emitted as ordinary <c>IrTextRun</c>s) and
/// <em>excludes</em> tabs, breaks, images, note references, and opaque inlines (each contributes 0).
/// This is the simplest rule that is stable under the N5 run-coalescing pass (coalescing never
/// changes a block's total emitted text length).</para>
/// <para><b>Cross-block ranges.</b> A comment range that spans multiple blocks produces one
/// <see cref="IrCommentTarget"/> per touched block: the first runs from its start offset to that
/// block's end, intermediate blocks run from 0 to their end, and the last runs from 0 to the close
/// offset (spec §12 open-question #2).</para>
/// <para>A <c>commentReference</c> for a comment that has no ranges records a single zero-length
/// target (<see cref="StartChar"/> == <see cref="EndChar"/>) at the reference's offset.</para>
/// </remarks>
internal sealed record IrCommentTarget(IrAnchor BlockAnchor, int StartChar, int EndChar);
