#nullable enable

using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// How <see cref="IrBlockAligner"/> classified a body-block pairing.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="MovedModified"/> covers both fuzzy lexical move-and-edit matches and content-equal paragraph
/// moves with a modeled format delta. In either form the destination carries a token diff: lexical changes
/// become insert/delete spans and formatting changes become FormatChanged spans with native property history.
/// Structural blocks remain plain moves until they have an equivalent in-move format projection.
/// </para>
/// </remarks>
internal enum IrAlignmentKind
{
    /// <summary>Same block: <c>ContentHash</c> AND <c>FormatFingerprint</c> equal, in document order.</summary>
    Unchanged,

    /// <summary>Same text, different formatting: <c>ContentHash</c> equal, <c>FormatFingerprint</c> differs, in order.</summary>
    FormatOnly,

    /// <summary>Both sides present but neither hash-paired: an in-gap positional pairing whose token diff M2.2 runs.</summary>
    Modified,

    /// <summary>Relocated block without a projected in-move token change.</summary>
    Moved,

    /// <summary>Relocated paragraph with lexical and/or modeled formatting changes.</summary>
    MovedModified,

    /// <summary>Right-only block (no left counterpart): <c>Left</c> is null.</summary>
    Inserted,

    /// <summary>Left-only block (no right counterpart): <c>Right</c> is null.</summary>
    Deleted,

    /// <summary>One left paragraph split across N≥2 adjacent right paragraphs (M2.6). <c>Left</c> set,
    /// <c>Right</c> null, <see cref="IrAlignedBlock.MultiBlocks"/> = the N right blocks in right order.
    /// Emitted at the FIRST member right block's position; the other members get no entry of their own.</summary>
    Split,

    /// <summary>N≥2 adjacent left paragraphs merged into one right paragraph (M2.6). <c>Right</c> set,
    /// <c>Left</c> null, <see cref="IrAlignedBlock.MultiBlocks"/> = the N left blocks in left order.</summary>
    Merge,
}

/// <summary>
/// One entry in an <see cref="IrBlockAlignment"/>: a classified pairing of a left and/or right body
/// block. <see cref="IrAlignmentKind.Inserted"/> carries a null <see cref="Left"/>;
/// <see cref="IrAlignmentKind.Deleted"/> a null <see cref="Right"/>; every other 1:1 kind carries both.
/// <see cref="IrAlignmentKind.Split"/> carries a non-null <see cref="Left"/>, a null <see cref="Right"/>,
/// and <see cref="MultiBlocks"/> = the N≥2 right blocks in right order.
/// <see cref="IrAlignmentKind.Merge"/> carries a null <see cref="Left"/>, a non-null <see cref="Right"/>,
/// and <see cref="MultiBlocks"/> = the N≥2 left blocks in left order.
/// <para><see cref="BodyFullRewriteGroupId"/> is set only on the two standalone entries of one
/// body-level 1×1 full-lexical-rewrite gap. It is renderer-projection provenance, not an alignment
/// kind: its right Inserted and left Deleted entries share one positive id so the Word-shaped renderer
/// can retain their separate paragraph marks. Nested scopes are deliberately never marked.</para>
/// </summary>
internal sealed record IrAlignedBlock(
    IrAlignmentKind Kind, IrBlock? Left, IrBlock? Right,
    IrNodeList<IrBlock>? MultiBlocks = null,
    int? BodyFullRewriteGroupId = null);

/// <summary>
/// The result of aligning two documents' body block lists: a flat, document-ordered sequence of
/// classified <see cref="IrAlignedBlock"/> entries.
/// </summary>
/// <remarks>
/// <para><b>Entry order (pinned by the invariants tests).</b> Entries are emitted in RIGHT-document
/// order: walk the right body blocks in their original order, emitting each one's entry
/// (<see cref="IrAlignmentKind.Unchanged"/>/<see cref="IrAlignmentKind.FormatOnly"/>/
/// <see cref="IrAlignmentKind.Modified"/>/<see cref="IrAlignmentKind.Moved"/>/
/// <see cref="IrAlignmentKind.Inserted"/>) at its right position. <see cref="IrAlignmentKind.Deleted"/>
/// entries (left-only) are interleaved using the standard unified-diff <em>left-anchored</em>
/// convention: a deleted left block is emitted immediately after the entry of the nearest paired
/// left block that precedes it on the LEFT side (and deletions that precede every paired left block
/// are emitted at the very front, in left order). This makes the sequence read as a unified diff and
/// is fully deterministic.</para>
/// <para>Tables, section breaks, and opaque blocks align as WHOLE units in M2.1 — a table whose only
/// change is a cell-text edit surfaces as one <see cref="IrAlignmentKind.Modified"/> table entry;
/// row/cell-granular alignment is M2.2+.</para>
/// </remarks>
internal sealed record IrBlockAlignment(IrNodeList<IrAlignedBlock> Entries);
