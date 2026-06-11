#nullable enable

using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// How <see cref="IrBlockAligner"/> classified a body-block pairing.
/// </summary>
/// <remarks>
/// <para>
/// <b><see cref="MovedModified"/> is reserved for M2.2 and is NEVER produced in M2.1.</b> M2.1 move
/// detection is exact-content-hash only: a block can only be classified <see cref="Moved"/> when its
/// <c>ContentHash</c> matches an off-spine anchor exactly, so a "moved AND its text was edited" block
/// cannot be recognized as a move at all in M2.1 (it falls out as a Deleted+Inserted, or as a
/// positional Modified inside a gap). Similarity-based move detection that would surface
/// <see cref="MovedModified"/> is M2.2 territory (intra-block token diff + fuzzy move matching). The
/// kind is declared now so the enum is stable across the M2.1→M2.2 surface bump.
/// </para>
/// <para>
/// Conversely a block whose content hash matches an off-spine anchor but whose
/// <c>FormatFingerprint</c> differs (moved AND reformatted, same text) IS still classified plain
/// <see cref="Moved"/> in M2.1 — format equality does not refine the move kind here.
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

    /// <summary>Exact-content match that is off the in-order spine (relocated). <c>ContentHash</c> equal; format may differ.</summary>
    Moved,

    /// <summary>RESERVED for M2.2 (fuzzy moved+edited). Never produced in M2.1 — see type remarks.</summary>
    MovedModified,

    /// <summary>Right-only block (no left counterpart): <c>Left</c> is null.</summary>
    Inserted,

    /// <summary>Left-only block (no right counterpart): <c>Right</c> is null.</summary>
    Deleted,
}

/// <summary>
/// One entry in an <see cref="IrBlockAlignment"/>: a classified pairing of a left and/or right body
/// block. <see cref="Inserted"/> carries a null <see cref="Left"/>; <see cref="Deleted"/> a null
/// <see cref="Right"/>; every other kind carries both.
/// </summary>
internal sealed record IrAlignedBlock(IrAlignmentKind Kind, IrBlock? Left, IrBlock? Right);

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
