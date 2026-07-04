#nullable enable

using System.Collections.Generic;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// The kind of an <see cref="IrRevision"/> (M2.3 Task 1). Mirrors the public
/// <c>WmlComparer.WmlComparerRevisionType</c> (Docxodus/WmlComparer.cs ~line 3840) one-for-one so the
/// IR revisions surface is consumer-comparable to the shipped comparer's <c>GetRevisions</c> output.
/// </summary>
internal enum IrRevisionType
{
    /// <summary>Right-only content (a <c>w:ins</c>-grade insertion).</summary>
    Inserted,

    /// <summary>Left-only content (a <c>w:del</c>-grade deletion).</summary>
    Deleted,

    /// <summary>
    /// Relocated content. Source and destination revisions share a <see cref="IrRevision.MoveGroupId"/>;
    /// <see cref="IrRevision.IsMoveSource"/> distinguishes the two.
    /// </summary>
    Moved,

    /// <summary>
    /// Content unchanged in text but changed in modeled run formatting (a <c>w:rPrChange</c>-grade
    /// change). <see cref="IrRevision.FormatChange"/> carries the old/new modeled fields + changed names.
    /// </summary>
    FormatChanged,
}

/// <summary>
/// Details of a <see cref="IrRevisionType.FormatChanged"/> revision (M2.3 Task 1): the MODELED run-format
/// fields before and after, plus the names of the fields that differ. Field naming and shape mirror
/// <c>WmlComparer.FormatChangeDetails</c> (Docxodus/WmlComparer.cs ~line 3872 —
/// <c>OldProperties</c>/<c>NewProperties</c>/<c>ChangedPropertyNames</c>) so an adapter can map one to the
/// other with a rename only.
/// </summary>
/// <remarks>
/// <para><b>Modeled-only, by design.</b> The dictionaries enumerate ONLY the modeled
/// <see cref="IrRunFormat"/> fields (Bold/Italic/Underline/Size/Color/… — never the unmodeled rPr
/// digest), keyed by WmlComparer-friendly property names (<c>bold</c>, <c>italic</c>, <c>fontSize</c>,
/// <c>color</c>, …; see <c>WmlComparer.GetFriendlyPropertyName</c>). This is exactly what a
/// <c>w:rPrChange</c>-grade report can describe; an unmodeled-only flip is not a format change here, by
/// the same reasoning as <see cref="IrFormatComparison.ModeledOnly"/>.</para>
/// <para><b>Convention.</b> A field present on one side and absent on the other (e.g. bold added) appears
/// in only that side's dictionary and in <see cref="ChangedPropertyNames"/>. Both dictionaries omit a
/// field whose value is null on a side (no rPr child), mirroring WmlComparer's "absent key" convention.</para>
/// </remarks>
internal sealed record IrFormatChangeDetails(
    IReadOnlyDictionary<string, string> OldProperties,
    IReadOnlyDictionary<string, string> NewProperties,
    IReadOnlyList<string> ChangedPropertyNames,
    IrFormatChangeScope Scope = IrFormatChangeScope.Run);

/// <summary>
/// Which property container a <see cref="IrRevisionType.FormatChanged"/> revision describes
/// (block-format-change family, 2026-07-03). <see cref="Run"/> is the pre-campaign rPr-grade report;
/// <see cref="Paragraph"/> reports a pPr delta (<c>w:pPrChange</c>-grade); the table and section members
/// arrive with their phases. Mirrors the public <c>DocxDiffFormatChangeScope</c> member-for-member.
/// </summary>
internal enum IrFormatChangeScope { Run, Paragraph, TableCell, TableRow, Table, Section }

/// <summary>
/// One consumer-facing revision rendered from an <see cref="IrEditScript"/> by
/// <see cref="IrRevisionRenderer"/> (M2.3 Task 1). Mirrors the consumer-relevant shape of the public
/// <c>WmlComparer.WmlComparerRevision</c> (Docxodus/WmlComparer.cs ~line 3879):
/// <see cref="Type"/>↔<c>RevisionType</c>, <see cref="Text"/>, <see cref="Author"/>, <see cref="Date"/>,
/// <see cref="MoveGroupId"/>, <see cref="IsMoveSource"/>, <see cref="FormatChange"/>.
/// </summary>
/// <remarks>
/// <para><b>Extension over WmlComparerRevision.</b> <see cref="LeftAnchor"/>/<see cref="RightAnchor"/> are
/// NOT present on <c>WmlComparerRevision</c> — they are an IR-engine addition carrying the block anchors
/// (<c>kind:scope:unid</c>) the revision derives from, resolvable in the originating document's
/// <see cref="IrDocument.AnchorIndex"/>. They are valuable for consumers that want to locate a revision in
/// the document model (review UIs, blame). An adapter targeting the exact WmlComparerRevision surface
/// simply ignores them. <c>WmlComparerRevision</c>'s OOXML-specific members
/// (<c>ContentXElement</c>/<c>RevisionXElement</c>/<c>PartUri</c>/<c>PartContentType</c>) are deliberately
/// NOT mirrored: this surface is built from the IR edit script, not produced OOXML markup (that is M2.4).</para>
/// <para><b>Anchor presence by <see cref="Type"/>.</b> Each type's PRIMARY anchor is ALWAYS present; the
/// opposite anchor MAY also be present for a token-level revision. Inserted → <see cref="RightAnchor"/>
/// always (and <see cref="LeftAnchor"/> too when token-level in a Modified block); Deleted →
/// <see cref="LeftAnchor"/> always (and <see cref="RightAnchor"/> too when token-level); FormatChanged →
/// BOTH (content-equal block pair); Moved → EXCLUSIVE: source = <see cref="LeftAnchor"/> only, destination =
/// <see cref="RightAnchor"/> only. A nested token-op revision inside a Modified/MoveModify block carries BOTH
/// of the enclosing block's anchors (that block exists on both sides); a whole-block ins/del carries only its
/// primary anchor.</para>
/// </remarks>
internal sealed record IrRevision(
    IrRevisionType Type,
    string Text,
    string Author,
    string Date,
    int? MoveGroupId = null,
    bool? IsMoveSource = null,
    IrFormatChangeDetails? FormatChange = null,
    string? LeftAnchor = null,
    string? RightAnchor = null);
