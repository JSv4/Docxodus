#nullable enable

using System;
using System.Xml.Linq;

namespace Docxodus.Ir;

/// <summary>
/// Base type for inline-level IR content (the children of a paragraph). Inlines are pure value
/// records; their equality is full structural equality (no provenance is carried at inline level
/// in M1.1). Child lists use <see cref="IrNodeList{T}"/> for value-semantic equality.
/// </summary>
/// <remarks>
/// The M1.1 reader only emits <see cref="IrTextRun"/>, <see cref="IrTab"/>, <see cref="IrBreak"/>,
/// and <see cref="IrOpaqueInline"/>. The remaining inline kinds are modeled now so the type model
/// is complete for the M1.2 reader (hyperlinks, fields, note refs, images).
/// </remarks>
internal abstract record IrInline;

/// <summary>A run of literal text with its direct run formatting.</summary>
internal sealed record IrTextRun(string Text, IrRunFormat Format) : IrInline;

/// <summary>A tab character (`w:tab`) carrying the run formatting of its containing run.</summary>
internal sealed record IrTab(IrRunFormat Format) : IrInline;

/// <summary>A break (`w:br`) of the given <paramref name="Kind"/> (line, page, or column).</summary>
internal sealed record IrBreak(IrBreakKind Kind) : IrInline;

/// <summary>
/// A hyperlink (`w:hyperlink`). Exactly one of <paramref name="Target"/> (external URI) or
/// <paramref name="InternalTarget"/> (in-document anchor) is expected to be set.
/// </summary>
internal sealed record IrHyperlink(string? Target, IrAnchor? InternalTarget, IrNodeList<IrInline> Inlines) : IrInline;

/// <summary>
/// A field (`w:fldSimple` or the run-based field machinery), modeled as its instruction string
/// plus the cached result inlines that Word last computed for it.
/// </summary>
internal sealed record IrFieldRun(string Instruction, IrNodeList<IrInline> CachedResult) : IrInline;

/// <summary>A footnote/endnote reference (`w:footnoteReference`/`w:endnoteReference`).</summary>
internal sealed record IrNoteRef(IrNoteKind Kind, string NoteId) : IrInline;

/// <summary>
/// An inline image: the image part, a hash of its bytes, EMU dimensions, alt text, and the
/// addressable <see cref="Unid"/> of the source <c>w:drawing</c> (its <c>pt:Unid</c>, captured by the
/// reader; null when the drawing carried none). The Unid is the IR's <c>img</c>-anchor identity for
/// the markdown projection (M1.4-T2). It is equality-neutral metadata: two images with identical
/// bytes/extent/alt but different Unids are still the same VALUE for diff/hash purposes, so it is
/// excluded from record equality (the reader feeds image equality off bytes/extent/alt, not position).
/// </summary>
internal sealed record IrInlineImage(Uri PartUri, IrHash ImageBytesHash, long WidthEmu, long HeightEmu, string? AltText) : IrInline
{
    /// <summary>The source <c>w:drawing</c>'s <c>pt:Unid</c>, or null when absent. Equality-neutral
    /// (see type remarks): does not participate in the record's structural equality.</summary>
    public string? Unid { get; init; }

    public bool Equals(IrInlineImage? other) =>
        other is not null
        && PartUri == other.PartUri
        && ImageBytesHash == other.ImageBytesHash
        && WidthEmu == other.WidthEmu
        && HeightEmu == other.HeightEmu
        && AltText == other.AltText;

    public override int GetHashCode() =>
        HashCode.Combine(PartUri, ImageBytesHash, WidthEmu, HeightEmu, AltText);
}

/// <summary>
/// An unmodeled inline element preserved opaquely: its element name plus the canonical hash of
/// its source XML, so it is still diffable (same/different bytes) without being understood.
/// </summary>
internal sealed record IrOpaqueInline(XName ElementName, IrHash CanonicalHash) : IrInline;
