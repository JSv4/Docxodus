#nullable enable

using System;
using System.Collections.Generic;
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
/// <remarks>
/// <see cref="FromInlineSdt"/> marks a run the reader spliced out of an inline <c>w:sdt</c>/<c>w:smartTag</c>
/// content control. It is EQUALITY-NEUTRAL (excluded from <see cref="Equals(IrTextRun?)"/>/<see cref="GetHashCode"/>),
/// preserving the IR's content-transparency invariant for content controls (a run reads the same value
/// whether or not it came through an SDT wrapper — see <c>Read_InlineSdt_Spliced</c>). The markdown
/// emitter reads the flag to mirror the ORACLE, whose <c>GroupInlineRuns</c> walks only
/// <c>w:r</c>/<c>w:hyperlink</c>/<c>w:ins</c>/<c>w:del</c> children and so DROPS inline-SDT content from
/// the rendered markdown — though that text still counts toward TextPreview (the oracle's
/// <c>Descendants(w:t)</c>), which the IR matches because the run is present in the inline list.
/// </remarks>
internal sealed record IrTextRun(string Text, IrRunFormat Format) : IrInline
{
    /// <summary>True when this run was spliced from an inline <c>w:sdt</c>/<c>w:smartTag</c>. Equality-neutral.</summary>
    public bool FromInlineSdt { get; init; }

    public bool Equals(IrTextRun? other) =>
        other is not null && Text == other.Text && EqualityComparer<IrRunFormat>.Default.Equals(Format, other.Format);

    public override int GetHashCode() => HashCode.Combine(Text, Format);
}

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
/// <remarks>
/// <see cref="IsSimpleField"/> distinguishes a <c>w:fldSimple</c> (the inline self-contained form)
/// from the run-based <c>w:fldChar</c> begin/separate/end machinery. It is equality-participating: the
/// two forms render DIFFERENTLY in the markdown projection — the oracle's <c>GroupInlineRuns</c> walks
/// only <c>w:r</c>/<c>w:hyperlink</c>/<c>w:ins</c>/<c>w:del</c> children and so emits the run-based
/// field's result runs (they are direct <c>w:r</c> children) but DROPS a <c>w:fldSimple</c> entirely
/// (it is none of those). The emitter mirrors that by suppressing a simple field's cached text.
/// </remarks>
internal sealed record IrFieldRun(string Instruction, IrNodeList<IrInline> CachedResult) : IrInline
{
    /// <summary>True for a <c>w:fldSimple</c>; false for the run-based <c>w:fldChar</c> machinery.</summary>
    public bool IsSimpleField { get; init; }
}

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
/// A textbox body: the inner blocks of a <c>w:txbxContent</c> reachable from a <c>w:drawing</c>
/// (DrawingML <c>wps:txbx</c>) or a <c>w:pict</c> (VML <c>v:textbox</c>). The inner blocks are
/// FULLY modeled — each is anchored, hashed, fingerprinted, and registered in the document's
/// <c>AnchorIndex</c> in the containing scope, exactly as a body/cell block would be — so textbox
/// text is no longer opaque to <see cref="IrBlock.ContentHash"/> (the Phase-2 diff blind spot this
/// node closes).
/// </summary>
/// <remarks>
/// <para><b>One node per source <c>w:txbxContent</c>.</b> Word emits the same logical textbox twice
/// inside an <c>mc:AlternateContent</c> — a DrawingML <c>mc:Choice</c> (<c>wps:txbx</c>) and a VML
/// <c>mc:Fallback</c> (<c>v:textbox</c>). The reader does NOT pick one: it walks every descendant
/// <c>w:txbxContent</c> in document order and emits one <see cref="IrTextbox"/> per occurrence, so
/// the IR's flat text and anchor set mirror the ORACLE's <c>Descendants(w:t)</c>/
/// <c>DescendantsAndSelf</c> walks, which likewise traverse both the Choice and the Fallback.</para>
/// <para><b>Markdown-invisible, index/preview-visible.</b> The oracle's <c>GroupInlineRuns</c> walks
/// only <c>w:r</c>/<c>w:hyperlink</c>/<c>w:ins</c>/<c>w:del</c>, so textbox content is DROPPED from the
/// rendered markdown — but its <c>w:t</c> text still flows into <c>ComputeTextPreview</c>/
/// <c>ScopeHasContent</c> (both <c>Descendants(w:t)</c>) and its inner paragraphs are indexed via
/// <c>DescendantsAndSelf</c>. The emitter mirrors all three: it skips this node in run grouping,
/// includes its inner text in flat-text/preview, and descends into <see cref="Blocks"/> for the index.</para>
/// <para><b>Equality.</b> Inner blocks compose by value (<see cref="IrNodeList{T}"/>), so two textboxes
/// with the same modeled inner blocks are equal regardless of source representation (Choice vs Fallback).</para>
/// </remarks>
internal sealed record IrTextbox(IrNodeList<IrBlock> Blocks) : IrInline;

/// <summary>
/// An unmodeled inline element preserved opaquely: its element name plus the canonical hash of
/// its source XML, so it is still diffable (same/different bytes) without being understood.
/// </summary>
internal sealed record IrOpaqueInline(XName ElementName, IrHash CanonicalHash) : IrInline;
