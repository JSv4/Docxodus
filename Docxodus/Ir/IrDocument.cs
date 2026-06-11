#nullable enable

using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace Docxodus.Ir;

/// <summary>
/// A named sequence of block-level content (e.g. "body", a header/footer, or a note body). Scope
/// names follow the IR vocabulary: "body", "hdr1"/"ftr1"…, "fn", "en", "cmt".
/// </summary>
internal sealed record IrScope(string Name, IrNodeList<IrBlock> Blocks);

/// <summary>Which header/footer occurrence a part is bound to (`w:headerReference/@w:type`).</summary>
internal enum IrHeaderFooterKind { Default, First, Even }

/// <summary>A header or footer: its scope name, occurrence kind, and the scope holding its blocks.</summary>
internal sealed record IrHeaderFooter(string ScopeName, IrHeaderFooterKind Kind, IrScope Scope);

/// <summary>
/// The immutable root of a Document IR snapshot.
/// </summary>
/// <remarks>
/// Node-for-node value equality is defined over the content scopes and stores
/// (<see cref="Body"/>, <see cref="Headers"/>, <see cref="Footers"/>, <see cref="Footnotes"/>,
/// <see cref="Endnotes"/>, <see cref="Comments"/>) — these compose value equality via
/// <see cref="IrNodeList{T}"/>. <see cref="AnchorIndex"/> and <see cref="Sources"/> are derived
/// indexes / provenance pins that keep dictionary reference equality; do not rely on them for
/// document equality. Two reads of the same bytes produce equal scopes/stores (§8).
/// </remarks>
internal sealed record IrDocument
{
    public required IrScope Body { get; init; }
    public IrNodeList<IrHeaderFooter> Headers { get; init; } = IrNodeList.Empty<IrHeaderFooter>();
    public IrNodeList<IrHeaderFooter> Footers { get; init; } = IrNodeList.Empty<IrHeaderFooter>();
    public required IrNoteStore Footnotes { get; init; }
    public required IrNoteStore Endnotes { get; init; }
    public required IrCommentStore Comments { get; init; }
    public required IrStyleRegistry Styles { get; init; }
    public required IrNumberingRegistry Numbering { get; init; }
    public required IrThemeFonts ThemeFonts { get; init; }

    /// <summary>Derived index from <see cref="IrAnchor.ToString"/> to its block; reference-equal (not part of value equality).</summary>
    public required IReadOnlyDictionary<string, IrBlock> AnchorIndex { get; init; }

    /// <summary>Provenance pin from part URI to its source document; reference-equal (not part of value equality).</summary>
    public required IReadOnlyDictionary<Uri, XDocument> Sources { get; init; }

    /// <summary>Look up a block by anchor; returns null if no block carries that anchor.</summary>
    public IrBlock? FindByAnchor(IrAnchor anchor) =>
        AnchorIndex.TryGetValue(anchor.ToString(), out var b) ? b : null;
}
