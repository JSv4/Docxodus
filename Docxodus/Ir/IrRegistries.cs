#nullable enable

using System.Collections.Generic;
using System.Xml.Linq;

namespace Docxodus.Ir;

// ---------------------------------------------------------------------------
// Reference-equality note (consistent with IrDocument's dictionary policy).
//
// These registry records hold cloned XElement props (IrStyle.PPr/RPr,
// IrNumLevel.PPr) and IReadOnlyDictionary members. C# record equality compares
// member-by-member, but XElement compares by reference and IReadOnlyDictionary
// compares by reference too — so two builds of "the same" registry are already
// effectively reference-equal. We do NOT fight this: registries are derived
// indexes (like IrDocument.AnchorIndex / Sources), excluded from document value
// equality (IrDocument only composes value equality over its content scopes and
// stores). Callers must not rely on registry equality for document equality.
// ---------------------------------------------------------------------------

/// <summary>
/// A single resolved <c>w:style</c> definition. <see cref="PPr"/>/<see cref="RPr"/> are deep
/// clones of the style's <c>w:pPr</c>/<c>w:rPr</c> (detached from the source document) so the
/// effective-format resolver can read them without re-opening the package. They are
/// equality-excluded in spirit — see the reference-equality note at the top of this file.
/// </summary>
internal sealed record IrStyle(string Id, string? Name, string? BasedOn, string Type, bool IsDefault)
{
    /// <summary>Deep clone of the style's <c>w:pPr</c>, or null when absent.</summary>
    public XElement? PPr { get; init; }

    /// <summary>Deep clone of the style's <c>w:rPr</c>, or null when absent.</summary>
    public XElement? RPr { get; init; }
}

/// <summary>
/// Resolved style registry: every <c>w:style</c> keyed by its style id, the default paragraph
/// style id (the first <c>w:style type="paragraph" w:default="1"</c>), and deep clones of the
/// document defaults (<c>w:docDefaults/w:pPrDefault/w:pPr</c> and <c>w:rPrDefault/w:rPr</c>).
/// </summary>
internal sealed record IrStyleRegistry(
    IReadOnlyDictionary<string, IrStyle> Styles,
    string? DefaultParagraphStyleId,
    XElement? DocDefaultsPPr,
    XElement? DocDefaultsRPr)
{
    public static readonly IrStyleRegistry Empty =
        new(new Dictionary<string, IrStyle>(), null, null, null);
}

/// <summary>
/// One level (<c>w:lvl</c>) of an abstract numbering definition. <see cref="PPr"/> is a deep clone
/// of the level's <c>w:pPr</c> (level-specific indentation), or null when absent.
/// </summary>
internal sealed record IrNumLevel(int Ilvl, string NumberFormat, int? Start, string? LvlText)
{
    /// <summary>Deep clone of the level's <c>w:pPr</c>, or null when absent.</summary>
    public XElement? PPr { get; init; }
}

/// <summary>
/// An abstract numbering definition (<c>w:abstractNum</c>): its id and its levels keyed by
/// <c>w:ilvl</c>. <see cref="Levels"/> may be empty when the abstract num only carries a
/// <c>w:numStyleLink</c> indirection (not resolved at this tier — see TODO in the reader).
/// </summary>
internal sealed record IrAbstractNum(int AbstractNumId, IReadOnlyDictionary<int, IrNumLevel> Levels);

/// <summary>
/// A concrete numbering instance (<c>w:num</c>): its id, the abstract num it references, and any
/// per-level start overrides (<c>w:lvlOverride/w:startOverride</c>) keyed by <c>w:ilvl</c>.
/// </summary>
internal sealed record IrNum(int NumId, int AbstractNumId, IReadOnlyDictionary<int, int> StartOverrides);

/// <summary>
/// Resolved numbering registry: <c>w:num</c> instances and <c>w:abstractNum</c> definitions, each
/// keyed by its numeric id.
/// </summary>
internal sealed record IrNumberingRegistry(
    IReadOnlyDictionary<int, IrNum> Nums,
    IReadOnlyDictionary<int, IrAbstractNum> AbstractNums)
{
    public static readonly IrNumberingRegistry Empty =
        new(new Dictionary<int, IrNum>(), new Dictionary<int, IrAbstractNum>());
}

/// <summary>
/// Resolved theme fonts from <c>a:fontScheme</c>: the major (heading) and minor (body) ASCII/Latin
/// typefaces. Either may be null when the theme part or the corresponding face is absent.
/// </summary>
internal sealed record IrThemeFonts(string? MajorAscii, string? MinorAscii)
{
    public static readonly IrThemeFonts Empty = new(null, null);
}
