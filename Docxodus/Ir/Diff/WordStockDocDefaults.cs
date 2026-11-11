#nullable enable

using System.Xml.Linq;

namespace Docxodus;

/// <summary>
/// Microsoft Word's stock <c>w:docDefaults</c>, verbatim as Word's compare backfills them into its
/// output when the ORIGINAL document's styles part has none (Word never adopts the revised
/// document's). Two era variants exist, keyed on whether the original shipped a theme part:
/// a themeless original is seeded like a new document and gets <see cref="ModernXml"/> (the
/// Aptos-era blank-document defaults, matching <see cref="WordStockTheme"/>), while an original
/// with its own theme gets <see cref="ClassicXml"/> (the Calibri-era defaults). Each variant was
/// byte-identical across the independent Word-compare corpus outputs it appeared in.
/// </summary>
internal static class WordStockDocDefaults
{
    /// <summary>Aptos-era stock: theme-referencing fonts, kern 2, 12pt, spacing after=160 line=278.
    /// One deliberate deviation from Word's bytes: Word also writes
    /// <c>w14:ligatures standardContextual</c>, but the SDK's schema validator rejects that
    /// extension element inside docDefaults and it is rendering-neutral, so it is omitted.</summary>
    internal const string ModernXml =
        "<w:docDefaults xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
        "<w:rPrDefault><w:rPr>" +
        "<w:rFonts w:asciiTheme=\"minorHAnsi\" w:eastAsiaTheme=\"minorEastAsia\" w:hAnsiTheme=\"minorHAnsi\" w:cstheme=\"minorBidi\"/>" +
        "<w:kern w:val=\"2\"/><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/>" +
        "<w:lang w:val=\"en-US\" w:eastAsia=\"en-US\" w:bidi=\"ar-SA\"/>" +
        "</w:rPr></w:rPrDefault>" +
        "<w:pPrDefault><w:pPr><w:spacing w:after=\"160\" w:line=\"278\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault>" +
        "</w:docDefaults>";

    /// <summary>Calibri-era stock: theme-referencing fonts, 11pt, spacing after=160 line=259.</summary>
    internal const string ClassicXml =
        "<w:docDefaults xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
        "<w:rPrDefault><w:rPr>" +
        "<w:rFonts w:asciiTheme=\"minorHAnsi\" w:eastAsiaTheme=\"minorHAnsi\" w:hAnsiTheme=\"minorHAnsi\" w:cstheme=\"minorBidi\"/>" +
        "<w:sz w:val=\"22\"/><w:szCs w:val=\"22\"/>" +
        "<w:lang w:val=\"en-US\" w:eastAsia=\"en-US\" w:bidi=\"ar-SA\"/>" +
        "</w:rPr></w:rPrDefault>" +
        "<w:pPrDefault><w:pPr><w:spacing w:after=\"160\" w:line=\"259\" w:lineRule=\"auto\"/></w:pPr></w:pPrDefault>" +
        "</w:docDefaults>";

    private static readonly XElement MissingParagraphProperties = XElement.Parse(ModernXml)
        .Element(W.pPrDefault)!.Element(W.pPr)!;

    /// <summary>
    /// A styles part's document-default paragraph properties as Word applies them. Word reads a
    /// <c>w:pPrDefault</c> that is present, even empty, as written. When the element is missing (with
    /// or without <c>w:docDefaults</c>, or with no styles part at all) Word 2024 supplies its own
    /// 8 pt after and 1.16 lines, whatever the theme, compatibility mode or Normal template
    /// (issue #1652). Every reader of the paragraph defaults goes through here. The result is a
    /// detached copy, or null for a present <c>w:pPrDefault</c> without a <c>w:pPr</c>.
    /// </summary>
    internal static XElement? ParagraphProperties(XElement? stylesRoot)
    {
        var declared = stylesRoot?.Element(W.docDefaults)?.Element(W.pPrDefault);
        var properties = declared is null ? MissingParagraphProperties : declared.Element(W.pPr);
        return properties is null ? null : new XElement(properties);
    }
}
