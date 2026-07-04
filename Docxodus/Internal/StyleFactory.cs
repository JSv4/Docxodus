#nullable enable

using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Synthesizes character styles that <see cref="DocxSession"/> formatting ops reference by id.
/// <see cref="DocxSession.ApplyFormat"/> stamps inline code as <c>w:rStyle w:val="Code"</c>;
/// on a document that never defined a "Code" style that reference is a phantom and Word silently
/// renders the run as plain text. This ensures the style actually exists (find-or-create), so the
/// run renders monospace. Mirrors <see cref="NumberingFactory"/>: find-or-create + reuse, and the
/// styles part is flushed via <c>PutXDocument</c> because the session's <see cref="DocxSession.Save"/>
/// only persists the projected parts, not the styles part.
/// </summary>
internal static class StyleFactory
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>The run-style id that inline code references.</summary>
    public const string CodeStyleId = "Code";

    /// <summary>
    /// Ensure a character style with id <see cref="CodeStyleId"/> exists. If <em>any</em> style with
    /// that id is already defined it is left untouched (respect the document's own definition); only a
    /// missing style is synthesized, as a monospace character style.
    /// </summary>
    public static void EnsureCodeCharacterStyle(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart;
        if (main is null) return;

        var part = main.StyleDefinitionsPart;
        if (part is null)
        {
            part = main.AddNewPart<StyleDefinitionsPart>();
            part.PutXDocument(new XDocument(
                new XElement(W + "styles", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName))));
        }

        var root = part.GetXDocument().Root!;
        bool exists = root.Elements(W + "style")
            .Any(st => (string?)st.Attribute(W + "styleId") == CodeStyleId);
        if (exists) return;

        root.Add(new XElement(W + "style",
            new XAttribute(W + "type", "character"),
            new XAttribute(W + "styleId", CodeStyleId),
            new XAttribute(W + "customStyle", "1"),
            new XElement(W + "name", new XAttribute(W + "val", CodeStyleId)),
            new XElement(W + "rPr",
                new XElement(W + "rFonts",
                    new XAttribute(W + "ascii", "Consolas"),
                    new XAttribute(W + "hAnsi", "Consolas"),
                    new XAttribute(W + "cs", "Consolas")))));

        // Flush to the part stream — Save only persists the projected parts, not styles.
        part.PutXDocument();
    }

    /// <summary>
    /// Ensure a <em>paragraph</em> style with id <paramref name="styleId"/> exists, synthesizing a
    /// canonical definition for the well-known Word built-ins (<c>Title</c>, <c>Subtitle</c>,
    /// <c>Heading1</c>–<c>Heading9</c>) when missing. Returns <c>true</c> if the style now exists
    /// (was already defined, or was just created); <c>false</c> for an unrecognized custom id, which
    /// the caller surfaces as <c>UnknownStyle</c> (a document can't define an arbitrary style from a
    /// bare id). Headings carry an <c>outlineLvl</c> so the converter renders them as h1–h9. An
    /// already-defined style is left untouched. Like <see cref="EnsureCodeCharacterStyle"/>, the part
    /// is flushed via <c>PutXDocument</c> because <see cref="DocxSession.Save"/> only persists the
    /// projected parts.
    /// </summary>
    public static bool EnsureParagraphStyle(WordprocessingDocument doc, string styleId)
    {
        var main = doc.MainDocumentPart;
        if (main is null) return false;

        var part = main.StyleDefinitionsPart;
        if (part is null)
        {
            part = main.AddNewPart<StyleDefinitionsPart>();
            part.PutXDocument(new XDocument(
                new XElement(W + "styles", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName))));
        }

        var root = part.GetXDocument().Root!;
        bool exists = root.Elements(W + "style")
            .Any(st => (string?)st.Attribute(W + "styleId") == styleId);
        if (exists) return true;

        var def = BuiltInParagraphStyle(styleId);
        if (def is null) return false; // unknown custom id — leave it; caller reports UnknownStyle

        root.Add(def);
        part.PutXDocument();
        return true;
    }

    /// <summary>Canonical definition for a well-known built-in paragraph style, or null if unknown.</summary>
    private static XElement? BuiltInParagraphStyle(string styleId)
    {
        if (styleId == "Title")
            return ParagraphStyle(styleId, "Title", outlineLvl: null, bold: false, halfPoints: 56, color: "2E74B5");
        if (styleId == "Subtitle")
            return ParagraphStyle(styleId, "Subtitle", outlineLvl: null, bold: false, halfPoints: 28, color: "5A5A5A");
        if (styleId.StartsWith("Heading") &&
            int.TryParse(styleId.Substring("Heading".Length), out int n) && n is >= 1 and <= 9)
        {
            // Sizes taper from 16pt down to 11pt; accent colour for the top two levels.
            int[] half = { 32, 26, 24, 24, 22, 22, 22, 22, 22 };
            return ParagraphStyle(styleId, $"heading {n}", outlineLvl: n - 1,
                bold: true, halfPoints: half[n - 1], color: n <= 2 ? "2E74B5" : "1F4D78");
        }
        return null;
    }

    /// <summary>
    /// Build a CT_Style paragraph style. Child order follows the schema (name, basedOn, next,
    /// qFormat, pPr, rPr) so Word doesn't flag the file for repair. A non-null
    /// <paramref name="outlineLvl"/> makes it a heading (and adds keepNext/keepLines).
    /// </summary>
    private static XElement ParagraphStyle(
        string styleId, string name, int? outlineLvl, bool bold, int halfPoints, string color)
    {
        var pPr = new XElement(W + "pPr");
        if (outlineLvl.HasValue)
        {
            pPr.Add(new XElement(W + "keepNext"));
            pPr.Add(new XElement(W + "keepLines"));
        }
        pPr.Add(new XElement(W + "spacing",
            new XAttribute(W + "before", "240"), new XAttribute(W + "after", "0")));
        if (outlineLvl.HasValue)
            pPr.Add(new XElement(W + "outlineLvl", new XAttribute(W + "val", outlineLvl.Value.ToString())));

        var rPr = new XElement(W + "rPr");
        if (bold) rPr.Add(new XElement(W + "b"));
        rPr.Add(new XElement(W + "color", new XAttribute(W + "val", color)));
        rPr.Add(new XElement(W + "sz", new XAttribute(W + "val", halfPoints.ToString())));
        rPr.Add(new XElement(W + "szCs", new XAttribute(W + "val", halfPoints.ToString())));

        return new XElement(W + "style",
            new XAttribute(W + "type", "paragraph"),
            new XAttribute(W + "styleId", styleId),
            new XElement(W + "name", new XAttribute(W + "val", name)),
            new XElement(W + "basedOn", new XAttribute(W + "val", "Normal")),
            new XElement(W + "next", new XAttribute(W + "val", "Normal")),
            new XElement(W + "qFormat"),
            pPr,
            rPr);
    }
}
