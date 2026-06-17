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
}
