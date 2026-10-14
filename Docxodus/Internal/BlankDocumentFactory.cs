#nullable enable

using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Internal;

/// <summary>
/// Mints a complete, minimal-but-valid blank DOCX: a single empty paragraph, a Normal
/// paragraph style + document defaults (Calibri 11pt), document settings, and a US-Letter
/// portrait section with 1" margins. Used by <see cref="DocxSession.CreateBlankDocxBytes"/>
/// to seed editors that draft from scratch. Built with the strongly-typed Open XML SDK in
/// its own file to avoid element-name collisions with Docxodus' LINQ-to-XML helpers.
/// </summary>
internal static class BlankDocumentFactory
{
    public static byte[] CreateBytes()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();

            // US-Letter portrait, 1" (1440 twip) margins.
            var sectPr = new SectionProperties(
                new PageSize { Width = 12240U, Height = 15840U },
                new PageMargin { Top = 1440, Bottom = 1440, Left = 1440U, Right = 1440U, Header = 720U, Footer = 720U, Gutter = 0U });

            main.Document = new Document(new Body(new Paragraph(), sectPr));

            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(
                new DocDefaults(
                    new RunPropertiesDefault(new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" },
                        new FontSize { Val = "22" },
                        new FontSizeComplexScript { Val = "22" })),
                    new ParagraphPropertiesDefault()),
                new Style(
                    new StyleName { Val = "Normal" })
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Normal",
                    Default = true,
                });

            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            main.Document.Save();
        }
        return ms.ToArray();
    }
}
