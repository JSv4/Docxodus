#nullable enable

using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;

namespace Docxodus.Tests.Ir;

/// <summary>
/// Builders for the small programmatic DOCX fixtures the <see cref="Docxodus.Ir.IrReader"/>
/// tests exercise. Each fixture includes the parts CLAUDE.md flags as required for a
/// well-formed package built from scratch (StyleDefinitionsPart, DocumentSettingsPart).
/// </summary>
internal static class IrTestDocuments
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// A document whose body holds one simple text paragraph per supplied string.
    /// </summary>
    internal static WmlDocument Create(params string[] paragraphTexts)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            foreach (var text in paragraphTexts)
                body.Append(new Paragraph(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve })));

            main.Document.Save();
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    /// <summary>
    /// A document whose <c>w:body</c> inner XML is exactly <paramref name="bodyInnerXml"/> — the
    /// raw OOXML between <c>&lt;w:body&gt;</c> and <c>&lt;/w:body&gt;</c>. Lets a test express any
    /// body shape (tables, breaks, opaque elements, sectPr, revisions) directly.
    /// </summary>
    internal static WmlDocument FromBodyXml(string bodyInnerXml)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var documentXml =
                $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInnerXml}</w:body></w:document>";
            using (var partStream = main.GetStream(FileMode.Create, FileAccess.Write))
            using (var writer = new StreamWriter(partStream))
            {
                writer.Write(documentXml);
            }
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }

    /// <summary>
    /// A document whose <c>w:body</c> inner XML is <paramref name="bodyInnerXml"/> and whose
    /// <c>w:styles</c> inner XML (the content between <c>&lt;w:styles&gt;</c> and
    /// <c>&lt;/w:styles&gt;</c>) is <paramref name="stylesInnerXml"/>. Lets a test wire up a style
    /// chain (e.g. a style carrying <c>w:numPr</c>, optionally via <c>w:basedOn</c>) that a body
    /// paragraph references by <c>w:pStyle</c>.
    /// </summary>
    internal static WmlDocument FromBodyAndStylesXml(string bodyInnerXml, string stylesInnerXml)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var stylesXml = $"<w:styles xmlns:w=\"{W}\">{stylesInnerXml}</w:styles>";
            using (var stylesStream = stylesPart.GetStream(FileMode.Create, FileAccess.Write))
            using (var stylesWriter = new StreamWriter(stylesStream))
            {
                stylesWriter.Write(stylesXml);
            }

            var documentXml =
                $"<w:document xmlns:w=\"{W}\"><w:body>{bodyInnerXml}</w:body></w:document>";
            using (var partStream = main.GetStream(FileMode.Create, FileAccess.Write))
            using (var writer = new StreamWriter(partStream))
            {
                writer.Write(documentXml);
            }
        }
        return new WmlDocument("ir-test.docx", ms.ToArray());
    }
}
