#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Guards for how a CARRIED-THROUGH <c>word/fontTable.xml</c> is normalized in a DocxDiff output.
/// Word does not copy an input's fontTable verbatim — it regenerates every font declaration from its
/// own installed-font metadata (real <c>panose1</c> + <c>family</c> classes). An input table whose
/// declarations are degraded (e.g. <c>w:family w:val="auto"</c>, zeroed panose — common in
/// third-party-generated documents) makes LibreOffice pick a DIFFERENT substitute for an absent font
/// than it picks for Word's own compare output, so a byte-equivalent body renders to different glyph
/// metrics. The normalization rewrites the declarations of KNOWN stock fonts to Word's exact
/// metadata and leaves everything else (unknown fonts, embedded-font relationships) alone.
/// </summary>
public class DocxDiffFontTableNormalizeTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void CarriedFontTable_DegradedKnownFontDescriptor_NormalizedToWordMetadata()
    {
        const string degraded =
            "<w:font w:name=\"Times New Roman\">" +
            "<w:panose1 w:val=\"00000000000000000000\"/><w:charset w:val=\"00\"/>" +
            "<w:family w:val=\"auto\"/><w:pitch w:val=\"variable\"/>" +
            "</w:font>";
        var result = DocxDiff.Compare(Doc("old", degraded), Doc("new", degraded));

        var font = FontTableFont(result, "Times New Roman");
        Assert.NotNull(font);
        Assert.Equal("roman", (string?)font!.Element(W + "family")?.Attribute(W + "val"));
        Assert.Equal("02020603050405020304", (string?)font.Element(W + "panose1")?.Attribute(W + "val"));
        Assert.Equal("new", BodyText(RevisionProcessor.AcceptRevisions(result)));
        Assert.Equal("old", BodyText(RevisionProcessor.RejectRevisions(result)));
        Assert.Empty(SchemaErrors(result));
    }

    [Fact]
    public void CarriedFontTable_UnknownFontDeclaration_LeftUntouched()
    {
        const string custom =
            "<w:font w:name=\"Somebody Custom Face\">" +
            "<w:altName w:val=\"Custom Alt\"/><w:panose1 w:val=\"01010101010101010101\"/>" +
            "<w:charset w:val=\"00\"/><w:family w:val=\"auto\"/><w:pitch w:val=\"fixed\"/>" +
            "</w:font>";
        var result = DocxDiff.Compare(Doc("old", custom), Doc("new", custom));

        var font = FontTableFont(result, "Somebody Custom Face");
        Assert.NotNull(font);
        Assert.Equal("Custom Alt", (string?)font!.Element(W + "altName")?.Attribute(W + "val"));
        Assert.Equal("01010101010101010101", (string?)font.Element(W + "panose1")?.Attribute(W + "val"));
        Assert.Equal("auto", (string?)font.Element(W + "family")?.Attribute(W + "val"));
        Assert.Empty(SchemaErrors(result));
    }

    [Fact]
    public void CarriedFontTable_EmbeddedFontRelationship_SurvivesNormalization()
    {
        // A known font whose declaration carries an embedded-font reference: the descriptor is
        // normalized but the relationship-backed w:embedRegular child must ride through, or the
        // embedded font file becomes an orphan and the render loses the actual glyphs.
        const string embedded =
            "<w:font w:name=\"Calibri\">" +
            "<w:panose1 w:val=\"00000000000000000000\"/><w:family w:val=\"auto\"/>" +
            "<w:embedRegular xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rIdEmbedded1\"/>" +
            "</w:font>";
        var result = DocxDiff.Compare(Doc("old", embedded), Doc("new", embedded));

        var font = FontTableFont(result, "Calibri");
        Assert.NotNull(font);
        Assert.Equal("swiss", (string?)font!.Element(W + "family")?.Attribute(W + "val"));
        Assert.Equal("020F0502020204030204", (string?)font.Element(W + "panose1")?.Attribute(W + "val"));
        Assert.NotNull(font.Element(W + "embedRegular"));
    }

    /// <summary>Doc whose package carries a fontTable containing <paramref name="fontDeclaration"/>.</summary>
    private static WmlDocument Doc(string text, string fontDeclaration)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
            var fontTable = main.AddNewPart<FontTablePart>();
            using (var writer = new StreamWriter(fontTable.GetStream(FileMode.Create)))
            {
                writer.Write(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<w:fonts xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    fontDeclaration +
                    "</w:fonts>");
            }
            doc.Save();
        }
        return new WmlDocument("font-normalize.docx", stream.ToArray());
    }

    private static XElement? FontTableFont(WmlDocument doc, string fontName)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var word = WordprocessingDocument.Open(stream, false);
        var fontTable = word.MainDocumentPart!.FontTablePart;
        if (fontTable is null)
            return null;
        using var reader = new StreamReader(fontTable.GetStream());
        return XDocument.Parse(reader.ReadToEnd())
            .Descendants(W + "font")
            .FirstOrDefault(f => (string?)f.Attribute(W + "name") == fontName);
    }

    private static string BodyText(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var word = WordprocessingDocument.Open(stream, false);
        return string.Concat(word.MainDocumentPart!.Document!.Body!.Descendants<Text>().Select(t => t.Text));
    }

    private static System.Collections.Generic.IEnumerable<ValidationErrorInfo> SchemaErrors(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var word = WordprocessingDocument.Open(stream, false);
        return new OpenXmlValidator().Validate(word).ToList();
    }
}
