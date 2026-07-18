#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Style-definition provenance of <see cref="DocxDiff.Compare"/> output, decoded from the
/// Word-compare oracle corpus: the result's styles part is the ORIGINAL (left) document's —
/// docDefaults byte-identical to the left — while each style whose RAW definition formatting
/// differs between the sides has its CURRENT payload updated to the RIGHT document's EFFECTIVE
/// formatting (docDefaults + basedOn chain + own definition resolved), with the left's effective
/// payload archived in a tracked <c>w:rPrChange</c>/<c>w:pPrChange</c> INSIDE the style definition.
/// Styles whose definitions agree (modulo rsid noise) are untouched, even when the two documents'
/// docDefaults differ. Right-only styles are copied; left-only styles survive for deleted content.
/// </summary>
public class DocxDiffStyleProvenanceTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static WmlDocument Doc(string ddFont, string? normalFont, string text)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
            var normal = normalFont is null
                ? new Style(new StyleName { Val = "Normal" })
                : new Style(new StyleName { Val = "Normal" },
                    new StyleRunProperties(new RunFonts { Ascii = normalFont, HighAnsi = normalFont }));
            normal.Type = StyleValues.Paragraph;
            normal.StyleId = "Normal";
            normal.Default = true;
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                new DocDefaults(
                    new RunPropertiesDefault(new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = ddFont, HighAnsi = ddFont }, new FontSize { Val = "22" })),
                    new ParagraphPropertiesDefault()),
                normal);
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("doc.docx", stream.ToArray());
    }

    private static XDocument StylesOf(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        using var reader = new StreamReader(wdoc.MainDocumentPart!.StyleDefinitionsPart!.GetStream());
        return XDocument.Parse(reader.ReadToEnd());
    }

    private static XElement StyleOf(XDocument styles, string id) =>
        styles.Root!.Elements(W + "style").Single(s => (string?)s.Attribute(W + "styleId") == id);

    private static List<string> BodyTexts(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var body = wdoc.MainDocumentPart?.Document.Body;
        return body is null
            ? new List<string>()
            : body.Descendants<Paragraph>().Select(p => p.InnerText).ToList();
    }

    [Fact]
    public void Output_KeepsLeftDocDefaults()
    {
        var left = Doc("Courier New", null, "Shared line.");
        var right = Doc("Arial", null, "Shared line.");

        var result = DocxDiff.Compare(left, right);

        var dd = StylesOf(result).Root!
            .Element(W + "docDefaults")?.Element(W + "rPrDefault")?.Element(W + "rPr")
            ?.Element(W + "rFonts");
        Assert.Equal("Courier New", (string?)dd?.Attribute(W + "ascii"));
    }

    [Fact]
    public void SharedStyleWithEqualDefinitions_IsUntouched_EvenWhenDocDefaultsDiffer()
    {
        // Both Normals are formatting-empty; only docDefaults differ → Word records NO style change.
        var left = Doc("Courier New", null, "Shared line.");
        var right = Doc("Arial", null, "Shared line.");

        var result = DocxDiff.Compare(left, right);

        var normal = StyleOf(StylesOf(result), "Normal");
        Assert.Null(normal.Element(W + "rPr")?.Element(W + "rPrChange"));
        Assert.Null(normal.Element(W + "rPr")?.Element(W + "rFonts"));
    }

    [Fact]
    public void SharedStyleWithDifferingDefinition_UpdatesToRightEffective_AndTracksOldPayload()
    {
        var left = Doc("Courier New", "Consolas", "Shared line.");
        var right = Doc("Calibri", "Arial", "Shared line.");

        var result = DocxDiff.Compare(left, right);

        var normal = StyleOf(StylesOf(result), "Normal");
        var rPr = normal.Element(W + "rPr");
        Assert.NotNull(rPr);
        // Current payload = right's EFFECTIVE formatting (its own def wins over its docDefaults).
        Assert.Equal("Arial", (string?)rPr!.Element(W + "rFonts")?.Attribute(W + "ascii"));
        // Old payload archived in a tracked rPrChange, carrying the left's effective fonts.
        var change = rPr.Element(W + "rPrChange");
        Assert.NotNull(change);
        Assert.Equal("Consolas",
            (string?)change!.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "ascii"));
    }

    [Fact]
    public void RightOnlyStyle_IsCopied_AndLeftOnlyStyleSurvives()
    {
        static WmlDocument WithExtraStyle(string ddFont, string extraId, string text)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(new Paragraph(
                    new ParagraphProperties(new ParagraphStyleId { Val = extraId }),
                    new Run(new Text(text)))));
                mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
                    new DocDefaults(
                        new RunPropertiesDefault(new RunPropertiesBaseStyle(
                            new RunFonts { Ascii = ddFont, HighAnsi = ddFont }, new FontSize { Val = "22" })),
                        new ParagraphPropertiesDefault()),
                    new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
                    new Style(new StyleName { Val = extraId }, new StyleRunProperties(new Italic()))
                    {
                        Type = StyleValues.Paragraph,
                        StyleId = extraId,
                    });
                mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                doc.Save();
            }
            return new WmlDocument("doc.docx", stream.ToArray());
        }

        var left = WithExtraStyle("Courier New", "LeftOnly", "Old text entirely.");
        var right = WithExtraStyle("Arial", "RightOnly", "Completely new words.");

        var result = DocxDiff.Compare(left, right);

        var styles = StylesOf(result);
        Assert.Contains(styles.Root!.Elements(W + "style"), s => (string?)s.Attribute(W + "styleId") == "LeftOnly");
        Assert.Contains(styles.Root!.Elements(W + "style"), s => (string?)s.Attribute(W + "styleId") == "RightOnly");

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }
}
