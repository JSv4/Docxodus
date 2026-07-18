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
/// Dangling-numId repair in <see cref="DocxDiff.Compare"/> output, mirroring Microsoft Word: a
/// paragraph referencing a <c>w:numId</c> with no matching definition (or no numbering part at all —
/// tool-generated corpus documents ship this shape) renders as a plain unnumbered paragraph in
/// LibreOffice, while Word SYNTHESIZES a decimal multilevel definition on open — its compare oracle
/// carries exactly that (abstractNum: decimal, "%1.", 720-twip hanging indent; num numId→abstract).
/// The renderer now performs the same repair so numbered lists survive into the redline.
/// </summary>
public class DocxDiffNumberingRepairTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static WmlDocument Doc(params string[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                paragraphs.Select(t => new Paragraph(new Run(new Text(t))))));
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()));
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("doc.docx", stream.ToArray());
    }

    /// <summary>List items referencing numId 1 with NO numbering part in the package.</summary>
    private static WmlDocument DocWithDanglingListRefs(params string[] items)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                items.Select(t => new Paragraph(
                    new ParagraphProperties(new NumberingProperties(
                        new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = 1 })),
                    new Run(new Text(t))))));
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()));
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("list.docx", stream.ToArray());
    }

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
    public void DanglingNumIdInInsertedContent_GetsSynthesizedNumberingDefinition()
    {
        var left = Doc("Plain shared paragraph.");
        var right = DocWithDanglingListRefs("Apples arrive", "Bananas follow", "Oranges finish");

        var result = DocxDiff.Compare(left, right);

        using var stream = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var numberingPart = wdoc.MainDocumentPart!.NumberingDefinitionsPart;
        Assert.NotNull(numberingPart);
        using var reader = new StreamReader(numberingPart!.GetStream());
        var numbering = XDocument.Parse(reader.ReadToEnd());
        var num = numbering.Root!.Elements(W + "num")
            .FirstOrDefault(n => (string?)n.Attribute(W + "numId") == "1");
        Assert.NotNull(num);
        var abstractId = (string?)num!.Element(W + "abstractNumId")?.Attribute(W + "val");
        var abstractNum = numbering.Root!.Elements(W + "abstractNum")
            .FirstOrDefault(a => (string?)a.Attribute(W + "abstractNumId") == abstractId);
        Assert.NotNull(abstractNum);
        var lvl0 = abstractNum!.Elements(W + "lvl")
            .FirstOrDefault(l => (string?)l.Attribute(W + "ilvl") == "0");
        Assert.NotNull(lvl0);
        Assert.Equal("decimal", (string?)lvl0!.Element(W + "numFmt")?.Attribute(W + "val"));

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }

    [Fact]
    public void ResolvedNumIds_AreLeftUntouched()
    {
        // Both sides plain, no numbering anywhere — the repair must not invent a numbering part.
        var left = Doc("One paragraph.");
        var right = Doc("Another paragraph entirely different.");

        var result = DocxDiff.Compare(left, right);

        using var stream = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        Assert.Null(wdoc.MainDocumentPart!.NumberingDefinitionsPart);
    }
}
