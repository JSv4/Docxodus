#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Inserted content cloned from the RIGHT document can carry an inline <c>w:sectPr</c> whose
/// <c>w:headerReference</c>/<c>w:footerReference</c> r:ids only exist in the right package. The
/// part import intentionally skips header/footer references (correct for the shared-base
/// consolidate case), so a two-way compare of unrelated documents used to emit DANGLING references
/// — LibreOffice refuses to load such a package outright ("source file could not be loaded"), and
/// the SDK validator crashes on the relationship constraint. The renderer now strips any
/// header/footer reference whose id resolves to no relationship in the output part: per OOXML,
/// an absent reference falls back to section inheritance, so the document stays valid and loadable.
/// Regression coverage for the corpus pair verdana_italic_… → word_clean_strict01.
/// </summary>
public class DocxDiffDanglingSectionRefTests
{
    private static WmlDocument SimpleDoc(params string[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                paragraphs.Select(text => new Paragraph(new Run(new Text(text))))));
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()));
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("left.docx", stream.ToArray());
    }

    /// <summary>Two sections: paragraph 1 carries an inline sectPr with a headerReference to a
    /// right-package-only HeaderPart; trailing body sectPr closes section 2.</summary>
    private static WmlDocument DocWithSectionHeader(string firstText, string secondText)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var headerPart = mainPart.AddNewPart<HeaderPart>("rId77");
            headerPart.Header = new Header(new Paragraph(new Run(new Text("Right-only header"))));

            var firstPara = new Paragraph(
                new ParagraphProperties(
                    new SectionProperties(
                        new HeaderReference { Type = HeaderFooterValues.Default, Id = "rId77" })),
                new Run(new Text(firstText)));
            var body = new Body(
                firstPara,
                new Paragraph(new Run(new Text(secondText))),
                new SectionProperties());
            mainPart.Document = new Document(body);
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()));
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("right.docx", stream.ToArray());
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
    public void InsertedInlineSectPr_HeaderRefCollidingWithWrongTypeRelationship_IsStripped()
    {
        // The right's headerReference id may RESOLVE in the left package — to a relationship of the
        // wrong KIND (a hyperlink, the comments part, ...). LibreOffice refuses such a package and the
        // SDK validator throws on the relationship-type constraint, so the reference must be stripped
        // exactly like a dangling one (absent reference ⇒ OOXML section inheritance).
        WmlDocument LeftWithHyperlinkAt(string relId)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(new Paragraph(
                    new Hyperlink(new Run(new Text("left link"))) { Id = relId })));
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles(new DocDefaults(
                    new RunPropertiesDefault(new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                    new ParagraphPropertiesDefault()));
                mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                mainPart.AddHyperlinkRelationship(new System.Uri("https://example.com/left"), true, relId);
                doc.Save();
            }
            return new WmlDocument("left.docx", stream.ToArray());
        }

        var left = LeftWithHyperlinkAt("rId77");
        var right = DocWithSectionHeader("Sectioned new content.", "Closing new content.");

        var result = DocxDiff.Compare(left, right);

        using var s = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(s, false);
        var main = wdoc.MainDocumentPart!;
        var headerIds = new HashSet<string>(main.Parts
            .Where(p => p.OpenXmlPart is HeaderPart).Select(p => p.RelationshipId));
        var badRefs = main.Document.Body!.Descendants<HeaderReference>()
            .Select(h => h.Id?.Value)
            .Where(id => !string.IsNullOrEmpty(id) && !headerIds.Contains(id!))
            .ToList();
        Assert.Empty(badRefs);
    }

    [Fact]
    public void InsertedInlineSectPr_WithRightOnlyHeaderRef_ProducesNoDanglingReferences()
    {
        var left = SimpleDoc("Totally unrelated left content.");
        var right = DocWithSectionHeader("Sectioned new content.", "Closing new content.");

        var result = DocxDiff.Compare(left, right);

        using var stream = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var main = wdoc.MainDocumentPart!;
        var known = new HashSet<string>(main.Parts.Select(p => p.RelationshipId));
        foreach (var er in main.ExternalRelationships) known.Add(er.Id);
        foreach (var hr in main.HyperlinkRelationships) known.Add(hr.Id);

        var danglingRefs = main.Document.Body!
            .Descendants()
            .Where(e => e is HeaderReference or FooterReference)
            .Select(e => ((HeaderFooterReferenceType)e).Id?.Value)
            .Where(id => !string.IsNullOrEmpty(id) && !known.Contains(id!))
            .ToList();
        Assert.Empty(danglingRefs);

        // Round-trip contract still holds at body-text level.
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }
}
