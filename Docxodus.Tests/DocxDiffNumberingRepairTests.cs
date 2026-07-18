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
    public void MissingSettingsPart_IsBackfilledWithWordDefaultTabStop()
    {
        // A package with no word/settings.xml makes LibreOffice fall back to its own default tab
        // stop (1.25cm ≈ 709 twips) instead of Word's 720, drifting every tab-positioned run a
        // little further per stop. Word always writes a settings part; the renderer backfills one
        // carrying the 720-twip default so tab metrics match the oracle.
        static WmlDocument NoSettingsDoc(string text)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles(new DocDefaults(
                    new RunPropertiesDefault(new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                    new ParagraphPropertiesDefault()));
                doc.Save();   // deliberately NO DocumentSettingsPart
            }
            return new WmlDocument("nosettings.docx", stream.ToArray());
        }

        var left = NoSettingsDoc("Old words here.");
        var right = NoSettingsDoc("Entirely different new words.");

        var result = DocxDiff.Compare(left, right);

        using var s = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(s, false);
        var settings = wdoc.MainDocumentPart!.DocumentSettingsPart;
        Assert.NotNull(settings);
        using var reader = new StreamReader(settings!.GetStream());
        var xml = XDocument.Parse(reader.ReadToEnd());
        var tab = xml.Root!.Element(W + "defaultTabStop");
        Assert.Equal("720", (string?)tab?.Attribute(W + "val"));
    }

    [Fact]
    public void MissingThemePart_IsBackfilledFromRight()
    {
        // Theme colors (schemeClr bg1/tx1/accentN) resolve to BLACK without a theme part — right-
        // sourced charts and shapes render as black boxes when the left package ships no theme.
        // Word's oracle always carries one; backfill the right's (already transitional).
        static WmlDocument DocMaybeTheme(bool withTheme, string text)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles(new DocDefaults(
                    new RunPropertiesDefault(new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                    new ParagraphPropertiesDefault()));
                mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                if (withTheme)
                {
                    var theme = mainPart.AddNewPart<ThemePart>();
                    using var w = new StreamWriter(theme.GetStream(FileMode.Create), System.Text.Encoding.UTF8);
                    w.Write("<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"T\">" +
                        "<a:themeElements><a:clrScheme name=\"O\">" +
                        "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>" +
                        "<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>" +
                        "<a:dk2><a:srgbClr val=\"44546A\"/></a:dk2><a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2>" +
                        "<a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1><a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2>" +
                        "<a:accent3><a:srgbClr val=\"A5A5A5\"/></a:accent3><a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4>" +
                        "<a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5><a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6>" +
                        "<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink><a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>" +
                        "</a:clrScheme><a:fontScheme name=\"O\"><a:majorFont><a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:majorFont>" +
                        "<a:minorFont><a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/></a:minorFont></a:fontScheme>" +
                        "<a:fmtScheme name=\"O\"><a:fillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:fillStyleLst>" +
                        "<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln><a:ln><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:ln></a:lnStyleLst>" +
                        "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>" +
                        "<a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill></a:bgFillStyleLst>" +
                        "</a:fmtScheme></a:themeElements></a:theme>");
                }
                doc.Save();
            }
            return new WmlDocument("d.docx", stream.ToArray());
        }

        var left = DocMaybeTheme(false, "Old text here.");
        var right = DocMaybeTheme(true, "Entirely new replacement words.");

        var result = DocxDiff.Compare(left, right);

        using var s = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(s, false);
        Assert.NotNull(wdoc.MainDocumentPart!.ThemePart);
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
