#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using Docxodus;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// When the ORIGINAL (left) document's styles part has no <c>w:docDefaults</c>, Word's compare
/// output backfills Word's stock docDefaults — never the revised document's. Empirically (six
/// corpus oracles, two byte-identical groups): a left WITHOUT a theme part gets the MODERN stock
/// (sz 24, spacing after=160 line=278, kern, ligatures — the Aptos-era blank-document defaults,
/// consistent with the stock-theme backfill treating the doc as new-document seeding), while a left
/// WITH its own theme gets the CLASSIC stock of that theme's era (sz 22, line=259). A left that
/// already has docDefaults keeps them untouched.
/// </summary>
public class DocxDiffDocDefaultsBackfillTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static WmlDocument BuildDoc(string text, bool withDocDefaults, bool withTheme)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = withDocDefaults
                ? new Styles(new DocDefaults(
                    new RunPropertiesDefault(new RunPropertiesBaseStyle(
                        new RunFonts { Ascii = "Georgia" }, new FontSize { Val = "26" })),
                    new ParagraphPropertiesDefault()))
                : new Styles();
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            if (withTheme)
            {
                var theme = mainPart.AddNewPart<ThemePart>();
                using var w = new StreamWriter(theme.GetStream(FileMode.Create), new System.Text.UTF8Encoding(false));
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

    private static XElement? OutputDocDefaults(WmlDocument result)
    {
        using var s = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(s, false);
        var styles = wdoc.MainDocumentPart!.StyleDefinitionsPart;
        if (styles is null)
            return null;
        using var reader = new StreamReader(styles.GetStream());
        return XDocument.Parse(reader.ReadToEnd()).Root!.Element(W + "docDefaults");
    }

    [Fact]
    public void MissingDocDefaults_ThemelessLeft_GetsModernStock()
    {
        var left = BuildDoc("Old text here.", withDocDefaults: false, withTheme: false);
        var right = BuildDoc("Entirely new replacement words.", withDocDefaults: true, withTheme: false);

        var dd = OutputDocDefaults(DocxDiff.Compare(left, right));

        Assert.NotNull(dd);
        var rPr = dd!.Element(W + "rPrDefault")?.Element(W + "rPr");
        Assert.Equal("minorHAnsi", (string?)rPr?.Element(W + "rFonts")?.Attribute(W + "asciiTheme"));
        Assert.Equal("24", (string?)rPr?.Element(W + "sz")?.Attribute(W + "val"));
        Assert.NotNull(rPr?.Element(W + "kern"));
        var spacing = dd.Element(W + "pPrDefault")?.Element(W + "pPr")?.Element(W + "spacing");
        Assert.Equal("160", (string?)spacing?.Attribute(W + "after"));
        Assert.Equal("278", (string?)spacing?.Attribute(W + "line"));
        // NOT the right's docDefaults.
        Assert.Null(dd.Descendants(W + "rFonts").FirstOrDefault(f => (string?)f.Attribute(W + "ascii") == "Georgia"));
    }

    [Fact]
    public void MissingDocDefaults_ThemedLeft_GetsClassicStock()
    {
        var left = BuildDoc("Old text here.", withDocDefaults: false, withTheme: true);
        var right = BuildDoc("Entirely new replacement words.", withDocDefaults: true, withTheme: false);

        var dd = OutputDocDefaults(DocxDiff.Compare(left, right));

        Assert.NotNull(dd);
        var rPr = dd!.Element(W + "rPrDefault")?.Element(W + "rPr");
        Assert.Equal("minorHAnsi", (string?)rPr?.Element(W + "rFonts")?.Attribute(W + "asciiTheme"));
        Assert.Equal("22", (string?)rPr?.Element(W + "sz")?.Attribute(W + "val"));
        Assert.Null(rPr?.Element(W + "kern"));
        var spacing = dd.Element(W + "pPrDefault")?.Element(W + "pPr")?.Element(W + "spacing");
        Assert.Equal("259", (string?)spacing?.Attribute(W + "line"));
    }

    [Fact]
    public void ExistingDocDefaults_AreLeftUntouched()
    {
        var left = BuildDoc("Old text here.", withDocDefaults: true, withTheme: false);
        var right = BuildDoc("Entirely new replacement words.", withDocDefaults: false, withTheme: false);

        var dd = OutputDocDefaults(DocxDiff.Compare(left, right));

        Assert.NotNull(dd);
        var fonts = dd!.Element(W + "rPrDefault")?.Element(W + "rPr")?.Element(W + "rFonts");
        Assert.Equal("Georgia", (string?)fonts?.Attribute(W + "ascii"));
    }

    [Fact]
    public void MissingStylesPart_GetsPartWithModernStockDefaults()
    {
        // The tiff_image corpus shape: left ships NO styles part at all (and no theme). The
        // output must still carry a styles part whose docDefaults are the modern stock.
        static WmlDocument BareDoc(string text)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
                doc.Save();
            }
            return new WmlDocument("d.docx", stream.ToArray());
        }

        var left = BareDoc("Old text here.");
        var right = BareDoc("Entirely new replacement words.");

        var dd = OutputDocDefaults(DocxDiff.Compare(left, right));

        Assert.NotNull(dd);
        var rPr = dd!.Element(W + "rPrDefault")?.Element(W + "rPr");
        Assert.Equal("24", (string?)rPr?.Element(W + "sz")?.Attribute(W + "val"));
        Assert.Equal("278", (string?)dd.Element(W + "pPrDefault")?.Element(W + "pPr")
            ?.Element(W + "spacing")?.Attribute(W + "line"));
    }
}
