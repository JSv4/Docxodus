using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class MixedTabRunTests
{
    private const string Prefix = "Kestrel shelf";
    private const string Middle = "Copper lantern inventory";
    private const string Tail = "Window alcove";

    [Theory]
    [InlineData("body", "mixed")]
    [InlineData("body", "single")]
    [InlineData("hyperlink", "mixed")]
    [InlineData("hyperlink", "single")]
    [InlineData("textbox", "mixed")]
    [InlineData("textbox", "single")]
    [InlineData("header", "mixed")]
    [InlineData("header", "single")]
    [InlineData("textbox-hyperlink", "mixed")]
    [InlineData("textbox-hyperlink", "single")]
    [InlineData("textbox-inserted", "mixed")]
    [InlineData("textbox-inserted", "single")]
    public void Run_partition_does_not_change_tab_segments(string story, string partition)
    {
        var expected = Convert(Fixture(story, "split"));
        var actual = Convert(Fixture(story, partition));
        string[] Segments(XElement html) => html.Descendants()
            .Where(e => e.Attribute("data-docx-tab") != null)
            .Select(tab => tab.Ancestors().First(e =>
                ((string?)e.Attribute("style") ?? "").Contains("display: inline-flex")))
            .Select(e => e.ToString(SaveOptions.DisableFormatting)).ToArray();
        Assert.Equal(2, Segments(expected).Length);
        Assert.Equal(Segments(expected), Segments(actual));
        Assert.Contains(Tail, actual.Value);
        if (story.Contains("inserted"))
        {
            var insertion = Assert.Single(actual.Descendants(Xhtml.ins));
            Assert.Equal("Kestrel reviewer", (string?)insertion.Attribute("data-author"));
            Assert.Contains(Prefix, insertion.Value);
            Assert.Contains(Middle, insertion.Value);
            Assert.Contains(Tail, insertion.Value);
        }
        if (story.Contains("hyperlink"))
        {
            var link = Assert.Single(actual.Descendants(Xhtml.a), e => (string?)e.Attribute("href") == "#kestrel");
            Assert.Contains(Prefix, link.Value);
            Assert.Contains(Middle, link.Value);
            Assert.Contains(Tail, link.Value);
        }
    }

    private static XElement Convert(byte[] bytes) => WmlToHtmlConverter.ConvertToHtml(
        new WmlDocument("kestrel-mixed-tabs.docx", bytes), new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = false,
            RenderPagination = PaginationMode.Paginated,
            RenderHeadersAndFooters = true,
            RenderTrackedChanges = true,
        });

    private static byte[] Fixture(string story, string partition)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles,
                new XElement(W.docDefaults, new XElement(W.rPrDefault, new XElement(W.rPr,
                    new XElement(W.rFonts, new XAttribute(W.ascii, "Liberation Serif"), new XAttribute(W.hAnsi, "Liberation Serif")),
                    new XElement(W.sz, new XAttribute(W.val, "22"))))))));
            XElement Run(params object[] content) => new(W.r,
                new XElement(W.rPr, new XElement(W.b), new XElement(W.color, new XAttribute(W.val, "254C68"))), content);
            var prefix = new XElement(W.t, Prefix);
            var middle = new XElement(W.t, Middle);
            var tail = new XElement(W.t, Tail);
            var runs = partition switch
            {
                "single" => new[] { Run(prefix, new XElement(W.tab), middle, new XElement(W.tab), tail) },
                "mixed" => new[] { Run(prefix), Run(new XElement(W.tab), middle, new XElement(W.tab), tail) },
                _ => new[] { Run(prefix), Run(new XElement(W.tab)), Run(middle), Run(new XElement(W.tab)), Run(tail) },
            };
            var paragraph = new XElement(W.p, new XElement(W.pPr, new XElement(W.tabs,
                new XElement(W.tab, new XAttribute(W.val, "left"), new XAttribute(W.pos, "2800")),
                new XElement(W.tab, new XAttribute(W.val, "left"), new XAttribute(W.pos, "6000")))),
                story.Contains("hyperlink") ? new XElement(W.hyperlink, new XAttribute(W.anchor, "kestrel"), runs)
                : story.Contains("inserted") ? new XElement(W.ins, new XAttribute(W.id, "7"),
                    new XAttribute(W.author, "Kestrel reviewer"), new XAttribute(W.date, "2025-04-08T10:00:00Z"), runs)
                : runs);
            var section = new XElement(W.sectPr,
                new XElement(W.pgSz, new XAttribute(W._w, "12240"), new XAttribute(W.h, "15840")),
                new XElement(W.pgMar, new XAttribute(W.top, "720"), new XAttribute(W.right, "720"),
                    new XAttribute(W.bottom, "720"), new XAttribute(W.left, "720")));
            XElement bodyContent = paragraph;
            if (story.StartsWith("textbox"))
                bodyContent = new XElement(W.p, new XElement(W.r, new XElement(W.pict,
                    new XElement(VML.shape, new XAttribute("style", "width:510pt;height:94pt"),
                        new XElement(VML.textbox, new XAttribute("inset", "0,0,0,0"), new XElement(W.txbxContent, paragraph))))));
            if (story == "header")
            {
                var header = main.AddNewPart<HeaderPart>();
                header.PutXDocument(new XDocument(new XElement(W.hdr, paragraph)));
                section.AddFirst(new XElement(W.headerReference, new XAttribute(W.type, "default"), new XAttribute(R.id, main.GetIdOfPart(header))));
                bodyContent = new XElement(W.p, Run(new XElement(W.t, "Kestrel body")));
            }
            main.PutXDocument(new XDocument(new XElement(W.document, new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XAttribute(XNamespace.Xmlns + "v", VML.vml), new XAttribute(XNamespace.Xmlns + "r", R.r),
                new XElement(W.body, bodyContent, section))));
        }
        return stream.ToArray();
    }
}
