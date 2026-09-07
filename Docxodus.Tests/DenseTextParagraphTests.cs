#nullable enable
using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class DenseTextParagraphTests
{
    private static XElement Frame(int count = 600) => new XElement(W.p,
        new XElement(W.pPr, new XElement(W.spacing, new XAttribute(W.line, 48),
            new XAttribute(W.lineRule, "exact")), new XElement(W.shd,
            new XAttribute(W.val, "clear"), new XAttribute(W.fill, "000000"))),
        Enumerable.Range(0, count).Select(i => new XElement(W.r,
            new XElement(W.rPr, new XElement(W.rFonts, new XAttribute(W.ascii, "Courier New"),
                new XAttribute(W.hAnsi, "Courier New")),
                new XElement(W.color, new XAttribute(W.val, i % 2 == 0 ? "FF8844" : "FFFFFF")),
                new XElement(W.spacing, new XAttribute(W.val, -2)),
                new XElement(W.sz, new XAttribute(W.val, 6))),
            new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), " | <&#@>  "),
            i % 12 == 0 ? new XElement(W.br) : null)));

    [Fact]
    public void CompactResolvesOnlyDistinctFormatsAndExpandsEveryCharacter()
    {
        var frame = Frame();
        var originalText = string.Concat(frame.Descendants(W.t).Select(t => t.Value));
        var originalBreaks = frame.Descendants(W.br).Count();
        var dense = DenseTextParagraph.TryCompact(frame);
        Assert.NotNull(dense);
        Assert.Equal(600, frame.Elements(W.r).Count()); // the authoritative XML is untouched
        Assert.Equal(2, dense.Template.Elements(W.r).Count());
        var html = new XElement(Xhtml.p, dense.Template.Elements(W.r).Select(r => new XElement(Xhtml.span,
            new XAttribute("style", "color:#" + (string)r.Element(W.rPr)!.Element(W.color)!.Attribute(W.val)!),
            r.Element(W.t)!.Value, new XElement(Xhtml.br))));
        Assert.True(dense.TryExpand(html));
        Assert.Equal(originalText, html.Value.Replace('\u00a0', ' '));
        Assert.Equal(originalBreaks, html.Descendants(Xhtml.br).Count());
        Assert.Equal(600, html.Elements(Xhtml.span).Count());
    }

    [Fact]
    public void ComplexOrSmallParagraphsRemainOnTheNormalPath()
    {
        Assert.Null(DenseTextParagraph.TryCompact(Frame(10)));
        foreach (var extra in new[] { new XElement(W.tab), new XElement(W.drawing),
            new XElement(W.footnoteReference), new XElement(W.t, "unicode: ▀") })
        {
            var frame = Frame();
            frame.Elements(W.r).Last().Add(extra);
            var before = frame.ToString();
            Assert.Null(DenseTextParagraph.TryCompact(frame));
            Assert.Equal(before, frame.ToString());
        }
    }

    [Fact]
    public void IncrementalRenderingMatchesFullConversionForDenseAscii()
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var anchor = session.ListBlocks().Body.First().Id;
        var frame = Frame();
        frame.SetAttributeValue(PtOpenXml.Unid, anchor.Split(':').Last());
        Assert.True(session.Raw.ReplaceXml(anchor, frame.ToString()).Success);
        var block = XElement.Parse(HtmlConversionOps.RenderBlockHtml(session, anchor,
            new HtmlConversionOptions { FabricateCssClasses = false }));
        var full = XElement.Parse(HtmlConversionOps.ConvertToHtml(session.Save(persistAnchorIds: true),
            new HtmlConversionOptions { FabricateCssClasses = false, StampAnchors = true }));
        var expected = full.Descendants().Single(e => (string?)e.Attribute("data-anchor") == anchor.Split(':').Last());
        Assert.Equal(expected.Value.Replace('\u00a0', ' '), block.Value.Replace('\u00a0', ' '));
        Assert.Equal(expected.Descendants(Xhtml.br).Count(), block.Descendants(Xhtml.br).Count());
        Assert.Equal((string?)expected.Attribute("style"), (string?)block.Attribute("style"));
        Assert.Equal(expected.Elements(Xhtml.span).Select(e => (string?)e.Attribute("style")),
            block.Elements(Xhtml.span).Select(e => (string?)e.Attribute("style")));
        Assert.Contains("letter-spacing: -0.1pt", block.Elements(Xhtml.span).First().Attribute("style")!.Value);
    }
}
