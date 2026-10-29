using System;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class NestedInlineRenderingTests
{
    [Theory]
    [InlineData("hyperlink")]
    [InlineData("field")]
    [InlineData("mixed")]
    public async Task DeepInlineWrappers_RenderWithinBudget(string shape)
    {
        var paragraph = Paragraph(shape, 64);
        var document = IrTestDocuments.FromBodyXml(paragraph.ToString(SaveOptions.DisableFormatting));
        PackageValidation.AssertValid(document, FileFormatVersions.Office2013);

        // The old coalescer revisits every nested descendant recursively. At this depth it
        // cannot finish; a single traversal leaves ample room for slow CI hosts and font startup.
        var html = await Task.Run(() => WmlToHtmlConverter.ConvertToHtml(document, new()))
            .WaitAsync(TimeSpan.FromSeconds(10));

        Assert.Equal(1, html.Descendants(Xhtml.body).Single().Value.Split("Willow path.").Length - 1);
    }

    [Theory]
    [InlineData("hyperlink")]
    [InlineData("field")]
    [InlineData("mixed")]
    public void CoalescingPreservesWrappersFormattingAndInput(string shape)
    {
        var input = Paragraph(shape, 8);
        var original = new XElement(input);

        var output = WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(input);

        Assert.True(XNode.DeepEquals(original, input));
        Assert.Equal(Wrappers(input), Wrappers(output));
        var run = Assert.Single(output.Descendants(W.r));
        Assert.NotNull(run.Element(W.rPr)?.Element(W.b));
        Assert.Equal("Willow path.", Assert.Single(run.Elements(W.t)).Value);

        static string[] Wrappers(XElement paragraph) => paragraph.Descendants()
            .Where(e => e.Name != W.r && e.Name != W.rPr && e.Name != W.b && e.Name != W.t)
            .Select(e => e.Name + ":" + string.Join(";", e.Attributes().Select(a => a.ToString()))).ToArray();
    }

    [Fact]
    public void AnOuterBodyDoesNotBroadenParagraphCoalescingBeyondTextboxes()
    {
        var body = new XElement(W.body, new XElement(W.p,
            new XElement(W.r, new XElement(W.t, "Willow ")),
            new XElement(W.r, new XElement(W.t, "path."))));

        Assert.True(XNode.DeepEquals(body, WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(body)));
    }

    private static XElement Paragraph(string shape, int depth)
    {
        XElement Run(string text) => new(W.r, new XElement(W.rPr, new XElement(W.b)),
            new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), text));
        XElement content = new(W.hyperlink, new XAttribute(W.anchor, "Leaf"), Run("Willow "), Run("path."));
        for (int i = 0; i < depth; i++)
        {
            var kind = shape == "mixed" ? new[] { "hyperlink", "customXml", "sdt" }[i % 3] : shape;
            content = kind switch
            {
                "hyperlink" => new XElement(W.hyperlink, new XAttribute(W.anchor, "Target"), content),
                "field" => new XElement(W.fldSimple, new XAttribute(W.instr, " DOCPROPERTY Title "), content),
                "customXml" => new XElement(W.customXml, new XAttribute(W.element, "clause"), content),
                _ => new XElement(W.sdt, new XElement(W.sdtPr, new XElement(W.id, new XAttribute(W.val, i + 1))),
                    new XElement(W.sdtContent, content)),
            };
        }
        return new XElement(W.p, content);
    }
}
