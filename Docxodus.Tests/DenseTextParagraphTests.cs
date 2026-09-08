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
    public void MixedGeometryAndNonLiteralColorsRemainOnTheNormalPath()
    {
        foreach (var change in new System.Action<XElement>[] {
            r => r.Element(W.sz)!.SetAttributeValue(W.val, 8),
            r => r.Element(W.color)!.SetAttributeValue(W.themeColor, "accent1"),
            r => r.Add(new XElement(W.color, new XAttribute(W.val, "000000"))),
            r => r.Element(W.rFonts)!.Add(new XElement(W.b)),
        })
        {
            var frame = Frame();
            change(frame.Elements(W.r).Last().Element(W.rPr)!);
            Assert.Null(DenseTextParagraph.TryCompact(frame));
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

    [Fact]
    public void FormattingTemplateIsIndependentOfInkEncounterOrderAndBookkeepingIds()
    {
        var first = Frame();
        var second = Frame();
        second.ReplaceNodes(second.Element(W.pPr), second.Elements(W.r).Reverse().ToArray());
        UnidHelper.AssignToSelfAndDescendants(first);
        UnidHelper.AssignToSelfAndDescendants(second);
        second.SetAttributeValue(PtOpenXml.Unid, (string?)first.Attribute(PtOpenXml.Unid));
        Assert.Equal(DenseTextParagraph.TryCompact(first)!.Template.ToString(),
            DenseTextParagraph.TryCompact(second)!.Template.ToString());
    }

    [Fact]
    public void CachedFormattingExpandsTheCurrentFrameAndKeysOptionsAndGeometry()
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var anchor = session.ListBlocks().Body.First().Id;
        var options = new HtmlConversionOptions { FabricateCssClasses = false };
        void Replace(XElement frame)
        {
            frame.SetAttributeValue(PtOpenXml.Unid, anchor.Split(':').Last());
            Assert.True(session.Raw.ReplaceXml(anchor, frame.ToString()).Success);
        }
        Replace(Frame());
        _ = HtmlConversionOps.RenderBlockHtml(session, anchor, options);
        var cached = Assert.Single(session.DenseTextRenderTemplates).Value;

        var next = Frame();
        foreach (var text in next.Descendants(W.t)) text.Value = " DIFFERENT <&> FRAME ";
        next.ReplaceNodes(next.Element(W.pPr), next.Elements(W.r).Reverse().ToArray());
        Replace(next);
        var actual = HtmlConversionOps.RenderBlockHtml(session, anchor, options);
        Assert.Same(cached, Assert.Single(session.DenseTextRenderTemplates).Value);
        Assert.Equal(string.Concat(next.Descendants(W.t).Select(t => t.Value)),
            XElement.Parse(actual).Value.Replace('\u00a0', ' '));
        Assert.DoesNotContain("DXTEXT", actual);
        session.DenseTextRenderTemplates.Clear();
        Assert.Equal(actual, HtmlConversionOps.RenderBlockHtml(session, anchor, options));

        _ = HtmlConversionOps.RenderBlockHtml(session, anchor, options with { CssClassPrefix = "other-" });
        Assert.Equal(2, session.DenseTextRenderTemplates.Count);
        next.Element(W.pPr)!.Element(W.spacing)!.SetAttributeValue(W.line, 60);
        Replace(next);
        var changed = HtmlConversionOps.RenderBlockHtml(session, anchor, options);
        Assert.Equal(3, session.DenseTextRenderTemplates.Count);
        Assert.NotEqual((string?)XElement.Parse(actual).Attribute("style"),
            (string?)XElement.Parse(changed).Attribute("style"));
        session.DenseTextRenderTemplates.Clear();
        Assert.Equal(changed, HtmlConversionOps.RenderBlockHtml(session, anchor, options));

        session.DisposeRenderShell();
        Assert.Empty(session.DenseTextRenderTemplates);
    }

    [Fact]
    public void CachedFormattingIncludesNeighborContextAndHasABoundedLifetime()
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var anchor = session.ListBlocks().Body.First().Id;
        var frame = Frame();
        frame.SetAttributeValue(PtOpenXml.Unid, anchor.Split(':').Last());
        Assert.True(session.Raw.ReplaceXml(anchor, frame.ToString()).Success);
        var options = new HtmlConversionOptions { FabricateCssClasses = false };
        _ = HtmlConversionOps.RenderBlockHtml(session, anchor, options);
        Assert.Single(session.DenseTextRenderTemplates);
        // Adding a real neighbor changes the full conversion context, even
        // though the dense paragraph itself has not changed at all.
        var neighbor = new XElement(W.p, new XElement(W.r, new XElement(W.t, "neighbor")));
        frame = session.LiveDocument.MainDocumentPart!.GetXDocument().Descendants(W.p).First();
        frame.AddAfterSelf(neighbor);
        _ = HtmlConversionOps.RenderBlockHtml(session, anchor, options);
        Assert.Equal(2, session.DenseTextRenderTemplates.Count);
        neighbor.Descendants(W.t).Single().Value = "changed neighbor";
        var changed = HtmlConversionOps.RenderBlockHtml(session, anchor, options);
        Assert.Equal(3, session.DenseTextRenderTemplates.Count);
        session.DenseTextRenderTemplates.Clear();
        Assert.Equal(changed, HtmlConversionOps.RenderBlockHtml(session, anchor, options));
        for (int i = 0; i < 20; i++)
        {
            _ = HtmlConversionOps.RenderBlockHtml(session, anchor, options with { CssClassPrefix = $"frame-{i}-" });
            Assert.InRange(session.DenseTextRenderTemplates.Count, 1, 8);
        }
        // A field's cached result can depend on state outside the body XML;
        // adding one to a neighbor must bypass the formatting cache entirely.
        neighbor.Elements(W.r).Single().Add(new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin")));
        session.DenseTextRenderTemplates.Clear();
        _ = HtmlConversionOps.RenderBlockHtml(session, anchor, options);
        Assert.Empty(session.DenseTextRenderTemplates);
    }
}
