using System;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class StyleReferenceFieldRenderingTests
{
    [Theory]
    [InlineData(false, "\"Juniper Chapter\"", "Chapter")]
    [InlineData(true, "\"Juniper Chapter\"", "Chapter")]
    [InlineData(false, "Chapter", "Chapter")]
    [InlineData(true, "Chapter", "Chapter")]
    [InlineData(false, "1", "Heading1")]
    [InlineData(true, "1", "Heading1")]
    public void BodyFieldsFindTheNearestPrecedingExactParagraphStyle(bool simple, string argument, string style)
    {
        var html = Render(argument, simple, setup: (_, body) =>
        {
            body.AddFirst(Heading("ALDER SECTION", style), Heading("Wrong inherited heading", "Derived"));
            body.Add(Heading("BIRCH SECTION", style));
        });
        Assert.Equal("Before ALDER SECTION After", ResultParagraph(html).Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ABodyFieldSearchesForwardOnlyWhenNoPrecedingMatchExists(bool simple)
    {
        var html = Render("Chapter", simple, setup: (_, body) => body.Add(Heading("BIRCH SECTION", "Chapter")));
        Assert.Equal("Before BIRCH SECTION After", ResultParagraph(html).Value);
    }

    [Theory]
    [InlineData(false, "header", false)]
    [InlineData(true, "header", false)]
    [InlineData(false, "footer", false)]
    [InlineData(true, "footer", false)]
    [InlineData(false, "header", true)]
    [InlineData(true, "footer", true)]
    public void RunningStoriesCarryAStyleIdentityForPageSpecificResolution(bool simple, string story, bool last)
    {
        var html = Render("\"Juniper Chapter\"" + (last ? " \\l" : ""), simple, story,
            setup: (_, body) => body.AddFirst(Heading("ALDER SECTION", "Chapter")));
        var field = Assert.Single(html.Descendants(Xhtml.span), e => (string?)e.Attribute("data-field") == "STYLEREF");
        Assert.Equal("Chapter", (string?)field.Attribute("data-field-style"));
        Assert.Equal(last, (string?)field.Attribute("data-field-last") == "true");
        var source = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-paragraph-style") == "Chapter");
        Assert.Equal("ALDER SECTION", (string?)source.Attribute("data-styleref-text"));
    }

    [Theory]
    [InlineData(false, "body")]
    [InlineData(true, "body")]
    [InlineData(false, "header")]
    [InlineData(true, "footer")]
    public void LockedStyleFieldsPreserveTheirCacheAndHaveNoDynamicMarker(bool simple, string story)
    {
        var html = Render("Chapter", simple, story, locked: "true",
            setup: (_, body) => body.AddFirst(Heading("ALDER SECTION", "Chapter")));
        Assert.Equal("Before Old heading After", ResultParagraph(html).Value);
        Assert.DoesNotContain(html.Descendants(), e => (string?)e.Attribute("data-field") == "STYLEREF");
    }

    [Theory]
    [InlineData("MissingStyle")]
    [InlineData("InlineLabel")]
    [InlineData("Chapter \\n")]
    public void UnsupportedStyleReferencesPreserveTheCacheAndReportIt(string argument)
    {
        var html = Render(argument, false, placeholders: true,
            setup: (_, body) => body.AddFirst(Heading("ALDER SECTION", "Chapter")));
        Assert.Single(html.Descendants(), e => e.Attribute("data-content-type") != null).Remove();
        Assert.Equal("Before Old heading After", ResultParagraph(html).Value);
    }

    [Fact]
    public void AnExactStyleNameWinsBeforeAnAmbiguousStyleId()
    {
        var html = Render("Chapter", false, setup: (main, body) =>
        {
            main.StyleDefinitionsPart!.GetXDocument().Root!.Add(Style("Alias", "Chapter"));
            main.StyleDefinitionsPart.PutXDocument();
            body.AddFirst(Heading("ALDER SECTION", "Alias"), Heading("Wrong id match", "Chapter"));
        });
        Assert.Equal("Before ALDER SECTION After", ResultParagraph(html).Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AStyleResultReplacesAPropertyFieldInsideItsOldCache(bool simple)
    {
        var html = Render("Chapter", simple, mutate: paragraph =>
            paragraph.Descendants(W.t).Single(e => e.Value == "Old heading").Parent!.ReplaceWith(
                new XElement(W.fldSimple, new XAttribute(W.instr, "TITLE"),
                    new XElement(W.r, new XElement(W.t, "Old nested title")))),
            setup: (_, body) => body.AddFirst(Heading("ALDER SECTION", "Chapter")));
        Assert.Equal("Before ALDER SECTION After", ResultParagraph(html).Value);
    }

    [Theory]
    [InlineData("noBreakHyphen", "\u2011")]
    [InlineData("softHyphen", "\u00ad")]
    [InlineData("sym", "\u00a9")]
    public void AStyleReferenceRetainsSpecialVisibleHeadingCharacters(string kind, string expected)
    {
        var html = Render("Chapter", false, setup: (_, body) =>
        {
            var heading = Heading("", "Chapter");
            heading.Element(W.r)!.ReplaceNodes(new XElement(W.t, "ALDER"),
                new XElement(W.w + kind, kind == "sym" ? new[] {
                    new XAttribute(W.font, "Times New Roman"), new XAttribute(W._char, "00A9") } : Array.Empty<XAttribute>()),
                new XElement(W.t, "SECTION"));
            body.AddFirst(heading);
        });
        Assert.Equal("Before ALDER" + expected + "SECTION After", ResultParagraph(html).Value);
    }

    private static XElement ResultParagraph(XElement html) =>
        Assert.Single(html.Descendants(Xhtml.p), e => e.Value.StartsWith("Before", StringComparison.Ordinal));

    private static XElement Heading(string text, string style) => new XElement(W.p,
        new XElement(W.pPr, new XElement(W.pStyle, new XAttribute(W.val, style))), new XElement(W.r, new XElement(W.t, text)));

    private static XElement Style(string id, string name, string type = "paragraph") => new XElement(W.style,
        new XAttribute(W.type, type), new XAttribute(W.styleId, id), new XElement(W.name, new XAttribute(W.val, name)));

    private static XElement Render(string argument, bool simple, string story = "body", string? locked = null,
        bool placeholders = false, Action<MainDocumentPart, XElement>? setup = null, Action<XElement>? mutate = null) =>
        DocumentFieldRenderingTests.Render("STYLEREF " + argument, simple, story, "Old heading", locked,
            mutate: mutate, placeholders: placeholders, paginate: true, setup: (main, body) =>
            {
                main.StyleDefinitionsPart!.PutXDocument(new XDocument(new XElement(W.styles,
                    Style("Chapter", "Juniper Chapter"), Style("Heading1", "heading 1"),
                    Style("InlineLabel", "Inline Label", "character"),
                    new XElement(W.style, new XAttribute(W.type, "paragraph"), new XAttribute(W.styleId, "Derived"),
                        new XElement(W.name, new XAttribute(W.val, "Derived")),
                        new XElement(W.basedOn, new XAttribute(W.val, "Chapter")),
                        new XElement(W.pPr, new XElement(W.outlineLvl, new XAttribute(W.val, "0")))))));
                setup?.Invoke(main, body);
            });
}
