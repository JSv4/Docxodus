using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HtmlRubyTests
{
    private static string Run(string text, string properties = "") =>
        $"<w:r><w:rPr>{properties}</w:rPr><w:t>{text}</w:t></w:r>";

    private static string Ruby(string alignment = "center", bool multiple = false) => $"""
        <w:r><w:ruby><w:rubyPr><w:rubyAlign w:val="{alignment}"/><w:hps w:val="12"/>
          <w:hpsRaise w:val="22"/><w:hpsBaseText w:val="22"/><w:lid w:val="en-US"/></w:rubyPr>
          <w:rt>{(multiple ? Run("read", "<w:b/>") + Run("ing", "<w:i/>") : Run("reading"))}</w:rt>
          <w:rubyBase>{(multiple ? Run("CE", "<w:b/>") + Run("DAR") : Run("CEDAR"))}</w:rubyBase>
        </w:ruby></w:r>
        """;

    private static XElement Render(string content, bool placeholders = true)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse($"""
                <w:styles xmlns:w="{W.w}"><w:docDefaults><w:rPrDefault><w:rPr>
                  <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/>
                </w:rPr></w:rPrDefault></w:docDefaults></w:styles>
                """));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{W.w}\"/>"));
            main.PutXDocument(XDocument.Parse($"""
                <w:document xmlns:w="{W.w}"><w:body><w:p>{content}</w:p>
                  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
                    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="360" w:footer="360"/>
                  </w:sectPr>
                </w:body></w:document>
                """));
            Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(document));
        }
        return XElement.Parse(HtmlConversionOps.ConvertToHtml(stream.ToArray(), new HtmlConversionOptions
        {
            FabricateCssClasses = false, PaginationMode = 1, RenderUnsupportedContentPlaceholders = placeholders,
        }));
    }

    [Theory]
    [InlineData("center")]
    [InlineData("left")]
    [InlineData("distributeLetter")]
    [InlineData("distributeSpace")]
    public void HorizontalRubyKeepsBaseAndAnnotationInSemanticMarkup(string alignment)
    {
        var html = Render(Ruby(alignment));
        var ruby = Assert.Single(html.Descendants(), e => e.Name.LocalName == "ruby");
        var annotation = Assert.Single(ruby.Elements(), e => e.Name.LocalName == "rt");
        Assert.Equal("reading", annotation.Value);
        Assert.Equal("en-US", (string?)annotation.Attribute("lang"));
        Assert.Equal("CEDAR", string.Concat(ruby.Elements().Where(e => e.Name.LocalName is not ("rt" or "rp")).Select(e => e.Value)));
        Assert.DoesNotContain(ruby.DescendantsAndSelf(), e => e.Attribute("data-content-type") != null);
    }

    [Fact]
    public void MultipleRunsKeepTheirFormattingAndAnnotationSize()
    {
        var ruby = Assert.Single(Render(Ruby(multiple: true)).Descendants(), e => e.Name.LocalName == "ruby");
        var annotation = Assert.Single(ruby.Elements(), e => e.Name.LocalName == "rt");
        Assert.Equal("reading", annotation.Value);
        Assert.Contains(annotation.Descendants(), e => ((string?)e.Attribute("style"))?.Contains("font-weight: bold", StringComparison.Ordinal) == true);
        Assert.Contains(annotation.Descendants(), e => ((string?)e.Attribute("style"))?.Contains("font-style: italic", StringComparison.Ordinal) == true);
        Assert.Contains("font-size: 6pt", (string?)annotation.Attribute("style"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void UnsupportedVerticalAlignmentKeepsACompleteReadableFallback(bool placeholders)
    {
        var html = Render(Ruby("rightVertical"), placeholders);
        var body = Assert.Single(html.Elements(), e => e.Name.LocalName == "body");
        Assert.Contains("CEDAR", body.Value);
        Assert.Contains("reading", body.Value);
        Assert.DoesNotContain(body.Descendants(), e => e.Name.LocalName == "ruby");
        Assert.Equal(placeholders, body.Descendants().Any(e => (string?)e.Attribute("data-content-type") == "RubyAnnotation"));
    }

    [Fact]
    public void PlainRunControlKeepsBothStringsWithoutRubyMarkup()
    {
        var html = Render(Run("CEDAR reading"));
        Assert.Contains("CEDAR reading", html.Value);
        Assert.DoesNotContain(html.Descendants(), e => e.Name.LocalName is "ruby" or "rt");
    }

    [Theory]
    [InlineData(false, "11pt")]
    [InlineData(true, "15pt")]
    public void RubySizeOverridesAnnotationButOnlyDefaultsTheBase(bool explicitBaseSize, string baseSize)
    {
        var source = Ruby().Replace(Run("reading"), Run("reading", "<w:sz w:val=\"30\"/>"));
        if (explicitBaseSize) source = source.Replace(Run("CEDAR"), Run("CEDAR", "<w:sz w:val=\"30\"/>"));
        var ruby = Assert.Single(Render(source).Descendants(), e => e.Name.LocalName == "ruby");
        var annotation = ruby.Elements().Single(e => e.Name.LocalName == "rt");
        Assert.DoesNotContain(annotation.Descendants(), e => ((string?)e.Attribute("style"))?.Contains("font-size: 15pt", StringComparison.Ordinal) == true);
        Assert.Contains(ruby.Descendants(), e => e.Value == "CEDAR" && ((string?)e.Attribute("style"))?.Contains($"font-size: {baseSize}", StringComparison.Ordinal) == true);
    }
}
