using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Wp = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

public class HtmlAdjacentTextSpanTests
{
    private static Wp.Run Run(string text, params OpenXmlElement[] properties) =>
        new(new Wp.RunProperties(properties), new Wp.Text(text) { Space = SpaceProcessingModeValues.Preserve });

    private static byte[] Document(params OpenXmlElement[] content)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            // An empty paragraph default: single spacing. Without one Word supplies 1.16 lines (#1652),
            // and under a line multiple some output modes keep such runs apart, a stated multiple included.
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(new Wp.DocDefaults(new Wp.ParagraphPropertiesDefault()));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            main.Document = new Wp.Document(new Wp.Body(new Wp.Paragraph(content), new Wp.SectionProperties()));
        }
        return stream.ToArray();
    }

    private static XElement Convert(byte[] bytes, bool stampAnchors = true, bool fabricateClasses = true) =>
        XDocument.Parse(HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions
        { StampAnchors = stampAnchors, FabricateCssClasses = fabricateClasses })).Descendants(Xhtml.p).Single();

    [Theory]
    [InlineData(true, false, true)]
    [InlineData(true, true, true)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, true, false)]
    [InlineData(false, true, false)]
    public void IdenticallyRenderedRunsProduceOneTextSpan(bool identicalProperties, bool stampAnchors, bool fabricateClasses)
    {
        var bytes = Document(
            Run("Exam", identicalProperties ? new Wp.Bold() : new Wp.BoldComplexScript()),
            identicalProperties ? Run("ple", new Wp.Bold()) : Run("ple"));
        var paragraph = Convert(bytes, stampAnchors, fabricateClasses);
        Assert.Equal("Example", paragraph.Value);
        Assert.Equal(stampAnchors, paragraph.Attribute("data-anchor") != null);
        Assert.Equal("Example", Assert.Single(paragraph.Elements(Xhtml.span)).Value);
    }

    [Fact]
    public void ZeroPositionDoesNotSplitTextOrLoseSpacesAndEntities()
    {
        var paragraph = Convert(Document(Run("A & "), Run(" <B>", new Wp.Position { Val = "0" })));
        Assert.Equal("A &  <B>", Assert.Single(paragraph.Elements(Xhtml.span)).Value);
    }

    [Fact]
    public void DirectionalEntitiesKeepTheirCharactersWhenSpansJoin()
    {
        var paragraph = Convert(Document(Run("אב", new Wp.RightToLeftText()),
            Run("גד", new Wp.RightToLeftText())));
        Assert.Equal("\u200fאב\u200f\u200fגד\u200f", Assert.Single(paragraph.Elements(Xhtml.span)).Value);
    }

    [Theory]
    [InlineData("complex-size")]
    [InlineData("unused-fonts")]
    [InlineData("inherited-italic")]
    public void PropertiesWithoutAnHtmlEffectDoNotSplitLatinText(string kind)
    {
        OpenXmlElement property = kind switch
        {
            "complex-size" => new Wp.FontSizeComplexScript { Val = "32" },
            "unused-fonts" => new Wp.RunFonts { ComplexScript = "Unused", EastAsia = "Unused" },
            _ => new Wp.Italic { Val = false },
        };
        Assert.Equal("Example", Assert.Single(Convert(Document(Run("Exam", property), Run("ple")))
            .Elements(Xhtml.span)).Value);
    }

    [Fact]
    public void DifferentFormattingKeepsSeparateSpans()
    {
        var spans = Convert(Document(Run("Exam", new Wp.Bold()), Run("ple", new Wp.Italic())))
            .Elements(Xhtml.span).ToList();
        Assert.Equal(new[] { "Exam", "ple" }, spans.Select(s => s.Value));
        Assert.NotEqual((string?)spans[0].Attribute("class"), (string?)spans[1].Attribute("class"));
    }

    [Fact]
    public void BookmarkAndTabBoundariesSurvive()
    {
        var paragraph = Convert(Document(Run("Before"),
            new Wp.BookmarkStart { Name = "Target", Id = "1" },
            Run("Inside", new Wp.BoldComplexScript()), new Wp.BookmarkEnd { Id = "1" },
            new Wp.Run(new Wp.TabChar()), Run("After")));
        Assert.Contains(paragraph.Descendants(), e => (string?)e.Attribute("id") == "Target");
        Assert.Single(paragraph.Descendants(), e => e.Attribute("data-docx-tab") != null);
        Assert.Equal("BeforeInsideAfter", paragraph.Value);
    }

    [Fact]
    public void IdenticallyStyledRunBordersKeepTheirSeparateBoxes()
    {
        Wp.Border Border() => new() { Val = Wp.BorderValues.Single };
        var paragraph = Convert(Document(Run("Exam", Border()),
            Run("ple", Border(), new Wp.BoldComplexScript())));
        Assert.Equal(new[] { "Exam", "ple" }, paragraph.Elements(Xhtml.span).Select(s => s.Value));
    }
}
