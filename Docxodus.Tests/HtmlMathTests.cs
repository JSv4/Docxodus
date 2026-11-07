using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HtmlMathTests
{
    private const string RunX = "<m:r><m:t>x</m:t></m:r>";
    private const string RunTwo = "<m:r><m:t>2</m:t></m:r>";
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";

    [Theory]
    [InlineData("fraction", false)]
    [InlineData("fraction", true)]
    [InlineData("root", false)]
    [InlineData("root", true)]
    [InlineData("indexed-root", false)]
    [InlineData("indexed-root", true)]
    public void Equations_preserve_operands_and_structure_as_selectable_math(string shape, bool display)
    {
        var formula = shape == "fraction"
            ? "<m:f><m:num>" + RunX + "</m:num><m:den>" + RunTwo + "</m:den></m:f>"
            : "<m:rad><m:radPr><m:degHide m:val=\"" + (shape == "root" ? "1" : "0")
                + "\"/></m:radPr><m:deg>" + (shape == "root" ? "" : RunTwo) + "</m:deg><m:e>" + RunX + "</m:e></m:rad>";
        var html = Convert(formula, display);
        var math = Assert.Single(html.Descendants(MathMl + "math"));
        Assert.Equal(display ? "block" : "inline", (string?)math.Attribute("display"));
        var expression = Assert.Single(math.Descendants(MathMl + (shape == "fraction" ? "mfrac" : shape == "root" ? "msqrt" : "mroot")));
        Assert.Equal(shape == "root" ? "x" : "x2", expression.Value);
        Assert.DoesNotContain(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "MathEquation");
        Assert.DoesNotContain("\u00a0", math.Ancestors().First(e => e.Name.LocalName == "p").Value);
    }

    [Theory]
    [InlineData("x\u0301")]
    [InlineData("x y")]
    [InlineData("x\u200dy")]
    public void Token_forms_without_a_faithful_projection_keep_the_diagnostic(string text)
    {
        var html = Convert("<m:r><m:t xml:space=\"preserve\">" + text + "</m:t></m:r>", false);
        Assert.DoesNotContain(html.Descendants(), e => e.Name.Namespace == MathMl);
        Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "MathEquation");
    }

    [Theory]
    [InlineData("<m:smallFrac m:val=\"1\"/>")]
    [InlineData("<m:defJc m:val=\"right\"/>")]
    public void Unsupported_document_math_properties_keep_the_diagnostic(string property)
    {
        var html = Convert(RunX, true, property);
        Assert.DoesNotContain(html.Descendants(), e => e.Name.Namespace == MathMl);
        Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "MathEquation");
    }

    [Fact]
    public void The_document_math_font_is_preserved_for_the_shared_font_resolver()
    {
        var math = Assert.Single(Convert(RunX, false, "<m:mathFont m:val=\"Noto Sans Math\"/>").Descendants(MathMl + "math"));
        Assert.Contains("Noto Sans Math", (string?)math.Attribute("style"));
    }

    [Theory]
    [InlineData(24, false, false, "12pt")]
    [InlineData(48, false, false, "24pt")]
    [InlineData(24, true, false, "12pt")]
    [InlineData(24, true, true, "18pt")]
    public void Equation_size_comes_from_run_styles_not_large_neighbors_or_the_paragraph_mark(
        int defaultSize, bool mark, bool style, string expected)
    {
        var paragraphStart = "<w:pPr>" + (style ? "<w:pStyle w:val=\"MathBody\"/>" : "")
            + (mark ? "<w:rPr><w:sz w:val=\"60\"/></w:rPr>" : "") + "</w:pPr>"
            + "<w:r><w:rPr><w:sz w:val=\"48\"/></w:rPr><w:t>Large neighbor</w:t></w:r>";
        var styles = "<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"" + defaultSize
            + "\"/></w:rPr></w:rPrDefault></w:docDefaults>"
            + (style ? "<w:style w:type=\"paragraph\" w:styleId=\"MathBody\"><w:name w:val=\"Math Body\"/><w:rPr><w:sz w:val=\"36\"/></w:rPr></w:style>" : "");
        var math = Assert.Single(Convert(RunX, false, paragraphStart: paragraphStart, styles: styles).Descendants(MathMl + "math"));
        Assert.Contains("font-size: " + expected + ";", (string?)math.Attribute("style"));
    }

    [Fact]
    public void Literal_operators_do_not_become_scalable_delimiters_or_large_operators()
    {
        var math = Assert.Single(Convert("<m:r><m:t>(x)+∑</m:t></m:r>", true).Descendants(MathMl + "math"));
        Assert.All(math.Descendants(MathMl + "mo"), token =>
        {
            Assert.Equal("false", (string?)token.Attribute("stretchy"));
            Assert.Equal("false", (string?)token.Attribute("largeop"));
        });
    }

    [Fact]
    public void An_empty_equation_keeps_the_empty_paragraph_line()
    {
        var html = Convert("", true);
        Assert.DoesNotContain(html.Descendants(), e => e.Name.Namespace == MathMl);
        Assert.Contains("\u00a0", html.Value);
    }

    [Fact]
    public void A_decimal_denominator_is_one_number_without_operator_spacing()
    {
        var formula = "<m:f><m:num>" + RunX + "</m:num><m:den><m:r><m:t>31.4</m:t></m:r></m:den></m:f>";
        var math = Assert.Single(Convert(formula, false).Descendants(MathMl + "math"));
        Assert.Equal("31.4", Assert.Single(math.Descendants(MathMl + "mn")).Value);
        Assert.Empty(math.Descendants(MathMl + "mo"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Paginated_source_identity_does_not_turn_supported_math_into_a_placeholder(bool root)
    {
        var formula = root ? "<m:rad><m:radPr><m:degHide/></m:radPr><m:deg/><m:e>" + RunX + "</m:e></m:rad>"
            : "<m:f><m:num>" + RunX + "</m:num><m:den>" + RunTwo + "</m:den></m:f>";
        Assert.Single(Convert(formula, false, paginated: true).Descendants(MathMl + (root ? "msqrt" : "mfrac")));
    }

    [Fact]
    public void Supplementary_math_letters_keep_their_code_point()
    {
        var math = Assert.Single(Convert("<m:r><m:t>𝑥</m:t></m:r>", false).Descendants(MathMl + "math"));
        Assert.Equal("𝑥", Assert.Single(math.Descendants(MathMl + "mi")).Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Nested_fractions_keep_a_diagnostic_for_the_unmodeled_display_size_rule(bool display)
    {
        var fraction = "<m:f><m:num>" + RunX + "</m:num><m:den>" + RunTwo + "</m:den></m:f>";
        var html = Convert("<m:f><m:num>" + RunX + "</m:num><m:den>" + fraction + "</m:den></m:f>", display);
        if (display)
        {
            Assert.DoesNotContain(html.Descendants(), e => e.Name.Namespace == MathMl);
            Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "MathEquation");
        }
        else
            Assert.Equal(2, html.Descendants(MathMl + "mfrac").Count());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Unsupported_equation_structure_retains_its_diagnostic(bool display)
    {
        var html = Convert("<m:acc><m:e>" + RunX + "</m:e></m:acc>", display);
        Assert.DoesNotContain(html.Descendants(), e => e.Name.Namespace == MathMl);
        Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "MathEquation");
    }

    private static XElement Convert(string formula, bool display, string mathProperties = "", bool paginated = false,
        string paragraphStart = "", string styles = "")
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var math = "<m:oMath>" + formula + "</m:oMath>";
            if (display) math = "<m:oMathPara>" + math + "</m:oMathPara>";
            main.PutXDocument(XDocument.Parse("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
                + "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><w:body><w:p>" + paragraphStart + math + "</w:p><w:sectPr/></w:body></w:document>"));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse("<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" + styles + "</w:styles>"));
            var settings = main.AddNewPart<DocumentSettingsPart>();
            settings.PutXDocument(XDocument.Parse("<w:settings xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
                + "xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
                + (mathProperties == "" ? "" : "<m:mathPr>" + mathProperties + "</m:mathPr>") + "</w:settings>"));
            document.Save();
            Assert.Empty(new OpenXmlValidator().Validate(document).Select(e => e.Description));
        }
        return XElement.Parse(HtmlConversionOps.ConvertToHtml(stream.ToArray(), new HtmlConversionOptions
        {
            RenderUnsupportedContentPlaceholders = true,
            FabricateCssClasses = false,
            PaginationMode = paginated ? (int)PaginationMode.Paginated : (int)PaginationMode.None,
        }));
    }
}
