#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Regression guards for the final-output compatibility projection for malformed CSS font-family
/// lists.  The guard intentionally requires the exact list on direct runs in BOTH input packages;
/// it must not turn into a general font resolver or rewrite one-sided values / quoted lists with
/// a concrete fallback face.
/// </summary>
public class DocxDiffCssFontStackCompatibilityTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void SharedUnquotedCssFontStack_ProjectsDirectOutputRunsToArial()
    {
        const string stack = "Roboto, sans-serif";
        var result = DocxDiff.Compare(Doc("old", stack), Doc("new", stack));

        var fonts = DirectRunFonts(result).ToList();

        Assert.NotEmpty(fonts);
        Assert.All(fonts, f =>
        {
            Assert.Equal("Arial", (string?)f.Attribute(W + "ascii"));
            Assert.Equal("Arial", (string?)f.Attribute(W + "hAnsi"));
            Assert.Equal("Arial", (string?)f.Attribute(W + "cs"));
        });
        Assert.Equal("new", BodyText(RevisionProcessor.AcceptRevisions(result)));
        Assert.Equal("old", BodyText(RevisionProcessor.RejectRevisions(result)));
        Assert.Empty(SchemaErrors(result));
    }

    [Fact]
    public void SharedQuotedPrimaryWithOnlyGenericSansFallback_ProjectsDirectOutputRunsToArial()
    {
        const string stack = "\"Open Sans\", sans-serif";
        var result = DocxDiff.Compare(Doc("old", stack), Doc("new", stack));

        var fonts = DirectRunFonts(result).ToList();

        Assert.NotEmpty(fonts);
        Assert.All(fonts, f =>
        {
            Assert.Equal("Arial", (string?)f.Attribute(W + "ascii"));
            Assert.Equal("Arial", (string?)f.Attribute(W + "hAnsi"));
            Assert.Equal("Arial", (string?)f.Attribute(W + "cs"));
        });
        Assert.Equal("new", BodyText(RevisionProcessor.AcceptRevisions(result)));
        Assert.Equal("old", BodyText(RevisionProcessor.RejectRevisions(result)));
        Assert.Empty(SchemaErrors(result));
    }

    [Fact]
    public void SharedQuotedCssFontStackWithConcreteFallback_IsNotProjected()
    {
        const string stack = "\"Calibri\", Arial, sans-serif";
        var result = DocxDiff.Compare(Doc("old", stack), Doc("new", stack));

        var fonts = DirectRunFonts(result).ToList();

        Assert.NotEmpty(fonts);
        Assert.All(fonts, f => Assert.Equal(stack, (string?)f.Attribute(W + "ascii")));
    }

    [Fact]
    public void SharedQuotedPrimaryWithNonSansGenericFallback_IsNotProjected()
    {
        const string stack = "\"Open Sans\", serif";
        var result = DocxDiff.Compare(Doc("old", stack), Doc("new", stack));

        var fonts = DirectRunFonts(result).ToList();

        Assert.NotEmpty(fonts);
        Assert.All(fonts, f => Assert.Equal(stack, (string?)f.Attribute(W + "ascii")));
    }

    [Fact]
    public void OneSidedUnquotedCssFontStack_IsNotProjected()
    {
        const string stack = "Roboto, sans-serif";
        var result = DocxDiff.Compare(Doc("old", stack), Doc("new", "Arial"));

        var asciiFaces = DirectRunFonts(result)
            .Select(f => (string?)f.Attribute(W + "ascii"))
            .Where(f => f is not null)
            .ToList();

        Assert.Contains(stack, asciiFaces);
    }

    [Fact]
    public void NonTripletSharedCssFontStack_IsNotProjected()
    {
        const string stack = "Roboto, sans-serif";
        var result = DocxDiff.Compare(Doc("old", stack, "Roboto"), Doc("new", stack, "Roboto"));

        var fonts = DirectRunFonts(result).ToList();

        Assert.NotEmpty(fonts);
        Assert.All(fonts, f =>
        {
            Assert.Equal(stack, (string?)f.Attribute(W + "ascii"));
            Assert.Equal(stack, (string?)f.Attribute(W + "hAnsi"));
            Assert.Equal("Roboto", (string?)f.Attribute(W + "cs"));
        });
    }

    [Fact]
    public void SharedCssFontStackWithTrackedRunFormatChange_IsNotProjected()
    {
        const string stack = "Roboto, sans-serif";
        var result = DocxDiff.Compare(Doc("same", stack), Doc("same", stack, underline: true));

        var fonts = DirectRunFonts(result).ToList();

        Assert.NotEmpty(fonts);
        Assert.Contains(fonts, f => (string?)f.Attribute(W + "ascii") == stack);
        Assert.DoesNotContain(fonts, f => (string?)f.Attribute(W + "ascii") == "Arial");
    }

    private static WmlDocument Doc(string text, string asciiAndHAnsi, string? cs = null, bool underline = false)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var properties = new RunProperties(new RunFonts
            {
                Ascii = asciiAndHAnsi,
                HighAnsi = asciiAndHAnsi,
                ComplexScript = cs ?? asciiAndHAnsi,
            });
            if (underline)
                properties.Append(new Underline { Val = UnderlineValues.Single });
            main.Document = new Document(new Body(new Paragraph(new Run(properties, new Text(text)))));
            doc.Save();
        }
        return new WmlDocument("font-stack.docx", stream.ToArray());
    }

    private static IEnumerable<XElement> DirectRunFonts(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var word = WordprocessingDocument.Open(stream, false);
        using var reader = new StreamReader(word.MainDocumentPart!.GetStream());
        var xdoc = XDocument.Parse(reader.ReadToEnd());
        return xdoc.Descendants(W + "r")
            .Select(run => run.Element(W + "rPr")?.Element(W + "rFonts"))
            .Where(fonts => fonts is not null)
            .Select(fonts => fonts!)
            .ToList();
    }

    private static string BodyText(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var word = WordprocessingDocument.Open(stream, false);
        return string.Concat(word.MainDocumentPart!.Document!.Body!.Descendants<Text>().Select(t => t.Text));
    }

    private static IEnumerable<ValidationErrorInfo> SchemaErrors(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var word = WordprocessingDocument.Open(stream, false);
        return new OpenXmlValidator().Validate(word).ToList();
    }
}
