using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// HTML-to-DOCX reads every keyword of a <c>text-decoration</c> value (issue #1273). The DOCX-to-HTML
/// converter writes <c>underline double</c> for <c>w:u w:val="double"</c>, <c>line-through double</c>
/// for <c>w:dstrike</c> and <c>underline line-through</c> for a run with both lines, so a value with more
/// than one keyword must not lose its lines on the way back.
/// </summary>
public class HtmlToWmlTextDecorationTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static XElement RunProperties(WmlDocument wml, string text)
    {
        using var stream = new MemoryStream(wml.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(stream, false);
        var run = doc.MainDocumentPart!.GetXDocument().Descendants(W + "r").Single(r => r.Value == text);
        return run.Element(W + "rPr") ?? new XElement(W + "rPr");
    }

    private static XElement Convert(string bodyHtml, string text) =>
        RunProperties(HtmlToWmlConverter.ConvertHtmlToWml("", "", "",
            XElement.Parse($"<html><body>{bodyHtml}</body></html>"), HtmlToWmlConverter.GetDefaultSettings()), text);

    private static string Describe(XElement rPr) =>
        string.Join(" ", rPr.Elements()
            .Where(e => e.Name == W + "u" || e.Name == W + "strike" || e.Name == W + "dstrike")
            .Select(e => e.Name.LocalName + ((string?)e.Attribute(W + "val") is { } val ? "=" + val : "")));

    [Theory]
    [InlineData("underline", "u=single")]
    [InlineData("line-through", "strike")]
    [InlineData("none", "")]
    [InlineData("underline double", "u=double")]
    [InlineData("double underline", "u=double")]
    [InlineData("line-through double", "dstrike")]
    [InlineData("underline line-through", "strike u=single")]
    [InlineData("line-through underline", "strike u=single")]
    [InlineData("underline line-through double", "dstrike u=double")]
    [InlineData("underline dotted", "u=dotted")]
    [InlineData("underline dashed", "u=dash")]
    [InlineData("underline wavy", "u=wave")]
    [InlineData("underline solid red", "u=single")]
    [InlineData("UNDERLINE Double", "u=double")]
    [InlineData("overline", "")]
    public void TextDecoration_EmitsALinePerKeyword(string decoration, string expected) =>
        Assert.Equal(expected, Describe(Convert($"<p><span style='text-decoration: {decoration}'>x</span></p>", "x")));

    [Theory]
    [InlineData("<u>x</u>", "u=single")]
    [InlineData("<s>x</s>", "strike")]
    [InlineData("<u><s>x</s></u>", "strike u=single")]
    [InlineData("<u style='text-decoration: underline double'>x</u>", "u=double")]
    public void UnderlineAndStrikeElements_KeepTheirSingleLines(string inner, string expected) =>
        Assert.Equal(expected, Describe(Convert($"<p>{inner}</p>", "x")));

    private static byte[] Docx(string rPr)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
                writer.Write($"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:r><w:t>plain</w:t></w:r>" +
                    $"<w:r><w:rPr>{rPr}</w:rPr><w:t>x</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            doc.Save();
        }
        return stream.ToArray();
    }

    /// <summary>DOCX to HTML and back through the public converters keeps each run's lines.</summary>
    [Theory]
    [InlineData("<w:u w:val=\"single\"/>", "u=single")]
    [InlineData("<w:strike/>", "strike")]
    [InlineData("<w:u w:val=\"double\"/>", "u=double")]
    [InlineData("<w:dstrike/>", "dstrike")]
    [InlineData("<w:strike/><w:u w:val=\"single\"/>", "strike u=single")]
    [InlineData("<w:dstrike/><w:u w:val=\"double\"/>", "dstrike u=double")]
    public void DocxToHtmlToDocx_KeepsTheRunsLines(string rPr, string expected)
    {
        var html = XElement.Parse(HtmlConversionOps.ConvertToHtml(Docx(rPr),
            new HtmlConversionOptions { FabricateCssClasses = false }));
        foreach (var e in html.DescendantsAndSelf())
            e.Name = e.Name.LocalName;
        foreach (var a in html.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration).ToList())
            a.Remove();
        var css = HtmlToWmlConverter.CleanUpCss((string?)html.Descendants().FirstOrDefault(d => d.Name.LocalName == "style") ?? "");

        var back = HtmlToWmlConverter.ConvertHtmlToWml("", css, "", html, HtmlToWmlConverter.GetDefaultSettings());

        Assert.Equal(expected, Describe(RunProperties(back, "x")));
        Assert.Equal("", Describe(RunProperties(back, "plain")));
    }
}
