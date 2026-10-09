// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A <c>w:vertAlign</c> superscript or subscript is drawn at Word's size and raise, not the browser's
/// (issue #1016). Word reference PDFs put a superscript at about 0.65 of the run's size, raised about
/// 0.345 of the size, and a subscript at the same size lowered about 0.085. The browser's own
/// <c>vertical-align: super</c> raises about a third of the size plus a pixel, at 0.83 of the size.
/// </summary>
public class WmlToHtmlScriptPositionTests
{
    private const string Ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    private static byte[] Docx(string bodyXml)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
                writer.Write($"<w:document {Ns}><w:body>{bodyXml}<w:sectPr/></w:body></w:document>");
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            doc.Save();
        }
        return stream.ToArray();
    }

    private static string Run(string text, string? vertAlign = null) =>
        "<w:r><w:rPr>" + (vertAlign is null ? "" : $"<w:vertAlign w:val=\"{vertAlign}\"/>") +
        $"<w:sz w:val=\"24\"/></w:rPr><w:t>{text}</w:t></w:r>";

    private static XElement Html(byte[] docx, bool footnotes = false) =>
        XElement.Parse(HtmlConversionOps.ConvertToHtml(docx,
            new HtmlConversionOptions { FabricateCssClasses = false, RenderFootnotesAndEndnotes = footnotes }));

    private static Dictionary<string, string> Style(XElement element) =>
        ((string?)element.Attribute("style") ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries)
            .Select(d => d.Split(':', 2))
            .ToDictionary(d => d[0].Trim(), d => d[1].Trim());

    [Theory]
    [InlineData("superscript", "sup", "0.5308em")]
    [InlineData("subscript", "sub", "-0.1308em")]
    public void ScriptRun_IsSizedAndRaisedAsWordDrawsIt(string vertAlign, string element, string raise)
    {
        var html = Html(Docx($"<w:p>{Run("x")}{Run("2", vertAlign)}</w:p>"));
        var script = Assert.Single(html.Descendants(), e => e.Name.LocalName == element);
        var style = Style(script);

        // An em in vertical-align is the shrunken element's own size: 0.345 / 0.65 and -0.085 / 0.65
        // of it are 0.345 and 0.085 of the run's size.
        Assert.Equal("0.65em", style["font-size"]);
        Assert.Equal(raise, style["vertical-align"]);
    }

    [Fact]
    public void FootnoteReference_UsesTheSameSizeAndRaise()
    {
        var html = Html(File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles", "CA", "CA008-Footnote-Reference.docx")),
            footnotes: true);
        var reference = Assert.Single(html.Descendants(), e => (string?)e.Attribute("class") == "footnote-ref");
        var style = Style(Assert.Single(reference.Elements(), e => e.Name.LocalName == "sup"));

        Assert.Equal("0.65em", style["font-size"]);
        Assert.Equal("0.5308em", style["vertical-align"]);
    }

    /// <summary>The explicit size must not round-trip into a smaller <c>w:sz</c> when the HTML is
    /// converted back to a document.</summary>
    [Fact]
    public void Superscript_RoundTripsThroughHtmlWithItsOwnSize()
    {
        var html = Html(Docx($"<w:p>{Run("x")}{Run("2", "superscript")}</w:p>"));
        foreach (var e in html.DescendantsAndSelf())
            e.Name = e.Name.LocalName;
        foreach (var a in html.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration).ToList())
            a.Remove();
        var css = HtmlToWmlConverter.CleanUpCss((string?)html.Descendants().FirstOrDefault(d => d.Name.LocalName == "style") ?? "");

        var back = HtmlToWmlConverter.ConvertHtmlToWml("", css, "", html, HtmlToWmlConverter.GetDefaultSettings());

        using var stream = new MemoryStream(back.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(stream, false);
        var w = (XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var run = doc.MainDocumentPart!.GetXDocument().Descendants(w + "r").Single(r => r.Value == "2");
        var rPr = run.Element(w + "rPr")!;
        Assert.Equal("superscript", (string?)rPr.Element(w + "vertAlign")?.Attribute(w + "val"));
        Assert.Equal("24", (string?)rPr.Element(w + "sz")?.Attribute(w + "val"));
        Assert.Null(rPr.Element(w + "position"));
    }
}
