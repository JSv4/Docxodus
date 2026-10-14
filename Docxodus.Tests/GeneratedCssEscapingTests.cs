#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// The generated stylesheet is the VALUE of an <c>h:style</c> element, so serializing the XHTML
/// escapes any XML metacharacter in it. A CSS child combinator therefore reaches the browser as
/// <c>&amp;gt;</c> — not a valid selector, so the whole rule is silently dropped. That is exactly
/// how the paginated footnote layout broke: <c>.footnote-content &gt; p:first-of-type</c> never
/// applied, so every note rendered its number alone on one line with the text starting below it,
/// which reads as a layout bug rather than a dead stylesheet.
/// </summary>
public class GeneratedCssEscapingTests
{
    private static readonly XNamespace Xhtml = "http://www.w3.org/1999/xhtml";

    private static string RenderWithAllCssOn()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body(
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Body.")),
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 1 }))));
            var fn = main.AddNewPart<FootnotesPart>();
            using var s = fn.GetStream(FileMode.Create);
            using var w = new StreamWriter(s);
            w.Write("""
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
                  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
                  <w:footnote w:id="1"><w:p><w:r><w:t>Note.</w:t></w:r></w:p></w:footnote>
                </w:footnotes>
                """);
        }

        var bytes = ms.ToArray();
        using var ms2 = new MemoryStream();
        ms2.Write(bytes, 0, bytes.Length);
        ms2.Position = 0;
        using var doc2 = WordprocessingDocument.Open(ms2, true);

        // Every CSS generator on, so a `>` anywhere in the generated stylesheet is caught.
        var settings = new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = true,
            CssClassPrefix = "docx-",
            RenderFootnotesAndEndnotes = true,
            RenderHeadersAndFooters = true,
            RenderComments = true,
            RenderTrackedChanges = true,
            RenderAnnotations = true,
            RenderPagination = PaginationMode.Paginated,
        };
        return WmlToHtmlConverter.ConvertToHtml(doc2, settings).ToString(SaveOptions.DisableFormatting);
    }

    [Fact]
    public void NoGeneratedCssIsXmlEscaped()
    {
        var html = RenderWithAllCssOn();
        var styleContent = XElement.Parse(html)
            .Descendants(Xhtml + "style")
            .Select(e => e.Value)
            .FirstOrDefault();
        Assert.False(string.IsNullOrEmpty(styleContent), "no stylesheet was generated");

        // Read the RAW serialized text, not the parsed value: parsing un-escapes it, which is the
        // whole thing this test is about — the browser gets the raw form.
        var raw = Regex.Match(html, "<style[^>]*>(.*?)</style>", RegexOptions.Singleline).Groups[1].Value;
        Assert.False(string.IsNullOrEmpty(raw), "no raw stylesheet found");

        var escaped = new[] { "&gt;", "&lt;", "&amp;" }
            .Where(tok => raw.Contains(tok, StringComparison.Ordinal))
            .ToList();

        Assert.True(
            escaped.Count == 0,
            $"generated CSS contains XML-escaped characters {string.Join(", ", escaped)} — the browser " +
            "will drop those rules. Rewrite the selector without the raw character (e.g. use a " +
            "descendant selector instead of the `>` child combinator).");
    }

    [Fact]
    public void FootnoteFirstParagraphRuleSurvivesSerialization()
    {
        var raw = Regex.Match(RenderWithAllCssOn(), "<style[^>]*>(.*?)</style>", RegexOptions.Singleline)
            .Groups[1].Value;

        // The rule that keeps a note's number and its first line on one line must be intact and
        // usable — this is the one that was silently dead in paginated mode.
        Assert.Contains(".footnote-content p:first-of-type", raw, StringComparison.Ordinal);
        Assert.DoesNotContain(".footnote-content &gt; p", raw, StringComparison.Ordinal);
    }
}
