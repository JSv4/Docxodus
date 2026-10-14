// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="RevisionPresentation.Word"/> renders tracked changes the way Word prints All Markup
/// (issue #851): one colour per author for both insertions and deletions, underline and
/// strikethrough only, no fills, no pilcrows, and a left-margin change bar per changed paragraph.
/// </summary>
public class RevisionPresentationTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace X = "http://www.w3.org/1999/xhtml";

    private const string Body =
        "<w:p><w:r><w:t xml:space='preserve'>Kept </w:t></w:r>" +
        "<w:ins w:id='1' w:author='Alice' w:date='2026-01-01T00:00:00Z'><w:r><w:t>added</w:t></w:r></w:ins>" +
        "<w:del w:id='2' w:author='Bob' w:date='2026-01-01T00:00:00Z'><w:r><w:delText>removed</w:delText></w:r></w:del></w:p>" +
        "<w:p><w:r><w:t>Untouched paragraph.</w:t></w:r></w:p>" +
        "<w:p><w:pPr><w:ind w:left='720'/></w:pPr><w:ins w:id='3' w:author='Bob' w:date='2026-01-01T00:00:00Z'><w:r><w:t>Indented insert.</w:t></w:r></w:ins></w:p>" +
        "<w:p><w:pPr><w:rPr><w:del w:id='4' w:author='Alice' w:date='2026-01-01T00:00:00Z'/></w:rPr></w:pPr><w:r><w:t>Merged mark.</w:t></w:r></w:p>";

    [Fact]
    public void Word_InsertionAndDeletion_TakeTheirAuthorsColourWithNoFill()
    {
        var (html, css) = Convert(RevisionPresentation.Word);

        var ins = html.Descendants(X + "ins").First();
        var del = html.Descendants(X + "del").First();
        Assert.Contains("rev-author-0", Classes(ins));
        Assert.Contains("rev-author-1", Classes(del));
        Assert.Contains(".rev-author-0 { color: #C00000; }", css);
        Assert.Contains(".rev-author-1 { color: #0070C0; }", css);
        Assert.Contains("ins.rev-ins { text-decoration: underline; }", css);
        Assert.Contains("del.rev-del { text-decoration: line-through; }", css);
        Assert.DoesNotContain("#e6ffe6", css);
        Assert.DoesNotContain("#ffe6e6", css);
    }

    [Fact]
    public void Word_AuthorColors_OverrideThePalette()
    {
        var (_, css) = Convert(RevisionPresentation.Word,
            authorColors: new Dictionary<string, string> { ["Bob"] = "teal" });

        Assert.Contains(".rev-author-1 { color: teal; }", css);
        Assert.Contains(".rev-author-0 { color: #C00000; }", css);
    }

    [Fact]
    public void Word_DeletedParagraphMark_HasNoPilcrow()
    {
        var (html, _) = Convert(RevisionPresentation.Word);

        Assert.DoesNotContain("¶", html.Value);
    }

    [Fact]
    public void Word_ChangedParagraphsGetOneChangeBar_UnchangedOnesNone()
    {
        var (html, css) = Convert(RevisionPresentation.Word);

        var paragraphs = html.Descendants(X + "p").ToList();
        var changed = paragraphs.Where(p => Classes(p).Contains("rev-changed-line")).Select(p => p.Value).ToList();
        Assert.Contains(changed, v => v.Contains("added"));
        Assert.Contains(changed, v => v.Contains("Indented insert."));
        Assert.Contains(changed, v => v.Contains("Merged mark."));
        Assert.DoesNotContain(changed, v => v.Contains("Untouched paragraph."));
        Assert.Contains(".rev-changed-line::before", css);
    }

    [Fact]
    public void Word_IndentedChangedParagraph_PassesItsIndentToTheChangeBar()
    {
        var (html, _) = Convert(RevisionPresentation.Word);

        var indented = html.Descendants(X + "p").Single(p => p.Value.Contains("Indented insert."));
        Assert.Contains("--rev-change-bar-indent", (string?)indented.Attribute("style") ?? string.Empty);
    }

    [Fact]
    public void Docxodus_Default_KeepsItsOwnPresentation()
    {
        var (html, css) = Convert(RevisionPresentation.Docxodus);

        Assert.Contains("#e6ffe6", css);
        Assert.Contains("¶", html.Value);
        Assert.DoesNotContain(html.Descendants(), e => Classes(e).Any(c => c.StartsWith("rev-author-") || c == "rev-changed-line"));
        Assert.Equal(RevisionPresentation.Docxodus, new WmlToHtmlConverterSettings().RevisionPresentation);
    }

    [Fact]
    public void HtmlConversionOps_RevisionPresentationOne_SelectsTheWordPresentation()
    {
        var html = Docxodus.Internal.HtmlConversionOps.ConvertToHtml(Docx(),
            new Docxodus.Internal.HtmlConversionOptions { RenderTrackedChanges = true, RevisionPresentation = 1 });
        var defaultHtml = Docxodus.Internal.HtmlConversionOps.ConvertToHtml(Docx(),
            new Docxodus.Internal.HtmlConversionOptions { RenderTrackedChanges = true });

        Assert.Contains("rev-author-0", html);
        Assert.Contains("rev-changed-line", html);
        Assert.DoesNotContain("rev-author-", defaultHtml);
    }

    private static List<string> Classes(XElement e) =>
        ((string?)e.Attribute("class") ?? string.Empty).Split(' ', System.StringSplitOptions.RemoveEmptyEntries).ToList();

    private static byte[] Docx()
    {
        using var stream = new MemoryStream();
        using (var package = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            using var writer = new StreamWriter(main.GetStream(FileMode.Create));
            writer.Write($"<w:document xmlns:w='{WNs}'><w:body>{Body}<w:sectPr/></w:body></w:document>");
        }
        return stream.ToArray();
    }

    private static (XElement Html, string Css) Convert(RevisionPresentation presentation,
        Dictionary<string, string>? authorColors = null)
    {
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("rev.docx", Docx()),
            new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = false,
                RenderTrackedChanges = true,
                RevisionPresentation = presentation,
                AuthorColors = authorColors,
            });
        var parsed = XElement.Parse(html.ToString());
        return (parsed, string.Concat(parsed.Descendants(X + "style").Select(s => s.Value)));
    }
}
