// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A paragraph whose only content is a floating (anchored) drawing still holds its paragraph-mark line, as
/// Word and LibreOffice lay it out: only the drawing leaves the text flow (issue #880). The converter gives
/// such a paragraph the same placeholder line an empty paragraph gets; an inline drawing is in-flow content
/// and gets none.
/// </summary>
public class HtmlAnchorOnlyParagraphTests
{
    private const string Namespaces =
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
        "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"";

    private static string TextBox(string placement) =>
        "<w:r><w:drawing>" +
        (placement == "anchor"
            ? "<wp:anchor distT=\"0\" distR=\"114300\" distB=\"0\" distL=\"114300\" simplePos=\"0\" relativeHeight=\"10\" " +
              "behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\"><wp:simplePos x=\"0\" y=\"0\"/>" +
              "<wp:positionH relativeFrom=\"page\"><wp:posOffset>4572000</wp:posOffset></wp:positionH>" +
              "<wp:positionV relativeFrom=\"page\"><wp:posOffset>4572000</wp:posOffset></wp:positionV>" +
              "<wp:extent cx=\"1828800\" cy=\"914400\"/><wp:wrapNone/>"
            : "<wp:inline><wp:extent cx=\"1828800\" cy=\"914400\"/>") +
        "<wp:docPr id=\"1\" name=\"Text Box 1\"/><a:graphic><a:graphicData " +
        "uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:wsp><wps:txbx><w:txbxContent>" +
        "<w:p><w:r><w:t>Boxed</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>" +
        (placement == "anchor" ? "</wp:anchor>" : "</wp:inline>") +
        "</w:drawing></w:r>";

    private static byte[] Docx(string middleParagraphRuns)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
                writer.Write(
                    $"<w:document {Namespaces}><w:body>" +
                    "<w:p><w:r><w:t>Before</w:t></w:r></w:p>" +
                    $"<w:p>{middleParagraphRuns}</w:p>" +
                    "<w:p><w:r><w:t>After</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            doc.Save();
        }
        return stream.ToArray();
    }

    /// <summary>The middle paragraph's text outside the drawing: what holds its line in the flow.</summary>
    private static string FlowTextOfMiddleParagraph(byte[] docx, PaginationMode mode)
    {
        var html = HtmlConversionOps.ConvertToHtml(docx,
            new HtmlConversionOptions { FabricateCssClasses = false, PaginationMode = (int)mode });
        var root = XElement.Parse(html);
        var middle = root.Descendants().Where(e => e.Name.LocalName == "p")
            .Single(p => !p.Value.Contains("Before") && !p.Value.Contains("After"));
        // Text inside the drawing (its floating wrapper, or the text box's own source-anchored paragraphs) is
        // not flow text.
        return string.Concat(middle.DescendantNodes().OfType<XText>()
            .Where(t => !t.Ancestors().TakeWhile(a => a != middle).Any(a =>
                a.Attribute("data-docx-drawing-anchor") != null || a.Attribute("data-source-anchor-id") != null))
            .Select(t => t.Value));
    }

    [Theory]
    [InlineData(PaginationMode.Paginated)]
    [InlineData(PaginationMode.None)]
    public void ParagraphHoldingOnlyAFloatingTextBox_KeepsAPlaceholderLine(PaginationMode mode) =>
        Assert.Equal(" ", FlowTextOfMiddleParagraph(Docx(TextBox("anchor")), mode));

    [Fact]
    public void ParagraphHoldingAnInlineTextBox_GetsNoPlaceholder() =>
        Assert.Equal("", FlowTextOfMiddleParagraph(Docx(TextBox("inline")), PaginationMode.Paginated));

    [Fact]
    public void ParagraphWithTextBesideAFloatingTextBox_GetsNoPlaceholder() =>
        Assert.Equal("Caption", FlowTextOfMiddleParagraph(
            Docx("<w:r><w:t>Caption</w:t></w:r>" + TextBox("anchor")), PaginationMode.Paginated));
}
