#nullable enable

using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Content in a header/footer story must survive the render, including the runs that sit BEFORE a
/// tab.
/// </summary>
/// <remarks>
/// <para>
/// <c>ConvertParagraph</c> splits a paragraph at its first tab run: everything before goes through
/// <c>TransformElementsPrecedingTab</c> (which renders it and computes a span width), everything
/// after goes through the field-aware transform. The preceding set was filtered to runs carrying a
/// <c>PtOpenXml:TabWidth</c> annotation — but that annotation is applied by
/// <c>CalculateSpanWidthForTabs</c>, which only walks the MAIN document part. A header/footer run
/// therefore never has it, so every run before a tab was dropped from the output: not rendered by
/// the preceding-tab path, and not included in the succeeding-tab range either.
/// </para>
/// <para>
/// The visible symptom is a running foot of the form <c>Last Updated October 2025 [tab] PAGE</c>
/// rendering as just the page number. Found smoke-testing the NVCA model certificate of
/// incorporation, whose footer has exactly that shape; LibreOffice renders both parts.
/// </para>
/// <para>
/// The filter conflated two questions — "which runs contribute to the computed tab width?" and
/// "which runs get rendered?". Only the first should be filtered; a run with no width annotation
/// contributes zero width but still has text. Content must never be silently dropped.
/// </para>
/// </remarks>
public class PaginatedHeaderFooterContentTests
{
    /// <summary>A document whose footer paragraph is <c>text runs → tab → PAGE field</c>, the shape
    /// Word writes for a "label on the left, page number on the right" running foot. When
    /// <paramref name="tabStop"/> is given the paragraph declares that single stop, the way the
    /// NVCA charter's footer declares a centered one at 4680 twips.</summary>
    private static byte[] BuildDocWithTabbedFooter(TabStop? tabStop = null)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            body.Append(new Paragraph(new Run(new Text("Body paragraph."))));

            var footer = main.AddNewPart<FooterPart>();
            var p = new Paragraph();
            if (tabStop is not null)
                p.Append(new ParagraphProperties(new Tabs(tabStop)));
            // Split across runs exactly as Word does, to prove every one of them survives.
            p.Append(new Run(new Text("Last Updated ") { Space = SpaceProcessingModeValues.Preserve }));
            p.Append(new Run(new Text("October")));
            p.Append(new Run(new Text(" 2025") { Space = SpaceProcessingModeValues.Preserve }));
            p.Append(new Run(new TabChar()));
            p.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
            p.Append(new Run(new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve }));
            p.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
            p.Append(new Run(new Text("7")));
            p.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            footer.Footer = new Footer(p);

            body.Append(new SectionProperties(
                new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) },
                new PageSize { Width = 12240u, Height = 15840u },
                new PageMargin
                {
                    Top = 1440, Bottom = 1440, Left = 1440, Right = 1440,
                    Header = 720, Footer = 720, Gutter = 0,
                }));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static string PaginatedHtml(byte[] bytes)
    {
        using var ms = new MemoryStream();
        ms.Write(bytes);
        using var doc = WordprocessingDocument.Open(ms, true);
        return WmlToHtmlConverter.ConvertToHtml(doc, new WmlToHtmlConverterSettings
        {
            RenderPagination = PaginationMode.Paginated,
            RenderHeadersAndFooters = true,
            ImageHandler = image => new XElement(
                XName.Get("img", "http://www.w3.org/1999/xhtml"),
                new XAttribute("src", $"data:{image.ContentType};base64,{Convert.ToBase64String(image.ImageBytes)}")),
        }).ToString();
    }

    [Fact]
    public void PHF010_FooterTextBeforeATabSurvivesTheRender()
    {
        var html = PaginatedHtml(BuildDocWithTabbedFooter());

        // The whole label, across all three runs — and the field result it shares the line with.
        Assert.Contains("Last Updated", html);
        Assert.Contains("October", html);
        Assert.Contains("2025", html);
        Assert.Contains("7", html);
    }

    /// <summary>The declaration block of one CSS rule, e.g. <c>.page-header</c>.</summary>
    private static string RuleBody(string css, string selector)
    {
        var start = css.IndexOf(selector + " {", System.StringComparison.Ordinal);
        Assert.True(start >= 0, $"stylesheet has no `{selector}` rule");
        var open = css.IndexOf('{', start);
        var close = css.IndexOf('}', open);
        return css.Substring(open + 1, close - open - 1);
    }

    /// <summary>
    /// Issue #377 — the paginated stylesheet is the second owner of Word's band model, next to the
    /// paginator's per-page inline geometry, and the two must agree.
    /// </summary>
    /// <remarks>
    /// <c>w:header</c> is the distance from the paper's top edge to the TOP of the header story,
    /// which then grows downward, and <c>w:footer</c> the distance to the BOTTOM of the footer
    /// story, which grows upward. Bottom-aligning the header and top-aligning the footer — the
    /// stylesheet's former shape, paired with <c>top: 0</c>/<c>bottom: 0</c> anchors — instead
    /// pinned both stories to the MARGINS and pulled them toward the body by
    /// <c>margin − distance</c>.
    /// </remarks>
    [Fact]
    public void PHF011_RunningContentBandsAreAnchoredToTheirDeclaredDistances()
    {
        var html = PaginatedHtml(BuildDocWithTabbedFooter());

        // The distances themselves reach the client: 720 twips = 36 pt.
        Assert.Contains("data-header-height=\"36.0\"", html);
        Assert.Contains("data-footer-height=\"36.0\"", html);

        var header = RuleBody(html, ".page-header");
        var footer = RuleBody(html, ".page-footer");

        Assert.Contains("justify-content: flex-start;", header);
        Assert.Contains("justify-content: flex-end;", footer);

        // The paginator sets `top`/`bottom` per page from the section's own distances; a
        // stylesheet edge would silently win for any band it happened to leave unset.
        Assert.DoesNotContain("top:", header);
        Assert.DoesNotContain("bottom:", footer);
    }

    /// <summary>
    /// A relationship id is local to its owning OPC part. DB005 deliberately gives a header its
    /// own image relationship; resolving that id through MainDocumentPart either loses the image
    /// or collides with an unrelated main-story relationship.
    /// </summary>
    [Fact]
    public void PHF012_HeaderImagesResolveAgainstTheOwningStoryPart()
    {
        var bytes = File.ReadAllBytes("../../../../TestFiles/DB005-Headers-With-Images.docx");
        var html = PaginatedHtml(bytes);

        Assert.Contains("data:image/png;base64,", html);
        Assert.DoesNotContain("[UNSUPPORTED IMAGE]", html);
    }

    [Fact]
    public void PHF013_PublicProcessImageFindsTheHeaderOwnerWithoutConversionAnnotations()
    {
        using var ms = new MemoryStream();
        using (var created = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
        {
            var main = created.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text("body")))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var mainImage = main.AddImagePart(ImagePartType.Png, "rIdImage");
            using (var stream = mainImage.GetStream(FileMode.Create, FileAccess.Write))
                stream.Write(new byte[] { 1, 2, 3 });

            var header = main.AddNewPart<HeaderPart>();
            var headerImage = header.AddImagePart(ImagePartType.Png, "rIdImage");
            using (var stream = headerImage.GetStream(FileMode.Create, FileAccess.Write))
                stream.Write(new byte[] { 10, 11, 12 });
            using (var writer = new StreamWriter(header.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(
                    "<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                    "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                    "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                    "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"9525\" cy=\"9525\"/>" +
                    "<wp:docPr id=\"1\" name=\"header image\"/><a:graphic><a:graphicData>" +
                    "<pic:pic><pic:blipFill><a:blip r:embed=\"rIdImage\"/></pic:blipFill></pic:pic>" +
                    "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:hdr>");
            }

            main.Document.Body!.Append(new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header) }));
            main.Document.Save();
        }

        ms.Position = 0;
        using var document = WordprocessingDocument.Open(ms, true);
        var drawing = document.MainDocumentPart!.HeaderParts.Single()
            .GetXDocument().Descendants()
            .Single(element => element.Name.LocalName == "drawing");
        var image = WmlToHtmlConverter.ProcessImage(
            document,
            drawing,
            info => new XElement("probe", Convert.ToHexString(info.ImageBytes)));

        Assert.NotNull(image);
        Assert.Equal("0A0B0C", image!.Value);
    }

    [Fact]
    public void PHF014_DanglingHeaderHyperlinkKeepsItsVisibleChildren()
    {
        using var ms = new MemoryStream();
        using (var created = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
        {
            var main = created.AddMainDocumentPart();
            var header = main.AddNewPart<HeaderPart>();
            header.Header = new Header(new Paragraph(
                new Hyperlink(new Run(new Text("dangling header link"))) { Id = "rIdMissing" }));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("body"))),
                new SectionProperties(
                    new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(header) })));
            main.Document.Save();
            header.Header.Save();
        }

        var html = PaginatedHtml(ms.ToArray());

        Assert.Contains("dangling header link", html);
    }

    /// <summary>
    /// The wrapper span's declared width, in inches, for the aligned segment that ends at a tab.
    /// <c>TransformElementsPrecedingTab</c> emits the segment as one <c>inline-flex</c> box whose
    /// width is the whole advance to the tab stop, so this number IS the resolved tab geometry.
    /// </summary>
    private static decimal TabSegmentWidthInches(string html)
    {
        var match = System.Text.RegularExpressions.Regex.Match(
            html, @"display:\s*inline-flex[^""]*?width:\s*([0-9.]+)in");
        Assert.True(match.Success, "paginated HTML has no inline-flex tab segment:\n" + html);
        return decimal.Parse(match.Groups[1].Value, NumberFormatInfo.InvariantInfo);
    }

    /// <summary>
    /// A footer tab must advance to its declared stop. Tab geometry is annotated by
    /// <c>CalculateSpanWidthForTabs</c>, which walked the main document part alone, so every tab in
    /// a running story resolved to a zero-width advance — and the page number after it was painted
    /// on top of the label before it rather than at the stop (issue #688).
    /// </summary>
    /// <remarks>
    /// A LEFT stop makes the arithmetic exact and font-independent: the advance is
    /// <c>pos − pen</c>, and the segment's total width is <c>pen + (pos − pen)</c>, so the wrapper
    /// must declare exactly the stop's own position however wide the label measures.
    /// </remarks>
    [Fact]
    public void PHF015_AFooterTabAdvancesToItsDeclaredStop()
    {
        var html = PaginatedHtml(BuildDocWithTabbedFooter(
            new TabStop { Val = TabStopValues.Left, Position = 4680 }));

        Assert.Equal(3.25m, TabSegmentWidthInches(html));
    }

    /// <summary>
    /// The NVCA charter's own footer shape: a single CENTERED stop at 4680 twips, the midpoint of a
    /// 468pt text column, so the page number sits centered under the body. Centering spends half
    /// the following text's width, so the advance lands just short of the stop — but nowhere near
    /// zero, which is what the unannotated story used to produce.
    /// </summary>
    [Fact]
    public void PHF016_ACenteredFooterTabStopCentersOnItsPosition()
    {
        var html = PaginatedHtml(BuildDocWithTabbedFooter(
            new TabStop { Val = TabStopValues.Center, Position = 4680 }));

        var width = TabSegmentWidthInches(html);
        Assert.InRange(width, 3.0m, 3.25m);
    }
}
