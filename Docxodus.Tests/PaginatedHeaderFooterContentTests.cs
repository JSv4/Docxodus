#nullable enable

using System.IO;
using System.Linq;
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
    /// Word writes for a "label on the left, page number on the right" running foot.</summary>
    private static byte[] BuildDocWithTabbedFooter()
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
}
