#nullable enable

using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// The paginated header/footer registry must honor the flags that GATE the first/even stories, not
/// merely the presence of a <c>w:headerReference</c>/<c>w:footerReference</c> of that type.
/// </summary>
/// <remarks>
/// Word leaves the part AND its reference behind when "Different first page" / "Different odd &amp;
/// even pages" is switched back off — it drops only <c>w:titlePg</c> (per section) and
/// <c>w:evenAndOddHeaders</c> (document-global). A reference of type <c>first</c>/<c>even</c> is
/// therefore not on its own evidence that anything renders, and both Word and LibreOffice fall back
/// to the Default story in that state.
///
/// Found by smoke-testing a real filing template (NVCA model certificate of incorporation) whose
/// leftover even footer reads "DRAFT": LibreOffice rendered the Default footer with its page number
/// on every page, while the paginated view rendered "DRAFT" — and therefore no page number — on
/// every even page. <c>DocxSession.EnsureHeaderFooterVisible</c> is the write-side counterpart of
/// this rule.
/// </remarks>
public class PaginatedHeaderFooterGatingTests
{
    /// <summary>A one-section document with default + even + first header/footer stories, where the
    /// gating flags are set only as the arguments ask.</summary>
    private static byte[] BuildDoc(bool titlePg, bool evenAndOddHeaders)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            var settings = new Settings();
            if (evenAndOddHeaders) settings.Append(new EvenAndOddHeaders());
            main.AddNewPart<DocumentSettingsPart>().Settings = settings;

            body.Append(new Paragraph(new Run(new Text("Body paragraph."))));

            var sectPr = new SectionProperties(
                new PageSize { Width = 12240u, Height = 15840u },
                new PageMargin
                {
                    Top = 1440, Bottom = 1440, Left = 1440, Right = 1440,
                    Header = 720, Footer = 720, Gutter = 0,
                });

            foreach (var (kind, text) in new[]
                     {
                         (HeaderFooterValues.Default, "DEFAULT-STORY"),
                         (HeaderFooterValues.Even, "EVEN-STORY"),
                         (HeaderFooterValues.First, "FIRST-STORY"),
                     })
            {
                var hp = main.AddNewPart<HeaderPart>();
                hp.Header = new Header(new Paragraph(new Run(new Text(text))));
                var fp = main.AddNewPart<FooterPart>();
                fp.Footer = new Footer(new Paragraph(new Run(new Text(text))));
                sectPr.PrependChild(new FooterReference { Type = kind, Id = main.GetIdOfPart(fp) });
                sectPr.PrependChild(new HeaderReference { Type = kind, Id = main.GetIdOfPart(hp) });
            }

            if (titlePg) sectPr.Append(new TitlePage());
            body.Append(sectPr);
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static string PaginatedHtml(byte[] bytes)
    {
        // Resizable: the converter opens the package for editing (it annotates), and a MemoryStream
        // constructed over a fixed byte[] cannot grow.
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
    public void PHF001_EvenStoriesAreExcludedWhenEvenAndOddHeadersIsAbsent()
    {
        // The reference and the part exist; only the document-global flag is missing — exactly what
        // Word leaves behind when the option is switched back off.
        var html = PaginatedHtml(BuildDoc(titlePg: false, evenAndOddHeaders: false));

        Assert.Contains("DEFAULT-STORY", html);
        Assert.DoesNotContain("header-even", html);
        Assert.DoesNotContain("footer-even", html);
        Assert.DoesNotContain("EVEN-STORY", html);
    }

    [Fact]
    public void PHF002_EvenStoriesAreIncludedWhenEvenAndOddHeadersIsSet()
    {
        var html = PaginatedHtml(BuildDoc(titlePg: false, evenAndOddHeaders: true));

        Assert.Contains("header-even", html);
        Assert.Contains("footer-even", html);
        Assert.Contains("EVEN-STORY", html);
    }

    /// <summary>The first-page gate already existed; pin it so the two rules can't drift apart.</summary>
    [Fact]
    public void PHF003_FirstStoriesFollowTitlePg()
    {
        var without = PaginatedHtml(BuildDoc(titlePg: false, evenAndOddHeaders: false));
        Assert.DoesNotContain("header-first", without);
        Assert.DoesNotContain("FIRST-STORY", without);

        var with = PaginatedHtml(BuildDoc(titlePg: true, evenAndOddHeaders: false));
        Assert.Contains("header-first", with);
        Assert.Contains("footer-first", with);
        Assert.Contains("FIRST-STORY", with);
    }
}
