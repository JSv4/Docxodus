using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class ImplicitFinalParagraphTests
{
    private static XElement Convert(string fixture)
    {
        var path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "../../../../TestFiles/OversizedPageBlocks", fixture + ".docx"));
        using (var document = WordprocessingDocument.Open(path, false))
            PackageValidation.AssertValid(document, DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
        var serial = 0;
        return WmlToHtmlConverter.ConvertToHtml(new WmlDocument(path), new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = false,
            StampAnchors = true,
            SourceAnchorIdentityProvider = _ => "source-" + serial++,
        });
    }

    [Theory]
    [InlineData("exact-row-400")]
    [InlineData("table-end-default-line-24")]
    public void A_document_ending_in_a_table_has_a_noneditable_implied_final_paragraph(string fixture)
    {
        var html = Convert(fixture);
        var paragraph = Assert.Single(html.Descendants(Xhtml.p), p => !p.Ancestors(Xhtml.table).Any());
        Assert.Equal("true", (string?)paragraph.Attribute("data-docx-implied-table-end"));
        Assert.Equal("false", (string?)paragraph.Attribute("contenteditable"));
        Assert.Null(paragraph.Attribute("data-source-anchor-id"));
        Assert.Null(paragraph.Attribute("data-anchor"));
    }

    [Fact]
    public void An_explicit_paragraph_after_the_final_table_does_not_get_an_extra_one()
    {
        var html = Convert("tables-row-exact-taller-than-page-times-new-roman");
        Assert.Equal(2, html.Descendants(Xhtml.p).Count(p => !p.Ancestors(Xhtml.table).Any()));
        Assert.DoesNotContain(html.Descendants(), p => p.Attribute("data-docx-implied-table-end") != null);
    }
}
