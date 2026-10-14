#nullable enable
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Wp = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

public class LegacyPageNumberRenderingTests
{
    // Authored here from original literal content; no template or input DOCX is loaded.
    private static byte[] AuthoredDocument()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var header = main.AddNewPart<HeaderPart>();
            header.Header = new Wp.Header(new Wp.Paragraph(new Wp.Run(new Wp.PageNumber())));
            main.Document = new Wp.Document(new Wp.Body(
                new Wp.Paragraph(new Wp.Run(new Wp.Text("Cedar register "), new Wp.PageNumber())),
                new Wp.SectionProperties(new Wp.HeaderReference
                {
                    Type = Wp.HeaderFooterValues.Default, Id = main.GetIdOfPart(header),
                }, new Wp.PageSize { Width = 12240, Height = 15840 },
                    new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440,
                        Header = 720, Footer = 720 },
                    new Wp.PageNumberType { Format = Wp.NumberFormatValues.UpperRoman, Start = 7 })));
            header.Header.Save();
            main.Document.Save();
        }
        return stream.ToArray();
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(0, true)]
    [InlineData(1, false)]
    public void LegacyBlockHasVisibleNonEditableDecimalProjection(int paginationMode, bool stamp)
    {
        // Explicit stamping belongs to the block/native converter settings;
        // the full-document op stamps through its paginated mode instead.
        var html = stamp
            ? WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Authored.docx", AuthoredDocument()),
                new WmlToHtmlConverterSettings
                {
                    RenderPagination = (PaginationMode)paginationMode,
                    StampPageNumberFields = true, RenderHeadersAndFooters = true,
                }).ToString(SaveOptions.DisableFormatting)
            : HtmlConversionOps.ConvertToHtml(AuthoredDocument(), new HtmlConversionOptions
            {
                PaginationMode = paginationMode, RenderHeadersAndFooters = true,
            });
        var tree = XDocument.Parse(html);
        var blocks = tree.Descendants().Where(e => e.Name.LocalName == "span"
            && (string?)e.Attribute("contenteditable") == "false").ToList();
        Assert.Equal(2, blocks.Count);
        Assert.All(blocks, block => Assert.Equal("1", block.Value));
        var marked = paginationMode == (int)PaginationMode.Paginated || stamp;
        Assert.All(blocks, block =>
        {
            Assert.Equal(marked ? "PAGE" : null, (string?)block.Attribute("data-field"));
            Assert.Equal(marked ? "decimal" : null, (string?)block.Attribute("data-field-format"));
            Assert.Equal(marked ? "true" : null, (string?)block.Attribute("data-legacy-page-number"));
        });
    }
}
