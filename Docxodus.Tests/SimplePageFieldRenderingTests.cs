using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Wp = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

// Issue #1115: a w:fldSimple PAGE / NUMPAGES field must carry the same page-number marker as its
// complex-field form so the paginator can substitute each page's number.
public class SimplePageFieldRenderingTests
{
    // Authored here from original literal content; no template or input DOCX is loaded.
    private static byte[] AuthoredDocument()
    {
        static Wp.SimpleField Field(string instruction, string cached) =>
            new(new Wp.Run(new Wp.RunProperties(new Wp.Italic()), new Wp.Text(cached))) { Instruction = instruction };

        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream,
                   DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var header = main.AddNewPart<HeaderPart>();
            header.Header = new Wp.Header(new Wp.Paragraph(
                Field(" PAGE \\* Arabic ", "1"), new Wp.Run(new Wp.Text(" of ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
                Field(" NUMPAGES ", "1")));
            main.Document = new Wp.Document(new Wp.Body(
                new Wp.Paragraph(new Wp.Run(new Wp.Text("Maple register "))),
                new Wp.Paragraph(Field(" AUTHOR ", "Cedar")),
                new Wp.SectionProperties(new Wp.HeaderReference
                {
                    Type = Wp.HeaderFooterValues.Default, Id = main.GetIdOfPart(header),
                }, new Wp.PageSize { Width = 12240, Height = 15840 },
                    new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440,
                        Header = 720, Footer = 720 },
                    new Wp.PageNumberType { Format = Wp.NumberFormatValues.UpperRoman, Start = 12 })));
            header.Header.Save();
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static XDocument Convert(int paginationMode) =>
        XDocument.Parse(HtmlConversionOps.ConvertToHtml(AuthoredDocument(), new HtmlConversionOptions
        {
            PaginationMode = paginationMode, RenderHeadersAndFooters = true,
        }));

    [Fact]
    public void PaginatedSimplePageFieldsCarryTheComplexFieldMarker()
    {
        var fields = Convert((int)PaginationMode.Paginated).Descendants()
            .Where(e => e.Attribute("data-field") != null).ToList();

        Assert.Equal(new[] { "PAGE", "NUMPAGES" }, fields.Select(f => (string?)f.Attribute("data-field")));
        Assert.Equal(new[] { "Arabic", null }, fields.Select(f => (string?)f.Attribute("data-field-format")));
        // The cached result and its run formatting stay inside the marker.
        Assert.All(fields, field =>
        {
            Assert.Equal("1", field.Value);
            Assert.Contains(field.Descendants(), e => e.Name.LocalName == "span" && e.Value == "1");
        });
    }

    [Fact]
    public void OtherSimpleFieldsAndUnpaginatedOutputStayUnmarked()
    {
        Assert.DoesNotContain(Convert((int)PaginationMode.None).Descendants(), e => e.Attribute("data-field") != null);

        var paginated = Convert((int)PaginationMode.Paginated);
        Assert.Contains(paginated.Descendants(), e => e.Name.LocalName == "span" && e.Value == "Cedar");
        Assert.DoesNotContain(paginated.Descendants(), e => (string?)e.Attribute("data-field") == "AUTHOR");
    }
}
