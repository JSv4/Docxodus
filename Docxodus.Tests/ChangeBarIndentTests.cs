using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

public class ChangeBarIndentTests
{
    [Theory]
    [InlineData(null, "0pt")]
    [InlineData("0", "0pt")]
    [InlineData("720", "36pt")]
    public void WordChangeBarIndent_IsAlwaysALength(string? left, string expected)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var properties = new Wp.ParagraphProperties();
            if (left != null) properties.Append(new Wp.Indentation { Left = left });
            main.Document = new Wp.Document(new Wp.Body(new Wp.Paragraph(properties,
                new Wp.Run(new Wp.Text("Cedar ledger ")),
                new Wp.InsertedRun(new Wp.Run(new Wp.Text("updated.")))
                { Id = "1", Author = "Cedar Reviewer", Date = new DateTime(2026, 9, 10, 12, 0, 0, DateTimeKind.Utc) }),
                new Wp.SectionProperties()));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("cedar-ledger.docx", stream.ToArray()),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false, RenderTrackedChanges = true,
                RevisionPresentation = RevisionPresentation.Word });
        var paragraph = html.Descendants().Single(e => e.Name.LocalName == "p");
        var declarations = ((string?)paragraph.Attribute("style") ?? "").Split(';');
        Assert.Contains("--rev-change-bar-indent: " + expected, declarations);
    }
}
