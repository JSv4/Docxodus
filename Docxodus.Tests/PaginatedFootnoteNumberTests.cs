using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A paginated note's number is its own <c>w:footnoteRef</c> run, formatted like the note's other
/// runs, as Word draws it (issue #1003); it used to be a generic number set in the page's default
/// font and raised by CSS.
/// </summary>
public class PaginatedFootnoteNumberTests
{
    private static XElement NoteItem(string fixture)
    {
        var doc = new WmlDocument(Path.Combine("../../../../TestFiles", fixture));
        var html = WmlToHtmlConverter.ConvertToHtml(doc, new WmlToHtmlConverterSettings
        {
            RenderFootnotesAndEndnotes = true,
            RenderPagination = PaginationMode.Paginated,
        });
        return html.Descendants().First(e => (string?)e.Attribute("class") == "footnote-item");
    }

    [Fact]
    public void NoteNumber_IsRenderedFromTheNotesReferenceRun()
    {
        var item = NoteItem("CA/CA008-Footnote-Reference.docx");
        var number = item.Elements().First();

        Assert.Equal("footnote-number", (string?)number.Attribute("class"));
        Assert.Equal("true", (string?)number.Attribute("data-note-run"));
        // The FootnoteReference style's superscript, applied to the run that carries the number.
        var run = Assert.Single(number.Elements());
        Assert.Equal("pt-FootnoteReference", (string?)run.Attribute("class"));
        var sup = Assert.Single(run.Elements(), e => e.Name.LocalName == "sup");
        Assert.Equal("1", sup.Value);
    }

    [Fact]
    public void NoteNumber_IsNotRepeatedInTheNoteContent()
    {
        var item = NoteItem("CA/CA008-Footnote-Reference.docx");
        var content = item.Elements().Single(e => (string?)e.Attribute("class") == "footnote-content");

        Assert.Empty(content.Descendants().Where(e => e.Name.LocalName == "sup"));
        Assert.Equal("This is a test.", string.Concat(content.DescendantNodes().OfType<XText>().Select(t => t.Value)).Trim());
    }
}
