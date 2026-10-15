using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A paginated note's number is its own <c>w:footnoteRef</c> run, formatted like the note's other
/// runs, as Word draws it (issue #1003); it used to be a generic number set in the page's default
/// font and raised by CSS. Like Word's run, it is the first content of the note's first paragraph,
/// so that paragraph's own indentation governs the line the number sits on (issue #1130).
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
        var number = item.Descendants().Single(e => (string?)e.Attribute("class") == "footnote-number");

        Assert.Equal("footnote-number", (string?)number.Attribute("class"));
        Assert.Equal("true", (string?)number.Attribute("data-note-run"));
        // The FootnoteReference style's superscript, applied to the run that carries the number.
        var run = Assert.Single(number.Elements());
        Assert.Equal("pt-FootnoteReference", (string?)run.Attribute("class"));
        var sup = Assert.Single(run.Elements(), e => e.Name.LocalName == "sup");
        Assert.Equal("1", sup.Value);
    }

    [Fact]
    public void NoteNumber_IsTheFirstContentOfTheFirstNoteParagraph()
    {
        var item = NoteItem("CA/CA008-Footnote-Reference.docx");
        var content = item.Elements().Single(e => (string?)e.Attribute("class") == "footnote-content");
        var paragraph = content.Descendants().First(e => e.Name.LocalName == "p");

        Assert.Equal("footnote-number", (string?)paragraph.Elements().First().Attribute("class"));
        Assert.DoesNotContain(item.Elements(), e => (string?)e.Attribute("class") == "footnote-number");
    }

    [Fact]
    public void NoteNumber_IsNotRepeatedInTheNoteContent()
    {
        var item = NoteItem("CA/CA008-Footnote-Reference.docx");
        var content = item.Elements().Single(e => (string?)e.Attribute("class") == "footnote-content");

        Assert.Single(content.Descendants(), e => e.Name.LocalName == "sup");
        var text = content.DescendantNodes().OfType<XText>()
            .Where(t => !t.Ancestors().Any(a => (string?)a.Attribute("class") == "footnote-number"));
        Assert.Equal("This is a test.", string.Concat(text.Select(t => t.Value)).Trim());
    }
}
