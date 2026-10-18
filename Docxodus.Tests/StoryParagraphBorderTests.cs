using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace Docxodus.Tests;

public class StoryParagraphBorderTests
{
    [Theory]
    [InlineData("footnote", false, true)]
    [InlineData("footnote", true, true)]
    [InlineData("endnote", false, true)]
    [InlineData("endnote", true, true)]
    [InlineData("footnote", false, false)]
    [InlineData("footnote", true, false)]
    [InlineData("endnote", false, false)]
    [InlineData("endnote", true, false)]
    public void StoryBordersGroupParagraphsWithoutMovingNoteMarkers(string kind, bool paginated, bool bordered)
    {
        var path = Path.Combine("../../../../TestFiles/StoryParagraphBorders",
            $"paragraphs-story-borders-{kind}{(bordered ? "" : "-control")}.docx");
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument(path), new WmlToHtmlConverterSettings
        {
            StampCanonicalSourceAnchors = paginated, FabricateCssClasses = false, RenderHeadersAndFooters = true, RenderFootnotesAndEndnotes = true,
            RenderPagination = paginated ? PaginationMode.Paginated : PaginationMode.None,
        });
        XElement Paragraph(string marker) => Assert.Single(html.Descendants(),
            e => e.Name.LocalName == "p" && e.Value.Contains(marker));
        foreach (var story in new[] { "Body", "Header", "Footer", "Note" })
        {
            var first = Paragraph($"Cedar{story}One");
            var second = Paragraph($"Cedar{story}Two");
            var wrapper = first.Parent!;
            Assert.Equal(bordered, wrapper.Attribute("data-docx-exact-border") != null);
            if (bordered)
            {
                Assert.Same(wrapper, second.Parent);
                Assert.Contains("border-left: solid #246B7D 0.75pt", (string?)wrapper.Attribute("style") ?? "");
            }
            Assert.Null(Paragraph($"Cedar{story}Plain").Parent!.Attribute("data-docx-exact-border"));
        }
        var noteFirst = Paragraph("CedarNoteOne");
        if (paginated)
        {
            var number = Assert.Single(html.Descendants(), e => (string?)e.Attribute("class") == kind + "-number");
            Assert.Same(noteFirst, number.Ancestors().First(e => e.Name.LocalName == "p"));
            Assert.StartsWith(number.Value, noteFirst.Value);
            Assert.DoesNotContain(html.Descendants(), e => (string?)e.Attribute("class") == (kind == "footnote" ? "fn" : "en") + "-backref");
        }
        else
        {
            var backref = Assert.Single(html.Descendants(), e => (string?)e.Attribute("class") == (kind == "footnote" ? "fn" : "en") + "-backref");
            Assert.Same(Paragraph("CedarNotePlain"), backref.Parent);
        }
    }
}
