using System.Text.RegularExpressions;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionInlineTextTests
{
    [Theory]
    [InlineData("<w:tab/>", "\t", "\t")]
    [InlineData("<w:br/>", "\n", "  \n")]
    [InlineData("<w:cr/>", "\n", "  \n")]
    [InlineData("<w:noBreakHyphen/>", "\u2011", "\u2011")]
    [InlineData("<w:softHyphen/>", "\u00ad", "\u00ad")]
    [InlineData("<w:sym w:font=\"Symbol\" w:char=\"03A9\"/>", "\u03a9", "\u03a9")]
    [InlineData("<w:ptab w:alignment=\"right\" w:relativeTo=\"margin\" w:leader=\"none\"/>", "\t", "\t")]
    public void Visible_run_atoms_are_searchable_and_consumed_by_span_edits(string atom, string text, string markdown)
    {
        using var session = new DocxSession(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>A</w:t>" + atom + "<w:t>B</w:t></w:r></w:p>").DocumentByteArray);
        var projection = session.Project();
        var anchor = Assert.Single(projection.AnchorIndex.Values, a => a.Anchor.Kind == "p").Anchor.Id;

        Assert.Single(session.Grep("^" + Regex.Escape("A" + text + "B") + "$"));
        Assert.Contains("} A" + markdown + "B", projection.Markdown);
        var edit = session.ReplaceTextAtSpan(anchor, 1, 1, "X");

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Single(session.Grep("^AXB$"));
        Assert.Contains("} AXB", session.Project().Markdown);
    }
}
