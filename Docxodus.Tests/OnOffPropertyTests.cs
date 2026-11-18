using System.Linq;
using System.Xml.Linq;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// An on/off property's <c>w:val</c> (ST_OnOff) is read the same way by every Docxodus reader: <c>1</c>,
/// <c>true</c> and <c>on</c> are on and <c>0</c>, <c>false</c> and <c>off</c> are off, in any letter case
/// (issue #1786).
/// </summary>
public class OnOffPropertyTests
{
    [Theory]
    [InlineData(null, true)]
    [InlineData("1", true)]
    [InlineData("true", true)]
    [InlineData("on", true)]
    [InlineData("On", true)]
    [InlineData("TRUE", true)]
    [InlineData("0", false)]
    [InlineData("false", false)]
    [InlineData("off", false)]
    [InlineData("Off", false)]
    [InlineData("maybe", false)]
    public void GetBoolProp_reads_every_spelling_of_on_as_ToBoolean_does(string? value, bool expected)
    {
        var bold = new XElement(W.b);
        if (value is not null)
            bold.Add(new XAttribute(W.val, value));
        var properties = new XElement(W.rPr, bold);

        Assert.Equal(expected, WordprocessingMLUtil.GetBoolProp(properties, W.b));
        if (value is not null && value != "maybe")
            Assert.Equal(expected, bold.Attribute(W.val).ToBoolean());
    }

    [Theory]
    [InlineData("1", true)]
    [InlineData("true", true)]
    [InlineData("on", true)]
    [InlineData("On", true)]
    [InlineData("TRUE", true)]
    [InlineData("0", false)]
    [InlineData("false", false)]
    [InlineData("off", false)]
    [InlineData("OFF", false)]
    [InlineData("maybe", null)]
    [InlineData("", null)]
    [InlineData(null, null)]
    public void OnOffValue_reads_the_six_spellings_in_any_case(string? value, bool? expected) =>
        Assert.Equal(expected, WordprocessingMLUtil.OnOffValue(value));

    [Theory]
    [InlineData(" w:val=\"off\"", "Ferry timetable")]
    [InlineData(" w:val=\"Off\"", "Ferry timetable")]
    [InlineData(" w:val=\"On\"", "**Ferry timetable**")]
    [InlineData("", "**Ferry timetable**")]
    public void The_markdown_projection_reads_every_spelling_of_bold(string value, string expected)
    {
        var doc = IrTestDocuments.FromBodyXml($"<w:p><w:r><w:rPr><w:b{value}/></w:rPr><w:t>Ferry timetable</w:t></w:r></w:p>");
        var markdown = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings()).Markdown;
        Assert.Contains(expected, markdown);
        if (!expected.StartsWith("**"))
            Assert.DoesNotContain("**", markdown);
    }

    [Theory]
    [InlineData("True", true)]
    [InlineData("ON", true)]
    [InlineData("0", false)]
    [InlineData("maybe", false)]
    public void A_comment_resolved_flag_reads_every_spelling_of_on(string value, bool expected) =>
        Assert.Equal(expected, Docxodus.Internal.CommentOps.ParseDone(value));

    [Fact]
    public void A_toggle_written_FALSE_and_one_written_false_are_the_same_formatting()
    {
        // The IR reads a run's toggles; two spellings of off are no formatting change.
        static WmlDocument Bold(string value) => IrTestDocuments.FromBodyXml(
            $"<w:p><w:r><w:rPr><w:b w:val=\"{value}\"/></w:rPr><w:t>Ferry timetable</w:t></w:r></w:p>");
        Assert.Empty(DocxDiff.GetRevisions(Bold("FALSE"), Bold("false")));
        Assert.NotEmpty(DocxDiff.GetRevisions(Bold("TRUE"), Bold("false")));
    }

    [Fact]
    public void GetBoolProp_reads_an_absent_property_as_off() =>
        Assert.False(WordprocessingMLUtil.GetBoolProp(new XElement(W.rPr), W.b));

    [Theory]
    [InlineData(" w:val=\"on\"", 1)]
    [InlineData(" w:val=\"1\"", 1)]
    [InlineData(" w:val=\"off\"", 2)]
    public void A_paragraph_mark_hidden_with_any_spelling_of_on_joins_the_next_paragraph(string value, int paragraphs)
    {
        // A hidden paragraph mark folds its paragraph into the next one, whichever spelling turns it on.
        var doc = IrTestDocuments.FromBodyXml(
            $"<w:p><w:pPr><w:rPr><w:vanish{value}/></w:rPr></w:pPr><w:r><w:t>Cedar.</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t xml:space=\"preserve\"> Return the cedar folder before dusk.</w:t></w:r></w:p>");

        var html = WmlToHtmlConverter.ConvertToHtml(doc, new WmlToHtmlConverterSettings());
        var blocks = html.Descendants().Where(e => e.Name.LocalName == "p").ToList();

        Assert.Equal(paragraphs, blocks.Count);
        Assert.Contains("Return the cedar folder", blocks.Last().Value);
    }
}
