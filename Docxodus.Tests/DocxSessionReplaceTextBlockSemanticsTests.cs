#nullable enable

using System.Linq;
using System.Xml.Linq;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// #570 — ReplaceText must honor the block-level markdown semantics its parser
/// already understands, instead of consuming the markers and dropping them.
///
/// The projector escapes literal markers on the way out (a paragraph containing
/// literal "## x" projects as "\#\# x"), so an UNescaped "## " in a payload
/// genuinely declares a heading under the projector-symmetric contract —
/// exactly as InsertParagraph already treats it. And a payload that parses to
/// more than one block must fail with a typed error rather than silently
/// truncating to the first block.
/// </summary>
public class DocxSessionReplaceTextBlockSemanticsTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string? StyleIdOf(DocxSession session, string anchorId)
    {
        var p = XElement.Parse(session.Raw.GetXml(anchorId));
        return (string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");
    }

    private static string FirstAnchor(DocxSession session) =>
        session.Project().AnchorIndex.Keys.First();

    [Fact]
    public void DS570a_ReplaceText_HeadingPayload_AppliesHeadingStyle()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstAnchor(s);

        var r = s.ReplaceText(anchor, "## Replaced heading");
        Assert.True(r.Success, r.Error?.Message);

        var modified = Assert.Single(r.Modified!);
        Assert.Equal("h", modified.Kind);
        Assert.Equal("Heading2", StyleIdOf(s, modified.Id));
        Assert.Contains("Replaced heading", s.Project().Markdown);
        Assert.DoesNotContain("## Replaced heading".Replace("## ", "\\#\\# "), s.Project().Markdown);
    }

    [Fact]
    public void DS570b_ReplaceText_QuoteAndCodePayloads_ApplyTheirStyles()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());

        var quote = s.ReplaceText(FirstAnchor(s), "> quoted line");
        Assert.True(quote.Success, quote.Error?.Message);
        Assert.Equal("Quote", StyleIdOf(s, quote.Modified![0].Id));
        Assert.Contains("quoted line", s.Project().Markdown);

        var anchors = s.Project().AnchorIndex.Keys.ToList();
        var code = s.ReplaceText(anchors[1], "```\nvar x = 1;\n```");
        Assert.True(code.Success, code.Error?.Message);
        Assert.Equal("Code", StyleIdOf(s, code.Modified![0].Id));
    }

    [Fact]
    public void DS570c_ReplaceText_PlainPayload_PreservesExistingStyle()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstAnchor(s);
        var styled = s.SetParagraphStyle(anchor, "Heading1");
        Assert.True(styled.Success, styled.Error?.Message);

        var r = s.ReplaceText(styled.Modified![0].Id, "New text without markers");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal("Heading1", StyleIdOf(s, r.Modified![0].Id));
    }

    [Fact]
    public void DS570d_ReplaceText_EscapedMarkers_StayLiteralText()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstAnchor(s);

        var r = s.ReplaceText(anchor, "\\#\\# literal hashes");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Null(StyleIdOf(s, r.Modified![0].Id));

        var p = XElement.Parse(s.Raw.GetXml(r.Modified![0].Id));
        var text = string.Concat(p.Descendants(W + "t").Select(t => (string)t));
        Assert.Equal("## literal hashes", text);
    }

    [Fact]
    public void DS570e_ReplaceText_MultiBlockPayload_FailsTypedInsteadOfTruncating()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var before = s.Project().Markdown;

        var r = s.ReplaceText(FirstAnchor(s), "first block\n\nsecond block");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.UnsupportedMarkdownSyntax, r.Error!.Code);
        Assert.Contains("InsertParagraph", r.Error.Message);

        // The document must be untouched — no half-applied edit behind a failure.
        Assert.Equal(before, s.Project().Markdown);
    }

    [Fact]
    public void DS570f_ReplaceText_Tracked_HeadingPayload_EmitsPPrChange()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        s.SetTrackedChanges(TrackedChangeMode.RenderInline);
        var anchor = FirstAnchor(s);

        var r = s.ReplaceText(anchor, "## Tracked heading");
        Assert.True(r.Success, r.Error?.Message);

        var p = XElement.Parse(s.Raw.GetXml(r.Modified![0].Id));
        Assert.Equal("Heading2",
            (string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val"));
        Assert.NotNull(p.Element(W + "pPr")?.Element(W + "pPrChange"));
        Assert.NotEmpty(p.Descendants(W + "del"));
        Assert.NotEmpty(p.Descendants(W + "ins"));
    }

    [Fact]
    public void DS570g_ReplaceText_Tracked_HeadingNeedingSynthesis_IsRefused()
    {
        // The fixture defines only Normal: applying Heading2 would synthesize a
        // style, and tracked pPrChange has no styles-part before-image — the
        // same refusal SetParagraphStyle already gives.
        using var s = new DocxSession(DocxSessionTests.BuildDocWithoutCodeStyle());
        s.SetTrackedChanges(TrackedChangeMode.RenderInline);
        var before = s.Project().Markdown;

        var r = s.ReplaceText(FirstAnchor(s), "## Needs synthesis");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, r.Error!.Code);
        Assert.Equal(before, s.Project().Markdown);
    }

    [Fact]
    public void DS570h_ReplaceText_HeadingSynthesis_UndoRestoresStylesPart()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDocWithoutCodeStyle());
        var anchor = FirstAnchor(s);

        var r = s.ReplaceText(anchor, "## Synthesized");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal("Heading2", StyleIdOf(s, r.Modified![0].Id));

        Assert.True(s.Undo(), "undo must succeed");
        Assert.Contains("First paragraph.", s.Project().Markdown);
        // The synthesized Heading2 definition must be gone from the styles part.
        using var ms = new System.IO.MemoryStream(s.Save());
        using var doc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
        var styles = XElement.Load(doc.MainDocumentPart!.StyleDefinitionsPart!.GetStream());
        Assert.DoesNotContain(styles.Elements(W + "style"),
            st => (string?)st.Attribute(W + "styleId") == "Heading2");
    }
}
