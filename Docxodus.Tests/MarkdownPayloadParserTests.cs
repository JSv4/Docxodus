#nullable enable

// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Xml.Linq;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class MarkdownPayloadParserTests
{
    [Fact]
    public void DS010_PlainText()
    {
        var r = MarkdownPayloadParser.Parse("Hello world.");
        Assert.True(r.Success);
        Assert.Single(r.Blocks);
        var text = string.Concat(r.Blocks[0].RunElements.Descendants(W.t).Select(t => (string)t));
        Assert.Equal("Hello world.", text);
    }

    [Fact]
    public void DS011_EscapedAsterisk()
    {
        var r = MarkdownPayloadParser.Parse(@"Not \*bold\*.");
        Assert.True(r.Success);
        var text = string.Concat(r.Blocks[0].RunElements.Descendants(W.t).Select(t => (string)t));
        Assert.Equal("Not *bold*.", text);
    }

    [Fact]
    public void DS012_Bold()
    {
        var r = MarkdownPayloadParser.Parse("This is **bold** text.");
        Assert.True(r.Success);
        var boldRun = r.Blocks[0].RunElements.Single(run => run.Element(W.rPr)?.Element(W.b) is not null);
        Assert.Equal("bold", (string)boldRun.Element(W.t)!);
    }

    [Fact]
    public void DS013_Italic()
    {
        var r = MarkdownPayloadParser.Parse("This is *italic* text.");
        Assert.True(r.Success);
        var italicRun = r.Blocks[0].RunElements
            .Single(run => run.Element(W.rPr)?.Element(W.i) is not null);
        Assert.Equal("italic", (string)italicRun.Element(W.t)!);
    }

    [Fact]
    public void DS014_Code()
    {
        var r = MarkdownPayloadParser.Parse("Inline `code` here.");
        Assert.True(r.Success);
        var codeRun = r.Blocks[0].RunElements.Single(run =>
            (string?)run.Element(W.rPr)?.Element(W.rStyle)?.Attribute(W.val) == "Code");
        Assert.Equal("code", (string)codeRun.Element(W.t)!);
    }

    [Fact]
    public void DS015_Strike()
    {
        var r = MarkdownPayloadParser.Parse("Some ~~struck~~ text.");
        Assert.True(r.Success);
        var s = r.Blocks[0].RunElements.Single(run => run.Element(W.rPr)?.Element(W.strike) is not null);
        Assert.Equal("struck", (string)s.Element(W.t)!);
    }

    [Fact]
    public void DS016_Link()
    {
        var r = MarkdownPayloadParser.Parse("Visit [Docxodus](https://example.com/d) today.");
        Assert.True(r.Success);
        var link = r.Blocks[0].RunElements.Single(e => e.Name == W.hyperlink);
        Assert.Equal("Docxodus", string.Concat(link.Descendants(W.t).Select(t => (string)t)));
        Assert.Equal("https://example.com/d", (string?)link.Attribute(MarkdownPayloadParser.HrefAttr));
    }

    [Fact]
    public void DS017_Headings()
    {
        var r = MarkdownPayloadParser.Parse("# H1\n\n## H2\n\n###### H6");
        Assert.True(r.Success);
        Assert.Equal(3, r.Blocks.Count);
        Assert.Equal(ParserBlockKind.Heading1, r.Blocks[0].Kind);
        Assert.Equal(ParserBlockKind.Heading2, r.Blocks[1].Kind);
        Assert.Equal(ParserBlockKind.Heading6, r.Blocks[2].Kind);
    }

    [Fact]
    public void DS018_Blockquote()
    {
        var r = MarkdownPayloadParser.Parse("> Quoted text.");
        Assert.True(r.Success);
        Assert.Equal(ParserBlockKind.Quote, r.Blocks[0].Kind);
    }

    [Fact]
    public void DS019_FencedCode()
    {
        var r = MarkdownPayloadParser.Parse("```\ncode line 1\ncode line 2\n```");
        Assert.True(r.Success);
        Assert.Equal(ParserBlockKind.Code, r.Blocks[0].Kind);
        var text = string.Concat(r.Blocks[0].RunElements.Descendants(W.t).Select(t => (string)t));
        Assert.Contains("code line 1", text);
        Assert.Contains("code line 2", text);
    }

    [Fact]
    public void DS020_BulletedList()
    {
        var r = MarkdownPayloadParser.Parse("- First\n- Second\n- Third");
        Assert.True(r.Success);
        Assert.Equal(3, r.Blocks.Count);
        Assert.All(r.Blocks, b => Assert.Equal(ParserBlockKind.BulletItem, b.Kind));
    }

    [Fact]
    public void DS021_OrderedList()
    {
        var r = MarkdownPayloadParser.Parse("1. One\n2. Two");
        Assert.True(r.Success);
        Assert.Equal(2, r.Blocks.Count);
        Assert.All(r.Blocks, b => Assert.Equal(ParserBlockKind.OrderedItem, b.Kind));
    }

    [Fact]
    public void DS022_NestedList()
    {
        var r = MarkdownPayloadParser.Parse("- Top\n  - Nested\n  - Also nested\n- Top again");
        Assert.True(r.Success);
        Assert.Equal(4, r.Blocks.Count);
        Assert.Equal(0, r.Blocks[0].ListLevel);
        Assert.Equal(1, r.Blocks[1].ListLevel);
        Assert.Equal(1, r.Blocks[2].ListLevel);
        Assert.Equal(0, r.Blocks[3].ListLevel);
    }

    [Theory]
    [InlineData("| col1 | col2 |\n|---|---|\n| a | b |", EditErrorCode.TableInsertNotSupported)]
    [InlineData("See [^fn-abc].", EditErrorCode.FootnoteRefNotSupported)]
    [InlineData("Hello {#cmt:cmt:abcd} there.", EditErrorCode.CommentMarkerNotSupported)]
    [InlineData("![alt](docxodus://img/abcd)", EditErrorCode.ImageInsertNotSupported)]
    [InlineData("Hello {#p:body:abcd} there.", EditErrorCode.AnchorTokenInPayload)]
    public void DS023_RejectionCodes(string payload, EditErrorCode expected)
    {
        var r = MarkdownPayloadParser.Parse(payload);
        Assert.False(r.Success);
        Assert.Equal(expected, r.Error!.Code);
    }

    [Fact]
    public void DS024_NullPayload()
    {
        var r = MarkdownPayloadParser.Parse(null!);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.MalformedMarkdown, r.Error!.Code);
    }
}
