using System.Text.Json;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionTextValidationTests
{
    [Theory]
    [MemberData(nameof(InvalidCharacters))]
    public void InvalidTextIsRefusedBeforeItCanBreakSaveOrTransactions(bool emitPatch, int codePoint)
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes(),
            new DocxSessionSettings { EmitMarkdownPatch = emitPatch });
        var anchor = session.Project().AnchorIndex.Keys.First(id => id.StartsWith("p:body:", StringComparison.Ordinal));
        var version = session.Version;

        var result = session.ReplaceText(anchor, "bad" + (char)codePoint + "payload");

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.MalformedMarkdown, result.Error?.Code);
        Assert.Equal(version, session.Version);
        Assert.Null(session.LastInternalError);
        Assert.False(session.Undo());
        using (session.BeginTransaction()) { }
        using var reopened = new DocxSession(session.Save());
        Assert.True(reopened.ReplaceText(reopened.Project().AnchorIndex.Keys.First(), "Still editable").Success);
    }

    public static IEnumerable<object[]> InvalidCharacters =>
        from emitPatch in new[] { false, true }
        from codePoint in new[] { 0, 1, 11, 12, 0xd800, 0xdc00, 0xfffe, 0xffff }
        select new object[] { emitPatch, codePoint };

    public static IEnumerable<object[]> TextOperations =>
        from emitPatch in new[] { false, true }
        from operation in new[] { "replace", "range", "span", "insert-span", "formatted-span", "paragraph",
            "footnote", "endnote", "header", "footer", "comment", "reply", "update-comment", "cell", "toc-title", "toa-separator" }
        select new object[] { emitPatch, operation };

    [Theory]
    [MemberData(nameof(TextOperations))]
    public void TextEntryPointsRefuseInvalidPayloadWithoutConsumingHistory(bool emitPatch, string operation)
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes(),
            new DocxSessionSettings { EmitMarkdownPatch = emitPatch });
        var body = session.Project().AnchorIndex.Keys.First(id => id.StartsWith("p:body:", StringComparison.Ordinal));
        Assert.True(session.ReplaceText(body, "Original text").Success);
        var target = body;
        if (operation is "reply" or "update-comment")
        {
            var comment = session.AddComment(body, null, "Reviewer", "Comment");
            Assert.True(comment.Success);
            target = comment.Created.First(a => a.Kind == "cmt").Id;
        }
        if (operation == "cell")
        {
            Assert.True(session.InsertTable(body, Position.After, 1, 1).Success);
            target = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "tc").Anchor.Id;
        }
        var before = session.Project().Markdown;
        var version = session.Version;
        const string invalid = "bad\ud800text";

        var result = operation switch
        {
            "replace" => session.ReplaceText(body, invalid),
            "range" => Assert.Single(session.ReplaceTextRange(body, "Original", invalid)),
            "span" => session.ReplaceTextAtSpan(body, 0, 8, invalid),
            "insert-span" => session.ReplaceTextAtSpan(body, 0, 0, invalid),
            "formatted-span" => session.ReplaceTextAtSpanWithFormat(body, 0, 8, invalid, new FormatOp { Bold = true }),
            "paragraph" => session.InsertParagraph(body, Position.After, invalid),
            "footnote" => session.InsertFootnote(body, 0, invalid),
            "endnote" => session.InsertEndnote(body, 0, invalid),
            "header" => session.SetHeaderText(body, HeaderFooterKind.Default, invalid),
            "footer" => session.SetFooterText(body, HeaderFooterKind.Default, invalid),
            "comment" => session.AddComment(body, null, "Reviewer", invalid),
            "reply" => session.AddCommentReply(target, "Reviewer", invalid),
            "update-comment" => session.UpdateComment(target, invalid),
            "cell" => session.ReplaceCellContent(target, invalid),
            "toc-title" => session.InsertTableOfContents(body, Position.After, new TableOfContentsOptions { Title = invalid }),
            "toa-separator" => session.InsertTableOfAuthorities(body, Position.After,
                new TableOfAuthoritiesOptions { EntryPageSeparator = invalid }),
            _ => throw new ArgumentOutOfRangeException(nameof(operation)),
        };

        Assert.Equal(EditErrorCode.MalformedMarkdown, result.Error?.Code);
        Assert.Equal(version, session.Version);
        Assert.Null(session.LastInternalError);
        Assert.Equal(before, session.Project().Markdown);
        using (session.BeginTransaction()) { }
        using var reopened = new DocxSession(session.Save(persistAnchorIds: true));
        Assert.Equal(before, reopened.Project().Markdown);
        Assert.True(session.Undo());
        Assert.True(session.Redo());
        Assert.Equal(before, session.Project().Markdown);
    }

    [Theory]
    [InlineData("plain")]
    [InlineData("rich")]
    [InlineData("child")]
    [InlineData("date")]
    [InlineData("combo")]
    public void ContentControlTextIsValidatedBeforeAnyPayloadIsChanged(string operation)
    {
        using var session = new DocxSession(DocxSessionContentControlTests.BuildFixture(),
            new DocxSessionSettings { EmitMarkdownPatch = false });
        var controls = session.ListContentControls();
        var outer = controls.Single(c => c.Tag == "outer-tag");
        var child = Assert.Single(outer.NestedControlAnchorIds);
        var before = session.Save();
        const string invalid = "bad\udc00text";
        var preserve = new ContentControlFillOptions { NestedControls = ContentControlNestedPolicy.Preserve };

        var result = operation switch
        {
            "plain" => session.FillContentControlText(child, invalid),
            "rich" => session.FillContentControlRichText(outer.AnchorId, invalid, preserve),
            "child" => session.FillContentControlText(outer.AnchorId, "Valid outer text",
                new ContentControlFillOptions { NestedControls = ContentControlNestedPolicy.Preserve,
                    ChildFills = new Dictionary<string, string> { [child] = invalid } }),
            "date" => session.SetContentControlDate(controls.Single(c => c.Type == ContentControlType.Date).AnchorId,
                DateTimeOffset.UnixEpoch, invalid),
            "combo" => session.SelectContentControlItem(controls.Single(c => c.Type == ContentControlType.ComboBox).AnchorId, invalid),
            _ => throw new ArgumentOutOfRangeException(nameof(operation)),
        };

        Assert.Equal(EditErrorCode.MalformedMarkdown, result.Error?.Code);
        Assert.Null(session.LastInternalError);
        Assert.False(session.Undo());
        Assert.Equal(before, session.Save());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InvalidPayloadPreservesRedoAndABoundedUndoHistory(bool emitPatch)
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes(),
            new DocxSessionSettings { EmitMarkdownPatch = emitPatch, UndoDepth = 1 });
        var body = session.Project().AnchorIndex.Keys.First();
        Assert.True(session.ReplaceText(body, "Valid edit").Success);
        Assert.False(session.ReplaceText(body, "bad\0text").Success);
        Assert.True(session.Undo());
        Assert.False(session.ReplaceText(body, "bad\ud800text").Success);
        Assert.True(session.Redo());
        Assert.Contains("Valid edit", session.Project().Markdown);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ValidSupplementaryCharactersAndXmlWhitespaceStillSave(bool emitPatch)
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes(),
            new DocxSessionSettings { EmitMarkdownPatch = emitPatch });
        var body = session.Project().AnchorIndex.Keys.First();
        Assert.True(session.ReplaceText(body, "Text 🙂 𝄞").Success);
        Assert.True(session.ReplaceTextAtSpan(body, 0, 0, "\t\r\n🙂").Success);
        using var reopened = new DocxSession(session.Save());
        Assert.Contains("🙂", reopened.Project().Markdown);
        Assert.Contains("𝄞", reopened.Project().Markdown);
    }

    [Theory]
    [InlineData(1, 1)]
    [InlineData(2, 1)]
    [InlineData(2, 0)]
    public void SpanReplacementCannotLeaveHalfOfAnExistingSurrogatePair(int start, int length)
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes(),
            new DocxSessionSettings { EmitMarkdownPatch = false });
        var body = session.Project().AnchorIndex.Keys.First();
        Assert.True(session.ReplaceText(body, "A🙂B").Success);
        var version = session.Version;

        var result = session.ReplaceTextAtSpan(body, start, length, "x");

        Assert.Equal(EditErrorCode.OffsetOutOfRange, result.Error?.Code);
        Assert.Equal(version, session.Version);
        Assert.Contains("A🙂B", session.Project().Markdown);
        Assert.True(session.ReplaceTextAtSpan(body, 1, 2, "x").Success);
        Assert.NotEmpty(session.Save());
    }

    [Fact]
    public void FacadeReturnsTheSameBusinessErrorWithPatchesDisabled()
    {
        var handle = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(),
            new DocxSessionSettings { EmitMarkdownPatch = false });
        try
        {
            var session = SessionRegistry.Get(handle);
            var anchor = session.Project().AnchorIndex.Keys.First();
            using var result = JsonDocument.Parse(DocxSessionOps.ReplaceText(handle, anchor, "bad\ud800text"));
            Assert.Equal("malformed_markdown", result.RootElement.GetProperty("error").GetProperty("code").GetString());
            Assert.NotEmpty(session.Save());
        }
        finally { DocxSessionOps.CloseSession(handle); }
    }
}
