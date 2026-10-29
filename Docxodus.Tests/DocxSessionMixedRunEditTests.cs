using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionMixedRunEditTests
{
    [Theory]
    [InlineData("span", false, "tab")]
    [InlineData("span", true, "tab")]
    [InlineData("range", false, "tab")]
    [InlineData("range", true, "tab")]
    [InlineData("match", false, "tab")]
    [InlineData("match", true, "tab")]
    [InlineData("span", false, "br")]
    [InlineData("span", true, "br")]
    [InlineData("range", false, "br")]
    [InlineData("range", true, "br")]
    [InlineData("match", false, "br")]
    [InlineData("match", true, "br")]
    public void Replacement_PreservesAnInlineMarkerBetweenTextNodes(string operation, bool tracked, string marker)
    {
        // A tab is one character of the text (issue #1387), so a range across it takes it, and
        // keeps it only when the replacement writes it back. A break also takes one character.
        int c = TextOf(marker).IndexOf('c');
        var across = marker == "tab" ? "X\t" : "X\n";
        foreach (var (start, length, expected, replacement) in new[]
            { (1, 1, "aX|cd", "X"), (c, 1, "ab|Xd", "X"), (1, c, "aX|d", across) })
            Check(operation, tracked, marker, start, length, expected, replacement);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Control_AdjacentTextNodesRemainEditable(bool tracked) =>
        Check("span", tracked, null, 1, 2, "aXd");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RemovingAllText_ConsumesTheBreak(bool tracked) =>
        Check("span", tracked, "br", 0, 5, "", "");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RangeReplacement_UpdatesEveryMatchInTheOriginalMixedRun(bool tracked)
    {
        using var session = new DocxSession(DocxSessionSpanRevisionTests.Build(new XElement(W.r,
            new XElement(W.t, "ab"), new XElement(W.tab), new XElement(W.t, "ab"),
            new XElement(W.br), new XElement(W.t, "ab"))), new DocxSessionSettings
        {
            TrackedChanges = tracked ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept,
        });
        var anchor = session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p").Anchor.Id;
        var edits = session.ReplaceTextRange(anchor, "b", "B");
        Assert.Equal(3, edits.Count);
        Assert.All(edits, edit => Assert.True(edit.Success, edit.Error?.Message));
        Assert.Equal("aB|aB|aB", Text(RevisionProcessor.AcceptRevisions(new WmlDocument("a.docx", session.Save())).DocumentByteArray));
        Assert.Equal(tracked ? "ab|ab|ab" : "aB|aB|aB",
            Text(RevisionProcessor.RejectRevisions(new WmlDocument("r.docx", session.Save())).DocumentByteArray));
    }

    private static void Check(string operation, bool tracked, string? marker, int start, int length, string expected,
        string replacement = "X")
    {
        var input = DocxSessionSpanRevisionTests.Build(new XElement(W.r,
            new XElement(W.rPr, new XElement(W.b)), new XElement(W.t, "ab"),
            marker is null ? null : new XElement(W.w + marker), new XElement(W.t, "cd")));
        PackageValidation.AssertValid(input, FileFormatVersions.Office2019);
        using var session = new DocxSession(input, new DocxSessionSettings
        {
            TrackedChanges = tracked ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept,
            RevisionAuthor = "Editor",
        });
        var anchor = session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p").Anchor.Id;
        var edit = operation switch
        {
            "range" => Assert.Single(session.ReplaceTextRange(anchor, TextOf(marker).Substring(start, length), replacement)),
            "match" => session.ReplaceMatch(Assert.Single(session.Grep(TextOf(marker).Substring(start, length))), replacement),
            _ => session.ReplaceTextAtSpan(anchor, start, length, replacement),
        };
        Assert.True(edit.Success, edit.Error?.Message);
        var saved = session.Save();
        PackageValidation.AssertValid(saved, FileFormatVersions.Office2019);
        Assert.Equal(expected, Text(RevisionProcessor.AcceptRevisions(new WmlDocument("accepted.docx", saved)).DocumentByteArray));
        Assert.Equal(tracked ? (marker is null ? "abcd" : "ab|cd") : expected,
            Text(RevisionProcessor.RejectRevisions(new WmlDocument("rejected.docx", saved)).DocumentByteArray));
        using var accepted = new DocxSession(saved);
        Assert.True(accepted.AcceptAllRevisions().Success);
        Assert.Equal(expected, Text(accepted.Save()));
        using var rejected = new DocxSession(saved);
        Assert.True(rejected.RejectAllRevisions().Success);
        Assert.Equal(tracked ? (marker is null ? "abcd" : "ab|cd") : expected, Text(rejected.Save()));
        Assert.True(session.Undo());
        Assert.Equal(marker is null ? "abcd" : "ab|cd", Text(session.Save()));
    }

    /// <summary>The paragraph's text as the session counts it: a tab is '\t', a break '\n'.</summary>
    private static string TextOf(string? marker) => marker switch { "tab" => "ab\tcd", "br" => "ab\ncd", _ => "abcd" };

    private static string Text(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return string.Concat(document.MainDocumentPart!.GetXDocument().Descendants(W.r).SelectMany(r => r.Elements())
            .Select(e => e.Name == W.t ? e.Value : e.Name == W.tab || e.Name == W.br ? "|" : ""));
    }
}
