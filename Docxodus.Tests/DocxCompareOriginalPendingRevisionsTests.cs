// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// What the <see cref="DocxCompare"/> front door does with tracked changes the ORIGINAL already carries
/// from another author (issue #845). It compares the accepted view of each document, so an original's
/// pending change is resolved as accepted before the diff: nothing by the earlier author survives, only
/// the differences between the two accepted views are marked (by the compare author), and the result
/// round-trips — accept all gives the revised document's accepted view, reject all the original's.
/// Word's compare instead keeps an original's pending insertions marked as insertions (so its accept-all
/// can differ from the revised document); that difference is deliberate and documented in
/// <c>docs/ooxml_corner_cases.md</c>.
/// </summary>
public class DocxCompareOriginalPendingRevisionsTests
{
    private const string Rev = "w:author=\"Alice\" w:date=\"2026-01-01T00:00:00Z\"";

    private static string Run(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";
    private static string Ins(string text) => $"<w:ins w:id=\"1\" {Rev}>{Run(text)}</w:ins>";
    private static string Del(string text) =>
        $"<w:del w:id=\"2\" {Rev}><w:r><w:delText xml:space=\"preserve\">{text}</w:delText></w:r></w:del>";
    private static WmlDocument Paragraph(string inner) => IrTestDocuments.FromParts($"<w:p>{inner}</w:p>");

    private static XElement Body(byte[] bytes)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return package.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    private static string Text(byte[] bytes) =>
        string.Concat(Body(bytes).Descendants(W.t).Select(t => t.Value));

    /// <summary>The body text with each tracked change drawn as [+inserted] or [-deleted].</summary>
    private static string Markup(byte[] bytes) => string.Concat(Body(bytes).Descendants(W.p).First()
        .Elements().Where(e => e.Name != W.pPr).Select(e =>
            e.Name == W.ins ? $"[+{string.Concat(e.Descendants(W.t).Select(t => t.Value))}]" :
            e.Name == W.del ? $"[-{string.Concat(e.Descendants(W.delText).Select(t => t.Value))}]" :
            string.Concat(e.Descendants(W.t).Select(t => t.Value))));

    public static TheoryData<string, string, string, string, string, string> Scenarios => new()
    {
        // name, original body, revised body, expected markup, expected accept-all, expected reject-all
        { "pending insertion, revised accepted it",
            Run("The quick ") + Ins("brown ") + Run("fox."), Run("The quick brown fox."),
            "The quick brown fox.", "The quick brown fox.", "The quick brown fox." },
        { "pending insertion, revised rejected it",
            Run("The quick ") + Ins("brown ") + Run("fox."), Run("The quick fox."),
            "The quick [-brown ]fox.", "The quick fox.", "The quick brown fox." },
        { "pending insertion, revised accepted it and edited",
            Run("The quick ") + Ins("brown ") + Run("fox."), Run("The quick brown fox jumps."),
            "The quick brown fox[+ jumps].", "The quick brown fox jumps.", "The quick brown fox." },
        { "pending insertion, revised still pending",
            Run("The quick ") + Ins("brown ") + Run("fox."), Run("The quick ") + Ins("brown ") + Run("fox."),
            "The quick brown fox.", "The quick brown fox.", "The quick brown fox." },
        { "pending deletion, revised accepted it",
            Run("The ") + Del("lazy ") + Run("dog."), Run("The dog."),
            "The dog.", "The dog.", "The dog." },
        { "pending deletion, revised rejected it",
            Run("The ") + Del("lazy ") + Run("dog."), Run("The lazy dog."),
            "The [+lazy ]dog.", "The lazy dog.", "The dog." },
        { "pending deletion, revised accepted it and edited",
            Run("The ") + Del("lazy ") + Run("dog."), Run("The dog barks."),
            "The dog[+ barks].", "The dog barks.", "The dog." },
    };

    [Theory]
    [MemberData(nameof(Scenarios))]
    public void FrontDoor_ResolvesTheOriginalsPendingChangesAsAccepted(
        string scenario, string originalBody, string revisedBody,
        string expectedMarkup, string expectedAccepted, string expectedRejected)
    {
        _ = scenario;
        var redline = DocxCompare.Compare(Paragraph(originalBody), Paragraph(revisedBody)).DocumentByteArray;

        Assert.Equal(expectedMarkup, Markup(redline));
        Assert.Equal(expectedAccepted, Text(DocxDiffOps.AcceptRevisions(redline)));
        Assert.Equal(expectedRejected, Text(DocxDiffOps.RejectRevisions(redline)));
        // No change by the earlier author survives; every remaining change is the comparison's own.
        Assert.DoesNotContain(Body(redline).Descendants().Attributes(W.author), author => author.Value == "Alice");
    }
}
