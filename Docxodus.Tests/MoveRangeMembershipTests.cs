using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A move range (<c>w:moveFromRangeStart</c>…<c>w:moveFromRangeEnd</c>, and the <c>w:moveTo…</c> pair) is a
/// named container: ECMA-376 §17.13.5.24 and §17.13.5.28 link content explicitly marked as moved. Content inside the range that carries no
/// <c>w:moveFrom</c>/<c>w:moveTo</c> mark is ordinary content and survives both accept and reject. Both of
/// Docxodus's resolvers, <see cref="RevisionProcessor"/> and <see cref="DocxSession"/>, must agree on that,
/// and on a move whose content is split across several ranges sharing one name (issue #1340).
/// </summary>
public class MoveRangeMembershipTests
{
    private const string Stamp = "w:author=\"Alex\" w:date=\"2026-05-01T09:00:00Z\"";

    private static string Run(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private static string FromStart(int id) => $"<w:moveFromRangeStart w:id=\"{id}\" w:name=\"move1\" {Stamp}/>";
    private static string FromEnd(int id) => $"<w:moveFromRangeEnd w:id=\"{id}\"/>";
    private static string ToStart(int id) => $"<w:moveToRangeStart w:id=\"{id}\" w:name=\"move1\" {Stamp}/>";
    private static string ToEnd(int id) => $"<w:moveToRangeEnd w:id=\"{id}\"/>";
    private static string MoveFrom(int id, string text) =>
        $"<w:moveFrom w:id=\"{id}\" {Stamp}><w:r><w:delText xml:space=\"preserve\">{text}</w:delText></w:r></w:moveFrom>";
    private static string MoveTo(int id, string text) => $"<w:moveTo w:id=\"{id}\" {Stamp}>{Run(text)}</w:moveTo>";
    private static string MovedToMark(int id) => $"<w:pPr><w:rPr><w:moveTo w:id=\"{id}\" {Stamp}/></w:rPr></w:pPr>";

    // "amount" moves out of "Old: amount end" into "Fee  due".
    private static readonly string Source =
        "<w:p>" + Run("Old: ") + FromStart(1) + MoveFrom(2, "amount") + FromEnd(1) + Run(" end") + "</w:p>";

    private static string Destination(string rangeBody) =>
        "<w:p>" + Run("Fee ") + rangeBody + Run(" due") + "</w:p>";

    public static TheoryData<string, string, string[], string[]> UnmarkedContentInsideARange => new()
    {
        {
            // The run typed at the end of the moved text stepped out of w:moveTo but not past the range end.
            "untracked text after the w:moveTo, before the range end",
            Source + Destination(ToStart(3) + MoveTo(4, "amount") + Run("XY") + ToEnd(3)),
            new[] { "Old:  end", "Fee amountXY due" },
            new[] { "Old: amount end", "Fee XY due" }
        },
        {
            "untracked text between two w:moveTo elements of one range",
            Source + Destination(ToStart(3) + MoveTo(4, "amo") + Run("XY") + MoveTo(6, "unt") + ToEnd(3)),
            new[] { "Old:  end", "Fee amoXYunt due" },
            new[] { "Old: amount end", "Fee XY due" }
        },
        {
            // The mirror image on the source side, which accept used to drop.
            "untracked text inside the source range",
            "<w:p>" + Run("Old: ") + FromStart(1) + MoveFrom(2, "amount") + Run("XY") + FromEnd(1) + Run(" end") + "</w:p>" +
            Destination(ToStart(3) + MoveTo(4, "amount") + ToEnd(3)),
            new[] { "Old: XY end", "Fee amount due" },
            new[] { "Old: amountXY end", "Fee  due" }
        },
        {
            // A destination range that runs on into the next paragraph: text typed at the end of the first
            // moved paragraph and at the start of the second lies inside the range.
            "untracked text inside a range that spans two paragraphs",
            "<w:p>" + Run("Old: ") + FromStart(1) + MoveFrom(2, "amount") + MoveFrom(5, "more") + FromEnd(1) + Run(" end") + "</w:p>" +
            "<w:p>" + MovedToMark(7) + Run("Fee ") + ToStart(3) + MoveTo(4, "amount") + Run("XY") + "</w:p>" +
            "<w:p>" + Run("ZW") + MoveTo(6, "more") + ToEnd(3) + Run(" due") + "</w:p>",
            new[] { "Old:  end", "Fee amountXY", "ZWmore due" },
            new[] { "Old: amountmore end", "Fee XYZW due" }
        },
    };

    [Theory]
    [MemberData(nameof(UnmarkedContentInsideARange))]
    public void RevisionProcessor_KeepsUnmarkedContentInsideAMoveRange(
        string shape, string body, string[] accepted, string[] rejected)
    {
        var document = Build(body);
        Assert.True(accepted.SequenceEqual(Texts(RevisionProcessor.AcceptRevisions(document))), shape);
        Assert.True(rejected.SequenceEqual(Texts(RevisionProcessor.RejectRevisions(document))), shape);
    }

    [Theory]
    [MemberData(nameof(UnmarkedContentInsideARange))]
    public void Session_AgreesWithRevisionProcessorOnUnmarkedContentInsideAMoveRange(
        string shape, string body, string[] accepted, string[] rejected)
    {
        var bytes = Build(body).DocumentByteArray;
        Assert.True(accepted.SequenceEqual(SessionResolveAll(bytes, accept: true)), shape);
        Assert.True(rejected.SequenceEqual(SessionResolveAll(bytes, accept: false)), shape);
    }

    public static TheoryData<string, string, string[], string[]> MoveSplitAcrossSameNamedRanges => new()
    {
        {
            "two destination ranges in one paragraph",
            Source + Destination(ToStart(3) + MoveTo(4, "amo") + ToEnd(3) + Run("XY") +
                                 ToStart(5) + MoveTo(6, "unt") + ToEnd(5)),
            new[] { "Old:  end", "Fee amoXYunt due" },
            new[] { "Old: amount end", "Fee XY due" }
        },
        {
            "two destination ranges, the first starting in an earlier paragraph",
            "<w:p>" + Run("Old: ") + FromStart(1) + MoveFrom(2, "first") + MoveFrom(8, "amount") + FromEnd(1) + Run(" end") + "</w:p>" +
            "<w:p>" + MovedToMark(9) + ToStart(3) + MoveTo(4, "first") + "</w:p>" +
            "<w:p>" + MoveTo(7, "amo") + ToEnd(3) + Run("XY") + ToStart(5) + MoveTo(6, "unt") + ToEnd(5) + Run(" due") + "</w:p>",
            new[] { "Old:  end", "first", "amoXYunt due" },
            new[] { "Old: firstamount end", "XY due" }
        },
    };

    [Theory]
    [MemberData(nameof(MoveSplitAcrossSameNamedRanges))]
    public void RevisionProcessor_ResolvesAMoveSplitAcrossSameNamedRanges(
        string shape, string body, string[] accepted, string[] rejected)
    {
        var document = Build(body);
        Assert.True(accepted.SequenceEqual(Texts(RevisionProcessor.AcceptRevisions(document))), shape);
        Assert.True(rejected.SequenceEqual(Texts(RevisionProcessor.RejectRevisions(document))), shape);
    }

    [Theory]
    [MemberData(nameof(MoveSplitAcrossSameNamedRanges))]
    public void Session_ListsAndResolvesAMoveSplitAcrossSameNamedRangesAsOneMove(
        string shape, string body, string[] accepted, string[] rejected)
    {
        var bytes = Build(body).DocumentByteArray;
        using (var session = new DocxSession(bytes))
        {
            var move = Assert.Single(session.ListRevisions());
            Assert.Equal("move", move.Type);
            Assert.True(move.ResolutionStatus == RevisionResolutionStatus.Supported,
                $"{shape}: {move.Diagnostic?.Code}");
        }

        Assert.True(accepted.SequenceEqual(SessionResolveAll(bytes, accept: true)), shape);
        Assert.True(rejected.SequenceEqual(SessionResolveAll(bytes, accept: false)), shape);
        using (var session = new DocxSession(bytes))
        {
            Assert.True(session.RejectRevision(session.ListRevisions()[0].Id).Success, shape);
            Assert.Equal(rejected, ParagraphTexts(session.Save()));
            Assert.Empty(session.ListRevisions());
        }
    }

    [Fact]
    public void Session_KeepsNestedSameNamedRangesAmbiguous()
    {
        // Two destination containers of one name that surround the same text: the spec resolves that
        // by "last valid container", which a selective resolver cannot apply safely.
        var bytes = Build(Source + Destination(
            ToStart(3) + MoveTo(4, "amo") + ToStart(5) + MoveTo(6, "unt") + ToEnd(5) + ToEnd(3))).DocumentByteArray;
        using var session = new DocxSession(bytes);

        var move = Assert.Single(session.ListRevisions());
        Assert.Equal(RevisionResolutionStatus.Ambiguous, move.ResolutionStatus);
        Assert.Equal("ambiguous_move_topology", move.Diagnostic?.Code);
    }

    [Theory]
    [InlineData("RA001-Tracked-Revisions-01.docx")]
    [InlineData("RA001-Tracked-Revisions-02.docx")]
    [InlineData("RP/RP001-Tracked-Revisions-01.docx")]
    [InlineData("RP/RP015-MoveFrom-MoveTo.docx")]
    [InlineData("FA/RevTracking/014-MovedParagraph.docx")]
    public void WordAuthoredMoves_SelectiveResolutionPreservesProcessorEndpoints(string fixture)
    {
        var document = new WmlDocument(Path.Combine("../../../../TestFiles/", fixture));
        foreach (bool accept in new[] { true, false })
        {
            using var session = new DocxSession(document.DocumentByteArray);
            var moves = session.ListRevisions().Where(revision => revision.Type == "move").ToArray();
            Assert.NotEmpty(moves);
            foreach (var move in moves)
            {
                Assert.True(move.ResolutionStatus == RevisionResolutionStatus.Supported, move.Diagnostic?.Code);
                var result = accept ? session.AcceptRevision(move.Id) : session.RejectRevision(move.Id);
                Assert.True(result.Success, result.Error?.Message);
            }
            var resolved = new WmlDocument("resolved-moves.docx", session.Save());
            Assert.Equal(
                Texts(accept ? RevisionProcessor.AcceptRevisions(document) : RevisionProcessor.RejectRevisions(document)),
                Texts(accept ? RevisionProcessor.AcceptRevisions(resolved) : RevisionProcessor.RejectRevisions(resolved)));
        }
    }

    private static List<string> SessionResolveAll(byte[] bytes, bool accept)
    {
        using var session = new DocxSession(bytes);
        var result = accept ? session.AcceptAllRevisions() : session.RejectAllRevisions();
        Assert.True(result.Success, result.Error?.Message);
        return ParagraphTexts(session.Save());
    }

    private static List<string> Texts(WmlDocument document) => ParagraphTexts(document.DocumentByteArray);

    private static List<string> ParagraphTexts(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.GetXDocument().Descendants(W.p)
            .Select(p => string.Concat(p.Descendants(W.t).Select(t => t.Value)))
            .Where(text => text.Length > 0).ToList();
    }

    private static WmlDocument Build(string bodyXml)
    {
        var xml = $"<w:document xmlns:w=\"{W.w.NamespaceName}\"><w:body>{bodyXml}" +
                  "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>";
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.PutXDocument(XDocument.Parse(xml));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles)));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
        }
        return new WmlDocument("move-range.docx", stream.ToArray());
    }
}
