using System.Linq;
using System.Xml.Linq;
using Docxodus;
using DocumentFormat.OpenXml;
using Xunit;
using static Docxodus.Tests.DocxSessionClonedRevisionTests;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1313: session edits that ignored tracked-change mode. A cross-reference is recorded as
/// an insertion; a paragraph merge is recorded as a deleted paragraph mark, so accepting it gives
/// what an untracked merge gives and rejecting it gives back both paragraphs; and a split no longer
/// cuts a field in two.
/// </summary>
public class DocxSessionTrackedEditGapTests
{
    // ─── InsertCrossReference ──────────────────────────────────────────────

    private const string CrossReferenceBody =
        "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"tgt\"/><w:r><w:t>Target text</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>" +
        "<w:p><w:r><w:t xml:space=\"preserve\">Intro: </w:t></w:r></w:p>";

    private static string SecondParagraph(DocxSession session) =>
        session.Project().AnchorIndex.Values.Select(v => v.Anchor).Where(a => a.Kind == "p").Skip(1).First().Id;

    [Fact]
    public void TrackedCrossReference_IsOneInsertion_RejectRemovesItAcceptKeepsIt()
    {
        using var session = Open(CrossReferenceBody, TrackedChangeMode.RenderInline);

        var result = session.InsertCrossReference(SecondParagraph(session), 7, "tgt");

        Assert.True(result.Success, result.Error?.Message);
        var paragraph = Body(session).Elements(W.p).Last();
        Assert.Empty(paragraph.Descendants(W.fldSimple));
        Assert.All(paragraph.Descendants().Where(e => e.Name == W.fldChar || e.Name == W.instrText),
            field => Assert.NotNull(field.Ancestors(W.ins).FirstOrDefault()));
        Assert.Contains(" REF tgt ", string.Concat(paragraph.Descendants(W.instrText).Select(t => t.Value)));
        Assert.Single(session.ListRevisions());
        AssertRevisionIdsDistinct(session, expectedCount: 1);
        AssertSchemaValid(session);
        AssertOutcomes(session,
            accepted: new[] { "p: Target text", "p: Intro: |Target text" },
            rejected: new[] { "p: Target text", "p: Intro: " });
    }

    // ─── MergeParagraphs ───────────────────────────────────────────────────

    private const string TwoParagraphs =
        "<w:p><w:pPr><w:jc w:val=\"center\"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>First.</w:t></w:r></w:p>" +
        "<w:p><w:pPr><w:jc w:val=\"right\"/></w:pPr><w:r><w:t>Second.</w:t></w:r></w:p>";

    private static (string First, string Second) Paragraphs(DocxSession session)
    {
        var anchors = session.Project().AnchorIndex.Values.Select(v => v.Anchor).Where(a => a.Kind == "p").ToList();
        return (anchors[0].Id, anchors[1].Id);
    }

    [Fact]
    public void TrackedMerge_DeletesTheFirstParagraphMark_AndKeepsBothParagraphsUntilReview()
    {
        using var session = Open(TwoParagraphs, TrackedChangeMode.RenderInline);
        var (first, second) = Paragraphs(session);

        var result = session.MergeParagraphs(first, second);

        Assert.True(result.Success, result.Error?.Message);
        Assert.Empty(result.Removed);
        var paragraphs = Body(session).Elements(W.p).ToList();
        Assert.Equal(2, paragraphs.Count);
        Assert.NotNull(paragraphs[0].Element(W.pPr)!.Element(W.rPr)!.Element(W.del));
        AssertRevisionIdsDistinct(session, expectedCount: null);
        AssertSchemaValid(session);
        AssertOutcomes(session,
            accepted: new[] { "p[center]: First.| |Second." },
            rejected: new[] { "p[center]: First.", "p[right]: Second." });
    }

    /// <summary>The joined paragraph takes the first paragraph's properties, as an untracked merge
    /// does, its paragraph-mark formatting included; rejecting gives the second its own back.</summary>
    [Fact]
    public void TrackedMerge_AcceptedGivesTheUntrackedMerge_RejectedGivesBackTheSecondsProperties()
    {
        using var untracked = Open(TwoParagraphs, TrackedChangeMode.Accept);
        var (a, b) = Paragraphs(untracked);
        Assert.True(untracked.MergeParagraphs(a, b).Success);
        using var tracked = Open(TwoParagraphs, TrackedChangeMode.RenderInline);
        var (first, second) = Paragraphs(tracked);
        Assert.True(tracked.MergeParagraphs(first, second).Success);

        Assert.True(tracked.AcceptAllRevisions().Success);
        Assert.Equal(Properties(Body(untracked)), Properties(Body(tracked)));

        using var rejected = Open(TwoParagraphs, TrackedChangeMode.RenderInline);
        var (c, d) = Paragraphs(rejected);
        Assert.True(rejected.MergeParagraphs(c, d).Success);
        Assert.True(rejected.RejectAllRevisions().Success);
        using var original = Open(TwoParagraphs, TrackedChangeMode.Accept);
        Assert.Equal(Properties(Body(original)), Properties(Body(rejected)));
    }

    /// <summary>A section break belongs to the paragraph mark that ends the section. The second
    /// paragraph's mark survives the merge, so its section break stays where it was.</summary>
    [Fact]
    public void TrackedMerge_IntoASectionEndingParagraph_KeepsTheSectionBreak()
    {
        using var session = Open(
            "<w:p><w:r><w:t>First.</w:t></w:r></w:p>" +
            "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr><w:r><w:t>Second.</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Next section.</w:t></w:r></w:p>",
            TrackedChangeMode.RenderInline);
        var (first, second) = Paragraphs(session);

        Assert.True(session.MergeParagraphs(first, second).Success);
        Assert.True(session.AcceptAllRevisions().Success);

        var paragraphs = Body(session).Elements(W.p).ToList();
        Assert.Equal(2, paragraphs.Count);
        Assert.NotNull(paragraphs[0].Element(W.pPr)?.Element(W.sectPr));
    }

    [Fact]
    public void TrackedMerge_OfASectionEndingFirstParagraph_IsRefusedWithoutChanges()
    {
        using var session = Open(
            "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr><w:r><w:t>First.</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Second.</w:t></w:r></w:p>",
            TrackedChangeMode.RenderInline);
        var (first, second) = Paragraphs(session);
        var before = Body(session).ToString();

        var result = session.MergeParagraphs(first, second);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, result.Error!.Code);
        Assert.Equal(before, Body(session).ToString());
    }

    [Fact]
    public void TrackedMerge_WhenTheFirstMarkIsAlreadyDeleted_IsRefusedWithoutChanges()
    {
        using var session = Open(
            "<w:p><w:pPr><w:rPr><w:del w:id=\"9\" w:author=\"Alice\" w:date=\"2025-01-02T09:00:00Z\"/></w:rPr></w:pPr>" +
            "<w:r><w:t>First.</w:t></w:r></w:p><w:p><w:r><w:t>Second.</w:t></w:r></w:p>",
            TrackedChangeMode.RenderInline);
        var (first, second) = Paragraphs(session);
        var before = Body(session).ToString();

        var result = session.MergeParagraphs(first, second);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnresolvedStructuralRevision, result.Error!.Code);
        Assert.Equal(before, Body(session).ToString());
    }

    /// <summary>Enter then Backspace under tracking: the first paragraph's mark is both the
    /// editor's insertion and the editor's deletion, and either review gives back one paragraph.</summary>
    [Fact]
    public void TrackedSplitThenMerge_ReviewsBackToTheOriginalParagraph()
    {
        using var session = Open(
            "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t xml:space=\"preserve\">Terms apply</w:t></w:r></w:p>",
            TrackedChangeMode.RenderInline);
        var split = session.SplitParagraph(FirstParagraph(session), 6);
        Assert.True(split.Success, split.Error?.Message);

        var merge = session.MergeParagraphs(split.Modified[0].Id, split.Created[0].Id);

        Assert.True(merge.Success, merge.Error?.Message);
        AssertRevisionIdsDistinct(session, expectedCount: null);
        AssertOutcomes(session,
            accepted: new[] { "p[center]: Terms |apply" },
            rejected: new[] { "p[center]: Terms |apply" });
    }

    // ─── SplitParagraph and complex fields ─────────────────────────────────

    /// <summary>A tracked cross-reference's field sits inside its insertion, which a split can
    /// cut; the field inside it is kept whole all the same.</summary>
    [Fact]
    public void SplitInsideATrackedCrossReferencesResult_IsRefused()
    {
        using var session = Open(CrossReferenceBody, TrackedChangeMode.RenderInline);
        Assert.True(session.InsertCrossReference(SecondParagraph(session), 7, "tgt").Success);

        var result = session.SplitParagraph(SecondParagraph(session), 10);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnsupportedInlineBoundary, result.Error!.Code);
    }

    /// <summary>"See " + a REF field whose result is "Clause twelve" (offsets 4..17) + " now".</summary>
    private const string FieldParagraph =
        "<w:p><w:r><w:t xml:space=\"preserve\">See </w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
        "<w:r><w:instrText xml:space=\"preserve\"> REF x </w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
        "<w:r><w:t>Clause twelve</w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" +
        "<w:r><w:t xml:space=\"preserve\"> now</w:t></w:r></w:p>";

    [Theory]
    [InlineData(TrackedChangeMode.Accept, 10)]
    [InlineData(TrackedChangeMode.RenderInline, 10)]
    [InlineData(TrackedChangeMode.Accept, 5)]
    [InlineData(TrackedChangeMode.Accept, 16)]
    public void SplitInsideAFieldResult_IsRefusedWithoutChanges(TrackedChangeMode mode, int offset)
    {
        using var session = Open(FieldParagraph, mode);
        var before = Body(session).ToString();

        var result = session.SplitParagraph(FirstParagraph(session), offset);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnsupportedInlineBoundary, result.Error!.Code);
        Assert.Contains("field", result.Error.Message);
        Assert.Equal(before, Body(session).ToString());
    }

    /// <summary>Splitting just before or just after the field keeps the whole field, from its
    /// begin to its end, in one paragraph.</summary>
    [Theory]
    [InlineData(TrackedChangeMode.Accept, 4, "See ", "Clause twelve now")]
    [InlineData(TrackedChangeMode.Accept, 17, "See Clause twelve", " now")]
    [InlineData(TrackedChangeMode.RenderInline, 17, "See Clause twelve", " now")]
    public void SplitAtAFieldsEdge_KeepsTheFieldInOneParagraph(
        TrackedChangeMode mode, int offset, string firstText, string secondText)
    {
        using var session = Open(FieldParagraph, mode);

        var result = session.SplitParagraph(FirstParagraph(session), offset);

        Assert.True(result.Success, result.Error?.Message);
        var paragraphs = Body(session).Elements(W.p).ToList();
        Assert.Equal(new[] { firstText, secondText },
            paragraphs.Select(p => string.Concat(p.Descendants(W.t).Select(t => t.Value))));
        var withField = Assert.Single(paragraphs, p => p.Descendants(W.fldChar).Any());
        Assert.Equal(new[] { "begin", "separate", "end" },
            withField.Descendants(W.fldChar).Select(f => (string?)f.Attribute(W.fldCharType)));
    }

    /// <summary>Enter at the end of a line that ends in a field (a cross-reference, a page
    /// number): the new paragraph starts empty and the field stays whole on the line above.</summary>
    [Theory]
    [InlineData(TrackedChangeMode.Accept)]
    [InlineData(TrackedChangeMode.RenderInline)]
    public void SplitAtTheEndOfAParagraphEndingInAField_LeavesTheFieldWhole(TrackedChangeMode mode)
    {
        using var session = Open(FieldParagraph.Replace(
            "<w:r><w:t xml:space=\"preserve\"> now</w:t></w:r>", ""), mode);

        var result = session.SplitParagraph(FirstParagraph(session), 17);

        Assert.True(result.Success, result.Error?.Message);
        var paragraphs = Body(session).Elements(W.p).ToList();
        Assert.Equal(new[] { "begin", "separate", "end" },
            paragraphs[0].Descendants(W.fldChar).Select(f => (string?)f.Attribute(W.fldCharType)));
        Assert.Empty(paragraphs[1].Descendants(W.fldChar));
    }

    /// <summary>Control: a field that already runs on into the next paragraph (a table of contents,
    /// say) is not this paragraph's to keep whole, so a split inside its result still goes ahead.</summary>
    [Fact]
    public void SplitInsideAFieldThatSpansParagraphs_StillSplits()
    {
        using var session = Open(
            "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
            "<w:r><w:instrText xml:space=\"preserve\"> TOC \\o \"1-3\" </w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>Entry one</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Entry two</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>",
            TrackedChangeMode.Accept);

        var result = session.SplitParagraph(FirstParagraph(session), 5);

        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(3, Body(session).Elements(W.p).Count());
    }

    // ─── Helpers ───────────────────────────────────────────────────────────

    /// <summary>Each paragraph's properties without the projector's bookkeeping attributes.</summary>
    private static string[] Properties(XElement body) =>
        body.Elements(W.p).Select(p =>
        {
            var pPr = new XElement(p.Element(W.pPr) ?? new XElement(W.pPr));
            foreach (var element in pPr.DescendantsAndSelf())
                element.Attributes().Where(attribute => attribute.Name.Namespace == PtOpenXml.pt).Remove();
            return pPr.ToString(SaveOptions.DisableFormatting);
        }).ToArray();

    private static void AssertSchemaValid(DocxSession session) =>
        Assert.Empty(PackageValidation.Errors(session.Save(), FileFormatVersions.Office2019)
            .Where(PackageValidation.IsSchemaError).Select(PackageValidation.Keys.Description));
}
