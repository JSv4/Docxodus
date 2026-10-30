using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;
using static Docxodus.Tests.DocxBackendReconciliationTests;

namespace Docxodus.Tests;

public class DocxDiffMathAfterTabTests
{
    private static readonly XNamespace Math = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private static WmlDocument Doc(string upper, bool tab, bool nary)
    {
        var separator = tab ? "<w:r><w:tab/></w:r>" : "<w:r><w:t xml:space=\"preserve\"> </w:t></w:r>";
        var math = nary
            ? $"<m:nary><m:naryPr><m:chr m:val=\"∑\"/><m:ctrlPr><w:rPr/></m:ctrlPr></m:naryPr>" +
              $"<m:sub><m:r><m:t>i=1</m:t></m:r></m:sub><m:sup><m:r><m:t>{upper}</m:t></m:r></m:sup><m:e/></m:nary>"
            : $"<m:r><m:t>{upper}</m:t></m:r>";
        return IrTestDocuments.FromParts($"<w:p><w:r><w:t>III.</w:t></w:r>{separator}<m:oMath>{math}</m:oMath></w:p>",
            rootAttributes: $" xmlns:m=\"{Math}\"");
    }

    private static XElement Body(byte[] bytes)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return package.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    private static WmlDocument AtomicDoc(string value, string prefix, string beforeMath = "", string afterMath = "", bool split = false) =>
        IrTestDocuments.FromParts($"<w:p><w:r><w:t>III.</w:t></w:r>{prefix}" +
            (split ? "</w:p><w:p>" : "") + beforeMath +
            $"<m:oMath><m:r><m:t>{value}</m:t></m:r></m:oMath>{afterMath}</w:p>",
            rootAttributes: $" xmlns:m=\"{Math}\" xmlns:v=\"urn:schemas-microsoft-com:vml\"");

    private static string[] ContentOrder(XElement body)
    {
        var items = new System.Collections.Generic.List<string>();
        foreach (var element in body.Descendants().Where(e => e.Name == W.t || e.Name == W.tab || e.Name == W.br ||
                     e.Name == Math + "oMath" || e.Name == W.bookmarkStart || e.Name == W.bookmarkEnd))
        {
            if (element.Name == W.t && element.Value.Length == 0) continue;
            if (element.Name == W.t && items.Count > 0 && items[^1].StartsWith("t:"))
                items[^1] += element.Value;
            else
                items.Add(element.Name.LocalName + ":" + (element.Name == W.bookmarkStart
                    ? (string?)element.Attribute(W.name) : element.Value));
        }
        return items.ToArray();
    }

    [Theory]
    [InlineData("<w:r><w:tab/></w:r>")]
    [InlineData("<w:r><w:t/></w:r><w:r><w:tab/></w:r>")]
    [InlineData("<w:r><w:pict><v:shapetype id=\"s\"/></w:pict></w:r><w:r><w:tab/></w:r>")]
    [InlineData("<w:r><w:tab/><w:br/></w:r>")]
    [InlineData("<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:tab/><w:br/></w:r></w:sdtContent></w:sdt>")]
    public void AdjacentAtoms_AreOwnedByTheirTokenSpan(string prefix)
    {
        var left = AtomicDoc("OLD", prefix);
        var right = AtomicDoc("NEW", prefix);
        var result = DocxCompare.Compare(left, right).DocumentByteArray;
        Assert.Equal(ContentOrder(Body(right.DocumentByteArray)), ContentOrder(Body(DocxDiffOps.AcceptRevisions(result))));
        Assert.Equal(ContentOrder(Body(left.DocumentByteArray)), ContentOrder(Body(DocxDiffOps.RejectRevisions(result))));
        NoNewValidationErrors(left.DocumentByteArray, result);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BookmarkAtAtomicBoundary_KeepsBothProjectedPositions(bool after)
    {
        const string start = "<w:bookmarkStart w:id=\"5\" w:name=\"equation\"/>";
        const string end = "<w:bookmarkEnd w:id=\"5\"/>";
        var left = AtomicDoc("OLD", "<w:r><w:tab/></w:r>", after ? "" : start, after ? start + end : end);
        var right = AtomicDoc("NEW", "<w:r><w:tab/></w:r>", after ? "" : start, after ? start + end : end);
        var result = DocxCompare.Compare(left, right).DocumentByteArray;
        Assert.Equal(ContentOrder(Body(right.DocumentByteArray)), ContentOrder(Body(DocxDiffOps.AcceptRevisions(result))));
        Assert.Equal(ContentOrder(Body(left.DocumentByteArray)), ContentOrder(Body(DocxDiffOps.RejectRevisions(result))));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SplitAndMergeBetweenAtoms_PreserveBothViews(bool merge)
    {
        var joined = AtomicDoc("OLD", "<w:r><w:tab/></w:r>");
        var split = AtomicDoc("NEW", "<w:r><w:tab/></w:r>", split: true);
        var left = merge ? split : joined;
        var right = merge ? joined : split;
        var result = DocxCompare.Compare(left, right).DocumentByteArray;
        var accepted = Body(DocxDiffOps.AcceptRevisions(result));
        var rejected = Body(DocxDiffOps.RejectRevisions(result));
        Assert.Equal(ContentOrder(Body(right.DocumentByteArray)), ContentOrder(accepted));
        Assert.Equal(ContentOrder(Body(left.DocumentByteArray)), ContentOrder(rejected));
        Assert.Equal(merge ? 1 : 2, accepted.Elements(W.p).Count());
        Assert.Equal(merge ? 2 : 1, rejected.Elements(W.p).Count());
    }

    [Theory]
    [InlineData("<w:r><w:t/></w:r><w:r><w:tab/></w:r>")]
    [InlineData("<w:sdt><w:sdtPr/><w:sdtContent><w:r><w:tab/><w:br/></w:r></w:sdtContent></w:sdt>")]
    public void TwoReviewers_PreserveAtomicBoundariesAndAuthors(string prefix)
    {
        const string markers = "<w:bookmarkStart w:id=\"5\" w:name=\"after\"/><w:bookmarkEnd w:id=\"5\"/>";
        string Tail(string word) => markers + $"<w:r><w:t xml:space=\"preserve\"> shared tail {word}</w:t></w:r>";
        var original = AtomicDoc("OLD", prefix, afterMath: Tail("original"));
        var alice = AtomicDoc("NEW", prefix, afterMath: Tail("original"));
        var bob = AtomicDoc("OLD", prefix, afterMath: Tail("updated"));
        var expected = AtomicDoc("NEW", prefix, afterMath: Tail("updated"));
        var result = DocxDiff.Consolidate(original, new[]
        {
            new DocxDiffReviewer { Author = "Alice", Document = alice },
            new DocxDiffReviewer { Author = "Bob", Document = bob },
        }).DocumentByteArray;
        Assert.Equal(ContentOrder(Body(expected.DocumentByteArray)), ContentOrder(Body(DocxDiffOps.AcceptRevisions(result))));
        Assert.Equal(ContentOrder(Body(original.DocumentByteArray)), ContentOrder(Body(DocxDiffOps.RejectRevisions(result))));
        Assert.Contains(Body(result).Descendants(W.ins), e => (string?)e.Attribute(W.author) == "Alice" && e.Descendants(Math + "oMath").Any());
        Assert.Contains(Body(result).Descendants(W.ins), e => (string?)e.Attribute(W.author) == "Bob" && e.Value.Contains("updated"));
        NoNewValidationErrors(original.DocumentByteArray, result);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    public void OneSidedBookmarkAroundRetainedAtom_ResolvesWithItsOwner(bool added, bool consolidate)
    {
        const string tab = "<w:r><w:tab/></w:r>";
        const string marked = "<w:bookmarkStart w:id=\"5\" w:name=\"tab\"/>" + tab + "<w:bookmarkEnd w:id=\"5\"/>";
        string Tail(string word) => $"<w:r><w:t xml:space=\"preserve\"> shared tail {word}</w:t></w:r>";
        string originalPrefix = added ? tab : marked;
        string revisedPrefix = added ? marked : tab;
        var original = AtomicDoc("OLD", originalPrefix, afterMath: Tail("original"));
        var revised = AtomicDoc("NEW", revisedPrefix, afterMath: Tail("original"));
        var other = AtomicDoc("OLD", originalPrefix, afterMath: Tail("updated"));
        var expected = AtomicDoc("NEW", revisedPrefix, afterMath: Tail(consolidate ? "updated" : "original"));
        var result = (consolidate ? DocxDiff.Consolidate(original, new[]
        {
            new DocxDiffReviewer { Author = "Alice", Document = revised },
            new DocxDiffReviewer { Author = "Bob", Document = other },
        }) : DocxCompare.Compare(original, revised)).DocumentByteArray;
        Assert.Equal(ContentOrder(Body(expected.DocumentByteArray)), ContentOrder(Body(DocxDiffOps.AcceptRevisions(result))));
        Assert.Equal(ContentOrder(Body(original.DocumentByteArray)), ContentOrder(Body(DocxDiffOps.RejectRevisions(result))));
    }

    [Fact]
    public void RetainedAtomicSources_ExcludeDiscardedFieldInstructionBranches()
    {
        var document = AtomicDoc("OLD",
            "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
            "<w:hyperlink><w:r><w:tab/></w:r></w:hyperlink>" +
            "<w:r><w:instrText> PAGE </w:instrText><w:fldChar w:fldCharType=\"separate\"/><w:t>1</w:t><w:fldChar w:fldCharType=\"end\"/></w:r>" +
            "<w:r><w:tab/></w:r>");
        var retained = IrReader.Read(document, new IrReaderOptions { RetainSources = true });
        var detached = IrReader.Read(document, new IrReaderOptions { RetainSources = false });
        var paragraph = Assert.Single(retained.AnchorIndex.Values.OfType<IrParagraph>());
        var other = Assert.Single(detached.AnchorIndex.Values.OfType<IrParagraph>());
        Assert.Equal(paragraph.ContentHash, other.ContentHash);
        var tokens = IrDiffTokenizer.Tokenize(paragraph, new DocxDiffSettings().ToIrDiffSettings());
        Assert.Equal(2, tokens.Count(t => t.Kind is not (IrDiffTokenKind.Word or IrDiffTokenKind.Separator)));
        Assert.All(tokens.Where(t => t.Kind is not (IrDiffTokenKind.Word or IrDiffTokenKind.Separator)),
            token => Assert.Equal(token.StartChar, token.EndChar));
        var sources = paragraph.Source.Element!.Annotation<IrSourceContent>()!;
        Assert.Equal(new[] { W.tab, Math + "oMath" }, sources.Entries.Where(entry => entry.Atomic).Select(entry => entry.Source.Name));
    }

    [Theory]
    [InlineData(true, true, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, true)]
    [InlineData(true, false, true)]
    [InlineData(false, true, true)]
    public void ChangedMath_IsTrackedAndBothViewsRoundTrip(bool tab, bool nary, bool consolidate)
    {
        var left = Doc("s", tab, nary);
        var right = Doc("s+1", tab, nary);
        var redline = consolidate
            ? DocxDiff.Consolidate(left, new[] { new DocxDiffReviewer { Author = "Reviewer", Document = right } })
            : DocxCompare.Compare(left, right);
        var body = Body(redline.DocumentByteArray);
        var equations = body.Descendants(Math + "oMath").ToList();

        Assert.Contains(equations, e => e.Value.Contains("s+1") && e.Ancestors(W.ins).Any());
        Assert.Contains(equations, e => !e.Value.Contains("s+1") && e.Ancestors(W.del).Any());
        var accepted = Body(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray));
        var rejected = Body(DocxDiffOps.RejectRevisions(redline.DocumentByteArray));
        Assert.Equal(Body(right.DocumentByteArray).Value, accepted.Value);
        Assert.Equal(Body(left.DocumentByteArray).Value, rejected.Value);
        Assert.Equal(tab ? 1 : 0, accepted.Descendants(W.tab).Count());
        Assert.Equal(tab ? 1 : 0, rejected.Descendants(W.tab).Count());
        NoNewValidationErrors(left.DocumentByteArray, redline.DocumentByteArray);
    }
}
