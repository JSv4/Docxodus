using System;
using System.IO;
using Docxodus.Internal;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

public class TableVerticalMergeRenderingTests
{
    [Theory]
    [InlineData("aligned", 1, null)]
    [InlineData("shifted", 1, null)]
    [InlineData("shifted", 2, null)]
    [InlineData("omitted", 1, null)]
    [InlineData("omitted", 2, null)]
    [InlineData("shifted", 1, "sdt")]
    [InlineData("omitted", 2, "sdt")]
    [InlineData("shifted", 1, "customXml")]
    [InlineData("omitted", 2, "customXml")]
    public void MergeFollowsTheSameGridInterval(string prefix, int span, string? wrapper)
    {
        var first = prefix == "aligned" ? new[] { Cell("A"), Cell("B") } : new[] { Cell("A", 2) };
        var second = prefix == "omitted" ? new XElement[0] : new[] { Cell("C"), Cell("D") };
        var start = Cell("Merged", span, "restart");
        var continuation = Cell("", span, "continue");
        var row1 = Row(0, first.Append(start).ToArray());
        var row2 = Row(prefix == "omitted" ? 2 : 0, second.Append(continuation).ToArray());
        if (wrapper != null)
        {
            XElement Wrap(XElement cell) => wrapper == "sdt"
                ? new XElement(W.sdt, new XElement(W.sdtPr), new XElement(W.sdtContent, new XElement(cell)))
                : new XElement(W.customXml, new XAttribute(W.element, "entry"), new XElement(cell));
            start.ReplaceWith(Wrap(start));
            continuation.ReplaceWith(Wrap(continuation));
        }
        var html = Render(row1, row2);
        var merged = html.Descendants(Xhtml.td).Single(c => c.Value.Contains("Merged"));
        Assert.Equal(2, (int?)merged.Attribute("rowspan"));
        Assert.Equal(span == 1 ? null : span, (int?)merged.Attribute("colspan"));
        Assert.Equal(first.Length + second.Length + 1, html.Descendants(Xhtml.td).Count(c => !string.IsNullOrWhiteSpace(c.Value)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnmatchedContinuationRetainsItsCellContent(bool mismatchedSpan)
    {
        var html = mismatchedSpan
            ? Render(Row(0, Cell("Start", 2, "restart")), Row(0, Cell("Orphan", 1, "continue"), Cell("Tail")))
            : Render(Row(0, Cell("Orphan", 1, "continue")));
        Assert.Contains(html.Descendants(Xhtml.td), c => c.Value.Contains("Orphan"));
        Assert.DoesNotContain(html.Descendants(Xhtml.td), c => (int?)c.Attribute("rowspan") > 1);
    }

    [Fact]
    public void APlainCellBreaksTheMergeChain()
    {
        var html = Render(Row(0, Cell("Start", 1, "restart")), Row(0, Cell("", 1, "continue")),
            Row(0, Cell("Break")), Row(0, Cell("Orphan", 1, "continue")), Row(0, Cell("Tail", 1, "continue")));
        Assert.Equal(2, (int?)html.Descendants(Xhtml.td).Single(c => c.Value.Contains("Start")).Attribute("rowspan"));
        Assert.Equal(new[] { "Start", "Break", "Orphan", "Tail" },
            html.Descendants(Xhtml.td).Select(c => c.Value.Trim()));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OmittedStarterColumnsStayBeforeTheMergedCell(bool commonPrefix)
    {
        var html = Render(Row(1, Cell("Merged", 1, "restart")),
            commonPrefix ? Row(1, Cell("", 1, "continue")) : Row(0, Cell("Left"), Cell("", 1, "continue")));
        var rows = html.Descendants(Xhtml.tr).ToArray();
        var cells = rows[0].Elements(Xhtml.td).ToArray();
        Assert.Equal(commonPrefix ? 1 : 2, cells.Length);
        if (!commonPrefix)
            Assert.Empty(cells[0].Value);
        Assert.Contains("Merged", cells[^1].Value);
        Assert.Equal(2, (int?)cells[^1].Attribute("rowspan"));
    }

    [Fact]
    public void TallMergeOwnershipAllocatesLinearly()
    {
        static long Measure(int count)
        {
            var table = new XElement(W.tbl, Enumerable.Range(0, count)
                .Select(i => Row(0, Cell("", 1, i == 0 ? "restart" : "continue"))));
            var before = GC.GetAllocatedBytesForCurrentThread();
            var spans = TableGridModel.VerticalMergeSpans(table);
            var allocated = GC.GetAllocatedBytesForCurrentThread() - before;
            Assert.Equal(count, spans[table.Elements(W.tr).First().Element(W.tc)!]);
            Assert.Equal(count - 1, spans.Values.Count(span => span == 0));
            return allocated;
        }
        Measure(10); // JIT outside the allocation comparison.
        Assert.True(Measure(2000) < 3 * Measure(1000), "Doubling a merge must not quadruple grid allocations.");
    }

    [Theory]
    [InlineData("tc")]
    [InlineData("tr")]
    public void StandaloneTableFragmentsRetainTheirContent(string kind)
    {
        var bytes = CreateDocument(Row(0, Cell("OrdinaryLeft"), Cell("OrdinaryRight")));
        using var session = new DocxSession(bytes);
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == kind).Anchor.Id;
        var options = new HtmlConversionOptions { FabricateCssClasses = false };
        Assert.Contains("OrdinaryLeft", HtmlConversionOps.RenderBlockHtml(bytes, anchor, options));
    }

    private static XElement Cell(string text, int span = 1, string? merge = null) =>
        new(W.tc, new XElement(W.tcPr,
            new XElement(W.tcW, new XAttribute(W._w, 2000 * span), new XAttribute(W.type, "dxa")),
            span == 1 ? null : new XElement(W.gridSpan, new XAttribute(W.val, span)),
            merge == null ? null : new XElement(W.vMerge,
                merge == "continue" ? null : new XAttribute(W.val, merge))),
            new XElement(W.p, new XElement(W.r, new XElement(W.t, text))));

    private static XElement Row(int before, params XElement[] cells) =>
        new(W.tr, before == 0 ? null : new XElement(W.trPr, new XElement(W.gridBefore, new XAttribute(W.val, before))), cells);

    private static XElement Render(params XElement[] rows)
    {
        using var stream = new MemoryStream();
        stream.Write(CreateDocument(rows));
        using var doc = WordprocessingDocument.Open(stream, true);
        return WmlToHtmlConverter.ConvertToHtml(doc,
            new WmlToHtmlConverterSettings { RenderPagination = PaginationMode.Paginated, FabricateCssClasses = false });
    }

    private static byte[] CreateDocument(params XElement[] rows)
    {
        using var stream = new MemoryStream();
        stream.Write(DocxSession.CreateBlankDocxBytes());
        using (var doc = WordprocessingDocument.Open(stream, true))
        {
            var columns = rows.Max(row => TableGridModel.RowGrid(row)[^1].End);
            var main = doc.MainDocumentPart!;
            main.GetXDocument().Descendants(W.p).First().ReplaceWith(new XElement(W.tbl,
                new XElement(W.tblPr, new XElement(W.tblW, new XAttribute(W._w, "0"), new XAttribute(W.type, "auto"))),
                new XElement(W.tblGrid, Enumerable.Range(0, columns)
                    .Select(_ => new XElement(W.gridCol, new XAttribute(W._w, "2000")))), rows));
            main.PutXDocument();
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
        return stream.ToArray();
    }
}
