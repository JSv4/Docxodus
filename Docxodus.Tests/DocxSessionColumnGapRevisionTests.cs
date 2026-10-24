using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionColumnGapRevisionTests
{
    [Theory]
    [InlineData(true, true)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(false, false)]
    public void ColumnInsertion_AcrossAnOmittedRow_IsOneResolvableAction(bool leading, bool tracked) =>
        Check(leading, tracked, gap: true);

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Control_ColumnInsertion_AcrossOrdinaryRows_IsOneResolvableAction(bool leading) =>
        Check(leading, tracked: true, gap: false);

    [Fact]
    public void BareGrid_DoesNotChooseBetweenTwoCompleteColumnCandidates()
    {
        string Cell(string text, string? author = null) => "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/>"
            + (author is null ? "" : "<w:cellIns w:id=\"" + (author == "First" ? "2" : "3")
                + "\" w:author=\"" + author + "\" w:date=\"2026-01-01T00:00:00Z\"/>")
            + "</w:tcPr><w:p><w:r><w:t>" + text + "</w:t></w:r></w:p></w:tc>";
        string Columns(int count) => string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"2000\"/>", count));
        var input = IrTestDocuments.FromBodyXml("<w:tbl><w:tblPr/><w:tblGrid>" + Columns(4)
            + "<w:tblGridChange w:id=\"1\"><w:tblGrid>" + Columns(3) + "</w:tblGrid></w:tblGridChange></w:tblGrid><w:tr>"
            + Cell("", "First") + Cell("A") + Cell("", "Second") + Cell("B") + "</w:tr></w:tbl><w:p/>").DocumentByteArray;
        PackageValidation.AssertValid(input, FileFormatVersions.Office2019);
        using var session = new DocxSession(input);
        Assert.All(session.ListRevisions(), revision => Assert.Equal(RevisionResolutionStatus.Ambiguous, revision.ResolutionStatus));
        Assert.False(session.RejectAllRevisions().Success);
        Assert.False(session.Undo());
        Assert.True(XNode.DeepEquals(Xml(input), Xml(session.Save())));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ColumnInsertion_CrossesBothAMergedRowAndAnOmittedRow(bool tracked)
    {
        string Cell(string text, int span = 1) => "<w:tc><w:tcPr><w:tcW w:w=\"" + (span * 2000)
            + "\" w:type=\"dxa\"/>" + (span == 1 ? "" : "<w:gridSpan w:val=\"" + span + "\"/>")
            + "</w:tcPr><w:p><w:r><w:t>" + text + "</w:t></w:r></w:p></w:tc>";
        string Row(string suffix) => "<w:tr>" + Cell("A" + suffix) + Cell("B" + suffix) + Cell("C" + suffix) + "</w:tr>";
        var input = IrTestDocuments.FromBodyXml("<w:tbl><w:tblPr/><w:tblGrid>"
            + string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"2000\"/>", 3)) + "</w:tblGrid>" + Row("1")
            + "<w:tr>" + Cell("Merged", 3) + "</w:tr><w:tr><w:trPr><w:gridBefore w:val=\"2\"/></w:trPr>"
            + Cell("C3") + "</w:tr>" + Row("4") + "</w:tbl><w:p/>").DocumentByteArray;
        Check(input, "A1", Position.After, 3, tracked);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ColumnResolution_DoesNotAbsorbSameStampFormatting(bool tableShading)
    {
        var inserted = InsertedDocument(out var original);
        var input = Update(inserted, root =>
        {
            var stamp = root.Descendants(W.cellIns).First();
            var table = root.Descendants(W.tbl).Single();
            var changeName = tableShading ? W.tblPrChange : W.trPrChange;
            var propertiesName = tableShading ? W.tblPr : W.trPr;
            var change = new XElement(changeName, new XAttribute(W.id, 900), new XAttribute(W.author, stamp.Attribute(W.author)!.Value),
                new XAttribute(W.date, stamp.Attribute(W.date)!.Value), new XElement(propertiesName));
            if (tableShading)
                table.Element(W.tblPr)!.Add(new XElement(W.shd, new XAttribute(W.val, "clear"), new XAttribute(W.fill, "E2F0D9")), change);
            else table.Elements(W.tr).Last().AddFirst(new XElement(W.trPr,
                new XElement(W.trHeight, new XAttribute(W.val, 480)), change));
        });
        PackageValidation.AssertValid(input, FileFormatVersions.Office2019);
        using var session = new DocxSession(input);
        var column = Assert.Single(session.ListRevisions(), revision => revision.Family == RevisionFamily.CellInsert);
        Assert.True(session.RejectRevision(column.Id).Success);
        AssertGrid(session.Save(), 2);
        if (tableShading)
            Assert.Equal("E2F0D9", (string?)Xml(session.Save()).Descendants(W.tblPr).First().Element(W.shd)?.Attribute(W.fill));
        else Assert.Equal(480, (int?)Xml(session.Save()).Descendants(W.trHeight).Single().Attribute(W.val));
        var formatting = Assert.Single(session.ListRevisions());
        Assert.Equal(RevisionFamily.PropertiesChange, formatting.Family);
        Assert.True(session.RejectRevision(formatting.Id).Success);
        Assert.True(XNode.DeepEquals(Xml(original), Xml(session.Save())));
    }

    [Theory]
    [InlineData("author")]
    [InlineData("date")]
    [InlineData("grid")]
    public void ColumnGrouping_RequiresMatchingPropertiesAndSavedGrid(string mismatch)
    {
        var input = Update(InsertedDocument(out _), root =>
        {
            if (mismatch == "grid")
                root.Descendants(W.tblGridChange).Single().Element(W.tblGrid)!.Elements(W.gridCol).First()
                    .SetAttributeValue(W._w, 3000);
            else root.Descendants(W.trPrChange).Single().SetAttributeValue(mismatch == "author" ? W.author : W.date,
                mismatch == "author" ? "Another Reviewer" : "2025-01-01T00:00:00Z");
        });
        PackageValidation.AssertValid(input, FileFormatVersions.Office2019);
        using var session = new DocxSession(input);
        Assert.Contains(session.ListRevisions(), revision => revision.ResolutionStatus == RevisionResolutionStatus.Ambiguous);
        Assert.False(session.RejectAllRevisions().Success);
        Assert.False(session.Undo());
        Assert.True(XNode.DeepEquals(Xml(input), Xml(session.Save())));
    }

    private static byte[] InsertedDocument(out byte[] original)
    {
        original = Document(leading: true, gap: true);
        using var session = new DocxSession(original, new DocxSessionSettings
        {
            RevisionAuthor = "Layout Reviewer", TrackedChanges = TrackedChangeMode.RenderInline,
        });
        var anchor = session.Project().AnchorIndex.Values.First(a => a.Anchor.Kind == "tc" && a.TextPreview.StartsWith("A1")).Anchor.Id;
        Assert.True(session.InsertTableColumn(anchor, Position.Before).Success);
        return session.Save();
    }

    private static byte[] Update(byte[] bytes, Action<XElement> edit)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            edit(main.GetXDocument().Root!);
            main.PutXDocument();
        }
        return stream.ToArray();
    }

    private static void Check(bool leading, bool tracked, bool gap) =>
        Check(Document(leading, gap), leading ? "A1" : "B1", leading ? Position.Before : Position.After, 2, tracked);

    private static void Check(byte[] input, string targetText, Position position, int originalColumns, bool tracked)
    {
        PackageValidation.AssertValid(input, FileFormatVersions.Office2019);
        using var session = new DocxSession(input, new DocxSessionSettings
        {
            RevisionAuthor = "Layout Reviewer",
            TrackedChanges = tracked ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept,
        });
        var anchor = session.Project().AnchorIndex.Values.First(a => a.Anchor.Kind == "tc"
            && a.TextPreview.StartsWith(targetText)).Anchor.Id;
        var result = session.InsertTableColumn(anchor, position);
        Assert.True(result.Success, result.Error?.Message);
        var saved = session.Save();
        PackageValidation.AssertValid(saved, FileFormatVersions.Office2019);
        AssertGrid(saved, originalColumns + 1);
        if (tracked)
        {
            var revision = Assert.Single(session.ListRevisions());
            Assert.Equal(RevisionResolutionStatus.Supported, revision.ResolutionStatus);
        }
        else Assert.Empty(session.ListRevisions());

        foreach (var accept in new[] { false, true }) foreach (var selective in new[] { false, true })
        {
            using var review = new DocxSession(saved);
            var resolved = selective && tracked
                ? (accept ? review.AcceptRevision(Assert.Single(review.ListRevisions()).Id)
                    : review.RejectRevision(Assert.Single(review.ListRevisions()).Id))
                : (accept ? review.AcceptAllRevisions() : review.RejectAllRevisions());
            Assert.True(resolved.Success, resolved.Error?.Message);
            var output = review.Save();
            PackageValidation.AssertValid(output, FileFormatVersions.Office2019);
            Assert.Empty(review.ListRevisions());
            AssertGrid(output, accept || !tracked ? originalColumns + 1 : originalColumns);
            if (!accept && tracked) Assert.True(XNode.DeepEquals(Xml(input), Xml(output)));
        }
        Assert.True(session.Undo());
        Assert.True(XNode.DeepEquals(Xml(input), Xml(session.Save())));
    }

    private static byte[] Document(bool leading, bool gap)
    {
        string Cell(string text) => "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>"
            + text + "</w:t></w:r></w:p></w:tc>";
        string Row(string suffix) => "<w:tr>" + Cell("A" + suffix) + Cell("B" + suffix) + "</w:tr>";
        var middle = gap ? "<w:tr><w:trPr><w:" + (leading ? "gridBefore" : "gridAfter")
            + " w:val=\"1\"/></w:trPr>" + Cell(leading ? "B2" : "A2") + "</w:tr>" : Row("2");
        return IrTestDocuments.FromBodyXml("<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/>"
            + "<w:gridCol w:w=\"2000\"/></w:tblGrid>" + Row("1") + middle + Row("3") + "</w:tbl><w:p/>").DocumentByteArray;
    }

    private static XElement Xml(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return new XElement(document.MainDocumentPart!.GetXDocument().Root!);
    }

    private static void AssertGrid(byte[] bytes, int columns)
    {
        var table = Xml(bytes).Descendants(W.tbl).Single();
        Assert.Equal(columns, table.Element(W.tblGrid)!.Elements(W.gridCol).Count());
        foreach (var row in table.Elements(W.tr))
        {
            var properties = row.Element(W.trPr);
            int covered = ((int?)properties?.Element(W.gridBefore)?.Attribute(W.val) ?? 0)
                + ((int?)properties?.Element(W.gridAfter)?.Attribute(W.val) ?? 0)
                + row.Elements(W.tc).Sum(c => (int?)c.Element(W.tcPr)?.Element(W.gridSpan)?.Attribute(W.val) ?? 1);
            Assert.Equal(columns, covered);
        }
    }
}
