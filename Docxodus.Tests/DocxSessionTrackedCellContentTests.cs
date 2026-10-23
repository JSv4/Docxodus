// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="DocxSession.ReplaceCellContent"/> under tracked changes (issue #1311): the replacement
/// is recorded as native revisions that reject-all and accept-all resolve to the old and new cell,
/// or refused with nothing written when the cell's shape has no reversible encoding.
/// </summary>
public class DocxSessionTrackedCellContentTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string Width = "<w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/></w:tcPr>";

    private static byte[] TableWithCell(string cellBody) =>
        IrTestDocuments.FromBodyXml(
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid>"
            + $"<w:tr><w:tc>{Width}{cellBody}</w:tc></w:tr></w:tbl>"
            + "<w:p><w:r><w:t>after</w:t></w:r></w:p>").DocumentByteArray;

    private static DocxSession Tracked(byte[] bytes) => new(bytes, new DocxSessionSettings
    {
        RevisionAuthor = "Bob",
        TrackedChanges = TrackedChangeMode.RenderInline,
    });

    private static string CellAnchor(DocxSession session) =>
        session.Project().AnchorIndex.Values.Select(v => v.Anchor).First(a => a.Kind == "tc").Id;

    private static readonly XNamespace R =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>The cell as saved, for comparison: without the session's <c>pt:</c> bookkeeping,
    /// and with each relationship id replaced by its target, since fresh ids are random.</summary>
    private static string CellXml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var cell = new XElement(main.GetXDocument().Descendants(W + "tc").First());
        foreach (var element in cell.DescendantsAndSelf())
            element.Attributes().Where(a => a.Name.Namespace == PtOpenXml.pt || a.IsNamespaceDeclaration).Remove();
        foreach (var id in cell.Descendants().Attributes(R + "id"))
            id.Value = main.HyperlinkRelationships.Single(h => h.Id == id.Value).Uri.ToString();
        return cell.ToString(SaveOptions.DisableFormatting);
    }

    private static byte[] UntrackedResult(byte[] baseline, string payload)
    {
        using var session = new DocxSession(baseline);
        Assert.True(session.ReplaceCellContent(CellAnchor(session), payload).Success);
        return session.Save();
    }

    public static TheoryData<string, string> Replacements => new()
    {
        { "<w:p><w:r><w:t>Old cell</w:t></w:r></w:p>", "New cell" },
        {
            "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>Old one</w:t></w:r></w:p>"
                + "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Old two</w:t></w:r></w:p>",
            "New one\n\nNew **two**\n\nNew three"
        },
        { "<w:p><w:r><w:t>Old one</w:t></w:r></w:p><w:p><w:r><w:t>Old two</w:t></w:r></w:p>", "" },
        { "<w:p><w:r><w:t>Old cell</w:t></w:r></w:p>", "- first\n- second" },
        { "<w:p><w:r><w:t>Old cell</w:t></w:r></w:p>", "See [the site](https://example.com) now" },
    };

    [Theory]
    [MemberData(nameof(Replacements))]
    public void TrackedReplaceCellContent_RecordsRevisions_RejectRestoresAndAcceptKeeps(
        string cellBody, string payload)
    {
        var baseline = TableWithCell(cellBody);
        using (var session = Tracked(baseline))
        {
            var result = session.ReplaceCellContent(CellAnchor(session), payload);

            Assert.True(result.Success, result.Error?.Message);
            var revisions = session.ListRevisions();
            Assert.Contains(revisions, r => r.Family == RevisionFamily.ContentDelete);
            Assert.All(revisions, r =>
            {
                Assert.Equal("Bob", r.Author);
                Assert.Equal(RevisionResolutionStatus.Supported, r.ResolutionStatus);
            });
            PackageValidation.AssertValid(session.Save(), FileFormatVersions.Office2019);

            Assert.True(session.RejectAllRevisions().Success);
            Assert.Equal(CellXml(baseline), CellXml(session.Save()));
        }

        using (var session = Tracked(baseline))
        {
            Assert.True(session.ReplaceCellContent(CellAnchor(session), payload).Success);
            Assert.True(session.AcceptAllRevisions().Success);
            Assert.Equal(CellXml(UntrackedResult(baseline, payload)), CellXml(session.Save()));
        }
    }

    /// <summary>The cell's last paragraph keeps its mark, so the mark's own run formatting stays
    /// through accept, as when typing over a cell in Word; the untracked edit writes a fresh mark.</summary>
    [Fact]
    public void TrackedReplaceCellContent_KeepsTheLastParagraphMarksFormatting()
    {
        var baseline = TableWithCell("<w:p><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t>Old</w:t></w:r></w:p>");
        using var session = Tracked(baseline);

        Assert.True(session.ReplaceCellContent(CellAnchor(session), "New").Success);
        Assert.True(session.AcceptAllRevisions().Success);

        var paragraph = Assert.Single(XElement.Parse(CellXml(session.Save())).Elements(W + "p"));
        Assert.NotNull(paragraph.Element(W + "pPr")?.Element(W + "rPr")?.Element(W + "b"));
        Assert.Equal("New", string.Concat(paragraph.Descendants(W + "t").Select(t => t.Value)));
    }

    /// <summary>Another author's pending revisions in the cell are deleted over, not resolved:
    /// rejecting only Bob's revisions gives back the cell with Alice's intact.</summary>
    [Fact]
    public void TrackedReplaceCellContent_OverAnotherAuthorsRevisions_RejectsToThem()
    {
        const string Alice = "w:author=\"Alice\" w:date=\"2025-01-02T09:00:00Z\"";
        var baseline = TableWithCell(
            $"<w:p><w:ins w:id=\"70\" {Alice}><w:r><w:t>inserted</w:t></w:r></w:ins><w:r><w:t> plain</w:t></w:r></w:p>"
            + $"<w:p><w:r><w:t>kept </w:t></w:r><w:del w:id=\"71\" {Alice}><w:r><w:delText>deleted</w:delText></w:r></w:del></w:p>");
        using var session = Tracked(baseline);

        Assert.True(session.ReplaceCellContent(CellAnchor(session), "New one\n\nNew two").Success);
        PackageValidation.AssertValid(session.Save(), FileFormatVersions.Office2019);
        while (session.ListRevisions().FirstOrDefault(r => r.Author == "Bob") is { } bob)
            Assert.True(session.RejectRevision(bob.Id).Success);

        Assert.Equal(CellXml(baseline), CellXml(session.Save()));
    }

    /// <summary>Filling a row inserted under tracking: the new cell's paragraph is Bob's own
    /// insertion, and reject-all still removes the whole row.</summary>
    [Fact]
    public void TrackedReplaceCellContent_InATrackedNewRow_RejectsToTheOriginalTable()
    {
        var baseline = TableWithCell("<w:p><w:r><w:t>A1</w:t></w:r></w:p>");
        using var session = Tracked(baseline);
        var row = session.InsertTableRow(CellAnchor(session), Position.After);

        Assert.True(session.ReplaceCellContent(row.Created[0].Id, "Filled").Success);
        PackageValidation.AssertValid(session.Save(), FileFormatVersions.Office2019);
        Assert.True(session.RejectAllRevisions().Success);

        using var stream = new MemoryStream(session.Save());
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Single(document.MainDocumentPart!.GetXDocument().Descendants(W + "tr"));
        Assert.Equal(CellXml(baseline), CellXml(session.Save()));
    }

    [Fact]
    public void UntrackedReplaceCellContent_WritesNoRevisions()
    {
        using var session = new DocxSession(TableWithCell("<w:p><w:r><w:t>Old cell</w:t></w:r></w:p>"));

        Assert.True(session.ReplaceCellContent(CellAnchor(session), "New cell").Success);

        Assert.Empty(session.ListRevisions());
        var cell = XElement.Parse(CellXml(session.Save()));
        Assert.DoesNotContain(cell.Descendants(), e => e.Name == W + "ins" || e.Name == W + "del");
        Assert.Equal("New cell", string.Concat(cell.Descendants(W + "t").Select(t => t.Value)));
    }

    /// <summary>Untracked replacement removes only the cell's own paragraphs, so a cell holding a
    /// nested table or a content control has no tracked encoding whose accept matches it; one whose
    /// last paragraph holds a bookmark cannot keep that paragraph's mark, as tracked
    /// <c>ReplaceText</c> cannot. Each is refused and the document is left exactly as it was.</summary>
    [Theory]
    [InlineData("<w:p><w:r><w:t>Old</w:t></w:r></w:p>"
        + "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>"
        + "<w:tr><w:tc><w:p><w:r><w:t>Nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>")]
    [InlineData("<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>Control</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p/>")]
    [InlineData("<w:p><w:bookmarkStart w:id=\"7\" w:name=\"mark\"/><w:r><w:t>Old</w:t></w:r><w:bookmarkEnd w:id=\"7\"/></w:p>")]
    public void TrackedReplaceCellContent_RefusesAShapeWithNoReversibleEncoding(string cellBody)
    {
        var baseline = TableWithCell(cellBody);
        using var session = Tracked(baseline);
        long version = session.Version;

        var result = session.ReplaceCellContent(CellAnchor(session), "New cell");

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, result.Error!.Code);
        Assert.Equal(version, session.Version);
        Assert.Equal(CellXml(baseline), CellXml(session.Save()));
    }
}
