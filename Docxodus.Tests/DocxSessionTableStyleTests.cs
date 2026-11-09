using System.Linq;
using Docxodus.Tests.Ir;
using Xunit;
using static Docxodus.Tests.DocxSessionClonedRevisionTests;

namespace Docxodus.Tests;

public class DocxSessionTableStyleTests
{
    internal const string Styles = "<w:style w:type=\"table\" w:styleId=\"TableA\"><w:name w:val=\"Table A\"/></w:style>" +
        "<w:style w:type=\"table\" w:styleId=\"TableB\"><w:name w:val=\"Table B\"/><w:tblPr><w:tblBorders><w:top w:val=\"double\" w:sz=\"12\" w:color=\"336699\"/></w:tblBorders></w:tblPr></w:style>" +
        "<w:style w:type=\"paragraph\" w:styleId=\"ParagraphStyle\"><w:name w:val=\"Paragraph\"/></w:style>";
    internal const string Content = "<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:tbl>" +
        "<w:tblPr><w:tblStyle w:val=\"TableA\"/><w:tblW w:w=\"2400\" w:type=\"dxa\"/><w:tblBorders><w:bottom w:val=\"single\" w:sz=\"8\"/></w:tblBorders></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell</w:t></w:r></w:p>" +
        "</w:tc></w:tr></w:tbl><w:p/>";

    private static string Anchor(DocxSession session, string kind) =>
        session.Project().AnchorIndex.Values.First(a => a.Anchor.Kind == kind).Anchor.Id;

    [Theory]
    [InlineData("tbl", true)]
    [InlineData("tc", false)]
    public void SetTableStyle_RecordsOneReversiblePropertyChange(string kind, bool accept)
    {
        using var session = Open(Content, TrackedChangeMode.RenderInline, stylesXml: Styles);
        var before = Body(session).ToString();

        var edit = session.SetTableStyle(Anchor(session, kind), "TableB");

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal("TableB", (string?)Body(session).Descendants(W.tblStyle).First().Attribute(W.val));
        var revision = Assert.Single(session.ListRevisions());
        Assert.Equal(RevisionFamily.PropertiesChange, revision.Family);
        Assert.Equal("Bob", revision.Author);
        PackageValidation.AssertValid(session.Save(), DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
        Assert.True(session.Undo());
        Assert.Equal(before, Body(session).ToString());
        Assert.True(session.Redo());
        var resolved = accept ? session.AcceptRevision(revision.Id) : session.RejectRevision(revision.Id);
        Assert.True(resolved.Success, resolved.Error?.Message);
        Assert.Empty(session.ListRevisions());
        Assert.Equal(accept ? "TableB" : "TableA", (string?)Body(session).Descendants(W.tblStyle).Single().Attribute(W.val));
        Assert.Equal("2400", (string?)Body(session).Descendants(W.tblW).Single().Attribute(W._w));
        Assert.Equal("single", (string?)Body(session).Descendants(W.tblBorders).Single().Element(W.bottom)?.Attribute(W.val));
        PackageValidation.AssertValid(session.Save(), DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
    }

    [Theory]
    [InlineData("")]
    [InlineData("Missing")]
    [InlineData("ParagraphStyle")]
    public void InvalidTableStyle_RefusesBeforeMutationOrHistory(string style)
    {
        using var session = Open(Content, TrackedChangeMode.RenderInline, stylesXml: Styles);
        var before = Body(session).ToString();
        var edit = session.SetTableStyle(Anchor(session, "tbl"), style);
        Assert.False(edit.Success);
        Assert.Equal(EditErrorCode.UnknownStyle, edit.Error!.Code);
        Assert.Equal(before, Body(session).ToString());
        Assert.False(session.Undo());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InsertTable_WithStyle_InheritsBordersUnlessExplicitlyBorderless(bool borderless)
    {
        using var session = Open(Content, TrackedChangeMode.Accept, stylesXml: Styles);
        var edit = session.InsertTable(FirstParagraph(session), Position.Before, 1, 1,
            new TableInsertOptions { StyleId = "TableB", Borderless = borderless });
        Assert.True(edit.Success, edit.Error?.Message);
        var properties = Body(session).Descendants(W.tblPr).First();
        Assert.Equal("TableB", (string?)properties.Element(W.tblStyle)?.Attribute(W.val));
        Assert.Equal(borderless, properties.Element(W.tblBorders) is not null);
        PackageValidation.AssertValid(session.Save(), DocumentFormat.OpenXml.FileFormatVersions.Microsoft365);
    }

    [Fact]
    public void SameStyle_IsANoOp_AndPendingChangesAreNotOverwritten()
    {
        using var session = Open(Content, TrackedChangeMode.RenderInline, stylesXml: Styles);
        var table = Anchor(session, "tbl");
        var version = session.Version;
        Assert.True(session.SetTableStyle(table, "TableA").Success);
        Assert.Equal(version, session.Version);
        Assert.Empty(session.ListRevisions());
        Assert.False(session.Undo());
        Assert.True(session.SetTableStyle(table, "TableB").Success);
        var before = Body(session).ToString();
        version = session.Version;
        var revision = Assert.Single(session.ListRevisions()).Id;
        Assert.True(session.SetTableStyle(table, "TableB").Success);
        Assert.Equal(version, session.Version);
        Assert.Equal(revision, Assert.Single(session.ListRevisions()).Id);
        var refused = session.SetTableStyle(table, "TableA");
        Assert.False(refused.Success);
        Assert.Equal(EditErrorCode.UnresolvedStructuralRevision, refused.Error!.Code);
        Assert.Equal(before, Body(session).ToString());
        Assert.True(session.Undo());
        Assert.Empty(session.ListRevisions());
        Assert.False(session.Undo());
    }

    [Theory]
    [InlineData("Missing")]
    [InlineData("ParagraphStyle")]
    public void InsertTable_InvalidStyle_RefusesWithoutHistory(string style)
    {
        using var session = Open(Content, TrackedChangeMode.Accept, stylesXml: Styles);
        var before = Body(session).ToString();
        var result = session.InsertTable(FirstParagraph(session), Position.After, 1, 1,
            new TableInsertOptions { StyleId = style });
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnknownStyle, result.Error!.Code);
        Assert.Equal(before, Body(session).ToString());
        Assert.False(session.Undo());
    }

    [Theory]
    [InlineData("tbl")]
    [InlineData("tc")]
    [InlineData("p")]
    public void NestedCell_TargetsItsOwnTable_AndPreservesDirectFormatting(string kind)
    {
        var inner = "<w:tbl><w:tblPr><w:tblStyle w:val=\"TableA\"/></w:tblPr>" +
            "<w:tblGrid><w:gridCol w:w=\"1200\"/></w:tblGrid><w:tr><w:tc><w:p/>" +
            "</w:tc></w:tr></w:tbl><w:p/>";
        using var session = Open(Content.Replace("</w:tc>", inner + "</w:tc>"), TrackedChangeMode.Accept, stylesXml: Styles);
        var nested = session.Project().AnchorIndex.Values.Where(a => a.Anchor.Kind == "tbl").Last().Anchor.Id;
        var cell = session.GetTableMetadata(nested).Metadata!.Rows[0].Cells[0].Anchor.Id;

        var target = kind == "tbl" ? nested : kind == "tc" ? cell :
            session.GetTableMetadata(nested).Metadata!.Rows[0].Cells[0].ParagraphAnchors[0].Id;
        Assert.True(session.SetTableStyle(target, "TableB").Success);

        Assert.Equal(new[] { "TableA", "TableB" }, Body(session).Descendants(W.tblStyle).Select(s => (string?)s.Attribute(W.val)));
        Assert.Equal("2400", (string?)Body(session).Descendants(W.tblW).Single().Attribute(W._w));
        Assert.Empty(session.ListRevisions());
    }
}
