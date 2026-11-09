using System.Linq;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class WmlToMarkdownTableRevisionTests
{
    private static string Row(string text, bool deleted = false, int revisionId = 1)
    {
        var properties = deleted ? $"<w:trPr><w:del w:id=\"{revisionId}\" w:author=\"Alice\"/></w:trPr>" : "";
        return $"<w:tr>{properties}{Cell(text)}</w:tr>";
    }

    private static string Cell(string text) => $"<w:tc><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";

    private static byte[] Table(string rows, int columns = 1) => IrTestDocuments.FromBodyXml(
        "<w:tbl><w:tblPr/><w:tblGrid>" + string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"2400\"/>", columns)) +
        "</w:tblGrid>" + rows +
        "</w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p>").DocumentByteArray;

    [Theory]
    [InlineData(TrackedChangeMode.Accept, TableRenderMode.GfmWithOpaqueFallback)]
    [InlineData(TrackedChangeMode.StripDeletions, TableRenderMode.AlwaysGfm)]
    [InlineData(TrackedChangeMode.Accept, TableRenderMode.AlwaysOpaque)]
    public void DeletedRows_AreOmittedFromTableProjection(TrackedChangeMode mode, TableRenderMode tableMode)
    {
        // Wrapped rows follow the same rule, and a pending deletion must not supply the header.
        var input = Table(Row("DeletedHeader", true) + Row("Kept") +
            "<w:sdt><w:sdtPr/><w:sdtContent>" + Row("DeletedWrapped", true, 2) + "</w:sdtContent></w:sdt>" +
            Row("Tail"));
        using var session = new DocxSession(input, new DocxSessionSettings
        {
            ProjectionSettings = new WmlToMarkdownConverterSettings { TrackedChanges = mode, TableMode = tableMode },
        });
        var projection = session.Project();
        var table = projection.AnchorIndex.Values.First(a => a.Anchor.Kind == "tbl").Anchor.Id;
        var markdown = session.ProjectAnchor(table).Markdown;

        Assert.DoesNotContain("Deleted", markdown);
        if (tableMode == TableRenderMode.AlwaysOpaque)
            Assert.Contains("rows: 2", markdown);
        else
            Assert.Equal(new[] { "| Kept |", "| --- |", "| Tail |" },
                markdown.Split('\n').Where(line => line.StartsWith('|')).Select(line => line.TrimEnd('\r')));
        Assert.Contains(markdown.Trim(), projection.Markdown);
        Assert.Equal(4, projection.AnchorIndex.Values.Count(a => a.Anchor.Kind == "tr"));
    }

    [Theory]
    [InlineData(TableRenderMode.GfmWithOpaqueFallback)]
    [InlineData(TableRenderMode.AlwaysOpaque)]
    public void WhollyDeletedTable_HasNoProjectedShell(TableRenderMode mode)
    {
        using var session = new DocxSession(Table(Row("Gone", true)), new DocxSessionSettings
        {
            ProjectionSettings = new WmlToMarkdownConverterSettings { TableMode = mode },
        });
        var projection = session.Project();
        Assert.DoesNotContain("{#tbl:", projection.Markdown);
        Assert.Contains("After", projection.Markdown);
        Assert.Contains(projection.AnchorIndex.Values, a => a.Anchor.Kind == "tbl");
    }

    [Fact]
    public void DeletedMergedRow_DoesNotForceSurvivingRowsOpaque()
    {
        var merged = Row("Merged", true).Replace("<w:tc>", "<w:tc><w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr>");
        var kept = $"<w:tr>{Cell("Kept")}{Cell("Tail")}</w:tr>";
        var empty = $"<w:tr>{Cell("")}{Cell("")}</w:tr>";
        using var session = new DocxSession(Table(merged + kept + empty, columns: 2));

        var markdown = session.Project().Markdown;

        Assert.Contains("| Kept | Tail |", markdown);
        Assert.Contains("|   |   |", markdown); // A genuine empty row still exists.
        Assert.DoesNotContain("```table", markdown);
    }

    [Fact]
    public void InlineProjection_KeepsDeletedRowsAvailable()
    {
        using var session = new DocxSession(Table(Row("Deleted", true) + Row("Kept")), new DocxSessionSettings
        {
            ProjectionSettings = new WmlToMarkdownConverterSettings { TrackedChanges = TrackedChangeMode.RenderInline },
        });
        Assert.Contains("| Deleted |", session.Project().Markdown);
        Assert.Contains("| Kept |", session.Project().Markdown);
    }

    [Fact]
    public void TextboxInDeletedRow_DoesNotLeakAfterTheTable()
    {
        var shape = "<w:r><w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" id=\"Box\" " +
            "style=\"width:100pt;height:30pt\"><v:textbox><w:txbxContent><w:p><w:r><w:t>HiddenBox</w:t>" +
            "</w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>";
        var row = Row("Deleted", true).Replace("</w:p>", shape + "</w:p>");
        using var session = new DocxSession(Table(row + Row("Kept")));
        Assert.DoesNotContain("HiddenBox", session.Project().Markdown);
        Assert.Contains("| Kept |", session.Project().Markdown);
    }

    [Fact]
    public void DeletingLastRow_RemovesProjectedBlockAndDoesNotBecomeFollowingPatchPredecessor()
    {
        using var session = new DocxSession(Table(Row("Gone")),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchors = session.Project().AnchorIndex.Values.ToArray();
        var table = anchors.First(a => a.Anchor.Kind == "tbl").Anchor.Id;
        var cell = anchors.First(a => a.Anchor.Kind == "tc").Anchor.Id;
        var after = anchors.First(a => a.TextPreview == "After").Anchor.Id;

        var edit = session.DeleteTableRow(cell);

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Contains(table, edit.Patch!.RemovedAnchorIds);
        Assert.Empty(session.ProjectAnchor(table).Markdown);
        Assert.True(session.Undo());
        Assert.Contains("| Gone |", session.ProjectAnchor(table).Markdown);
        Assert.True(session.Redo());
        var rejected = session.RejectAllRevisions();
        Assert.True(rejected.Success, rejected.Error?.Message);
        Assert.Contains("| Gone |", session.ProjectAnchor(table).Markdown);
        Assert.True(session.Undo());
        var followingEdit = session.ReplaceText(after, "Following");
        Assert.True(followingEdit.Success, followingEdit.Error?.Message);
        Assert.All(followingEdit.Patch!.Blocks, block => Assert.NotEqual(table, block.AfterAnchorId));
    }

    [Fact]
    public void TrackedDeleteRow_ProjectAndPatchMatchAcceptedView_RejectRestoresRow()
    {
        using var session = new DocxSession(Table(Row("Header") + Row("DeleteMe") + Row("Tail")),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var before = session.Project().Markdown;
        var cell = session.Project().AnchorIndex.Values.Where(a => a.Anchor.Kind == "tc").Skip(1).First().Anchor.Id;

        var edit = session.DeleteTableRow(cell);

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Equal(RevisionFamily.RowDelete, Assert.Single(session.ListRevisions()).Family);
        var rows = session.Project().Markdown.Split('\n').Where(line => line.StartsWith('|')).ToArray();
        Assert.Equal(3, rows.Length); // Header, separator, surviving body row.
        Assert.DoesNotContain("DeleteMe", session.Project().Markdown);
        Assert.DoesNotContain("|   |", edit.Patch!.Markdown);
        Assert.True(session.RejectAllRevisions().Success);
        Assert.Equal(before, session.Project().Markdown);
    }
}
