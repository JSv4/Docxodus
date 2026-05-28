#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using Docxodus;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionAnnotationWriteTests
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");

    private static byte[] LoadFixture(string name) =>
        File.ReadAllBytes(Path.Combine(TestFilesDir.FullName, name));

    // Smallest known-good fixture used throughout the suite.
    private const string Fixture = "DA001-TemplateDocument.docx";

    [Fact]
    public void AW001_AddAnnotation_ByAnchorAndSpan_PersistsBookmarkAndCustomXml()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body)
            .First(a => a.Anchor.Kind == "p");

        var annotation = new DocumentAnnotation
        {
            Id = "ann-001",
            LabelId = "RISK",
            Label = "Risk",
            Color = "#FFEB3B",
            Author = "tester",
        };

        var result = session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 4), annotation);

        Assert.True(result.Success);
        Assert.Equal("ann-001", result.AnnotationId);
        Assert.Single(result.Modified);
        Assert.Equal(firstP.Anchor.Id, result.Modified[0].Id);

        var listed = session.ListAnnotations();
        Assert.Single(listed, a => a.Id == "ann-001" && a.LabelId == "RISK");
    }

    [Fact]
    public void AW002_AddAnnotation_NullSpan_BookmarksWholeBlock()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body)
            .First(a => a.Anchor.Kind == "p" && a.TextPreview.Length > 0);

        var result = session.AddAnnotation(firstP.Anchor.Id, span: null,
            new DocumentAnnotation { Id = "ann-whole", LabelId = "L", Label = "L", Color = "#FFF" });

        Assert.True(result.Success);
        var listed = session.ListAnnotations().Single(a => a.Id == "ann-whole");

        // TextPreview is capped at ~80 chars (with trailing "…" when truncated).
        // AnnotatedText is the full run text of the block.  When the paragraph is
        // short the two are equal; when it is long, AnnotatedText starts with the
        // preview characters (minus the ellipsis sentinel).
        if (firstP.TextPreview.EndsWith("…"))
            Assert.StartsWith(firstP.TextPreview[..^1], listed.AnnotatedText,
                StringComparison.Ordinal);
        else
            Assert.Equal(firstP.TextPreview, listed.AnnotatedText);
    }

    [Fact]
    public void AW003_AddAnnotation_AnchorNotFound_ReturnsError()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var result = session.AddAnnotation("p:body:DEADBEEFDEADBEEF", new CharSpan(0, 1),
            new DocumentAnnotation { Id = "x", LabelId = "L", Label = "L", Color = "#000" });

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.AnchorNotFound, result.Error!.Code);
    }

    [Fact]
    public void AW004_AddAnnotation_EmptySpan_ReturnsError()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");

        var result = session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 0),
            new DocumentAnnotation { Id = "x", LabelId = "L", Label = "L", Color = "#000" });

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.EmptyAnnotationSpan, result.Error!.Code);
    }

    [Fact]
    public void AW005_AddAnnotation_OutOfRangeSpan_ReturnsError()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body)
            .First(a => a.Anchor.Kind == "p" && a.TextPreview.Length > 0);

        var result = session.AddAnnotation(firstP.Anchor.Id,
            new CharSpan(0, firstP.TextPreview.Length + 9999),
            new DocumentAnnotation { Id = "x", LabelId = "L", Label = "L", Color = "#000" });

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.OffsetOutOfRange, result.Error!.Code);
    }

    [Fact]
    public void AW006_AddAnnotation_DuplicateCallerId_ReturnsError()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");

        var first = session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation { Id = "dup", LabelId = "L", Label = "L", Color = "#000" });
        Assert.True(first.Success);

        var second = session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation { Id = "dup", LabelId = "L", Label = "L", Color = "#000" });
        Assert.False(second.Success);
        Assert.Equal(EditErrorCode.DuplicateAnnotationId, second.Error!.Code);
    }

    [Fact]
    public void AW007_AddAnnotation_AutoId_GeneratesUnique16Hex()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");

        var result = session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation { LabelId = "L", Label = "L", Color = "#000" });

        Assert.True(result.Success);
        Assert.NotNull(result.AnnotationId);
        Assert.Equal(16, result.AnnotationId!.Length);
        Assert.Matches("^[0-9a-f]{16}$", result.AnnotationId);
    }

    [Fact]
    public void AW008_AddAnnotation_UndoRollsBack()
    {
        using var session = new DocxSession(LoadFixture(Fixture));
        var firstP = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");

        session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 1),
            new DocumentAnnotation { Id = "undoable", LabelId = "L", Label = "L", Color = "#000" });
        Assert.Single(session.ListAnnotations(), a => a.Id == "undoable");

        Assert.True(session.Undo());
        Assert.DoesNotContain(session.ListAnnotations(), a => a.Id == "undoable");

        Assert.True(session.Redo());
        Assert.Single(session.ListAnnotations(), a => a.Id == "undoable");
    }

    [Fact]
    public void AW009_AddAnnotation_SaveAndReopen_Persists()
    {
        byte[] saved;
        using (var session = new DocxSession(LoadFixture(Fixture)))
        {
            var firstP = session.AnchorsByScope(ProjectionScopes.Body).First(a => a.Anchor.Kind == "p");
            var r = session.AddAnnotation(firstP.Anchor.Id, new CharSpan(0, 4),
                new DocumentAnnotation { Id = "persist", LabelId = "PERSIST", Label = "P", Color = "#0F0",
                    Metadata = new System.Collections.Generic.Dictionary<string, string> { ["k"] = "v" } });
            Assert.True(r.Success);
            saved = session.Save();
        }

        using var reopened = new DocxSession(saved);
        var found = reopened.ListAnnotations().Single(a => a.Id == "persist");
        Assert.Equal("PERSIST", found.LabelId);
        Assert.Equal("v", found.Metadata["k"]);
    }
}
