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
}
