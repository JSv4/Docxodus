#nullable enable

using System.IO;
using System.Linq;
using Docxodus.Ir.Diff;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="IrMarkupRenderer.PruneOrphanedComments"/> removes comment definitions that no marker
/// references — the orphans a LEFT-cloned comments part leaves behind when the diff fully replaces the
/// content a LEFT comment annotated (its <c>commentRangeStart</c>/<c>commentReference</c> gone everywhere,
/// not even preserved inside a <c>w:del</c>). Word's compare output never emits such an orphan. Definitions
/// that ARE still referenced — including a LEFT comment whose marker survives inside a <c>w:del</c> so
/// reject can restore it — must be kept.
/// </summary>
public class DocxDiffCommentPruneTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
    private const string W15 = "http://schemas.microsoft.com/office/word/2012/wordml";

    /// <summary>Build a package whose body references <paramref name="referencedId"/> only, but whose
    /// comments part defines both that id and <paramref name="orphanId"/> (with a w14:paraId + a matching
    /// commentsExtended entry) — the dangling shape a LEFT-cloned comments part produces.</summary>
    private static byte[] BuildWithOrphan(string referencedId, string orphanId, bool orphanInDelMarker)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            // Body references the live comment; the orphan's marker is either absent entirely, or (control
            // case) present but wrapped in a w:del so it is still referenced and MUST survive.
            var orphanMarker = orphanInDelMarker
                ? $"<w:del w:id=\"9\" w:author=\"A\" w:date=\"2020-01-01T00:00:00Z\">" +
                  $"<w:commentRangeStart w:id=\"{orphanId}\"/><w:r><w:delText>x</w:delText></w:r>" +
                  $"<w:commentRangeEnd w:id=\"{orphanId}\"/><w:r><w:commentReference w:id=\"{orphanId}\"/></w:r></w:del>"
                : "";
            using (var w = new StreamWriter(main.GetStream(FileMode.Create), new System.Text.UTF8Encoding(false)))
                w.Write($"<w:document xmlns:w=\"{W}\"><w:body><w:p>" +
                        $"<w:commentRangeStart w:id=\"{referencedId}\"/><w:r><w:t>live</w:t></w:r>" +
                        $"<w:commentRangeEnd w:id=\"{referencedId}\"/><w:r><w:commentReference w:id=\"{referencedId}\"/></w:r>" +
                        orphanMarker + "</w:p></w:body></w:document>");

            var comments = main.AddNewPart<WordprocessingCommentsPart>();
            using (var w = new StreamWriter(comments.GetStream(FileMode.Create), new System.Text.UTF8Encoding(false)))
                w.Write($"<w:comments xmlns:w=\"{W}\" xmlns:w14=\"{W14}\">" +
                        $"<w:comment w:id=\"{referencedId}\" w:author=\"A\" w:date=\"2020-01-01T00:00:00Z\"><w:p w14:paraId=\"00000001\"><w:r><w:t>live note</w:t></w:r></w:p></w:comment>" +
                        $"<w:comment w:id=\"{orphanId}\" w:author=\"A\" w:date=\"2020-01-01T00:00:00Z\"><w:p w14:paraId=\"00000002\"><w:r><w:t>orphan note</w:t></w:r></w:p></w:comment>" +
                        "</w:comments>");

            var ex = main.AddNewPart<WordprocessingCommentsExPart>();
            using (var w = new StreamWriter(ex.GetStream(FileMode.Create), new System.Text.UTF8Encoding(false)))
                w.Write($"<w15:commentsEx xmlns:w15=\"{W15}\">" +
                        "<w15:commentEx w15:paraId=\"00000001\" w15:done=\"0\"/>" +
                        "<w15:commentEx w15:paraId=\"00000002\" w15:done=\"0\"/></w15:commentsEx>");
        }
        return ms.ToArray();
    }

    private static (string[] defIds, string[] exParaIds) Read(byte[] bytes, System.Action<MainDocumentPart> mutate)
    {
        using var ms = new MemoryStream();
        ms.Write(bytes, 0, bytes.Length); ms.Position = 0;
        using var doc = WordprocessingDocument.Open(ms, true);
        var main = doc.MainDocumentPart!;
        mutate(main);
        var cRoot = System.Xml.Linq.XDocument.Load(main.WordprocessingCommentsPart!.GetStream()).Root!;
        var defs = cRoot.Elements().Where(e => e.Name.LocalName == "comment")
            .Select(e => e.Attribute(System.Xml.Linq.XName.Get("id", W))!.Value).ToArray();
        var exRoot = System.Xml.Linq.XDocument.Load(main.WordprocessingCommentsExPart!.GetStream()).Root!;
        var exIds = exRoot.Elements().Select(e => e.Attribute(System.Xml.Linq.XName.Get("paraId", W15))!.Value).ToArray();
        return (defs, exIds);
    }

    [Fact]
    public void PrunesUnreferencedDefinition_AndItsThreadingEntry()
    {
        var (defs, exIds) = Read(BuildWithOrphan("0", "5", orphanInDelMarker: false),
            IrMarkupRenderer.PruneOrphanedComments);

        Assert.Equal(new[] { "0" }, defs);                 // orphan "5" removed, referenced "0" kept
        Assert.Equal(new[] { "00000001" }, exIds);         // orphan's commentsExtended entry pruned too
    }

    [Fact]
    public void KeepsDefinition_WhenItsMarkerSurvivesInsideADeletion()
    {
        // Control: the "orphan" id IS referenced — inside a w:del, so reject restores it. It must NOT be pruned.
        var (defs, _) = Read(BuildWithOrphan("0", "5", orphanInDelMarker: true),
            IrMarkupRenderer.PruneOrphanedComments);

        Assert.Contains("0", defs);
        Assert.Contains("5", defs);
    }
}
