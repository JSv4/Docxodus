#nullable enable

using System.IO;
using System.Linq;
using Docxodus;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Accepting a document whose block-level content control's runs are ALL deleted (the SDT's
/// paragraph carries a deleted mark and every run is inside <c>w:del</c>) must not throw —
/// the legacy content-control reattachment assumed every recorded run survives accept
/// (<c>First(...)</c> over surviving runs) and crashed with "Sequence contains no matching
/// element" on five Word-compare corpus outputs carrying SDT controls in deleted regions.
/// </summary>
public class RevisionProcessorSdtDeletedRunsTests
{
    [Fact]
    public void Accept_SdtWithAllRunsDeleted_DoesNotThrow()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var deletedPara = new Paragraph(
                new ParagraphProperties(new ParagraphMarkRunProperties(
                    new Deleted { Author = "A", Date = System.DateTime.Parse("2026-01-01") })),
                new DeletedRun(
                    new Run(new DeletedText("gone text") { Space = SpaceProcessingModeValues.Preserve }))
                { Author = "A", Date = System.DateTime.Parse("2026-01-01") });
            var sdt = new SdtBlock(
                new SdtProperties(new SdtId { Val = 77 }),
                new SdtContentBlock(deletedPara));
            main.Document = new Document(new Body(
                sdt,
                new Paragraph(new Run(new Text("survivor paragraph text")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults());
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        var input = new WmlDocument("d.docx", stream.ToArray());

        var accepted = RevisionProcessor.AcceptRevisions(input);   // must not throw

        using var s = new MemoryStream(accepted.DocumentByteArray);
        using var d = WordprocessingDocument.Open(s, false);
        var text = d.MainDocumentPart!.Document.Body!.InnerText;
        Assert.Contains("survivor paragraph text", text);
        Assert.DoesNotContain("gone text", text);
    }
}
