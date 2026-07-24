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
    private static readonly System.DateTime RevisionDate = System.DateTime.Parse("2026-01-01");

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

    [Fact]
    public void Accept_DeletedParagraphBeforeEmptyBlockSdt_PreservesTheSdtAsItsOwnBlock()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var deleted = DeletedParagraph("removed before control");
            var control = new SdtBlock(
                new SdtProperties(new Tag { Val = "empty-block-control" }),
                new SdtContentBlock());
            main.Document = new Document(new Body(
                deleted,
                control,
                new Paragraph(new Run(new Text("following paragraph")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults());
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }

        var accepted = RevisionProcessor.AcceptRevisions(new WmlDocument("d.docx", stream.ToArray()));

        using var acceptedStream = new MemoryStream(accepted.DocumentByteArray);
        using var acceptedDoc = WordprocessingDocument.Open(acceptedStream, false);
        var body = acceptedDoc.MainDocumentPart!.Document.Body!;
        var blocks = body.ChildElements.ToList();
        Assert.Collection(blocks,
            first =>
            {
                var sdt = Assert.IsType<SdtBlock>(first);
                Assert.Equal("empty-block-control", sdt.SdtProperties!.GetFirstChild<Tag>()!.Val!.Value);
            },
            second => Assert.Equal("following paragraph", Assert.IsType<Paragraph>(second).InnerText));
    }

    [Fact]
    public void Accept_DeletedInnerParagraph_DoesNotMergeAcrossTheBlockSdtBoundary()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var control = new SdtBlock(
                new SdtProperties(new Tag { Val = "outer-control" }),
                new SdtContentBlock(
                    DeletedParagraph("removed inside control"),
                    new Paragraph(new Run(new Text("live inside control")))));
            main.Document = new Document(new Body(
                control,
                new Paragraph(new Run(new Text("outside control")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults());
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }

        var accepted = RevisionProcessor.AcceptRevisions(new WmlDocument("d.docx", stream.ToArray()));

        using var acceptedStream = new MemoryStream(accepted.DocumentByteArray);
        using var acceptedDoc = WordprocessingDocument.Open(acceptedStream, false);
        var body = acceptedDoc.MainDocumentPart!.Document.Body!;
        var blocks = body.ChildElements.ToList();
        var acceptedControl = Assert.IsType<SdtBlock>(blocks[0]);
        var content = acceptedControl.SdtContentBlock!;
        var inside = Assert.Single(content.Elements<Paragraph>());
        Assert.Equal("live inside control", inside.InnerText);
        Assert.Equal("outside control", Assert.IsType<Paragraph>(blocks[1]).InnerText);
    }

    [Fact]
    public void Accept_InlineSdtWithoutRevisions_PreservesTheInlineControl()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var inlineControl = new SdtRun(
                new SdtProperties(new Tag { Val = "inline-control" }),
                new SdtContentRun(new Run(new Text("controlled text"))));
            main.Document = new Document(new Body(
                new Paragraph(
                    new Run(new Text("before ")),
                    inlineControl,
                    new Run(new Text(" after")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults());
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }

        var accepted = RevisionProcessor.AcceptRevisions(new WmlDocument("d.docx", stream.ToArray()));

        using var acceptedStream = new MemoryStream(accepted.DocumentByteArray);
        using var acceptedDoc = WordprocessingDocument.Open(acceptedStream, false);
        var paragraph = acceptedDoc.MainDocumentPart!.Document.Body!.GetFirstChild<Paragraph>()!;
        var acceptedInlineControl = Assert.Single(paragraph.Elements<SdtRun>());
        Assert.Equal("inline-control", acceptedInlineControl.SdtProperties!.GetFirstChild<Tag>()!.Val!.Value);
        Assert.Equal("controlled text", acceptedInlineControl.InnerText);
        Assert.Equal("before controlled text after", paragraph.InnerText);
    }

    private static Paragraph DeletedParagraph(string text) =>
        new(
            new ParagraphProperties(new ParagraphMarkRunProperties(
                new Deleted { Author = "A", Date = RevisionDate })),
            new DeletedRun(
                new Run(new DeletedText(text) { Space = SpaceProcessingModeValues.Preserve }))
            { Author = "A", Date = RevisionDate });
}
