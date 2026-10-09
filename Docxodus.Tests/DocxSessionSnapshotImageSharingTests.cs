// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Undo snapshots share each image's bytes by reference instead of copying them on every mutation
/// (issue #965), and the undo memory budget counts each shared array once. Before, a text edit on a
/// document with large pictures copied every picture into its snapshot, and the budget charged every
/// snapshot for them, so a few edits exhausted it.
/// </summary>
public class DocxSessionSnapshotImageSharingTests
{
    private const int ImageBytes = 1024 * 1024;

    /// <summary>A document with <paramref name="paragraphs"/> paragraphs and one embedded 1 MiB PNG.</summary>
    private static byte[] DocumentWithImage(int paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            var image = main.AddImagePart(ImagePartType.Png);
            using (var input = new MemoryStream(Png(64, 64, ImageBytes))) image.FeedData(input);
            var body = new StringBuilder(
                "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"914400\" cy=\"914400\"/><wp:docPr id=\"1\" name=\"Picture 1\"/>" +
                "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic>" +
                "<pic:nvPicPr><pic:cNvPr id=\"1\" name=\"image1.png\"/><pic:cNvPicPr/></pic:nvPicPr>" +
                $"<pic:blipFill><a:blip r:embed=\"{main.GetIdOfPart(image)}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
                "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/></a:xfrm><a:prstGeom prst=\"rect\"/></pic:spPr>" +
                "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>");
            for (var p = 0; p < paragraphs; p++)
                body.Append($"<w:p><w:r><w:t>Paragraph {p} of the document.</w:t></w:r></w:p>");
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
                writer.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                    "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                    "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                    $"<w:body>{body}<w:sectPr/></w:body></w:document>");
            doc.Save();
        }
        return stream.ToArray();
    }

    /// <summary>A PNG whose IHDR declares the size, padded to <paramref name="totalBytes"/>.</summary>
    private static byte[] Png(int width, int height, int totalBytes, byte fill = 0)
    {
        var bytes = new byte[totalBytes];
        Array.Fill(bytes, fill);
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, (byte)'I', (byte)'H', (byte)'D', (byte)'R' }
            .CopyTo(bytes, 0);
        bytes[16] = (byte)(width >> 24); bytes[17] = (byte)(width >> 16); bytes[18] = (byte)(width >> 8); bytes[19] = (byte)width;
        bytes[20] = (byte)(height >> 24); bytes[21] = (byte)(height >> 16); bytes[22] = (byte)(height >> 8); bytes[23] = (byte)height;
        return bytes;
    }

    /// <summary>The text paragraphs, after the first one, which holds the picture.</summary>
    private static List<string> Paragraphs(DocxSession session) =>
        session.Project().AnchorIndex.Values.Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p")
            .Select(t => t.Anchor.Id).Skip(1).ToList();

    private static byte[] LiveImageBytes(DocxSession session)
    {
        var image = Assert.Single(session.ListImages());
        using var output = new MemoryStream();
        using (var input = session.LiveDocument.GetPackage().GetPart(new Uri(image.TargetPartUri!, UriKind.Relative))
                   .GetStream(FileMode.Open, FileAccess.Read))
            input.CopyTo(output);
        return output.ToArray();
    }

    [Fact]
    public void ConsecutiveSnapshots_ShareAnUnchangedImagesBytes()
    {
        using var session = new DocxSession(DocumentWithImage(5));
        var first = Assert.Single(session.TakeSnapshot().ImageParts).Bytes;
        Assert.True(session.ReplaceText(Paragraphs(session)[0], "An edited paragraph.").Success);
        var second = Assert.Single(session.TakeSnapshot().ImageParts).Bytes;

        Assert.Same(first, second);
    }

    [Fact]
    public void TextEdits_OnADocumentWithALargeImage_KeepEveryUndoStepWithinTheBudget()
    {
        // Four MiB holds the one image once plus many small XML snapshots, but not one image per step.
        using var session = new DocxSession(DocumentWithImage(5),
            new DocxSessionSettings { UndoDepth = 20, UndoMemoryBudgetBytes = 4L * ImageBytes });
        var paragraphs = Paragraphs(session);
        for (var i = 0; i < 10; i++)
            Assert.True(session.ReplaceText(paragraphs[i % paragraphs.Count], $"Edit {i}.").Success);

        Assert.Equal(10, session.UndoCount);
        Assert.False(session.UndoHistoryTrimmedForMemory);
        Assert.InRange(session.UndoMemoryBytes, ImageBytes, 2L * ImageBytes);
    }

    [Fact]
    public void UndoAndRedo_AcrossAReplacedImageAndTextEdits_RestoreEachImage()
    {
        using var session = new DocxSession(DocumentWithImage(5));
        var original = LiveImageBytes(session);
        var replacement = Png(32, 32, 2048, fill: 7);
        var imageId = Assert.Single(session.ListImages()).Id;

        Assert.True(session.ReplaceText(Paragraphs(session)[0], "Before the replacement.").Success);
        Assert.True(session.ReplaceImage(imageId, replacement).Success);
        Assert.True(session.ReplaceText(Paragraphs(session)[1], "After the replacement.").Success);
        Assert.Equal(replacement, LiveImageBytes(session));

        Assert.True(session.Undo());
        Assert.Equal(replacement, LiveImageBytes(session));
        Assert.True(session.Undo());
        Assert.Equal(original, LiveImageBytes(session));
        Assert.True(session.Undo());
        Assert.Equal(original, LiveImageBytes(session));

        Assert.True(session.Redo());
        Assert.True(session.Redo());
        Assert.Equal(replacement, LiveImageBytes(session));
    }

    /// <summary>An undo that restores image topology reopens the package; the restored parts must
    /// still share the snapshot's arrays, or every older snapshot's copy is counted a second time.</summary>
    [Fact]
    public void UndoThatRestoresAnImage_KeepsSharingTheSnapshotsBytes()
    {
        using var session = new DocxSession(DocumentWithImage(5));
        var original = Assert.Single(session.TakeSnapshot().ImageParts).Bytes;
        Assert.True(session.ReplaceImage(Assert.Single(session.ListImages()).Id, Png(32, 32, 2048, fill: 7)).Success);

        Assert.True(session.Undo());

        Assert.Same(original, Assert.Single(session.TakeSnapshot().ImageParts).Bytes);
    }

    [Fact]
    public void UndoRing_CountsAPayloadSharedByManyEntriesOnce()
    {
        var shared = new byte[1000];
        var ring = new UndoRing<string>(10, budgetBytes: 1_000_000, costOf: _ => 10,
            sharedPayloadsOf: s => s == "own" ? new[] { ((object)new byte[500], 500L) } : new[] { ((object)shared, 1000L) });

        ring.RecordPreOp("a");
        ring.RecordPreOp("b");
        ring.RecordPreOp("c");
        Assert.Equal(3 * 10 + 1000, ring.RetainedBytes);

        var state = ring.CaptureState();
        ring.RecordPreOp("own");
        Assert.Equal(4 * 10 + 1000 + 500, ring.RetainedBytes);

        ring.RestoreState(state);
        Assert.Equal(3 * 10 + 1000, ring.RetainedBytes);

        ring.PopForUndo();
        ring.PopForUndo();
        ring.PopForUndo();
        Assert.Equal(0, ring.RetainedBytes);
    }

    [Fact]
    public void UndoRing_EvictsByTheDistinctSharedTotal()
    {
        var shared = new byte[1000];
        var ring = new UndoRing<string>(10, budgetBytes: 1100, costOf: _ => 10,
            sharedPayloadsOf: _ => new[] { ((object)shared, 1000L) });

        for (var i = 0; i < 8; i++) ring.RecordPreOp($"edit {i}");

        Assert.Equal(8, ring.UndoCount);
        Assert.False(ring.EvictedForMemory);
    }
}
