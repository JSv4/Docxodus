// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionImageTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";
    private static readonly XNamespace MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";

    private static string[] Paragraphs(DocxSession session, string scope = "body") =>
        session.Project().AnchorIndex.Values.Where(target => target.Anchor.Scope == scope
            && target.Anchor.Kind is "p" or "h" or "li")
            .Select(target => target.Anchor.Id).Distinct().ToArray();

    [Fact]
    public void IM001_InsertInspectMutateRemove_RoundTripsSchemaValidDrawing()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = Paragraphs(session)[0];
        var insert = session.InsertImage(anchor, 5, Png(2, 3), new ImageInsertOptions
        {
            AltText = "diagram", Title = "title", WidthPoints = 72,
        });
        Assert.True(insert.Success, insert.Error?.Message);
        var image = Assert.Single(session.ListImages());
        Assert.Equal(insert.ImageId, image.Id);
        Assert.Equal(ImageBinaryFormat.Png, image.Format);
        Assert.Equal(2, image.IntrinsicWidthPixels);
        Assert.Equal(3, image.IntrinsicHeightPixels);
        Assert.Equal(72, image.RenderedWidthPoints!.Value, 6);
        Assert.Equal(108, image.RenderedHeightPoints!.Value, 6);
        Assert.Equal(new CharSpan(5, 0), image.Span);
        Assert.True(image.ContentTypeMatchesBytes);

        Assert.True(session.SetImageMetadata(image.Id, "updated", null).Success);
        Assert.True(session.SetImageDimensions(image.Id, 36, null).Success);
        image = Assert.Single(session.ListImages());
        Assert.Equal("updated", image.AltText);
        Assert.Equal(36, image.RenderedWidthPoints!.Value, 6);
        Assert.Equal(54, image.RenderedHeightPoints!.Value, 6);

        var saved = session.Save(true);
        using (var stream = new MemoryStream(saved))
        using (var document = WordprocessingDocument.Open(stream, false))
            Assert.Empty(new OpenXmlValidator().Validate(document).Where(IsRealValidationError));

        Assert.True(session.RemoveImage(image.Id).Success);
        Assert.Empty(session.ListImages());
        Assert.Empty(ImageRelationships(session.Save(true)).SelectMany(pair => pair.Relationships));
    }

    [Theory]
    [MemberData(nameof(SupportedFormats))]
    public void IM002_SupportedMagicFormats_AreAcceptedAndReported(byte[] bytes,
        ImageBinaryFormat expected)
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var result = session.InsertImage(Paragraphs(session)[0], 0, bytes);
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(expected, Assert.Single(session.ListImages()).Format);
    }

    public static IEnumerable<object[]> SupportedFormats()
    {
        yield return new object[] { Png(2, 3), ImageBinaryFormat.Png };
        yield return new object[] { Jpeg(3, 2), ImageBinaryFormat.Jpeg };
        yield return new object[] { Gif(4, 5), ImageBinaryFormat.Gif };
        yield return new object[] { Bmp(6, 7), ImageBinaryFormat.Bmp };
        yield return new object[] { Tiff(8, 9), ImageBinaryFormat.Tiff };
    }

    [Fact]
    public void IM003_DedupAcrossOwners_AndUndoRestoreExactTopology()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var body = Paragraphs(session);
        Assert.True(session.SetHeaderText(body[0], HeaderFooterKind.Default, "header").Success);
        var header = Assert.Single(Paragraphs(session, "hdr1"));
        var bytes = Png(11, 13);
        var bodyInsert = session.InsertImage(body[0], 0, bytes);
        var headerInsert = session.InsertImage(header, 0, bytes);
        Assert.True(bodyInsert.Success && headerInsert.Success);
        var before = ImageRelationships(session.Save(true));
        Assert.Equal(2, before.Count);
        Assert.Single(before.SelectMany(owner => owner.Relationships).Select(rel => rel.TargetUri).Distinct());
        Assert.Equal(2, DocumentPropertyIds(session.Save(true)).Distinct().Count());
        var exact = before.SelectMany(owner => owner.Relationships
            .Select(rel => (owner.OwnerUri, rel.RelId, rel.TargetUri))).OrderBy(value => value).ToArray();

        Assert.True(session.RemoveImage(bodyInsert.ImageId!).Success);
        Assert.Single(session.ListImages());
        Assert.True(session.Undo());
        var restored = ImageRelationships(session.Save(true)).SelectMany(owner => owner.Relationships
            .Select(rel => (owner.OwnerUri, rel.RelId, rel.TargetUri))).OrderBy(value => value).ToArray();
        Assert.Equal(exact, restored);
        Assert.True(session.Redo());
        Assert.Single(session.ListImages());
    }

    [Fact]
    public void IM004_ReplaceUndoRedo_RestoresBytesContentTypeRelIdAndTargetUri()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var insert = session.InsertImage(Paragraphs(session)[0], 0, Png(2, 3));
        var original = Assert.Single(session.ListImages());
        Assert.True(session.ReplaceImage(original.Id, Jpeg(7, 5)).Success);
        var replaced = Assert.Single(session.ListImages());
        Assert.Equal(ImageBinaryFormat.Jpeg, replaced.Format);
        Assert.True(session.Undo());
        var undone = Assert.Single(session.ListImages());
        Assert.Equal(ImageBinaryFormat.Png, undone.Format);
        Assert.Equal(original.RelationshipId, undone.RelationshipId);
        Assert.Equal(original.TargetPartUri, undone.TargetPartUri);
        Assert.True(session.Redo());
        var redone = Assert.Single(session.ListImages());
        Assert.Equal(replaced.RelationshipId, redone.RelationshipId);
        Assert.Equal(replaced.TargetPartUri, redone.TargetPartUri);
    }

    [Fact]
    public void IM005_NoOpsAndRejectedInputs_DoNotCreateHistory()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var png = Png(2, 3);
        var insert = session.InsertImage(Paragraphs(session)[0], 0, png);
        var image = Assert.Single(session.ListImages());
        int undo = session.UndoCount;
        Assert.True(session.ReplaceImage(image.Id, png).Success);
        Assert.True(session.SetImageDimensions(image.Id, image.RenderedWidthPoints, image.RenderedHeightPoints, false).Success);
        Assert.True(session.SetImageMetadata(image.Id, image.AltText, image.Title).Success);
        Assert.Equal(undo, session.UndoCount);

        Assert.Equal(EditErrorCode.InvalidImageData, session.ReplaceImage(image.Id, Array.Empty<byte>()).Error!.Code);
        Assert.Equal(EditErrorCode.UnsupportedImageFormat, session.ReplaceImage(image.Id,
            new byte[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 }).Error!.Code);
        Assert.Equal(EditErrorCode.InvalidImageDimensions,
            session.SetImageDimensions(image.Id, double.PositiveInfinity, null).Error!.Code);
        Assert.Equal(undo, session.UndoCount);
    }

    [Fact]
    public void IM006_FloatingSubsetRoundTrips_AndUnsupportedTokensAreReadOnly()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var layout = new FloatingImageLayout
        {
            HorizontalOffsetEmu = 12345, VerticalOffsetEmu = -23456,
            WrapMode = ImageWrapMode.Square, WrapSide = ImageWrapSide.Right,
            DistanceLeftEmu = 100, BehindDocument = true, AllowOverlap = false,
        };
        Assert.True(session.InsertImage(Paragraphs(session)[0], 0, Png(4, 5),
            new ImageInsertOptions { Placement = ImagePlacement.Floating, FloatingLayout = layout }).Success);
        var image = Assert.Single(session.ListImages());
        Assert.True(image.FloatingLayoutSupported,
            image.UnsupportedReason + ": " + image.FloatingLayout?.RawHorizontalPosition);
        Assert.Equal(layout, image.FloatingLayout);

        var mutated = MutatePackage(session.Save(true), document =>
        {
            var anchor = document.MainDocumentPart!.GetXDocument().Descendants(WP + "anchor").Single();
            anchor.SetAttributeValue("behindDoc", "banana");
            anchor.Element(WP + "wrapSquare")!.ReplaceWith(new XElement(WP + "wrapTight",
                new XAttribute("wrapText", "right")));
            document.MainDocumentPart.PutXDocument();
        });
        using var probe = new DocxSession(mutated);
        var unsupported = Assert.Single(probe.ListImages());
        Assert.False(unsupported.CanMutate);
        Assert.False(unsupported.FloatingLayoutSupported);
        Assert.Equal(ImageWrapMode.Tight, unsupported.FloatingLayout!.WrapMode);
        Assert.Equal("banana", unsupported.FloatingLayout.RawFlagTokens!["behindDoc"]);
    }

    [Fact]
    public void IM006A_FloatingPositionAndWrapExtrasAreEnumeratedButReadOnly()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.InsertImage(Paragraphs(seed)[0], 0, Png(4, 5),
            new ImageInsertOptions { Placement = ImagePlacement.Floating }).Success);
        var mutated = MutatePackage(seed.Save(true), document =>
        {
            var main = document.MainDocumentPart!;
            var anchor = main.GetXDocument().Descendants(WP + "anchor").Single();
            anchor.Element(WP + "positionH")!.Add(
                new XElement(WP14 + "pctPosHOffset", "50000"));
            anchor.Element(WP + "wrapSquare")!.SetAttributeValue("distL", "123");
            main.PutXDocument();
        });

        using var session = new DocxSession(mutated);
        var image = Assert.Single(session.ListImages());
        Assert.False(image.CanMutate);
        Assert.False(image.FloatingLayoutSupported);
        Assert.Contains("pctPosHOffset", image.FloatingLayout!.RawHorizontalPosition);
        Assert.Contains("distL=\"123\"", image.FloatingLayout.RawWrapMode);
        Assert.Equal(EditErrorCode.UnsupportedImageMarkup,
            session.SetImageFloatingLayout(image.Id, new FloatingImageLayout()).Error!.Code);
    }

    [Fact]
    public void IM007_ContentTypeMagicMismatchIsReportedBroken()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.InsertImage(Paragraphs(seed)[0], 0, Png(2, 3)).Success);
        var bytes = MutatePackage(seed.Save(true), document =>
        {
            var image = document.MainDocumentPart!.ImageParts.Single();
            using var input = new MemoryStream(Gif(4, 5), writable: false);
            image.FeedData(input);
        });
        using var session = new DocxSession(bytes);
        var image = Assert.Single(session.ListImages());
        Assert.Equal(ImageBinaryFormat.Gif, image.Format);
        Assert.False(image.ContentTypeMatchesBytes);
        Assert.True(image.IsBroken);
    }

    [Fact]
    public void IM008_RawReplaceRemovingDrawing_CleansRelationshipAndUndoRestoresIt()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = Paragraphs(session)[0];
        Assert.True(session.InsertImage(anchor, 0, Png(2, 3)).Success);
        var before = Assert.Single(session.ListImages());
        var paragraph = XElement.Parse(session.Raw.GetXml(anchor));
        paragraph.Descendants(W + "drawing").Remove();
        Assert.True(session.Raw.ReplaceXml(anchor, paragraph.ToString()).Success);
        Assert.Empty(session.ListImages());
        Assert.Empty(ImageRelationships(session.Save(true)).SelectMany(owner => owner.Relationships));
        Assert.True(session.Undo());
        Assert.Equal(before.TargetPartUri, Assert.Single(session.ListImages()).TargetPartUri);
    }

    [Fact]
    public void IM009_JsonBoundaryRejectsMalformedValuesAndUsesBase64()
    {
        var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var probe = new DocxSession(bytes);
        var anchor = Paragraphs(probe)[0];
        int handle = Docxodus.Internal.DocxSessionOps.OpenSession(bytes, null);
        try
        {
            using var malformed = JsonDocument.Parse(Docxodus.Internal.DocxSessionOps.InsertImage(
                handle, anchor, 0, Convert.ToBase64String(Png(2, 3)), "{\"widthPoints\":\"wide\"}"));
            Assert.False(malformed.RootElement.GetProperty("success").GetBoolean());
            using var malformedLayout = JsonDocument.Parse(Docxodus.Internal.DocxSessionOps.InsertImage(
                handle, anchor, 0, Convert.ToBase64String(Png(2, 3)), "{\"floatingLayout\":false}"));
            Assert.False(malformedLayout.RootElement.GetProperty("success").GetBoolean());
            Assert.Equal("invalid_image_layout",
                malformedLayout.RootElement.GetProperty("error").GetProperty("code").GetString());
            using var badBase64 = JsonDocument.Parse(Docxodus.Internal.DocxSessionOps.InsertImage(
                handle, anchor, 0, "***", "{}"));
            Assert.Equal("invalid_image_data", badBase64.RootElement.GetProperty("error").GetProperty("code").GetString());
            using var images = JsonDocument.Parse(Docxodus.Internal.DocxSessionOps.ListImages(handle));
            Assert.Empty(images.RootElement.EnumerateArray());
        }
        finally { Docxodus.Internal.DocxSessionOps.CloseSession(handle); }
    }

    [Fact]
    public void IM010_ExternalLinkIsReadOnly_AndUndoRestoresExactRelationship()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.InsertImage(Paragraphs(seed)[0], 0, Png(2, 3)).Success);
        const string relationshipId = "rIdExternalImage37";
        const string target = "https://example.test/image.png";
        var linkedBytes = MutatePackage(seed.Save(true), document =>
        {
            var main = document.MainDocumentPart!;
            var blip = main.GetXDocument().Descendants(A + "blip").Single();
            blip.Attribute(R + "embed")!.Remove();
            blip.SetAttributeValue(R + "link", relationshipId);
            main.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                new Uri(target), relationshipId);
            main.PutXDocument();
        });

        using var session = new DocxSession(linkedBytes);
        var image = Assert.Single(session.ListImages());
        Assert.True(image.IsLinked);
        Assert.False(image.CanMutate);
        Assert.Equal(relationshipId, image.LinkedRelationshipId);
        Assert.Equal(target, image.LinkedTarget);
        Assert.Equal(EditErrorCode.LinkedImageReadOnly,
            session.ReplaceImage(image.Id, Png(4, 5)).Error!.Code);

        Assert.True(session.ReplaceText(Paragraphs(session)[1], "changed").Success);
        Assert.True(session.Undo());
        image = Assert.Single(session.ListImages());
        Assert.Equal(relationshipId, image.LinkedRelationshipId);
        Assert.Equal(target, image.LinkedTarget);
    }

    [Fact]
    public void IM011_LegacyVmlOccurrenceKeepsSharedRelationshipWhenModernImageIsRemoved()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.InsertImage(Paragraphs(seed)[0], 0, Png(2, 3)).Success);
        var mixed = MutatePackage(seed.Save(true), document =>
        {
            var main = document.MainDocumentPart!;
            var root = main.GetXDocument();
            var relationshipId = (string)root.Descendants(A + "blip").Single().Attribute(R + "embed")!;
            root.Descendants(W + "p").First().Add(new XElement(W + "r",
                new XElement(W + "pict",
                    new XElement(V + "shape", new XAttribute("alt", "legacy"),
                        new XElement(V + "imagedata", new XAttribute(R + "id", relationshipId))))));
            main.PutXDocument();
        });

        using var session = new DocxSession(mixed);
        var images = session.ListImages();
        Assert.Equal(2, images.Count);
        var modern = Assert.Single(images.Where(value => value.MarkupKind == ImageMarkupKind.ModernDrawing));
        var legacy = Assert.Single(images.Where(value => value.MarkupKind == ImageMarkupKind.LegacyVml));
        Assert.False(legacy.CanMutate);
        Assert.Equal(modern.RelationshipId, legacy.RelationshipId);
        Assert.True(session.RemoveImage(modern.Id).Success);
        legacy = Assert.Single(session.ListImages());
        Assert.Equal(ImageMarkupKind.LegacyVml, legacy.MarkupKind);
        Assert.False(legacy.IsBroken);
        Assert.Single(ImageRelationships(session.Save(true)).SelectMany(owner => owner.Relationships));
    }

    [Fact]
    public void IM012_MultiPictureDrawingEnumeratesStableReadOnlySubOccurrences()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.InsertImage(Paragraphs(seed)[0], 0, Png(2, 3)).Success);
        var multi = MutatePackage(seed.Save(true), document =>
        {
            var main = document.MainDocumentPart!;
            var root = main.GetXDocument();
            var blip = root.Descendants(A + "blip").Single();
            blip.AddAfterSelf(new XElement(blip));
            main.PutXDocument();
        });

        using var session = new DocxSession(multi);
        var images = session.ListImages();
        Assert.Equal(2, images.Count);
        Assert.All(images, image =>
        {
            Assert.Equal(ImageMarkupKind.UnsupportedDrawing, image.MarkupKind);
            Assert.False(image.CanMutate);
            Assert.Contains(":sub", image.Id);
        });
        Assert.NotEqual(images[0].Id, images[1].Id);
        Assert.Equal(images.Select(image => image.Id), session.ListImages().Select(image => image.Id));
    }

    [Fact]
    public void IM013_CommentStoryInsertSaveReopenAndUndoPreserveExactTopology()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var body = Paragraphs(seed);
        var comment = seed.AddComment(body[0], null, "Alice", "comment image");
        Assert.True(comment.Success, comment.Error?.Message);
        var commentParagraph = comment.Created.Single(anchor => anchor.Kind == "p" && anchor.Scope == "cmt").Id;
        var inserted = seed.InsertImage(commentParagraph, 7, Png(7, 9));
        Assert.True(inserted.Success, inserted.Error?.Message);
        var commentImage = Assert.Single(seed.ListImages(ProjectionScopes.Comments));
        Assert.Equal("cmt", commentImage.Scope);
        Assert.Equal("/word/comments.xml", commentImage.OwningPartUri);

        var saved = seed.Save(true);
        var exact = ImageRelationships(saved).SelectMany(owner => owner.Relationships
            .Select(relationship => (owner.OwnerUri, relationship.RelId, relationship.TargetUri)))
            .OrderBy(value => value).ToArray();
        using var reopened = new DocxSession(saved);
        commentImage = Assert.Single(reopened.ListImages(ProjectionScopes.Comments));
        Assert.Equal(inserted.ImageId, commentImage.Id);
        Assert.True(reopened.RemoveImage(commentImage.Id).Success);
        Assert.Empty(reopened.ListImages(ProjectionScopes.Comments));
        Assert.True(reopened.Undo());
        Assert.Equal(commentImage.Id, Assert.Single(reopened.ListImages(ProjectionScopes.Comments)).Id);
        var restored = ImageRelationships(reopened.Save(true)).SelectMany(owner => owner.Relationships
            .Select(relationship => (owner.OwnerUri, relationship.RelId, relationship.TargetUri)))
            .OrderBy(value => value).ToArray();
        Assert.Equal(exact, restored);
    }

    [Fact]
    public void IM014_UnrelatedUndoRedoKeepsSdkGraphAndMediaLayerUntouched()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = Paragraphs(session);
        Assert.True(session.InsertImage(paragraphs[0], 0, Png(19, 23)).Success);
        var expectedTopology = FlatImageRelationships(session.Save(true));
        var expectedPayloads = ImagePartPayloads(session.Save(true));

        Assert.True(session.ReplaceText(paragraphs[1], "unrelated text edit").Success);
        var documentField = typeof(DocxSession).GetField("_doc",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!;
        var graphBeforeUndo = documentField.GetValue(session);
        Assert.True(session.Undo());
        Assert.Same(graphBeforeUndo, documentField.GetValue(session));
        Assert.Equal(expectedTopology, FlatImageRelationships(session.Save(true)));
        Assert.Equal(expectedPayloads, ImagePartPayloads(session.Save(true)));

        var graphBeforeRedo = documentField.GetValue(session);
        Assert.True(session.Redo());
        Assert.Same(graphBeforeRedo, documentField.GetValue(session));
        Assert.Equal(expectedTopology, FlatImageRelationships(session.Save(true)));
        Assert.Equal(expectedPayloads, ImagePartPayloads(session.Save(true)));
    }

    [Fact]
    public void IM015_FooterAndTableCellOccurrencesOwnCorrectPartsAndRoundTrip()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var body = Paragraphs(session);
        Assert.True(session.SetFooterText(body[0], HeaderFooterKind.Default, "footer").Success);
        var footer = Assert.Single(Paragraphs(session, "ftr1"));
        var table = session.InsertTable(body[0], Position.After, 1, 1,
            new TableInsertOptions { CellContents = new[] { "cell" } });
        Assert.True(table.Success, table.Error?.Message);
        // #450 returns canonical structural identities via TableAnchors; images remain
        // paragraph-addressed, so select the newly materialized cell paragraph.
        var cellParagraph = Assert.Single(Paragraphs(session).Except(body));
        var bytes = Png(29, 31);
        var footerInsert = session.InsertImage(footer, 0, bytes,
            new ImageInsertOptions { AltText = "footer image" });
        var cellInsert = session.InsertImage(cellParagraph, 4, bytes,
            new ImageInsertOptions { AltText = "cell image" });
        Assert.True(footerInsert.Success, footerInsert.Error?.Message);
        Assert.True(cellInsert.Success, cellInsert.Error?.Message);

        var images = session.ListImages();
        var footerImage = Assert.Single(images.Where(image => image.Id == footerInsert.ImageId));
        var cellImage = Assert.Single(images.Where(image => image.Id == cellInsert.ImageId));
        Assert.Equal("ftr1", footerImage.Scope);
        Assert.StartsWith("/word/footer", footerImage.OwningPartUri);
        Assert.Equal(new CharSpan(0, 0), footerImage.Span);
        Assert.Equal("body", cellImage.Scope);
        Assert.Equal("/word/document.xml", cellImage.OwningPartUri);
        Assert.Equal(cellParagraph, cellImage.AnchorId);
        Assert.Equal(new CharSpan(4, 0), cellImage.Span);

        var saved = session.Save(true);
        Assert.Single(ImagePartPayloads(saved));
        Assert.Equal(2, FlatImageRelationships(saved).Length);
        Assert.Single(FlatImageRelationships(saved).Select(value => value.TargetUri).Distinct());
        using var reopened = new DocxSession(saved);
        var reopenedImages = reopened.ListImages();
        Assert.Contains(reopenedImages, image => image.Id == footerInsert.ImageId
            && image.AltText == "footer image");
        Assert.Contains(reopenedImages, image => image.Id == cellInsert.ImageId
            && image.AltText == "cell image" && image.AnchorId == cellParagraph);
    }

    [Fact]
    public void IM016_AlternateContentDrawingAndVmlFallbackAreBothReadOnly()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.InsertImage(Paragraphs(seed)[0], 0, Png(2, 3)).Success);
        var compatible = MutatePackage(seed.Save(true), document =>
        {
            var main = document.MainDocumentPart!;
            var root = main.GetXDocument();
            var drawing = root.Descendants(W + "drawing").Single();
            var relationshipId = (string)drawing.Descendants(A + "blip").Single()
                .Attribute(R + "embed")!;
            drawing.ReplaceWith(new XElement(MC + "AlternateContent",
                new XElement(MC + "Choice", new XAttribute("Requires", "wp14"),
                    new XElement(drawing)),
                new XElement(MC + "Fallback",
                    new XElement(W + "pict",
                        new XElement(V + "shape",
                            new XElement(V + "imagedata",
                                new XAttribute(R + "id", relationshipId)))))));
            main.PutXDocument();
        });

        using var session = new DocxSession(compatible);
        var images = session.ListImages();
        Assert.Equal(2, images.Count);
        var modern = Assert.Single(images.Where(image => image.MarkupKind == ImageMarkupKind.ModernDrawing));
        var legacy = Assert.Single(images.Where(image => image.MarkupKind == ImageMarkupKind.LegacyVml));
        Assert.False(modern.CanMutate);
        Assert.False(legacy.CanMutate);
        Assert.Contains("AlternateContent", modern.UnsupportedReason);
        Assert.Equal(modern.RelationshipId, legacy.RelationshipId);
        Assert.Equal(EditErrorCode.UnsupportedImageMarkup,
            session.SetImageMetadata(modern.Id, "changed", null).Error!.Code);
    }

    [Fact]
    public void IM017_UpdateCommentSweepsImagesRemovedWithOldBody()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var added = session.AddComment(Paragraphs(session)[0], null, "Alice", "old body");
        Assert.True(added.Success, added.Error?.Message);
        var commentAnchor = Assert.Single(added.Created.Where(anchor => anchor.Kind == "cmt"));
        var paragraphAnchor = Assert.Single(added.Created.Where(anchor => anchor.Kind == "p"));
        Assert.True(session.InsertImage(paragraphAnchor.Id, 0, Png(3, 4)).Success);
        Assert.Single(session.ListImages(ProjectionScopes.Comments));

        var updated = session.UpdateComment(commentAnchor.Id, "replacement body");
        Assert.True(updated.Success, updated.Error?.Message);
        Assert.Empty(session.ListImages(ProjectionScopes.Comments));
        Assert.DoesNotContain(FlatImageRelationships(session.Save(true)),
            relationship => relationship.OwnerUri == "/word/comments.xml");
    }

    [Fact]
    public void IM018_SaveSweepsPreExistingOrphanImageRelationships()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.InsertImage(Paragraphs(seed)[0], 0, Png(2, 3)).Success);
        var orphaned = MutatePackage(seed.Save(true), document =>
        {
            var main = document.MainDocumentPart!;
            main.GetXDocument().Descendants(W + "drawing").Remove();
            main.PutXDocument();
        });
        Assert.Single(FlatImageRelationships(orphaned));

        using var session = new DocxSession(orphaned);
        Assert.Empty(session.ListImages());
        Assert.Empty(FlatImageRelationships(session.Save(true)));
        Assert.Empty(FlatImageRelationships(session.Save(false)));
    }

    [Fact]
    public void IM019_LiveRelationshipNamedByAnUnmodeledAttribute_SurvivesSaveAndRender()
    {
        // The negative direction of IM018. The sweep runs on EVERY Save — including the one
        // inside ConvertToHtml(session), i.e. on a pure RENDER — so a whitelist of known
        // reference attributes would silently and irrecoverably delete media named any other
        // way. o:relid is a real OLE/VML spelling and is outside {r:embed, r:link, r:id}.
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.InsertImage(Paragraphs(seed)[0], 0, Png(2, 3)).Success);
        var oleReferenced = MutatePackage(seed.Save(true), document =>
        {
            var main = document.MainDocumentPart!;
            var root = main.GetXDocument();
            var relationshipId = (string)root.Descendants(A + "blip").Single().Attribute(R + "embed")!;
            // Drop the only modeled reference and re-name the SAME relationship through o:relid.
            root.Descendants(W + "drawing").Remove();
            root.Descendants(W + "p").First().Add(new XElement(W + "r",
                new XElement(W + "object",
                    new XElement(V + "shape",
                        new XElement(O + "OLEObject", new XAttribute(O + "relid", relationshipId))))));
            main.PutXDocument();
        });
        var expected = Assert.Single(FlatImageRelationships(oleReferenced));

        using var session = new DocxSession(oleReferenced);
        Assert.Equal(expected, Assert.Single(FlatImageRelationships(session.Save(true))));
        Assert.Equal(expected, Assert.Single(FlatImageRelationships(session.Save(false))));

        // A render must not be a mutation of last resort either.
        _ = Docxodus.Internal.HtmlConversionOps.ConvertToHtml(session, new Docxodus.Internal.HtmlConversionOptions());
        Assert.Equal(expected, Assert.Single(FlatImageRelationships(session.Save(true))));
    }

    [Fact]
    public void IM020_FloatingInsertAndLayoutMutation_ProduceSchemaValidAnchors()
    {
        // The whole floating write path — the wp:anchor builder and ApplyFloatingLayout's three
        // ReplaceWith calls — had no passing-path coverage. CT_Anchor's child sequence is
        // simplePos, positionH, positionV, extent, effectExtent?, wrap-choice, docPr,
        // cNvGraphicFramePr?, graphic, and this repo has repeatedly shipped schema-order defects
        // that surface only as a Word "unreadable content" repair.
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var inserted = session.InsertImage(Paragraphs(session)[0], 3, Png(4, 5),
            new ImageInsertOptions
            {
                Placement = ImagePlacement.Floating,
                AltText = "floater",
                FloatingLayout = new FloatingImageLayout
                {
                    HorizontalRelativeFrom = ImageHorizontalReference.Page,
                    HorizontalOffsetEmu = 914400,
                    VerticalRelativeFrom = ImageVerticalReference.Margin,
                    VerticalOffsetEmu = -457200,
                    WrapMode = ImageWrapMode.Square,
                    WrapSide = ImageWrapSide.Left,
                    DistanceTopEmu = 45720,
                    BehindDocument = true,
                },
            });
        Assert.True(inserted.Success, inserted.Error?.Message);
        AssertSchemaValid(session.Save(true));
        AssertSchemaValid(session.Save(false));
        AssertAnchorChildOrder(session.Save(false));

        // The passing path of SetImageFloatingLayout: same op, different layout shape
        // (alignment instead of offset, wrapNone instead of wrapSquare).
        var replacement = new FloatingImageLayout
        {
            HorizontalRelativeFrom = ImageHorizontalReference.Margin,
            HorizontalOffsetEmu = null,
            HorizontalAlignment = ImageHorizontalAlignment.Right,
            VerticalRelativeFrom = ImageVerticalReference.Line,
            VerticalOffsetEmu = null,
            VerticalAlignment = ImageVerticalAlignment.Top,
            WrapMode = ImageWrapMode.None,
            DistanceLeftEmu = 91440,
            RelativeHeight = 42,
            LayoutInCell = false,
        };
        var image = Assert.Single(session.ListImages());
        Assert.True(image.CanMutate, image.UnsupportedReason);
        var applied = session.SetImageFloatingLayout(image.Id, replacement);
        Assert.True(applied.Success, applied.Error?.Message);

        image = Assert.Single(session.ListImages());
        Assert.True(image.FloatingLayoutSupported, image.UnsupportedReason);
        Assert.Equal(replacement, image.FloatingLayout);
        AssertSchemaValid(session.Save(true));
        AssertSchemaValid(session.Save(false));
        AssertAnchorChildOrder(session.Save(false));

        Assert.True(session.Undo());
        Assert.Equal(ImageWrapMode.Square, Assert.Single(session.ListImages()).FloatingLayout!.WrapMode);
    }

    [Fact]
    public void IM021_SvgBlipExtensionOccurrenceIsEnumeratedButReadOnly()
    {
        // An SVG picture stores its raster fallback in a:blip/@r:embed and the real art in an
        // asvg:svgBlip extension. Descendants(a:blip) still counts one blip, so without an
        // extension check ReplaceImage would swap only the fallback while reporting success.
        XNamespace asvg = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.InsertImage(Paragraphs(seed)[0], 0, Png(2, 3)).Success);
        var svg = MutatePackage(seed.Save(true), document =>
        {
            var main = document.MainDocumentPart!;
            var root = main.GetXDocument();
            var blip = root.Descendants(A + "blip").Single();
            var svgPart = main.AddImagePart("image/svg+xml", "rIdSvgArt");
            using (var input = new MemoryStream(
                System.Text.Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\"/>")))
                svgPart.FeedData(input);
            blip.Add(new XElement(A + "extLst",
                new XElement(A + "ext",
                    new XAttribute("uri", "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"),
                    new XElement(asvg + "svgBlip", new XAttribute(R + "embed", "rIdSvgArt")))));
            main.PutXDocument();
        });

        using var session = new DocxSession(svg);
        var image = Assert.Single(session.ListImages());
        Assert.Equal(ImageMarkupKind.ModernDrawing, image.MarkupKind);
        Assert.False(image.CanMutate);
        Assert.Contains("svgBlip", image.UnsupportedReason);
        Assert.Equal(EditErrorCode.UnsupportedImageMarkup,
            session.ReplaceImage(image.Id, Png(9, 9)).Error!.Code);
        Assert.Equal(EditErrorCode.UnsupportedImageMarkup,
            session.RemoveImage(image.Id).Error!.Code);
        Assert.Equal(EditErrorCode.UnsupportedImageMarkup,
            session.SetImageDimensions(image.Id, 36, null).Error!.Code);

        // Refusing must also leave both media parts intact through the normalizing save.
        Assert.Equal(2, FlatImageRelationships(session.Save(true)).Length);
    }

    [Fact]
    public void IM022_ReplaceImageIsDimensionPreserving_AndTheReFitRecipeWorks()
    {
        // ReplaceImage deliberately rewrites only r:embed: the rendered box is a layout decision
        // the caller made, not a property of the bytes. This pins BOTH halves of that contract —
        // the box survives, and the caller can re-fit because ListImages reports the NEW
        // intrinsic pixels and preserveAspect:false writes an exact box.
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(session.InsertImage(Paragraphs(session)[0], 0, Png(100, 100),
            new ImageInsertOptions { WidthPoints = 75 }).Success);
        var image = Assert.Single(session.ListImages());
        Assert.Equal(75, image.RenderedWidthPoints!.Value, 6);
        Assert.Equal(75, image.RenderedHeightPoints!.Value, 6);

        Assert.True(session.ReplaceImage(image.Id, Png(4000, 3000)).Success);
        image = Assert.Single(session.ListImages());
        Assert.Equal(75, image.RenderedWidthPoints!.Value, 6);
        Assert.Equal(75, image.RenderedHeightPoints!.Value, 6);
        // The recovery input: the new intrinsic ratio is readable immediately after the replace.
        Assert.Equal(4000, image.IntrinsicWidthPixels);
        Assert.Equal(3000, image.IntrinsicHeightPixels);

        double ratio = image.IntrinsicHeightPixels!.Value / (double)image.IntrinsicWidthPixels!.Value;
        Assert.True(session.SetImageDimensions(image.Id, 75, 75 * ratio, preserveAspect: false).Success);
        image = Assert.Single(session.ListImages());
        Assert.Equal(75, image.RenderedWidthPoints!.Value, 6);
        Assert.Equal(56.25, image.RenderedHeightPoints!.Value, 6);

        // preserveAspect keeps scaling from the CURRENT rendered box, which is now the new ratio.
        Assert.True(session.SetImageDimensions(image.Id, 150, null).Success);
        image = Assert.Single(session.ListImages());
        Assert.Equal(150, image.RenderedWidthPoints!.Value, 6);
        Assert.Equal(112.5, image.RenderedHeightPoints!.Value, 6);
    }

    [Fact]
    public void IM023_RenderInlineTrackedModeRejectsEveryImageMutation()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = Paragraphs(session)[0];
        Assert.True(session.InsertImage(anchor, 0, Png(2, 3)).Success);
        var image = Assert.Single(session.ListImages());

        session.SetTrackedChanges(TrackedChangeMode.RenderInline);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported,
            session.InsertImage(anchor, 0, Png(4, 5)).Error!.Code);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported,
            session.ReplaceImage(image.Id, Png(4, 5)).Error!.Code);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported,
            session.SetImageDimensions(image.Id, 36, null).Error!.Code);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported,
            session.SetImageMetadata(image.Id, "alt", null).Error!.Code);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported,
            session.SetImageFloatingLayout(image.Id, new FloatingImageLayout()).Error!.Code);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported,
            session.RemoveImage(image.Id).Error!.Code);

        // Rejection is not a mutation: listing and the document are unchanged.
        Assert.Equal(image, Assert.Single(session.ListImages()));

        // …and the ops come back once tracking is off.
        session.SetTrackedChanges(TrackedChangeMode.Accept);
        Assert.True(session.SetImageMetadata(image.Id, "alt", null).Success);
    }

    [Fact]
    public void IM024_FootnoteAndEndnoteStoriesOwnTheirImagesAndRoundTrip()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var body = Paragraphs(session);
        var footnote = session.InsertFootnote(body[0], 0, "see the mark");
        Assert.True(footnote.Success, footnote.Error?.Message);
        var endnote = session.InsertEndnote(body[1], 0, "and the other mark");
        Assert.True(endnote.Success, endnote.Error?.Message);

        var footnoteAnchor = Assert.Single(Paragraphs(session, "fn"));
        var endnoteAnchor = Assert.Single(Paragraphs(session, "en"));
        Assert.True(session.InsertImage(footnoteAnchor, 0, Png(6, 7)).Success);
        Assert.True(session.InsertImage(endnoteAnchor, 0, Png(8, 9)).Success);

        var footnoteImage = Assert.Single(session.ListImages(ProjectionScopes.Footnotes));
        var endnoteImage = Assert.Single(session.ListImages(ProjectionScopes.Endnotes));
        Assert.Equal("/word/footnotes.xml", footnoteImage.OwningPartUri);
        Assert.Equal("/word/endnotes.xml", endnoteImage.OwningPartUri);
        Assert.Equal(6, footnoteImage.IntrinsicWidthPixels);
        Assert.Equal(8, endnoteImage.IntrinsicWidthPixels);

        var saved = session.Save(true);
        AssertSchemaValid(saved);
        var owners = FlatImageRelationships(saved).Select(value => value.OwnerUri).ToArray();
        Assert.Contains("/word/footnotes.xml", owners);
        Assert.Contains("/word/endnotes.xml", owners);

        // Reopening must resolve both to their own story owners, not to the main part.
        using var reopened = new DocxSession(saved);
        Assert.Equal("/word/footnotes.xml",
            Assert.Single(reopened.ListImages(ProjectionScopes.Footnotes)).OwningPartUri);
        Assert.Equal("/word/endnotes.xml",
            Assert.Single(reopened.ListImages(ProjectionScopes.Endnotes)).OwningPartUri);

        Assert.True(session.RemoveImage(footnoteImage.Id).Success);
        Assert.DoesNotContain(FlatImageRelationships(session.Save(true)),
            relationship => relationship.OwnerUri == "/word/footnotes.xml");
        Assert.True(session.Undo());
        Assert.Equal(footnoteImage.TargetPartUri,
            Assert.Single(session.ListImages(ProjectionScopes.Footnotes)).TargetPartUri);
    }

    [Fact]
    public void IM025_RealDecodablePngSurvivesInsertSaveReopenByteIdentically()
    {
        // Every other fixture in this file is a synthetic header stub with no pixel data, so
        // nothing else here proves a genuinely decodable image survives the round trip.
        var real = File.ReadAllBytes(Path.Combine("../../../../TestFiles/", "img.png"));
        Assert.True(real.Length > 15000);

        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(session.InsertImage(Paragraphs(session)[0], 0, real).Success);
        var image = Assert.Single(session.ListImages());
        Assert.Equal(ImageBinaryFormat.Png, image.Format);
        Assert.Equal(180, image.IntrinsicWidthPixels);
        Assert.Equal(174, image.IntrinsicHeightPixels);
        Assert.True(image.ContentTypeMatchesBytes);
        Assert.False(image.IsBroken);

        var saved = session.Save(false);
        AssertSchemaValid(saved);
        var payload = Assert.Single(ImagePartPayloads(saved));
        Assert.Equal("image/png", payload.ContentType);
        Assert.Equal(Convert.ToBase64String(real), payload.Bytes);

        using var reopened = new DocxSession(saved);
        var reread = Assert.Single(reopened.ListImages());
        Assert.Equal(180, reread.IntrinsicWidthPixels);
        Assert.Equal(174, reread.IntrinsicHeightPixels);
        Assert.False(reread.IsBroken);
    }

    [Fact]
    public void IM026_NegativeFloatingOffsetsSerializeAsInvariantJsonUnderAnyCulture()
    {
        // horizontalOffsetEmu/verticalOffsetEmu are legitimately negative, and a culture whose
        // NegativeSign is not "-" would otherwise emit text JSON.parse rejects.
        var hostile = (System.Globalization.CultureInfo)
            System.Globalization.CultureInfo.InvariantCulture.Clone();
        hostile.NumberFormat.NegativeSign = "−";

        // Run on a DEDICATED thread rather than assigning CultureInfo.CurrentCulture on the
        // calling one. xUnit runs collections in parallel over a shared pool, and CurrentCulture
        // rides along on pool-thread reuse, so a bare assignment here would leak "−" into
        // unrelated tests and flake far away from this file.
        string json = string.Empty;
        Exception? failure = null;
        var worker = new System.Threading.Thread(() =>
        {
            try
            {
                Assert.Equal("−1", (-1L).ToString());   // the culture really is hostile
                using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
                Assert.True(session.InsertImage(Paragraphs(session)[0], 0, Png(4, 5),
                    new ImageInsertOptions
                    {
                        Placement = ImagePlacement.Floating,
                        FloatingLayout = new FloatingImageLayout
                        {
                            HorizontalOffsetEmu = -12345, VerticalOffsetEmu = -23456,
                        },
                    }).Success);
                json = Docxodus.Internal.DocxSessionJson.SerializeImages(session.ListImages());
            }
            catch (Exception ex) { failure = ex; }
        });
        worker.CurrentCulture = hostile;
        worker.CurrentUICulture = hostile;
        worker.Start();
        worker.Join();
        if (failure is not null) throw new Xunit.Sdk.XunitException(failure.ToString());

        Assert.DoesNotContain("−", json);
        using var parsed = JsonDocument.Parse(json);
        var layout = parsed.RootElement[0].GetProperty("floatingLayout");
        Assert.Equal(-12345, layout.GetProperty("horizontalOffsetEmu").GetInt64());
        Assert.Equal(-23456, layout.GetProperty("verticalOffsetEmu").GetInt64());
    }

    private static void AssertSchemaValid(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator().Validate(document).Where(IsRealValidationError));
    }

    /// <summary>CT_Anchor is a strict sequence. The validator catches order slips, but naming the
    /// expected sequence makes a failure say WHICH child moved.</summary>
    private static void AssertAnchorChildOrder(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var anchor = document.MainDocumentPart!.GetXDocument().Descendants(WP + "anchor").Single();
        var names = anchor.Elements().Select(element => element.Name.LocalName).ToArray();
        Assert.Equal(new[] { "simplePos", "positionH", "positionV", "extent", "effectExtent" },
            names.Take(5).ToArray());
        Assert.StartsWith("wrap", names[5], StringComparison.Ordinal);
        Assert.Equal(new[] { "docPr", "cNvGraphicFramePr", "graphic" }, names.Skip(6).ToArray());
        foreach (var required in new[] { "distT", "distB", "distL", "distR", "simplePos",
            "relativeHeight", "behindDoc", "locked", "layoutInCell", "allowOverlap" })
            Assert.NotNull(anchor.Attribute(required));
    }

    private static bool IsRealValidationError(ValidationErrorInfo error) =>
        !(error.Description ?? string.Empty).Contains("powertools.codeplex.com", StringComparison.Ordinal);

    private static List<(string OwnerUri, List<(string RelId, string TargetUri)> Relationships)>
        ImageRelationships(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var owners = StoryOwners(document);
        return owners.Select(owner => (owner.Uri.ToString(), owner.Parts
            .Where(pair => pair.OpenXmlPart is ImagePart)
            .Select(pair => (pair.RelationshipId, pair.OpenXmlPart.Uri.ToString())).ToList())).ToList();
    }

    private static (string OwnerUri, string RelId, string TargetUri)[] FlatImageRelationships(byte[] bytes) =>
        ImageRelationships(bytes).SelectMany(owner => owner.Relationships
            .Select(relationship => (owner.OwnerUri, relationship.RelId, relationship.TargetUri)))
            .OrderBy(value => value).ToArray();

    private static uint[] DocumentPropertyIds(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return StoryOwners(document).SelectMany(owner => owner.GetXDocument().Descendants(WP + "docPr"))
            .Select(element => (uint)element.Attribute("id")!).ToArray();
    }

    private static List<(string PartUri, string ContentType, string Bytes)> ImagePartPayloads(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return StoryOwners(document).SelectMany(owner => owner.Parts)
            .Where(pair => pair.OpenXmlPart is ImagePart)
            .Select(pair => (ImagePart)pair.OpenXmlPart)
            .GroupBy(part => part.Uri.ToString(), StringComparer.Ordinal)
            .Select(group => group.First())
            .OrderBy(part => part.Uri.ToString(), StringComparer.Ordinal)
            .Select(part =>
            {
                using var input = part.GetStream(FileMode.Open, FileAccess.Read);
                using var output = new MemoryStream();
                input.CopyTo(output);
                return (part.Uri.ToString(), part.ContentType, Convert.ToBase64String(output.ToArray()));
            }).ToList();
    }

    private static List<OpenXmlPart> StoryOwners(WordprocessingDocument document)
    {
        var main = document.MainDocumentPart!;
        var owners = new List<OpenXmlPart> { main };
        owners.AddRange(main.HeaderParts);
        owners.AddRange(main.FooterParts);
        if (main.FootnotesPart is not null) owners.Add(main.FootnotesPart);
        if (main.EndnotesPart is not null) owners.Add(main.EndnotesPart);
        if (main.WordprocessingCommentsPart is not null) owners.Add(main.WordprocessingCommentsPart);
        return owners;
    }

    private static byte[] MutatePackage(byte[] bytes, Action<WordprocessingDocument> mutate)
    {
        var stream = new MemoryStream(); stream.Write(bytes); stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true)) mutate(document);
        return stream.ToArray();
    }

    private static byte[] Png(int width, int height)
    {
        var bytes = new byte[24];
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0, 0, 0, 13, (byte)'I', (byte)'H', (byte)'D', (byte)'R' }.CopyTo(bytes, 0);
        WriteBig(bytes, 16, width); WriteBig(bytes, 20, height); return bytes;
    }

    private static byte[] Jpeg(int width, int height) => new byte[]
    {
        0xFF, 0xD8, 0xFF, 0xC0, 0, 17, 8,
        (byte)(height >> 8), (byte)height, (byte)(width >> 8), (byte)width,
        3, 1, 0x11, 0, 2, 0x11, 0, 3, 0x11, 0, 0xFF, 0xD9,
    };

    private static byte[] Gif(int width, int height) => new byte[]
    { (byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a',
      (byte)width, (byte)(width >> 8), (byte)height, (byte)(height >> 8) };

    private static byte[] Bmp(int width, int height)
    {
        var bytes = new byte[54]; bytes[0] = (byte)'B'; bytes[1] = (byte)'M';
        bytes[14] = 40; WriteLittle(bytes, 18, width); WriteLittle(bytes, 22, height); return bytes;
    }

    private static byte[] Tiff(int width, int height)
    {
        var bytes = new byte[38]; bytes[0] = (byte)'I'; bytes[1] = (byte)'I'; bytes[2] = 42;
        bytes[4] = 8; bytes[8] = 2;
        WriteTiffEntry(bytes, 10, 256, width); WriteTiffEntry(bytes, 22, 257, height); return bytes;
    }

    private static void WriteTiffEntry(byte[] bytes, int offset, int tag, int value)
    { bytes[offset] = (byte)tag; bytes[offset + 1] = (byte)(tag >> 8); bytes[offset + 2] = 4;
      bytes[offset + 4] = 1; WriteLittle(bytes, offset + 8, value); }
    private static void WriteBig(byte[] bytes, int offset, int value)
    { bytes[offset] = (byte)(value >> 24); bytes[offset + 1] = (byte)(value >> 16);
      bytes[offset + 2] = (byte)(value >> 8); bytes[offset + 3] = (byte)value; }
    private static void WriteLittle(byte[] bytes, int offset, int value)
    { bytes[offset] = (byte)value; bytes[offset + 1] = (byte)(value >> 8);
      bytes[offset + 2] = (byte)(value >> 16); bytes[offset + 3] = (byte)(value >> 24); }
}
