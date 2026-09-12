// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Issue #762: the per-occurrence operation matrix, WebP writes, the remaining wrap
/// forms, linked/VML/SVG/AlternateContent coverage, the explicit embed-linked conversion, and
/// native tracked image mutations.</summary>
public class DocxSessionImageCoverageTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";
    private static readonly XNamespace MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XNamespace ASVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main";
    private static readonly XNamespace PT = "http://powertools.codeplex.com/2011";

    [Fact]
    public void IM762a_CapabilitiesAndOccurrencesExposeTheOperationMatrix()
    {
        var capabilities = DocxSession.GetImageCapabilities();
        Assert.Equal(2, capabilities.SchemaVersion);
        var webp = Assert.Single(capabilities.Formats, format => format.Format == ImageBinaryFormat.Webp);
        Assert.True(webp.CanInsert);
        Assert.True(webp.CanReplace);
        Assert.Contains("embed_linked", capabilities.Operations);
        Assert.Equal(new[] { ImageWrapMode.None, ImageWrapMode.Square, ImageWrapMode.Tight,
            ImageWrapMode.Through, ImageWrapMode.TopAndBottom }, capabilities.MutableWrapModes);
        Assert.Equal(new[] { "insert", "replace", "embed_linked", "set_dimensions", "set_metadata",
            "set_floating_layout", "remove" }, capabilities.TrackedOperations);
        var linked = Assert.Single(capabilities.Markups, markup => markup.Markup == "linked_picture");
        Assert.Contains("embed_linked", linked.Operations);
        Assert.DoesNotContain("replace", linked.Operations);
        Assert.Empty(Assert.Single(capabilities.Markups, markup => markup.Markup == "multi_picture").Operations);

        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(session.InsertImage(Paragraphs(session)[0], 0, Png(2, 3)).Success);
        var image = Assert.Single(session.ListImages());
        Assert.True(image.CanMutate);
        Assert.All(image.Operations.All.Where(support => support.Operation != "embed_linked"
                && support.Operation != "set_floating_layout"),
            support => Assert.True(support.CanMutate, support.Reason));
        Assert.Equal("picture is already embedded", image.Operations.EmbedLinked.Reason);
        Assert.Equal("inline picture has no floating layout", image.Operations.SetFloatingLayout.Reason);
        Assert.Equal(EditErrorCode.UnsupportedImageMarkup, session.EmbedLinkedImage(image.Id, Png(4, 4)).Error!.Code);

        // The wire shape carries the matrix and the polygon, and the polygon parses back.
        using var json = JsonDocument.Parse(DocxSessionJson.SerializeImages(session.ListImages()));
        var operations = json.RootElement[0].GetProperty("operations").EnumerateArray().ToList();
        Assert.Equal(new[] { "replace", "embed_linked", "set_dimensions", "set_metadata", "set_floating_layout", "remove" },
            operations.Select(entry => entry.GetProperty("operation").GetString()));
        Assert.False(operations[1].GetProperty("canMutate").GetBoolean());
        Assert.Equal("picture is already embedded", operations[1].GetProperty("reason").GetString());
        using var capabilitiesJson = JsonDocument.Parse(DocxSessionJson.SerializeImageCapabilities(capabilities));
        Assert.Equal(6, capabilitiesJson.RootElement.GetProperty("markups").GetArrayLength());
        Assert.Equal(7, capabilitiesJson.RootElement.GetProperty("trackedOperations").GetArrayLength());
        var parsed = DocxSessionJson.ParseFloatingImageLayout(
            """{"wrapMode":"through","wrapPolygon":{"points":[{"x":0,"y":0},{"x":21600,"y":0},{"x":10800,"y":21600}],"edited":true}}""");
        Assert.Equal(new ImageWrapPolygon(new[] { new ImageWrapPoint(0, 0), new ImageWrapPoint(21600, 0),
            new ImageWrapPoint(10800, 21600) }, Edited: true), parsed.WrapPolygon);
    }

    [Fact]
    public void IM762b_WebPIsWrittenAsAProperMediaPart()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = Paragraphs(session)[0];
        Assert.True(session.InsertImage(anchor, 0, Webp(16, 8)).Success);
        var image = Assert.Single(session.ListImages());
        Assert.Equal(ImageBinaryFormat.Webp, image.Format);
        Assert.Equal("image/webp", image.ContentType);
        Assert.True(image.ContentTypeMatchesBytes);
        Assert.Equal(16, image.IntrinsicWidthPixels);
        Assert.Equal(8, image.IntrinsicHeightPixels);
        Assert.EndsWith(".webp", image.MediaFileName, StringComparison.Ordinal);
        Assert.Equal(12, image.RenderedWidthPoints);

        // A PNG picture can become WebP and back; each swap lands on its own correctly typed part.
        Assert.True(session.InsertImage(Paragraphs(session)[1], 0, Png(2, 3)).Success);
        var png = Assert.Single(session.ListImages(), value => value.Format == ImageBinaryFormat.Png);
        Assert.True(session.ReplaceImage(png.Id, Webp(4, 4)).Success);
        var saved = session.Save(true);
        AssertSchemaValid(saved);
        using var reopened = new DocxSession(saved);
        Assert.All(reopened.ListImages(), value =>
        {
            Assert.Equal(ImageBinaryFormat.Webp, value.Format);
            Assert.Equal("image/webp", value.ContentType);
            Assert.True(value.CanMutate, value.UnsupportedReason);
        });
        using var document = WordprocessingDocument.Open(new MemoryStream(saved), false);
        Assert.All(document.MainDocumentPart!.ImageParts, part =>
        {
            Assert.Equal("image/webp", part.ContentType);
            Assert.EndsWith(".webp", part.Uri.OriginalString, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void IM762c_TightThroughAndTopAndBottomWrapRoundTrip()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var tight = new FloatingImageLayout { WrapMode = ImageWrapMode.Tight, WrapSide = ImageWrapSide.Left };
        Assert.True(session.InsertImage(Paragraphs(session)[0], 0, Png(4, 5),
            new ImageInsertOptions { Placement = ImagePlacement.Floating, FloatingLayout = tight }).Success);
        var image = Assert.Single(session.ListImages());
        Assert.True(image.FloatingLayoutSupported, image.UnsupportedReason);
        Assert.True(image.CanMutate, image.UnsupportedReason);
        // Without vertices the session writes the picture rectangle and reads it back as such.
        Assert.Equal(tight with { WrapPolygon = ImageWrapPolygon.Rectangle }, image.FloatingLayout);
        AssertSchemaValid(session.Save(true));

        // Writing the same layout again is a no-op even though the caller left the polygon out.
        int undo = session.UndoCount;
        Assert.True(session.SetImageFloatingLayout(image.Id, tight).Success);
        Assert.Equal(undo, session.UndoCount);

        var outline = new ImageWrapPolygon(new[] { new ImageWrapPoint(0, 10800), new ImageWrapPoint(10800, 0),
            new ImageWrapPoint(21600, 10800), new ImageWrapPoint(10800, 21600), new ImageWrapPoint(0, 10800) }, Edited: true);
        var through = new FloatingImageLayout { WrapMode = ImageWrapMode.Through, WrapPolygon = outline,
            HorizontalAlignment = ImageHorizontalAlignment.Center, HorizontalOffsetEmu = null };
        Assert.True(session.SetImageFloatingLayout(image.Id, through).Success);
        image = Assert.Single(session.ListImages());
        Assert.Equal(through, image.FloatingLayout);
        var saved = session.Save(true);
        AssertSchemaValid(saved);
        using (var document = WordprocessingDocument.Open(new MemoryStream(saved), false))
        {
            var wrap = document.MainDocumentPart!.GetXDocument().Descendants(WP + "wrapThrough").Single();
            Assert.Equal("1", (string?)wrap.Element(WP + "wrapPolygon")!.Attribute("edited"));
            Assert.Equal(5, wrap.Element(WP + "wrapPolygon")!.Elements().Count());
            Assert.Equal(WP + "start", wrap.Element(WP + "wrapPolygon")!.Elements().First().Name);
        }

        var topAndBottom = new FloatingImageLayout { WrapMode = ImageWrapMode.TopAndBottom };
        Assert.True(session.SetImageFloatingLayout(image.Id, topAndBottom).Success);
        image = Assert.Single(session.ListImages());
        Assert.Equal(topAndBottom, image.FloatingLayout);
        AssertSchemaValid(session.Save(true));

        Assert.Equal(EditErrorCode.InvalidImageLayout, session.SetImageFloatingLayout(image.Id,
            new FloatingImageLayout { WrapMode = ImageWrapMode.Square, WrapPolygon = outline }).Error!.Code);
        Assert.Equal(EditErrorCode.InvalidImageLayout, session.SetImageFloatingLayout(image.Id,
            new FloatingImageLayout { WrapMode = ImageWrapMode.Tight,
                WrapPolygon = new ImageWrapPolygon(new[] { new ImageWrapPoint(0, 0), new ImageWrapPoint(1, 1) }) }).Error!.Code);

        // A Word-written tight wrap with distances on the wrap element itself stays report-only.
        var unmodeled = MutatePackage(session.Save(true), document =>
        {
            var main = document.MainDocumentPart!;
            main.GetXDocument().Descendants(WP + "wrapTopAndBottom").Single().SetAttributeValue("distT", "100");
            main.PutXDocument();
        });
        using var probe = new DocxSession(unmodeled);
        var reported = Assert.Single(probe.ListImages());
        Assert.False(reported.FloatingLayoutSupported);
        Assert.False(reported.Operations.SetFloatingLayout.CanMutate);
        Assert.True(reported.Operations.Replace.CanMutate);
        Assert.Contains("distT", reported.FloatingLayout!.RawWrapMode);
    }

    [Fact]
    public void IM762d_LinkedPicture_TakesEmbedLinkedInsteadOfReplace()
    {
        const string relationshipId = "rIdExternalImage37";
        const string target = "https://example.test/image.png";
        using var session = new DocxSession(LinkedFixture(relationshipId, target));
        var image = Assert.Single(session.ListImages());
        Assert.True(image.IsLinked);
        Assert.False(image.CanMutate);
        Assert.Contains("embed_linked", image.UnsupportedReason);
        Assert.False(image.Operations.Replace.CanMutate);
        Assert.True(image.Operations.EmbedLinked.CanMutate);
        Assert.True(image.Operations.SetMetadata.CanMutate);
        Assert.True(image.Operations.SetDimensions.CanMutate);
        Assert.True(image.Operations.Remove.CanMutate);
        Assert.Equal(EditErrorCode.LinkedImageReadOnly, session.ReplaceImage(image.Id, Png(4, 5)).Error!.Code);

        // Metadata and geometry never touched the media, so they were always safe to edit.
        Assert.True(session.SetImageMetadata(image.Id, "described", null).Success);
        Assert.True(session.SetImageDimensions(image.Id, 36, null).Success);
        image = Assert.Single(session.ListImages());
        Assert.True(image.IsLinked);
        Assert.Equal("described", image.AltText);
        Assert.Equal(36, image.RenderedWidthPoints);

        var embedded = session.EmbedLinkedImage(image.Id, Png(6, 7));
        Assert.True(embedded.Success, embedded.Error?.Message);
        image = Assert.Single(session.ListImages());
        Assert.False(image.IsLinked);
        Assert.True(image.IsEmbedded);
        Assert.True(image.CanMutate, image.UnsupportedReason);
        Assert.Equal(6, image.IntrinsicWidthPixels);
        Assert.Equal("described", image.AltText);
        Assert.Equal(36, image.RenderedWidthPoints);
        var saved = session.Save(true);
        AssertSchemaValid(saved);
        using (var document = WordprocessingDocument.Open(new MemoryStream(saved), false))
        {
            var main = document.MainDocumentPart!;
            Assert.Empty(main.ExternalRelationships);
            var blip = main.GetXDocument().Descendants(A + "blip").Single();
            Assert.Null(blip.Attribute(R + "link"));
            Assert.Equal(Assert.Single(main.ImageParts).Uri, main.GetPartById((string)blip.Attribute(R + "embed")!).Uri);
        }

        // Undo restores the external relationship exactly and drops the part the embed created.
        Assert.True(session.Undo());
        image = Assert.Single(session.ListImages());
        Assert.Equal(relationshipId, image.LinkedRelationshipId);
        Assert.Equal(target, image.LinkedTarget);
        using (var document = WordprocessingDocument.Open(new MemoryStream(session.Save(true)), false))
        {
            Assert.Empty(document.MainDocumentPart!.ImageParts);
            Assert.Equal(target, Assert.Single(document.MainDocumentPart.ExternalRelationships).Uri.ToString());
        }
        Assert.True(session.Redo());
        Assert.False(Assert.Single(session.ListImages()).IsLinked);
    }

    [Fact]
    public void IM762e_LegacyVml_ReplaceResizeDescribeRemove()
    {
        using var session = new DocxSession(VmlFixture("position:absolute;width:36pt;height:27pt;z-index:1"));
        var images = session.ListImages();
        var modern = Assert.Single(images, value => value.MarkupKind == ImageMarkupKind.ModernDrawing);
        var legacy = Assert.Single(images, value => value.MarkupKind == ImageMarkupKind.LegacyVml);
        Assert.True(legacy.CanMutate, legacy.UnsupportedReason);
        Assert.Equal(36, legacy.RenderedWidthPoints);
        Assert.Equal(27, legacy.RenderedHeightPoints);
        Assert.False(legacy.Operations.SetFloatingLayout.CanMutate);
        Assert.Contains("VML", legacy.Operations.SetFloatingLayout.Reason);
        Assert.Equal(modern.RelationshipId, legacy.RelationshipId);

        // Replacing the VML occurrence re-points only its own reference; the modern picture keeps the shared part.
        Assert.True(session.ReplaceImage(legacy.Id, Png(9, 9)).Success);
        images = session.ListImages();
        legacy = Assert.Single(images, value => value.MarkupKind == ImageMarkupKind.LegacyVml);
        modern = Assert.Single(images, value => value.MarkupKind == ImageMarkupKind.ModernDrawing);
        Assert.Equal(9, legacy.IntrinsicWidthPixels);
        Assert.Equal(2, modern.IntrinsicWidthPixels);
        Assert.NotEqual(modern.RelationshipId, legacy.RelationshipId);

        Assert.True(session.SetImageDimensions(legacy.Id, 72, null).Success);
        Assert.True(session.SetImageMetadata(legacy.Id, "vml alt", "vml title").Success);
        legacy = Assert.Single(session.ListImages(), value => value.MarkupKind == ImageMarkupKind.LegacyVml);
        Assert.Equal(72, legacy.RenderedWidthPoints);
        Assert.Equal(54, legacy.RenderedHeightPoints);
        Assert.Equal("vml alt", legacy.AltText);
        Assert.Equal("vml title", legacy.Title);
        Assert.Equal(EditErrorCode.UnsupportedImageMarkup,
            session.SetImageFloatingLayout(legacy.Id, new FloatingImageLayout()).Error!.Code);
        var saved = session.Save(true);
        AssertSchemaValid(saved);
        using (var document = WordprocessingDocument.Open(new MemoryStream(saved), false))
        {
            var shape = document.MainDocumentPart!.GetXDocument().Descendants(V + "shape").Single();
            // Untouched style entries keep their text and order around the rewritten size.
            Assert.Equal("position:absolute;width:72pt;height:54pt;z-index:1", (string?)shape.Attribute("style"));
            Assert.Equal("vml title", (string?)shape.Element(V + "imagedata")!.Attribute(O + "title"));
            Assert.Equal(2, document.MainDocumentPart!.ImageParts.Count());
        }

        Assert.True(session.RemoveImage(legacy.Id).Success);
        modern = Assert.Single(session.ListImages());
        Assert.Equal(ImageMarkupKind.ModernDrawing, modern.MarkupKind);
        using (var document = WordprocessingDocument.Open(new MemoryStream(session.Save(true)), false))
        {
            Assert.Empty(document.MainDocumentPart!.GetXDocument().Descendants(W + "pict"));
            Assert.Single(document.MainDocumentPart!.ImageParts);
        }
        Assert.True(session.Undo());
        Assert.Equal(2, session.ListImages().Count);
    }

    [Fact]
    public void IM762f_ExtendedPicture_RefusesReplaceButRemovesBothPayloads()
    {
        using var session = new DocxSession(SvgFixture());
        var image = Assert.Single(session.ListImages());
        Assert.False(image.CanMutate);
        Assert.Contains("svgBlip", image.UnsupportedReason);
        Assert.False(image.Operations.Replace.CanMutate);
        Assert.False(image.Operations.EmbedLinked.CanMutate);
        Assert.True(image.Operations.SetDimensions.CanMutate);
        Assert.True(image.Operations.SetMetadata.CanMutate);
        Assert.True(image.Operations.Remove.CanMutate);
        Assert.Equal(EditErrorCode.UnsupportedImageMarkup, session.ReplaceImage(image.Id, Png(9, 9)).Error!.Code);
        Assert.True(session.SetImageDimensions(image.Id, 36, null).Success);
        Assert.Equal(36, Assert.Single(session.ListImages()).RenderedWidthPoints);
        Assert.Equal(2, FlatImageRelationships(session.Save(true)).Length);

        // Removing the drawing orphans the raster fallback AND the vector part; both are swept.
        Assert.True(session.RemoveImage(image.Id).Success);
        Assert.Empty(session.ListImages());
        Assert.Empty(FlatImageRelationships(session.Save(true)));
        Assert.True(session.Undo());
        Assert.Equal(2, FlatImageRelationships(session.Save(true)).Length);
        Assert.Contains("svgBlip", Assert.Single(session.ListImages()).UnsupportedReason);
    }

    [Fact]
    public void IM762g_AlternateContent_ChangesEveryBranchTogether()
    {
        using var session = new DocxSession(AlternateContentFixture());
        var images = session.ListImages();
        var modern = Assert.Single(images, value => value.MarkupKind == ImageMarkupKind.ModernDrawing);
        var fallback = Assert.Single(images, value => value.MarkupKind == ImageMarkupKind.LegacyVml);
        Assert.True(modern.CanMutate, modern.UnsupportedReason);
        Assert.False(modern.Operations.SetFloatingLayout.CanMutate);
        Assert.False(fallback.CanMutate);
        Assert.Contains("modern occurrence", fallback.UnsupportedReason);
        Assert.All(fallback.Operations.All, support => Assert.False(support.CanMutate));

        Assert.True(session.SetImageMetadata(modern.Id, "shared alt", "shared title").Success);
        Assert.True(session.SetImageDimensions(modern.Id, 72, null).Success);
        Assert.True(session.ReplaceImage(modern.Id, Png(9, 9)).Success);
        images = session.ListImages();
        modern = Assert.Single(images, value => value.MarkupKind == ImageMarkupKind.ModernDrawing);
        fallback = Assert.Single(images, value => value.MarkupKind == ImageMarkupKind.LegacyVml);
        Assert.Equal(modern.RelationshipId, fallback.RelationshipId);
        Assert.Equal(9, fallback.IntrinsicWidthPixels);
        Assert.Equal(("shared alt", "shared title"), (fallback.AltText, fallback.Title));
        Assert.Equal(("shared alt", "shared title"), (modern.AltText, modern.Title));
        Assert.Equal(72, modern.RenderedWidthPoints);
        Assert.Equal(72, fallback.RenderedWidthPoints);
        var saved = session.Save(true);
        AssertSchemaValid(saved);
        Assert.Single(FlatImageRelationships(saved));

        Assert.True(session.RemoveImage(modern.Id).Success);
        Assert.Empty(session.ListImages());
        using (var document = WordprocessingDocument.Open(new MemoryStream(session.Save(true)), false))
        {
            Assert.Empty(document.MainDocumentPart!.GetXDocument().Descendants(MC + "AlternateContent"));
            Assert.Empty(document.MainDocumentPart!.ImageParts);
        }
        Assert.True(session.Undo());
        Assert.Equal(2, session.ListImages().Count);
    }

    [Fact]
    public void IM762h_TrackedMode_MutationsAreDeletedPlusInsertedRuns()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = Paragraphs(seed)[0];
        Assert.True(seed.InsertImage(anchor, 0, Png(2, 3), new ImageInsertOptions { AltText = "before" }).Success);
        var baseline = seed.Save(true);

        using var session = new DocxSession(baseline, new DocxSessionSettings
        { TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = "Reviewer" });
        var original = Assert.Single(session.ListImages());
        Assert.True(original.CanMutate, original.UnsupportedReason);
        var replaced = session.ReplaceImage(original.Id, Png(6, 7));
        Assert.True(replaced.Success, replaced.Error?.Message);
        Assert.NotEqual(original.Id, replaced.ImageId);

        // Both runs are listed: the deletion is inspectable but closed, the insertion is live.
        var images = session.ListImages();
        Assert.Equal(2, images.Count);
        var deleted = Assert.Single(images, value => value.Id == original.Id);
        var inserted = Assert.Single(images, value => value.Id == replaced.ImageId);
        Assert.False(deleted.CanMutate);
        Assert.Contains("tracked deletion", deleted.UnsupportedReason);
        Assert.Equal(2, deleted.IntrinsicWidthPixels);
        Assert.True(inserted.CanMutate, inserted.UnsupportedReason);
        Assert.Equal(6, inserted.IntrinsicWidthPixels);
        Assert.Equal("before", inserted.AltText);
        Assert.Equal(2, FlatImageRelationships(session.Save(true)).Length);
        var types = session.ListRevisions().Select(revision => revision.Type).ToList();
        Assert.Contains("insert", types);
        Assert.Contains("delete", types);

        // The re-fit recipe: the picture is now the author's own insertion, so the next edits are
        // in place and add no second revision pair.
        Assert.True(session.SetImageDimensions(inserted.Id, 18, 21, preserveAspect: false).Success);
        Assert.True(session.SetImageMetadata(inserted.Id, "after", null).Success);
        images = session.ListImages();
        Assert.Equal(2, images.Count);
        inserted = Assert.Single(images, value => value.Id == inserted.Id);
        Assert.Equal((18d, 21d, "after"), (inserted.RenderedWidthPoints, inserted.RenderedHeightPoints, inserted.AltText));
        var saved = session.Save(true);
        AssertSchemaValid(saved);
        using (var document = WordprocessingDocument.Open(new MemoryStream(saved), false))
        {
            var body = document.MainDocumentPart!.GetXDocument().Root!;
            Assert.Single(body.Descendants(W + "del"));
            Assert.Single(body.Descendants(W + "ins"));
            Assert.Single(body.Descendants(W + "del").Single().Descendants(W + "drawing"));
            Assert.Single(body.Descendants(W + "ins").Single().Descendants(W + "drawing"));
            Assert.Equal("Reviewer", (string?)body.Descendants(W + "ins").Single().Attribute(W + "author"));
            var ids = body.Descendants(WP + "docPr").Select(element => (string?)element.Attribute("id")).ToList();
            Assert.Equal(ids.Distinct().Count(), ids.Count);
        }

        // Rejecting restores the original bytes, relationship, metadata and geometry.
        using (var rejecting = new DocxSession(saved))
        {
            Assert.True(rejecting.RejectAllRevisions().Success);
            var restored = Assert.Single(rejecting.ListImages());
            Assert.Equal(original.RelationshipId, restored.RelationshipId);
            Assert.Equal(original.TargetPartUri, restored.TargetPartUri);
            Assert.Equal((2, "before", original.RenderedWidthPoints),
                (restored.IntrinsicWidthPixels, restored.AltText, restored.RenderedWidthPoints));
            var relationship = Assert.Single(FlatImageRelationships(rejecting.Save(true)));
            Assert.Equal(original.TargetPartUri, relationship.TargetUri);
        }
        // Accepting yields the new picture and sweeps the media only the deleted run referenced.
        using (var accepting = new DocxSession(saved))
        {
            Assert.True(accepting.AcceptAllRevisions().Success);
            var accepted = Assert.Single(accepting.ListImages());
            Assert.Equal((6, "after", 18d), (accepted.IntrinsicWidthPixels, accepted.AltText, accepted.RenderedWidthPoints));
            Assert.True(accepted.CanMutate, accepted.UnsupportedReason);
            Assert.Single(FlatImageRelationships(accepting.Save(true)));
        }

        // Undo walks the whole tracked edit back and redo replays it, topology included.
        Assert.True(session.Undo());
        Assert.True(session.Undo());
        Assert.True(session.Undo());
        var undone = Assert.Single(session.ListImages());
        Assert.Equal(original, undone);
        Assert.Single(FlatImageRelationships(session.Save(true)));
        Assert.True(session.Redo());
        Assert.Equal(2, session.ListImages().Count);
        Assert.Equal(2, FlatImageRelationships(session.Save(true)).Length);
    }

    [Fact]
    public void IM762i_TrackedMode_RemoveAndInsertAreSingleRevisions()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = Paragraphs(seed)[0];
        Assert.True(seed.InsertImage(anchor, 0, Png(2, 3)).Success);
        using var session = new DocxSession(seed.Save(true), new DocxSessionSettings
        { TrackedChanges = TrackedChangeMode.RenderInline });
        var image = Assert.Single(session.ListImages());
        Assert.True(session.RemoveImage(image.Id).Success);
        var deleted = Assert.Single(session.ListImages());
        Assert.False(deleted.CanMutate);
        Assert.Contains("tracked deletion", deleted.UnsupportedReason);
        // The part stays until the deletion is accepted; rejecting brings the picture back untouched.
        Assert.Single(FlatImageRelationships(session.Save(true)));
        Assert.Equal("delete", Assert.Single(session.ListRevisions()).Type);

        // Two tracked inserts into a paragraph that already carries revisions both land, wrapped.
        Assert.True(session.InsertImage(anchor, 0, Png(4, 4)).Success);
        var second = session.InsertImage(anchor, 0, Png(5, 5));
        Assert.True(second.Success, second.Error?.Message);
        Assert.Equal(3, session.ListImages().Count);
        Assert.Equal(2, session.ListImages().Count(value => value.CanMutate));
        var saved = session.Save(true);
        AssertSchemaValid(saved);
        using (var document = WordprocessingDocument.Open(new MemoryStream(saved), false))
        {
            var body = document.MainDocumentPart!.GetXDocument().Root!;
            Assert.Equal(2, body.Descendants(W + "ins").Count());
            Assert.All(body.Descendants(W + "ins"), ins => Assert.Single(ins.Descendants(W + "drawing")));
        }
        using var rejecting = new DocxSession(saved);
        Assert.True(rejecting.RejectAllRevisions().Success);
        var restored = Assert.Single(rejecting.ListImages());
        Assert.Equal(2, restored.IntrinsicWidthPixels);
        Assert.Single(FlatImageRelationships(rejecting.Save(true)));
        using var accepting = new DocxSession(saved);
        Assert.True(accepting.AcceptAllRevisions().Success);
        // Offset 0 lands each new picture after the ones already at that boundary, as text does.
        Assert.Equal(new int?[] { 4, 5 }, accepting.ListImages().Select(value => value.IntrinsicWidthPixels).ToArray());
        Assert.Equal(2, FlatImageRelationships(accepting.Save(true)).Length);
    }

    [Fact]
    public void IM762j_TrackedMode_IsolatesThePictureAndRespectsForeignInsertions()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.InsertImage(Paragraphs(seed)[0], 2, Png(2, 3)).Success);
        // Fold the picture run into its text neighbours so one run carries text + picture + text,
        // and wrap the second paragraph's picture in another author's insertion.
        var fixture = MutatePackage(seed.Save(true), document =>
        {
            var main = document.MainDocumentPart!;
            var root = main.GetXDocument();
            var paragraph = root.Descendants(W + "p").First();
            var runs = paragraph.Elements(W + "r").ToList();
            var pictureRun = runs.Single(run => run.Element(W + "drawing") is not null);
            var drawingCopy = new XElement(pictureRun.Element(W + "drawing")!);
            // A copy must be its own occurrence: drop the persisted anchor bookkeeping it inherited.
            foreach (var element in drawingCopy.DescendantsAndSelf())
                element.Attributes().Where(attribute => attribute.Name.Namespace == PT).Remove();
            foreach (var properties in drawingCopy.Descendants().Where(element =>
                element.Name == WP + "docPr" || element.Name.LocalName == "cNvPr"))
                properties.SetAttributeValue("id", "2");
            var merged = new XElement(W + "r", new XElement(W + "rPr", new XElement(W + "b")));
            foreach (var run in runs)
            {
                foreach (var node in run.Elements().Where(element => element.Name != W + "rPr").ToList())
                {
                    node.Remove();
                    merged.Add(node);
                }
            }
            runs[0].AddBeforeSelf(merged);
            foreach (var run in runs) run.Remove();
            var second = root.Descendants(W + "p").Skip(1).First();
            var foreign = new XElement(W + "ins", new XAttribute(W + "id", "900"),
                new XAttribute(W + "author", "Someone Else"), new XAttribute(W + "date", "2024-01-01T00:00:00Z"),
                new XElement(W + "r", drawingCopy));
            second.AddFirst(foreign);
            main.PutXDocument();
        });

        using var session = new DocxSession(fixture, new DocxSessionSettings
        { TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = "Reviewer" });
        var images = session.ListImages();
        Assert.Equal(2, images.Count);
        var foreignImage = Assert.Single(images, value => value.AnchorId == Paragraphs(session)[1]);
        Assert.False(foreignImage.CanMutate);
        Assert.Contains("another author", foreignImage.UnsupportedReason);
        Assert.Equal(EditErrorCode.UnsupportedImageMarkup, session.RemoveImage(foreignImage.Id).Error!.Code);

        var mine = Assert.Single(images, value => value.AnchorId == Paragraphs(session)[0]);
        Assert.True(session.RemoveImage(mine.Id).Success);
        var saved = session.Save(true);
        AssertSchemaValid(saved);
        using var document = WordprocessingDocument.Open(new MemoryStream(saved), false);
        var first = document.MainDocumentPart!.GetXDocument().Descendants(W + "p").First();
        var children = first.Elements().Where(element => element.Name != W + "pPr").Select(element => element.Name.LocalName).ToList();
        Assert.Equal(new[] { "r", "del", "r" }, children);
        Assert.All(first.Elements(W + "r"), run => Assert.NotNull(run.Element(W + "rPr")?.Element(W + "b")));
        Assert.Equal(Paragraphs(session).Length, document.MainDocumentPart.GetXDocument().Descendants(W + "p").Count());
        Assert.Equal("Someone Else", (string?)document.MainDocumentPart.GetXDocument()
            .Descendants(W + "p").Skip(1).First().Element(W + "ins")!.Attribute(W + "author"));
        // Switching tracking off makes the foreign insertion an ordinary editable picture.
        session.SetTrackedChanges(TrackedChangeMode.Accept);
        Assert.True(Assert.Single(session.ListImages(), value => value.Id == foreignImage.Id).CanMutate);
    }

    [Fact]
    public void IM762k_TrackedMode_FooterStoryOwnsItsSwappedMedia()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.SetFooterText(Paragraphs(seed)[0], HeaderFooterKind.Default, "footer").Success);
        var footer = Assert.Single(Paragraphs(seed, "ftr1"));
        Assert.True(seed.InsertImage(footer, 0, Png(2, 3)).Success);
        using var session = new DocxSession(seed.Save(true), new DocxSessionSettings
        { TrackedChanges = TrackedChangeMode.RenderInline });
        var image = Assert.Single(session.ListImages());
        Assert.Equal("ftr1", image.Scope);
        var replaced = session.ReplaceImage(image.Id, Png(6, 7));
        Assert.True(replaced.Success, replaced.Error?.Message);
        var saved = session.Save(true);
        AssertSchemaValid(saved);
        // Both media parts hang off the footer part, never the main document.
        var relationships = FlatImageRelationships(saved);
        Assert.Equal(2, relationships.Length);
        Assert.All(relationships, value => Assert.StartsWith("/word/footer", value.OwnerUri, StringComparison.Ordinal));
        using var reopened = new DocxSession(saved, new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var listed = reopened.ListImages();
        Assert.Equal(2, listed.Count);
        Assert.All(listed, value => Assert.Equal("ftr1", value.Scope));
        Assert.Equal(6, Assert.Single(listed, value => value.Id == replaced.ImageId).IntrinsicWidthPixels);
        Assert.True(reopened.AcceptAllRevisions().Success);
        Assert.Single(FlatImageRelationships(reopened.Save(true)));
        Assert.True(reopened.Undo());
        Assert.Equal(2, FlatImageRelationships(reopened.Save(true)).Length);
    }

    // ─── Fixtures ────────────────────────────────────────────────────

    /// <summary>The client-facing fixture: an embedded PNG in the first paragraph and an external
    /// linked picture in the second. Regenerate <c>TestFiles/IM762-ImageCoverage.docx</c> from it.</summary>
    public static byte[] BuildCoverageFixture()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = Paragraphs(session);
        Assert.True(session.InsertImage(paragraphs[0], 0, Png(2, 3), new ImageInsertOptions { AltText = "embedded" }).Success);
        Assert.True(session.InsertImage(paragraphs[1], 0, Png(4, 5), new ImageInsertOptions { AltText = "linked" }).Success);
        return MutatePackage(session.Save(true), document =>
        {
            var main = document.MainDocumentPart!;
            var blip = main.GetXDocument().Descendants(A + "blip").Last();
            var embedId = (string)blip.Attribute(R + "embed")!;
            blip.Attribute(R + "embed")!.Remove();
            blip.SetAttributeValue(R + "link", "rIdLinkedPicture");
            main.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                new Uri("https://example.test/linked.png"), "rIdLinkedPicture");
            main.PutXDocument();
            main.DeletePart(embedId);
        });
    }

    private static byte[] LinkedFixture(string relationshipId, string target)
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.InsertImage(Paragraphs(seed)[0], 0, Png(2, 3)).Success);
        return MutatePackage(seed.Save(true), document =>
        {
            var main = document.MainDocumentPart!;
            var blip = main.GetXDocument().Descendants(A + "blip").Single();
            var embedId = (string)blip.Attribute(R + "embed")!;
            blip.Attribute(R + "embed")!.Remove();
            blip.SetAttributeValue(R + "link", relationshipId);
            main.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                new Uri(target), relationshipId);
            main.PutXDocument();
            main.DeletePart(embedId);
        });
    }

    private static byte[] VmlFixture(string style)
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.InsertImage(Paragraphs(seed)[0], 0, Png(2, 3)).Success);
        return MutatePackage(seed.Save(true), document =>
        {
            var main = document.MainDocumentPart!;
            var root = main.GetXDocument();
            var relationshipId = (string)root.Descendants(A + "blip").Single().Attribute(R + "embed")!;
            root.Descendants(W + "p").Skip(1).First().Add(new XElement(W + "r",
                new XElement(W + "pict",
                    new XElement(V + "shape", new XAttribute("style", style), new XAttribute("alt", "legacy"),
                        new XElement(V + "imagedata", new XAttribute(R + "id", relationshipId))))));
            main.PutXDocument();
        });
    }

    private static byte[] SvgFixture()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.InsertImage(Paragraphs(seed)[0], 0, Png(2, 3)).Success);
        return MutatePackage(seed.Save(true), document =>
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
                    new XElement(ASVG + "svgBlip", new XAttribute(R + "embed", "rIdSvgArt")))));
            main.PutXDocument();
        });
    }

    private static byte[] AlternateContentFixture()
    {
        using var seed = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(seed.InsertImage(Paragraphs(seed)[0], 0, Png(2, 3)).Success);
        return MutatePackage(seed.Save(true), document =>
        {
            var main = document.MainDocumentPart!;
            var root = main.GetXDocument();
            var drawing = root.Descendants(W + "drawing").Single();
            var relationshipId = (string)drawing.Descendants(A + "blip").Single().Attribute(R + "embed")!;
            drawing.ReplaceWith(new XElement(MC + "AlternateContent",
                new XAttribute(XNamespace.Xmlns + "wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"),
                new XElement(MC + "Choice", new XAttribute("Requires", "wp14"), new XElement(drawing)),
                new XElement(MC + "Fallback",
                    new XElement(W + "pict",
                        new XElement(V + "shape", new XAttribute("style", "width:1.5pt;height:2.25pt"),
                            new XElement(V + "imagedata", new XAttribute(R + "id", relationshipId)))))));
            main.PutXDocument();
        });
    }

    // ─── Helpers ─────────────────────────────────────────────────────

    private static string[] Paragraphs(DocxSession session, string scope = "body") =>
        session.Project().AnchorIndex.Values.Where(target => target.Anchor.Scope == scope
            && target.Anchor.Kind is "p" or "h" or "li")
            .Select(target => target.Anchor.Id).Distinct().ToArray();

    private static void AssertSchemaValid(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator().Validate(document)
            .Where(error => !(error.Description ?? string.Empty).Contains("powertools.codeplex.com", StringComparison.Ordinal))
            .Select(error => error.Description).ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors));
    }

    private static (string OwnerUri, string RelId, string TargetUri)[] FlatImageRelationships(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var owners = new List<OpenXmlPart> { main };
        owners.AddRange(main.HeaderParts);
        owners.AddRange(main.FooterParts);
        if (main.FootnotesPart is not null) owners.Add(main.FootnotesPart);
        if (main.EndnotesPart is not null) owners.Add(main.EndnotesPart);
        return owners.SelectMany(owner => owner.Parts
                .Where(pair => pair.OpenXmlPart is ImagePart)
                .Select(pair => (owner.Uri.ToString(), pair.RelationshipId, pair.OpenXmlPart.Uri.ToString())))
            .OrderBy(value => value).ToArray();
    }

    private static byte[] MutatePackage(byte[] bytes, Action<WordprocessingDocument> mutate)
    {
        var stream = new MemoryStream(); stream.Write(bytes); stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true)) mutate(document);
        return stream.ToArray();
    }

    internal static byte[] Png(int width, int height)
    {
        var bytes = new byte[24];
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0, 0, 0, 13, (byte)'I', (byte)'H', (byte)'D', (byte)'R' }.CopyTo(bytes, 0);
        bytes[16] = (byte)(width >> 24); bytes[17] = (byte)(width >> 16); bytes[18] = (byte)(width >> 8); bytes[19] = (byte)width;
        bytes[20] = (byte)(height >> 24); bytes[21] = (byte)(height >> 16); bytes[22] = (byte)(height >> 8); bytes[23] = (byte)height;
        return bytes;
    }

    /// <summary>A minimal lossless (VP8L) WebP header: RIFF/WEBP, the VP8L chunk, the 0x2F
    /// signature, and the 14-bit width-1/height-1 pair the header parser reads.</summary>
    internal static byte[] Webp(int width, int height)
    {
        var bytes = new byte[30];
        "RIFF"u8.CopyTo(bytes);
        bytes[4] = 22;
        "WEBP"u8.CopyTo(bytes.AsSpan(8));
        "VP8L"u8.CopyTo(bytes.AsSpan(12));
        bytes[16] = 10;
        bytes[20] = 0x2F;
        int w = width - 1, h = height - 1;
        bytes[21] = (byte)(w & 0xFF);
        bytes[22] = (byte)(((w >> 8) & 0x3F) | ((h & 0x03) << 6));
        bytes[23] = (byte)((h >> 2) & 0xFF);
        bytes[24] = (byte)((h >> 10) & 0x0F);
        return bytes;
    }
}
