// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;

namespace Docxodus;

public enum ImageBinaryFormat { Unknown, Png, Jpeg, Gif, Bmp, Tiff, Webp }
public enum ImageMarkupKind { ModernDrawing, LegacyVml, UnsupportedDrawing }
public enum ImagePlacement { Inline, Floating }
public enum ImageWrapMode { None, Square, Tight, Through, TopAndBottom, Unknown }
public enum ImageWrapSide { BothSides, Left, Right, Largest, Unknown }
public enum ImageHorizontalReference { Page, Margin, Column, Character, Unknown }
public enum ImageVerticalReference { Page, Margin, Paragraph, Line, Unknown }
public enum ImageHorizontalAlignment { Left, Center, Right, Inside, Outside, Unknown }
public enum ImageVerticalAlignment { Top, Center, Bottom, Inside, Outside, Unknown }

/// <summary>One vertex of a tight/through wrap outline in DrawingML's 21600-unit wrap space,
/// where (0,0) is the picture's top-left corner and (21600,21600) its bottom-right.</summary>
public readonly record struct ImageWrapPoint(long X, long Y);

/// <summary>The outline body text follows under tight and through wrap: a start vertex followed by
/// at least two line segments. Word derives it from picture transparency; the session has no
/// decoder, so it writes <see cref="Rectangle"/> unless the caller supplies vertices.
/// <see cref="Edited"/> mirrors <c>wp:wrapPolygon/@edited</c>, which tells Word whether to keep the
/// outline as written or recompute it from the picture.</summary>
public sealed record ImageWrapPolygon(IReadOnlyList<ImageWrapPoint> Points, bool Edited = false)
{
    /// <summary>The picture rectangle — what Word itself writes for an opaque picture.</summary>
    public static ImageWrapPolygon Rectangle { get; } = new(new[]
    {
        new ImageWrapPoint(0, 0), new ImageWrapPoint(0, 21600), new ImageWrapPoint(21600, 21600),
        new ImageWrapPoint(21600, 0), new ImageWrapPoint(0, 0),
    });

    public bool Equals(ImageWrapPolygon? other) =>
        other is not null && Edited == other.Edited && Points.SequenceEqual(other.Points);

    public override int GetHashCode()
    {
        var hash = default(HashCode);
        hash.Add(Edited);
        foreach (var point in Points) hash.Add(point);
        return hash.ToHashCode();
    }
}

/// <summary>Supported floating DrawingML layout. Offsets and wrap distances are exact EMUs;
/// position axes use either an offset or an alignment, never both.</summary>
public sealed record FloatingImageLayout
{
    public ImageHorizontalReference HorizontalRelativeFrom { get; init; } = ImageHorizontalReference.Column;
    public long? HorizontalOffsetEmu { get; init; } = 0;
    public ImageHorizontalAlignment? HorizontalAlignment { get; init; }
    public ImageVerticalReference VerticalRelativeFrom { get; init; } = ImageVerticalReference.Paragraph;
    public long? VerticalOffsetEmu { get; init; } = 0;
    public ImageVerticalAlignment? VerticalAlignment { get; init; }
    public ImageWrapMode WrapMode { get; init; } = ImageWrapMode.Square;
    public ImageWrapSide WrapSide { get; init; } = ImageWrapSide.BothSides;
    /// <summary>Tight/through wrap outline. Null on write means <see cref="ImageWrapPolygon.Rectangle"/>;
    /// a layout read from a tight/through anchor always carries the outline the document holds.</summary>
    public ImageWrapPolygon? WrapPolygon { get; init; }
    public long DistanceTopEmu { get; init; }
    public long DistanceBottomEmu { get; init; }
    public long DistanceLeftEmu { get; init; }
    public long DistanceRightEmu { get; init; }
    public uint RelativeHeight { get; init; } = 251658240;
    public bool BehindDocument { get; init; }
    public bool Locked { get; init; }
    public bool LayoutInCell { get; init; } = true;
    public bool AllowOverlap { get; init; } = true;
    /// <summary>Raw OOXML tokens are populated only when a report-only layout contains a token
    /// outside the mutable subset. They preserve inspection truth without making it writable.</summary>
    public string? RawHorizontalReference { get; init; }
    public string? RawVerticalReference { get; init; }
    public string? RawHorizontalPosition { get; init; }
    public string? RawVerticalPosition { get; init; }
    public string? RawWrapMode { get; init; }
    public string? RawWrapSide { get; init; }
    public string? RawRelativeSizeHorizontal { get; init; }
    public string? RawRelativeSizeVertical { get; init; }
    public IReadOnlyDictionary<string, string>? RawFlagTokens { get; init; }
}

/// <summary>Options for binary image insertion. Rendered dimensions are points. At 96 DPI the
/// default is exactly 0.75 point per intrinsic pixel.</summary>
public sealed record ImageInsertOptions
{
    public ImagePlacement Placement { get; init; } = ImagePlacement.Inline;
    public double? WidthPoints { get; init; }
    public double? HeightPoints { get; init; }
    public bool PreserveAspect { get; init; } = true;
    public string? AltText { get; init; }
    public string? Title { get; init; }
    public FloatingImageLayout? FloatingLayout { get; init; }
}

public sealed record ImageFormatCapability(
    ImageBinaryFormat Format, string ContentType, bool CanInspect,
    bool CanInsert, bool CanReplace, string? Limitation);

/// <summary>Which operations one markup family accepts, independent of any particular occurrence.
/// The live answer for a given picture is <see cref="ImageOccurrence.Operations"/>.</summary>
public sealed record ImageMarkupCapability(
    string Markup, IReadOnlyList<string> Operations, string? Limitation);

/// <summary>Versioned runtime facts for the native image surface. These are operational
/// capabilities, not decoder/network/file-I/O claims.</summary>
public sealed record ImageCapabilities(
    int SchemaVersion, string Runtime, IReadOnlyList<ImageFormatCapability> Formats,
    IReadOnlyList<string> Operations, IReadOnlyList<ImageWrapMode> MutableWrapModes,
    IReadOnlyList<ImageHorizontalReference> HorizontalReferences,
    IReadOnlyList<ImageVerticalReference> VerticalReferences,
    long MaxInputBytes, double MaxRenderedPoints, double DefaultDpi,
    bool UsesHeaderParsingOnly, bool AcceptsBinaryBytes,
    bool SupportsNetworkFetch, bool SupportsFileIo,
    IReadOnlyList<ImageMarkupCapability> Markups, IReadOnlyList<string> TrackedOperations);

/// <summary>Whether one image operation can run against one occurrence in the session's current
/// tracked-change mode, with the reason when it cannot.</summary>
public sealed record ImageOperationSupport(string Operation, bool CanMutate, string? Reason);

/// <summary>The per-occurrence operation matrix. <see cref="ImageOccurrence.CanMutate"/> is the
/// conjunction of the standard set — replace, set_dimensions, set_metadata, remove, and
/// set_floating_layout for a floating picture; <see cref="EmbedLinked"/> is the explicit conversion
/// a linked picture takes instead of replace.</summary>
public sealed record ImageOperationMatrix(
    ImageOperationSupport Replace, ImageOperationSupport EmbedLinked,
    ImageOperationSupport SetDimensions, ImageOperationSupport SetMetadata,
    ImageOperationSupport SetFloatingLayout, ImageOperationSupport Remove)
{
    public IEnumerable<ImageOperationSupport> All
    {
        get
        {
            yield return Replace;
            yield return EmbedLinked;
            yield return SetDimensions;
            yield return SetMetadata;
            yield return SetFloatingLayout;
            yield return Remove;
        }
    }
}

/// <summary>One native Word image occurrence. Rendered dimensions are points; floating offsets
/// and distances are exact EMUs. <see cref="CanMutate"/> summarizes the standard operations;
/// <see cref="Operations"/> answers for each one, including the explicit embed_linked path.</summary>
public sealed record ImageOccurrence(
    string Id, ImageMarkupKind MarkupKind, ImagePlacement? Placement,
    bool CanMutate, string? UnsupportedReason,
    string OwningPartUri, string Scope, string AnchorId, CharSpan Span,
    string? RelationshipId, string? TargetPartUri,
    string? LinkedRelationshipId, string? LinkedTarget,
    bool IsEmbedded, bool IsLinked, bool IsBroken,
    string? MediaFileName, string? ContentType, ImageBinaryFormat Format,
    bool? ContentTypeMatchesBytes,
    int? IntrinsicWidthPixels, int? IntrinsicHeightPixels,
    double? RenderedWidthPoints, double? RenderedHeightPoints,
    string? AltText, string? Title,
    FloatingImageLayout? FloatingLayout, bool FloatingLayoutSupported,
    ImageOperationMatrix Operations);

public sealed partial class DocxSession
{
    internal const long MaxImageInputBytes = 64L * 1024 * 1024;
    internal const double MaxImageRenderedPoints = 100000;
    internal const double ImageDefaultDpi = 96.0;
    private const long EmusPerPoint = 12700;
    private const string PictureGraphicDataUri =
        "http://schemas.openxmlformats.org/drawingml/2006/picture";
    private static readonly XNamespace ImageR =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace ImageV = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace ImageO = "urn:schemas-microsoft-com:office:office";
    private static readonly XName VmlImageData = ImageV + "imagedata";
    private static readonly XName VmlShape = ImageV + "shape";

    private enum ImageOperation { Replace, EmbedLinked, SetDimensions, SetMetadata, SetFloatingLayout, Remove }

    /// <summary>The markup shapes the matrix distinguishes. Overlays (linked media, a blip
    /// extension, placement, unmodeled layout tokens, revision containers) refine the answer per
    /// operation; the family decides which writers can address the occurrence at all.</summary>
    private enum ImageFamily { Picture, LegacyVml, AlternateContent, AlternateFallback, MultiPicture, Malformed }

    private readonly record struct ImageFacts(
        ImageFamily Family, string? FamilyReason, bool IsLinked, bool BlipExtension,
        ImagePlacement? Placement, bool FloatingSupported, string? LayoutReason,
        bool HasRenderedSize, string? ContainerReason, bool SharedReference);

    /// <summary><paramref name="Outer"/> is the unit a remove or a tracked swap operates on: the
    /// <c>w:drawing</c>, the <c>w:pict</c>, or — for a canonical picture with a VML fallback — the
    /// whole <c>mc:AlternateContent</c>, so every branch changes together.</summary>
    private sealed record ImageCandidate(
        OwnedPartRelationships.Owner Owner, XElement Outer, XElement? Container,
        XElement? Blip, ImageOccurrence Info);

    private const string BlipExtensionReason =
        "picture carries a blip extension (such as an SVG asvg:svgBlip) whose payload cannot be "
        + "changed together with the raster fallback";
    private const string LinkedReplaceReason =
        "external linked images are read-only; embed_linked converts the picture to an embedded one";

    private static readonly IReadOnlyList<ImageMarkupCapability> ImageMarkups = new[]
    {
        new ImageMarkupCapability("embedded_picture",
            new[] { "replace", "set_dimensions", "set_metadata", "set_floating_layout", "remove" }, null),
        new ImageMarkupCapability("linked_picture",
            new[] { "embed_linked", "set_dimensions", "set_metadata", "set_floating_layout", "remove" },
            "external media is never converted in place: replace is refused, and embed_linked takes "
            + "caller-supplied bytes to make the picture embedded"),
        new ImageMarkupCapability("extended_picture",
            new[] { "set_dimensions", "set_metadata", "set_floating_layout", "remove" },
            "a blip extension carrying its own media (SVG art, an artistic-effect original) cannot "
            + "be replaced together with the raster fallback"),
        new ImageMarkupCapability("legacy_vml",
            new[] { "replace", "set_dimensions", "set_metadata", "remove" },
            "VML positioning is not modeled, so set_floating_layout is refused"),
        new ImageMarkupCapability("alternate_content",
            new[] { "replace", "set_dimensions", "set_metadata", "remove" },
            "operations change every branch of the mc:AlternateContent at once and require the "
            + "branches to share one embedded media relationship; the VML fallback occurrence is "
            + "inspection-only and set_floating_layout is refused"),
        new ImageMarkupCapability("multi_picture", Array.Empty<string>(),
            "drawings holding several pictures, groups, or no identifiable blip are inspection-only"),
    };

    private static readonly IReadOnlyList<string> TrackedImageOperations = new[]
    {
        "insert", "replace", "embed_linked", "set_dimensions", "set_metadata", "set_floating_layout", "remove",
    };

    public static ImageCapabilities GetImageCapabilities()
    {
#if WASM_BUILD
        const string runtime = "browser-wasm";
#else
        const string runtime = "dotnet";
#endif
        return new ImageCapabilities(
            2,
            runtime,
            new[]
            {
                new ImageFormatCapability(ImageBinaryFormat.Png, "image/png", true, true, true, null),
                new ImageFormatCapability(ImageBinaryFormat.Jpeg, "image/jpeg", true, true, true, null),
                new ImageFormatCapability(ImageBinaryFormat.Gif, "image/gif", true, true, true, null),
                new ImageFormatCapability(ImageBinaryFormat.Bmp, "image/bmp", true, true, true, null),
                new ImageFormatCapability(ImageBinaryFormat.Tiff, "image/tiff", true, true, true, null),
                new ImageFormatCapability(ImageBinaryFormat.Webp, "image/webp", true, true, true,
                    "written as an image/webp media part; consumers without WebP support show a placeholder"),
                new ImageFormatCapability(ImageBinaryFormat.Unknown, "application/octet-stream", false, false, false,
                    "unrecognized bytes are rejected"),
            },
            new[] { "list", "insert", "replace", "embed_linked", "set_dimensions", "set_metadata",
                "set_floating_layout", "remove" },
            new[] { ImageWrapMode.None, ImageWrapMode.Square, ImageWrapMode.Tight,
                ImageWrapMode.Through, ImageWrapMode.TopAndBottom },
            new[] { ImageHorizontalReference.Page, ImageHorizontalReference.Margin,
                ImageHorizontalReference.Column, ImageHorizontalReference.Character },
            new[] { ImageVerticalReference.Page, ImageVerticalReference.Margin,
                ImageVerticalReference.Paragraph, ImageVerticalReference.Line },
            MaxImageInputBytes, MaxImageRenderedPoints, ImageDefaultDpi,
            UsesHeaderParsingOnly: true, AcceptsBinaryBytes: true,
            SupportsNetworkFetch: false, SupportsFileIo: false,
            ImageMarkups, TrackedImageOperations);
    }

    public IReadOnlyList<ImageOccurrence> ListImages(ProjectionScopes scopes = ProjectionScopes.All)
    {
        ThrowIfDisposed();
        return EnumerateImageCandidates(scopes).Select(candidate => candidate.Info).ToList();
    }

    public EditResult InsertImage(string anchorId, int characterOffset, byte[] imageBytes,
        ImageInsertOptions? options = null)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        options ??= new ImageInsertOptions();
        var binary = ValidateImageBytes(imageBytes, anchorId);
        if (binary.Error is not null) return binary.Error;
        if (ValidateInsertOptions(options, binary.Width, binary.Height, anchorId,
            out var widthEmu, out var heightEmu, out var layout) is { } optionError) return optionError;
        var anchor = FindAnchor(anchorId);
        if (anchor is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        var paragraph = anchor.Resolve(_doc!);
        if (paragraph is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        if (paragraph.Name != W.p)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "InsertImage requires a paragraph/heading/list-item anchor", anchorId);
        var textLength = ParagraphText(paragraph).Length;
        if (characterOffset < 0 || characterOffset > textLength)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"offset {characterOffset} outside paragraph of length {textLength}", anchorId);
        if (ValidateImageInsertionBoundary(paragraph, characterOffset) is { } boundaryError)
            return EditResult.Fail(EditErrorCode.UnsupportedInlineBoundary, boundaryError, anchorId);
        var owner = OwnedPartRelationships.FindOwner(_doc!, paragraph);
        if (owner is null) return EditResult.Fail(EditErrorCode.InternalError,
            "paragraph has no owning story part", anchorId);
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var relationship = OwnedPartRelationships.FindOrAddImagePart(
                _doc!, owner.Value.Part, imageBytes, binary.ContentType!, binary.Format);
            var docPrId = NextDocumentPropertyId();
            var drawing = BuildImageDrawing(relationship.RelationshipId, docPrId,
                widthEmu, heightEmu, options.AltText, options.Title,
                options.Placement, layout);
            var run = new XElement(W.r,
                new XElement(W.rPr, new XElement(W.noProof)),
                drawing);
            // Under render_inline the picture is a tracked insertion: rejecting the revision removes
            // the run, and the orphan sweep drops the media part only that run referenced.
            var inserted = _trackedChanges == TrackedChangeMode.RenderInline
                ? CreateRevisionEnvelope(W.ins, NewRevisionStamp(), run)
                : run;
            InsertInlineElementAtOffset(paragraph, characterOffset, inserted);
            UnidHelper.AssignToSelfAndDescendants(inserted);
            InvalidateProjectionCache();
            var imageId = ImagePublicId(owner.Value, drawing);
            return new EditResult { Success = true, ImageId = imageId,
                Modified = new[] { anchor.Anchor } };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    public EditResult ReplaceImage(string imageId, byte[] imageBytes)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var binary = ValidateImageBytes(imageBytes, null);
        if (binary.Error is not null) return binary.Error;
        if (ResolveMutableImage(imageId, ImageOperation.Replace, out var candidate) is { } imageError) return imageError;
        var currentPart = OwnedPartRelationships.ResolveImagePart(
            candidate.Owner.Part, candidate.Info.RelationshipId);
        if (currentPart is not null && currentPart.ContentType == binary.ContentType
            && OwnedPartRelationships.ReadPartBytes(currentPart).SequenceEqual(imageBytes))
            return ImageMutationSuccess(candidate, imageId);
        return MutateImage(candidate, imageId, ImageOperation.Replace, outer =>
        {
            var relationship = OwnedPartRelationships.FindOrAddImagePart(
                _doc!, candidate.Owner.Part, imageBytes, binary.ContentType!, binary.Format);
            foreach (var reference in EmbeddedImageReferences(outer))
                if (reference.Value == candidate.Info.RelationshipId) reference.Value = relationship.RelationshipId;
        });
    }

    /// <summary>The explicit conversion a linked picture needs before its media can be edited in
    /// place: the caller fetches the external image (the session never does) and hands over the
    /// bytes, which become an owner-local media part that the picture now embeds. The external
    /// relationship disappears with its last reference.</summary>
    public EditResult EmbedLinkedImage(string imageId, byte[] imageBytes)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var binary = ValidateImageBytes(imageBytes, null);
        if (binary.Error is not null) return binary.Error;
        if (ResolveMutableImage(imageId, ImageOperation.EmbedLinked, out var candidate) is { } imageError) return imageError;
        return MutateImage(candidate, imageId, ImageOperation.EmbedLinked, outer =>
        {
            var relationship = OwnedPartRelationships.FindOrAddImagePart(
                _doc!, candidate.Owner.Part, imageBytes, binary.ContentType!, binary.Format);
            foreach (var blip in outer.DescendantsAndSelf(A.blip))
            {
                blip.SetAttributeValue(ImageR + "embed", relationship.RelationshipId);
                blip.Attribute(ImageR + "link")?.Remove();
            }
            foreach (var imageData in outer.DescendantsAndSelf(VmlImageData))
            {
                imageData.SetAttributeValue(ImageR + "id", relationship.RelationshipId);
                imageData.Attribute(ImageR + "href")?.Remove();
            }
        });
    }

    public EditResult SetImageDimensions(string imageId, double? widthPoints,
        double? heightPoints, bool preserveAspect = true)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (ResolveMutableImage(imageId, ImageOperation.SetDimensions, out var candidate) is { } imageError) return imageError;
        if (ResolveRenderedDimensions(widthPoints, heightPoints, preserveAspect,
            candidate.Info.RenderedWidthPoints, candidate.Info.RenderedHeightPoints,
            out var widthEmu, out var heightEmu) is { } dimensionError) return dimensionError;
        if (HasExtents(candidate.Outer, widthEmu, heightEmu))
            return ImageMutationSuccess(candidate, imageId);
        return MutateImage(candidate, imageId, ImageOperation.SetDimensions,
            outer => SetImageExtents(outer, widthEmu, heightEmu));
    }

    public EditResult SetImageMetadata(string imageId, string? altText, string? title)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (ResolveMutableImage(imageId, ImageOperation.SetMetadata, out var candidate) is { } imageError) return imageError;
        if (!ValidXmlAttributeText(altText) || !ValidXmlAttributeText(title))
            return EditResult.Fail(EditErrorCode.InvalidImageData,
                "image metadata contains characters XML attributes cannot represent");
        if (MetadataTargets(candidate.Outer).All(target =>
                (string?)target.Element.Attribute(target.Name) == (target.IsTitle ? title : altText)))
            return ImageMutationSuccess(candidate, imageId);
        return MutateImage(candidate, imageId, ImageOperation.SetMetadata, outer =>
        {
            foreach (var target in MetadataTargets(outer))
                target.Element.SetAttributeValue(target.Name, target.IsTitle ? title : altText);
        });
    }

    public EditResult SetImageFloatingLayout(string imageId, FloatingImageLayout layout)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (layout is null) return EditResult.Fail(EditErrorCode.InvalidImageLayout,
            "floating layout is required");
        if (ValidateFloatingLayout(layout) is { } layoutError) return layoutError;
        if (ResolveMutableImage(imageId, ImageOperation.SetFloatingLayout, out var candidate) is { } imageError)
            return candidate is { Info.Placement: ImagePlacement.Inline }
                ? EditResult.Fail(EditErrorCode.InvalidImageLayout,
                    "floating layout can only be set on a floating image")
                : imageError;
        var normalized = NormalizeWrapPolygon(layout);
        if (candidate.Info.FloatingLayout == normalized)
            return ImageMutationSuccess(candidate, imageId);
        return MutateImage(candidate, imageId, ImageOperation.SetFloatingLayout,
            outer => ApplyFloatingLayout(outer.DescendantsAndSelf(WP.anchor).First(), normalized));
    }

    public EditResult RemoveImage(string imageId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (ResolveMutableImage(imageId, ImageOperation.Remove, out var candidate) is { } imageError) return imageError;
        return MutateImage(candidate, imageId, ImageOperation.Remove, null);
    }

    /// <summary>Every occurrence mutation funnels through here. Under render_inline a change to an
    /// existing picture is recorded the one way WordprocessingML can express it: the original run
    /// becomes a tracked deletion and a re-identified copy carrying the change becomes a tracked
    /// insertion (remove records the deletion alone). Accepting yields the new state and the orphan
    /// sweep drops media only the deleted run referenced; rejecting restores bytes, relationships,
    /// metadata and geometry because the original run was never edited. A picture already inside the
    /// session author's own insertion is edited in place, so replace-then-refit stays one revision.</summary>
    private EditResult MutateImage(ImageCandidate candidate, string imageId,
        ImageOperation operation, Action<XElement>? write)
    {
        var paragraph = candidate.Outer.Ancestors(W.p).First();
        var anchor = AnchorForElement(paragraph);
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var outer = candidate.Outer;
            if (_trackedChanges == TrackedChangeMode.RenderInline && !InsideOwnInsertion(outer, paragraph))
            {
                var run = IsolateImageRun(outer);
                var stamp = NewRevisionStamp();
                var deleted = CreateRevisionEnvelope(W.del, stamp);
                if (operation == ImageOperation.Remove)
                {
                    run.ReplaceWith(deleted);
                    deleted.Add(run);
                }
                else
                {
                    var copy = new XElement(run);
                    var inserted = CreateRevisionEnvelope(W.ins, stamp);
                    run.ReplaceWith(deleted, inserted);
                    deleted.Add(run);
                    inserted.Add(copy);
                    outer = ReidentifyImageRun(copy, outer.Name);
                    write!(outer);
                }
            }
            else if (operation == ImageOperation.Remove) RemoveImageElement(outer);
            else write!(outer);
            OwnedPartRelationships.SweepOrphanedImages(candidate.Owner.Part);
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                ImageId = operation == ImageOperation.Remove
                    ? imageId
                    : ImagePublicId(candidate.Owner, ImageIdentityElement(outer)),
                Modified = anchor is null ? Array.Empty<Anchor>() : new[] { anchor.Value },
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    private static XElement ImageIdentityElement(XElement outer) =>
        outer.Name == MC.AlternateContent ? outer.Descendants(W.drawing).First() : outer;

    private static bool InsideOwnInsertion(XElement outer, XElement paragraph) =>
        outer.Ancestors().TakeWhile(element => !ReferenceEquals(element, paragraph))
            .Any(element => element.Name == W.ins);

    /// <summary>Move any run content that is not the picture into sibling runs carrying the same
    /// properties, so the revision envelopes wrap the picture alone.</summary>
    private static XElement IsolateImageRun(XElement outer)
    {
        var run = outer.Parent!;
        var others = run.Elements()
            .Where(element => element.Name != W.rPr && !ReferenceEquals(element, outer)).ToList();
        if (others.Count == 0) return run;
        var properties = run.Element(W.rPr);
        foreach (var (content, before) in new[]
        {
            (others.Where(element => element.IsBefore(outer)).ToList(), true),
            (others.Where(element => element.IsAfter(outer)).ToList(), false),
        })
        {
            if (content.Count == 0) continue;
            var sibling = new XElement(W.r);
            if (properties is not null) sibling.Add(new XElement(properties));
            foreach (var node in content)
            {
                node.Remove();
                sibling.Add(node);
            }
            if (before) run.AddBeforeSelf(sibling);
            else run.AddAfterSelf(sibling);
            UnidHelper.AssignToSelfAndDescendants(sibling);
        }
        return run;
    }

    /// <summary>A copied run must be a new occurrence: fresh Unids (the deleted original keeps its
    /// id) and fresh drawing property ids (Word wants them unique per document).</summary>
    private XElement ReidentifyImageRun(XElement copy, XName outerName)
    {
        foreach (var element in copy.DescendantsAndSelf())
            element.Attribute(PtOpenXml.Unid)?.Remove();
        UnidHelper.AssignToSelfAndDescendants(copy);
        foreach (var docPr in copy.Descendants(WP.docPr).ToList())
        {
            var id = NextDocumentPropertyId();
            docPr.SetAttributeValue("id", id);
            var picture = docPr.Parent?.Descendants(Pic.cNvPr).FirstOrDefault();
            picture?.SetAttributeValue("id", id);
        }
        return copy.Elements().First(element => element.Name == outerName);
    }

    private static void RemoveImageElement(XElement outer)
    {
        var run = outer.Parent!;
        outer.Remove();
        if (run.Elements().Any(element => element.Name != W.rPr)) return;
        var wrapper = run.Parent;
        run.Remove();
        if (wrapper is not null && wrapper.Name == W.ins && !wrapper.HasElements) wrapper.Remove();
    }

    // ─── Writers shared by every family ──────────────────────────────

    /// <summary>The attributes that embed media into a picture: <c>a:blip/@r:embed</c> and
    /// <c>v:imagedata/@r:id</c>. Deliberately not name-blind — <c>a:hlinkClick/@r:id</c> on a
    /// clickable picture must never be re-pointed at an image part.</summary>
    private static IEnumerable<XAttribute> EmbeddedImageReferences(XElement outer) =>
        outer.DescendantsAndSelf().SelectMany(element =>
            element.Name == A.blip ? element.Attributes(ImageR + "embed")
            : element.Name == VmlImageData ? element.Attributes(ImageR + "id")
            : Enumerable.Empty<XAttribute>());

    /// <summary>Where alt text and title live per markup: DrawingML keeps both on
    /// <c>wp:docPr</c> and <c>pic:cNvPr</c>; VML keeps alt on the shape and the title on its image data.</summary>
    private readonly record struct MetadataTarget(XElement Element, XName Name, bool IsTitle);

    private static IEnumerable<MetadataTarget> MetadataTargets(XElement outer)
    {
        foreach (var element in outer.DescendantsAndSelf())
        {
            if (element.Name == WP.docPr || element.Name == Pic.cNvPr)
            {
                yield return new MetadataTarget(element, "descr", false);
                yield return new MetadataTarget(element, "title", true);
            }
            else if (element.Name == VmlShape) yield return new MetadataTarget(element, "alt", false);
            else if (element.Name == VmlImageData) yield return new MetadataTarget(element, ImageO + "title", true);
        }
    }

    private static bool HasExtents(XElement outer, long widthEmu, long heightEmu)
    {
        var width = widthEmu.ToString(CultureInfo.InvariantCulture);
        var height = heightEmu.ToString(CultureInfo.InvariantCulture);
        foreach (var container in DrawingContainers(outer))
        {
            var extent = container.Element(WP.extent);
            var transform = container.Descendants(A.xfrm).FirstOrDefault()?.Element(A.ext);
            if ((string?)extent?.Attribute("cx") != width || (string?)extent?.Attribute("cy") != height
                || (string?)transform?.Attribute("cx") != width || (string?)transform?.Attribute("cy") != height)
                return false;
        }
        foreach (var shape in outer.DescendantsAndSelf(VmlShape))
        {
            var (shapeWidth, shapeHeight) = ReadVmlStylePoints(shape);
            if (shapeWidth is null || shapeHeight is null
                || PointsToEmu(shapeWidth.Value) != widthEmu || PointsToEmu(shapeHeight.Value) != heightEmu)
                return false;
        }
        return true;
    }

    private static void SetImageExtents(XElement outer, long widthEmu, long heightEmu)
    {
        foreach (var container in DrawingContainers(outer)) SetDrawingExtents(container, widthEmu, heightEmu);
        foreach (var shape in outer.DescendantsAndSelf(VmlShape)) SetVmlStyleSize(shape, widthEmu, heightEmu);
    }

    private static IEnumerable<XElement> DrawingContainers(XElement outer) =>
        outer.DescendantsAndSelf(W.drawing).SelectMany(drawing => drawing.Elements()
            .Where(element => element.Name == WP.inline || element.Name == WP.anchor));

    private static long PointsToEmu(double points) =>
        (long)Math.Round(points * EmusPerPoint, MidpointRounding.AwayFromZero);

    /// <summary>VML sizes live in the shape's CSS <c>style</c> attribute. Only the width/height
    /// entries are rewritten; every other entry keeps its text and position.</summary>
    private static void SetVmlStyleSize(XElement shape, long widthEmu, long heightEmu)
    {
        var entries = ((string?)shape.Attribute("style") ?? string.Empty)
            .Split(';', StringSplitOptions.RemoveEmptyEntries)
            .Select(entry => entry.Trim()).Where(entry => entry.Length > 0).ToList();
        void Put(string key, long emu)
        {
            var value = $"{key}:{(emu / (double)EmusPerPoint).ToString("0.###", CultureInfo.InvariantCulture)}pt";
            int index = entries.FindIndex(entry =>
                entry.StartsWith(key + ":", StringComparison.OrdinalIgnoreCase));
            if (index >= 0) entries[index] = value;
            else entries.Add(value);
        }
        Put("width", widthEmu);
        Put("height", heightEmu);
        shape.SetAttributeValue("style", string.Join(";", entries));
    }

    private static (double? Width, double? Height) ReadVmlStylePoints(XElement? shape)
    {
        double? width = null, height = null;
        foreach (var entry in ((string?)shape?.Attribute("style") ?? string.Empty).Split(';'))
        {
            var separator = entry.IndexOf(':');
            if (separator < 0) continue;
            var key = entry[..separator].Trim();
            if (key.Equals("width", StringComparison.OrdinalIgnoreCase)) width = ParseCssPoints(entry[(separator + 1)..]);
            else if (key.Equals("height", StringComparison.OrdinalIgnoreCase)) height = ParseCssPoints(entry[(separator + 1)..]);
        }
        return width is > 0 && height is > 0 ? (width, height) : (null, null);
    }

    private static double? ParseCssPoints(string value)
    {
        value = value.Trim();
        foreach (var (unit, factor) in new[] { ("pt", 1.0), ("in", 72.0), ("cm", 72.0 / 2.54),
            ("mm", 72.0 / 25.4), ("px", 0.75), ("pc", 12.0) })
        {
            if (!value.EndsWith(unit, StringComparison.OrdinalIgnoreCase)) continue;
            return double.TryParse(value[..^unit.Length], NumberStyles.Float, CultureInfo.InvariantCulture,
                out var number) && double.IsFinite(number) ? number * factor : null;
        }
        return null;
    }

    // ─── Discovery ───────────────────────────────────────────────────

    private IReadOnlyList<ImageCandidate> EnumerateImageCandidates(ProjectionScopes scopes)
    {
        _ = AnchorIndex();
        var result = new List<ImageCandidate>();
        foreach (var owner in OwnedPartRelationships.StoryParts(_doc!))
        {
            if (!scopes.IncludesScope(owner.Scope)) continue;
            var root = owner.Part.GetXDocument().Root;
            if (root is null) continue;
            foreach (var drawing in root.Descendants(W.drawing))
                result.AddRange(BuildDrawingCandidates(owner, drawing));
            foreach (var pict in root.Descendants(W.pict))
                result.AddRange(BuildLegacyCandidates(owner, pict));
        }
        return result;
    }

    private IEnumerable<ImageCandidate> BuildDrawingCandidates(
        OwnedPartRelationships.Owner owner, XElement drawing)
    {
        var paragraph = drawing.Ancestors(W.p).FirstOrDefault();
        var anchor = paragraph is null ? null : AnchorForElement(paragraph);
        if (paragraph is null || anchor is null) yield break;
        var alternate = AlternateContentHost(drawing, paragraph);
        var canonicalAlternate = alternate is not null && IsCanonicalAlternateContent(alternate, drawing);
        var outer = canonicalAlternate ? alternate! : drawing;
        var containers = drawing.Elements().Where(element =>
            element.Name == WP.inline || element.Name == WP.anchor).ToList();
        var container = containers.Count == 1 ? containers[0] : null;
        var blips = drawing.Descendants(A.blip).ToList();
        var occurrences = blips.Count == 0
            ? new XElement?[] { null }
            : blips.Cast<XElement?>().ToArray();
        ImagePlacement? placement = container?.Name == WP.inline ? ImagePlacement.Inline
            : container?.Name == WP.anchor ? ImagePlacement.Floating : null;
        var extent = container?.Element(WP.extent);
        var (renderedWidth, renderedHeight) = ReadRenderedPoints(extent);
        var docPr = container?.Element(WP.docPr);
        FloatingImageLayout? floatingLayout = null;
        bool floatingSupported = placement != ImagePlacement.Floating;
        string? layoutUnsupported = null;
        if (placement == ImagePlacement.Floating && container is not null)
        {
            floatingSupported = TryReadFloatingLayout(container, out floatingLayout, out layoutUnsupported);
        }
        var containerReason = ImageContainerReason(outer, paragraph);
        bool hasMutableStructure = container?.Element(WP.docPr) is not null
            && container.Element(WP.extent) is not null
            && container.Descendants(A.xfrm).FirstOrDefault()?.Element(A.ext) is not null;
        var offset = ImageOffset(paragraph, drawing);
        for (int index = 0; index < occurrences.Length; index++)
        {
            var blip = occurrences[index];
            var graphicData = blip?.Ancestors(A.graphicData).FirstOrDefault();
            bool picturePayload = container is not null && blips.Count == 1
                && graphicData is not null
                && ReferenceEquals(graphicData, container.Element(A.graphic)?.Element(A.graphicData))
                && (string?)graphicData.Attribute("uri") == PictureGraphicDataUri
                && graphicData.Descendants(Pic._pic).Count() == 1;
            var occurrenceProperties = blip?.Ancestors(Pic._pic).FirstOrDefault()?
                .Descendants(Pic.cNvPr).FirstOrDefault();
            var alt = (string?)occurrenceProperties?.Attribute("descr")
                ?? (picturePayload ? (string?)docPr?.Attribute("descr") : null);
            var title = (string?)occurrenceProperties?.Attribute("title")
                ?? (picturePayload ? (string?)docPr?.Attribute("title") : null);
            var embedId = (string?)blip?.Attribute(ImageR + "embed");
            var linkId = (string?)blip?.Attribute(ImageR + "link");
            var imagePart = OwnedPartRelationships.ResolveImagePart(owner.Part, embedId);
            var linked = string.IsNullOrEmpty(linkId) ? null
                : owner.Part.ExternalRelationships.FirstOrDefault(relationship => relationship.Id == linkId);
            var media = ReadMediaInfo(imagePart);
            bool broken = (!string.IsNullOrEmpty(embedId) && imagePart is null)
                || (!string.IsNullOrEmpty(linkId) && linked is null)
                || (string.IsNullOrEmpty(embedId) && string.IsNullOrEmpty(linkId))
                || media.IsMalformed || media.ContentTypeMatchesBytes == false;
            // A blip extension that names its OWN image relationship is a SECOND payload behind the
            // same a:blip: an SVG keeps the vector art in a:extLst/asvg:svgBlip while a:blip/@r:embed
            // holds only the raster fallback, and Word's artistic effects keep the untouched original
            // in a14:imgProps/a14:imgLayer, which carries its own r:embed. Replacing only the
            // fallback would leave an SVG-aware renderer showing the OLD picture while the API
            // reported success, so replace and embed_linked are refused for such pictures; sizing,
            // metadata, layout and removal touch both payloads together and stay available.
            //
            // The test is STRUCTURAL rather than a name list, mirroring what the relationship sweep
            // does: carrying a relationship attribute is the property that makes an extension a
            // second payload. Extensions that only hint at rendering carry no relationship and stay
            // mutable — notably a14:useLocalDpi, which Word writes on most inserted pictures, so
            // testing for the mere PRESENCE of an a:extLst would refuse ordinary Word images.
            bool blipExtension = blip is not null
                && blip.Elements()
                    .Where(element => element.Name.LocalName == "extLst")
                    .Descendants()
                    .Any(element => element.Attributes()
                        .Any(attribute => attribute.Name.Namespace == ImageR));
            ImageFamily family;
            string? familyReason = null;
            if (!picturePayload)
            {
                family = blip is null ? ImageFamily.Malformed : ImageFamily.MultiPicture;
                familyReason = blip is null
                    ? "drawing contains no identifiable image blip"
                    : "drawing contains a non-canonical or multi-picture payload";
            }
            else if (!hasMutableStructure)
            {
                family = ImageFamily.Malformed;
                familyReason = "drawing lacks required picture properties or extents";
            }
            else family = canonicalAlternate ? ImageFamily.AlternateContent : ImageFamily.Picture;
            var facts = new ImageFacts(family, familyReason, !string.IsNullOrEmpty(linkId), blipExtension,
                placement, floatingSupported, layoutUnsupported, renderedWidth is not null, containerReason,
                canonicalAlternate && SharesOneEmbeddedReference(alternate!, embedId));
            var operations = BuildImageOperations(facts);
            var (canMutate, unsupported) = SummarizeOperations(operations, placement);
            var info = new ImageOccurrence(
                ImagePublicId(owner, drawing, occurrences.Length == 1 ? null : index),
                picturePayload ? ImageMarkupKind.ModernDrawing : ImageMarkupKind.UnsupportedDrawing,
                placement, canMutate, unsupported,
                owner.PartUri, owner.Scope, anchor.Value.Id, new CharSpan(offset, 0),
                embedId ?? linkId, imagePart?.Uri.ToString(), linkId, linked?.Uri.ToString(),
                !string.IsNullOrEmpty(embedId), !string.IsNullOrEmpty(linkId), broken,
                imagePart is null ? null : Path.GetFileName(imagePart.Uri.OriginalString),
                imagePart?.ContentType, media.Format, media.ContentTypeMatchesBytes,
                media.Width, media.Height, renderedWidth, renderedHeight, alt, title,
                floatingLayout, floatingSupported, operations);
            yield return new ImageCandidate(owner, outer, container, blip, info);
        }
    }

    private IEnumerable<ImageCandidate> BuildLegacyCandidates(
        OwnedPartRelationships.Owner owner, XElement pict)
    {
        var imageDataOccurrences = pict.Descendants(VmlImageData).ToList();
        if (imageDataOccurrences.Count == 0) yield break;
        var paragraph = pict.Ancestors(W.p).FirstOrDefault();
        var anchor = paragraph is null ? null : AnchorForElement(paragraph);
        if (paragraph is null || anchor is null) yield break;
        var alternate = AlternateContentHost(pict, paragraph);
        var containerReason = ImageContainerReason(pict, paragraph);
        for (int index = 0; index < imageDataOccurrences.Count; index++)
        {
            var imageData = imageDataOccurrences[index];
            var relationshipId = (string?)imageData.Attribute(ImageR + "id");
            var imagePart = OwnedPartRelationships.ResolveImagePart(owner.Part, relationshipId);
            var external = string.IsNullOrEmpty(relationshipId) ? null
                : owner.Part.ExternalRelationships.FirstOrDefault(relationship => relationship.Id == relationshipId);
            var media = ReadMediaInfo(imagePart);
            var shape = imageData.Ancestors(VmlShape).FirstOrDefault();
            var (renderedWidth, renderedHeight) = ReadVmlStylePoints(shape);
            ImageFamily family;
            string? familyReason = null;
            if (alternate is not null)
            {
                family = ImageFamily.AlternateFallback;
                familyReason = "VML fallback of an AlternateContent picture; operate on the modern occurrence";
            }
            else if (imageDataOccurrences.Count != 1 || shape is null || string.IsNullOrEmpty(relationshipId))
            {
                family = ImageFamily.Malformed;
                familyReason = "legacy VML picture is not a single v:shape/v:imagedata pair with an r:id relationship";
            }
            else family = ImageFamily.LegacyVml;
            var facts = new ImageFacts(family, familyReason, external is not null, false, null, false, null,
                renderedWidth is not null, containerReason, false);
            var operations = BuildImageOperations(facts);
            var (canMutate, unsupported) = SummarizeOperations(operations, null);
            var info = new ImageOccurrence(
                ImagePublicId(owner, pict, imageDataOccurrences.Count == 1 ? null : index),
                ImageMarkupKind.LegacyVml, null, canMutate, unsupported,
                owner.PartUri, owner.Scope, anchor.Value.Id, new CharSpan(ImageOffset(paragraph, pict), 0),
                relationshipId, imagePart?.Uri.ToString(), external is null ? null : relationshipId,
                external?.Uri.ToString(), imagePart is not null, external is not null,
                !string.IsNullOrEmpty(relationshipId) && (imagePart is null || media.IsMalformed
                    || media.ContentTypeMatchesBytes == false) && external is null,
                imagePart is null ? null : Path.GetFileName(imagePart.Uri.OriginalString),
                imagePart?.ContentType, media.Format, media.ContentTypeMatchesBytes,
                media.Width, media.Height, renderedWidth, renderedHeight, (string?)shape?.Attribute("alt"),
                (string?)imageData.Attribute(ImageO + "title"), null, false, operations);
            yield return new ImageCandidate(owner, pict, null, null, info);
        }
    }

    private static XElement? AlternateContentHost(XElement image, XElement paragraph) =>
        image.Ancestors().TakeWhile(element => !ReferenceEquals(element, paragraph))
            .FirstOrDefault(element => element.Name == MC.AlternateContent);

    /// <summary>The one AlternateContent shape the writers can keep coherent: a run holding exactly
    /// one choice whose sole content is this drawing, plus at most one fallback holding a single VML
    /// picture. Anything richer (several choices, a text box wrapping the picture) stays read-only.</summary>
    private static bool IsCanonicalAlternateContent(XElement alternate, XElement drawing)
    {
        if (alternate.Parent?.Name != W.r) return false;
        var choice = alternate.Element(MC.Choice);
        var fallback = alternate.Element(MC.Fallback);
        if (choice is null || alternate.Elements().Count() != (fallback is null ? 1 : 2)) return false;
        if (choice.Elements().Count() != 1 || !ReferenceEquals(choice.Elements().First(), drawing)) return false;
        if (fallback is null) return true;
        var pict = fallback.Elements().SingleOrDefault();
        return pict is not null && pict.Name == W.pict && pict.Descendants(VmlImageData).Count() == 1;
    }

    private static bool SharesOneEmbeddedReference(XElement alternate, string? embedId)
    {
        if (string.IsNullOrEmpty(embedId)) return false;
        var payloads = alternate.Descendants()
            .Where(element => element.Name == A.blip || element.Name == VmlImageData).ToList();
        return payloads.Count > 0 && payloads.All(payload =>
            payload.Attribute(ImageR + "link") is null && payload.Attribute(ImageR + "href") is null
            && (string?)payload.Attribute(payload.Name == A.blip ? ImageR + "embed" : ImageR + "id") == embedId);
    }

    /// <summary>Why the run holding this picture cannot be edited at all, or null. Hyperlinks,
    /// smart tags and content-control wrappers hold runs and revision envelopes alike, so a picture
    /// inside them is addressable; deleted, moved and field-generated content is not, and another
    /// author's insertion is not under render_inline, where editing it would rewrite their revision.</summary>
    private string? ImageContainerReason(XElement outer, XElement paragraph)
    {
        if (outer.Parent?.Name != W.r) return "image is not inside a run";
        foreach (var ancestor in outer.Parent.Ancestors().TakeWhile(element => !ReferenceEquals(element, paragraph)))
        {
            if (ancestor.Name == W.del || ancestor.Name == W.moveFrom) return "image is inside a tracked deletion";
            if (ancestor.Name == W.moveTo) return "image is inside a tracked move";
            if (ancestor.Name == W.ins && _trackedChanges == TrackedChangeMode.RenderInline
                && !string.Equals((string?)ancestor.Attribute(W.author), _revisionAuthor ?? "docxodus", StringComparison.Ordinal))
                return "image is inside another author's tracked insertion";
            if (ancestor.Name == W.fldSimple) return "image is the result of a simple field";
            if (ancestor.Name == MC.AlternateContent)
                return "image is inside markup-compatibility AlternateContent whose branches cannot be changed together";
        }
        if (paragraph.Descendants().Any(element => element.Name == W.fldChar || element.Name == W.instrText))
            return "image is in a paragraph containing a complex field";
        return null;
    }

    private static ImageOperationMatrix BuildImageOperations(ImageFacts facts)
    {
        ImageOperationSupport Support(ImageOperation operation, string name)
        {
            var blocker = ImageOperationBlocker(facts, operation);
            return new ImageOperationSupport(name, blocker is null, blocker);
        }
        return new ImageOperationMatrix(
            Support(ImageOperation.Replace, "replace"),
            Support(ImageOperation.EmbedLinked, "embed_linked"),
            Support(ImageOperation.SetDimensions, "set_dimensions"),
            Support(ImageOperation.SetMetadata, "set_metadata"),
            Support(ImageOperation.SetFloatingLayout, "set_floating_layout"),
            Support(ImageOperation.Remove, "remove"));
    }

    private static string? ImageOperationBlocker(ImageFacts facts, ImageOperation operation)
    {
        if (facts.FamilyReason is not null) return facts.FamilyReason;
        if (facts.ContainerReason is not null) return facts.ContainerReason;
        switch (operation)
        {
            case ImageOperation.Replace:
                if (facts.BlipExtension) return BlipExtensionReason;
                if (facts.IsLinked) return LinkedReplaceReason;
                if (facts.Family == ImageFamily.AlternateContent && !facts.SharedReference)
                    return "AlternateContent branches do not share one embedded media relationship";
                return null;
            case ImageOperation.EmbedLinked:
                if (!facts.IsLinked) return "picture is already embedded";
                if (facts.BlipExtension) return BlipExtensionReason;
                if (facts.Family == ImageFamily.AlternateContent) return "AlternateContent branches cannot be embedded together";
                return null;
            case ImageOperation.SetDimensions:
                if (facts.HasRenderedSize) return null;
                return facts.Family == ImageFamily.LegacyVml
                    ? "legacy VML shape has no parsable width/height style"
                    : "picture extent is missing or malformed";
            case ImageOperation.SetFloatingLayout:
                if (facts.Family == ImageFamily.LegacyVml) return "legacy VML positioning is not modeled";
                if (facts.Family == ImageFamily.AlternateContent) return "AlternateContent branches carry independent geometry";
                if (facts.Placement != ImagePlacement.Floating) return "inline picture has no floating layout";
                return facts.FloatingSupported ? null : facts.LayoutReason ?? "floating layout is read-only";
            default:
                return null;
        }
    }

    private static (bool CanMutate, string? Reason) SummarizeOperations(
        ImageOperationMatrix operations, ImagePlacement? placement)
    {
        var standard = new List<ImageOperationSupport>
            { operations.Replace, operations.SetDimensions, operations.SetMetadata };
        if (placement == ImagePlacement.Floating) standard.Add(operations.SetFloatingLayout);
        standard.Add(operations.Remove);
        var blocker = standard.FirstOrDefault(support => !support.CanMutate);
        return (blocker is null, blocker?.Reason);
    }

    private sealed record MediaInfo(ImageBinaryFormat Format, int? Width, int? Height,
        bool? ContentTypeMatchesBytes, bool IsMalformed);

    private static MediaInfo ReadMediaInfo(ImagePart? part)
    {
        if (part is null) return new(ImageBinaryFormat.Unknown, null, null, null, false);
        var declaredFormat = FormatFromContentType(part.ContentType);
        try
        {
            var bytes = OwnedPartRelationships.ReadPartBytes(part);
            var format = FormatFromMagicToken(ImageHeaderParser.DetectFormat(bytes));
            var dimensions = ImageHeaderParser.GetDimensions(bytes);
            return new(format, dimensions?.Width, dimensions?.Height,
                format != ImageBinaryFormat.Unknown && format == declaredFormat,
                format == ImageBinaryFormat.Unknown || dimensions is null);
        }
        catch
        {
            return new(ImageBinaryFormat.Unknown, null, null, false, true);
        }
    }

    private static ImageBinaryFormat FormatFromMagicToken(string? token) => token switch
    {
        "png" => ImageBinaryFormat.Png,
        "jpeg" => ImageBinaryFormat.Jpeg,
        "gif" => ImageBinaryFormat.Gif,
        "bmp" => ImageBinaryFormat.Bmp,
        "tiff" => ImageBinaryFormat.Tiff,
        "webp" => ImageBinaryFormat.Webp,
        _ => ImageBinaryFormat.Unknown,
    };

    private static ImageBinaryFormat FormatFromContentType(string? contentType) =>
        contentType?.ToLowerInvariant() switch
        {
            "image/png" => ImageBinaryFormat.Png,
            "image/jpeg" or "image/jpg" => ImageBinaryFormat.Jpeg,
            "image/gif" => ImageBinaryFormat.Gif,
            "image/bmp" or "image/x-ms-bmp" => ImageBinaryFormat.Bmp,
            "image/tiff" => ImageBinaryFormat.Tiff,
            "image/webp" => ImageBinaryFormat.Webp,
            _ => ImageBinaryFormat.Unknown,
        };

    private static (double? Width, double? Height) ReadRenderedPoints(XElement? extent)
    {
        if (!long.TryParse((string?)extent?.Attribute("cx"), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var width)
            || !long.TryParse((string?)extent?.Attribute("cy"), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var height)
            || width <= 0 || height <= 0) return (null, null);
        return (width / (double)EmusPerPoint, height / (double)EmusPerPoint);
    }

    private static string ImagePublicId(OwnedPartRelationships.Owner owner, XElement outer,
        int? subOccurrence = null) =>
        $"img:{owner.Scope}:{UnidHelper.ReadOrDeriveUnid(outer)}"
        + (subOccurrence is null ? string.Empty : $":sub{subOccurrence.Value}");

    private EditResult? ResolveMutableImage(string imageId, ImageOperation operation, out ImageCandidate candidate)
    {
        candidate = EnumerateImageCandidates(ProjectionScopes.All)
            .FirstOrDefault(item => string.Equals(item.Info.Id, imageId, StringComparison.Ordinal))!;
        if (candidate is null)
            return EditResult.Fail(EditErrorCode.ImageNotFound, $"image not found: {imageId}");
        var support = operation switch
        {
            ImageOperation.Replace => candidate.Info.Operations.Replace,
            ImageOperation.EmbedLinked => candidate.Info.Operations.EmbedLinked,
            ImageOperation.SetDimensions => candidate.Info.Operations.SetDimensions,
            ImageOperation.SetMetadata => candidate.Info.Operations.SetMetadata,
            ImageOperation.SetFloatingLayout => candidate.Info.Operations.SetFloatingLayout,
            _ => candidate.Info.Operations.Remove,
        };
        if (support.CanMutate) return null;
        return EditResult.Fail(
            operation == ImageOperation.Replace && candidate.Info.IsLinked
                ? EditErrorCode.LinkedImageReadOnly
                : EditErrorCode.UnsupportedImageMarkup,
            support.Reason ?? "image markup is read-only");
    }

    private sealed record ValidatedImageData(
        ImageBinaryFormat Format, string? ContentType, int Width, int Height, EditResult? Error);

    private static ValidatedImageData ValidateImageBytes(byte[]? bytes, string? anchorId)
    {
        if (bytes is null || bytes.Length == 0)
            return new(ImageBinaryFormat.Unknown, null, 0, 0,
                EditResult.Fail(EditErrorCode.InvalidImageData, "image bytes are empty", anchorId));
        if (bytes.LongLength > MaxImageInputBytes)
            return new(ImageBinaryFormat.Unknown, null, 0, 0,
                EditResult.Fail(EditErrorCode.ImageTooLarge,
                    $"image exceeds the {MaxImageInputBytes}-byte runtime limit", anchorId));
        var token = ImageHeaderParser.DetectFormat(bytes);
        var format = FormatFromMagicToken(token);
        if (format == ImageBinaryFormat.Unknown)
            return new(format, null, 0, 0,
                EditResult.Fail(EditErrorCode.UnsupportedImageFormat,
                    "image bytes are not PNG, JPEG, GIF, BMP, TIFF, or WebP", anchorId));
        var dimensions = ImageHeaderParser.GetDimensions(bytes);
        if (dimensions is null)
            return new(format, null, 0, 0,
                EditResult.Fail(EditErrorCode.InvalidImageData,
                    "image header is malformed or has unreadable dimensions", anchorId));
        var contentType = format switch
        {
            ImageBinaryFormat.Png => "image/png",
            ImageBinaryFormat.Jpeg => "image/jpeg",
            ImageBinaryFormat.Gif => "image/gif",
            ImageBinaryFormat.Bmp => "image/bmp",
            ImageBinaryFormat.Tiff => "image/tiff",
            ImageBinaryFormat.Webp => "image/webp",
            _ => null,
        };
        return new(format, contentType, dimensions.Value.Width, dimensions.Value.Height, null);
    }

    private static EditResult? ValidateInsertOptions(ImageInsertOptions options,
        int intrinsicWidth, int intrinsicHeight, string? anchorId,
        out long widthEmu, out long heightEmu, out FloatingImageLayout? layout)
    {
        widthEmu = 0;
        heightEmu = 0;
        layout = null;
        if (!ValidXmlAttributeText(options.AltText) || !ValidXmlAttributeText(options.Title))
            return EditResult.Fail(EditErrorCode.InvalidImageData,
                "image metadata contains characters XML attributes cannot represent", anchorId);
        if (!Enum.IsDefined(options.Placement))
            return EditResult.Fail(EditErrorCode.InvalidImageLayout,
                $"unknown image placement: {options.Placement}", anchorId);
        if (options.Placement == ImagePlacement.Inline && options.FloatingLayout is not null)
            return EditResult.Fail(EditErrorCode.InvalidImageLayout,
                "floatingLayout is only valid when placement is floating", anchorId);
        if (options.Placement == ImagePlacement.Floating)
        {
            layout = options.FloatingLayout ?? new FloatingImageLayout();
            if (ValidateFloatingLayout(layout, anchorId) is { } layoutError) return layoutError;
            layout = NormalizeWrapPolygon(layout);
        }
        double defaultWidth = intrinsicWidth * 72.0 / ImageDefaultDpi;
        double defaultHeight = intrinsicHeight * 72.0 / ImageDefaultDpi;
        return ResolveRenderedDimensions(options.WidthPoints, options.HeightPoints,
            options.PreserveAspect, defaultWidth, defaultHeight,
            out widthEmu, out heightEmu, allowDefaults: true, anchorId);
    }

    private static EditResult? ResolveRenderedDimensions(double? widthPoints, double? heightPoints,
        bool preserveAspect, double? currentWidthPoints, double? currentHeightPoints,
        out long widthEmu, out long heightEmu, bool allowDefaults = false, string? anchorId = null)
    {
        widthEmu = 0;
        heightEmu = 0;
        if (currentWidthPoints is null || currentHeightPoints is null
            || !ValidRenderedPoints(currentWidthPoints.Value)
            || !ValidRenderedPoints(currentHeightPoints.Value))
            return EditResult.Fail(EditErrorCode.InvalidImageDimensions,
                "current image dimensions are missing or invalid", anchorId);
        if (widthPoints is not null && !ValidRenderedPoints(widthPoints.Value)
            || heightPoints is not null && !ValidRenderedPoints(heightPoints.Value))
            return EditResult.Fail(EditErrorCode.InvalidImageDimensions,
                $"rendered dimensions must be finite, positive, and no greater than {MaxImageRenderedPoints} points", anchorId);
        if (widthPoints is null && heightPoints is null)
        {
            if (!allowDefaults)
                return EditResult.Fail(EditErrorCode.InvalidImageDimensions,
                    "at least one rendered dimension is required", anchorId);
            widthPoints = currentWidthPoints;
            heightPoints = currentHeightPoints;
        }
        else if (preserveAspect)
        {
            double scale = widthPoints is not null && heightPoints is not null
                ? Math.Min(widthPoints.Value / currentWidthPoints.Value,
                    heightPoints.Value / currentHeightPoints.Value)
                : widthPoints is not null
                    ? widthPoints.Value / currentWidthPoints.Value
                    : heightPoints!.Value / currentHeightPoints.Value;
            widthPoints = currentWidthPoints.Value * scale;
            heightPoints = currentHeightPoints.Value * scale;
        }
        else if (widthPoints is null || heightPoints is null)
        {
            return EditResult.Fail(EditErrorCode.InvalidImageDimensions,
                "widthPoints and heightPoints are both required when preserveAspect is false", anchorId);
        }
        if (!ValidRenderedPoints(widthPoints!.Value) || !ValidRenderedPoints(heightPoints!.Value))
            return EditResult.Fail(EditErrorCode.InvalidImageDimensions,
                "preserve-aspect calculation produced an invalid rendered dimension", anchorId);
        try
        {
            widthEmu = checked((long)Math.Round(widthPoints.Value * EmusPerPoint,
                MidpointRounding.AwayFromZero));
            heightEmu = checked((long)Math.Round(heightPoints.Value * EmusPerPoint,
                MidpointRounding.AwayFromZero));
        }
        catch (OverflowException)
        {
            return EditResult.Fail(EditErrorCode.InvalidImageDimensions,
                "rendered dimension overflows DrawingML EMUs", anchorId);
        }
        return widthEmu <= 0 || heightEmu <= 0
            ? EditResult.Fail(EditErrorCode.InvalidImageDimensions,
                "rendered dimensions round to zero EMUs", anchorId)
            : null;
    }

    private static bool ValidRenderedPoints(double points) =>
        double.IsFinite(points) && points > 0 && points <= MaxImageRenderedPoints;

    private static bool ValidXmlAttributeText(string? value)
    {
        if (value is null) return true;
        try { XmlConvert.VerifyXmlChars(value); return true; }
        catch (XmlException) { return false; }
    }

    private static EditResult? ValidateFloatingLayout(FloatingImageLayout layout,
        string? anchorId = null)
    {
        if (!Enum.IsDefined(layout.HorizontalRelativeFrom)
            || !Enum.IsDefined(layout.VerticalRelativeFrom)
            || !Enum.IsDefined(layout.WrapMode)
            || !Enum.IsDefined(layout.WrapSide)
            || layout.HorizontalAlignment is { } horizontal && !Enum.IsDefined(horizontal)
            || layout.VerticalAlignment is { } vertical && !Enum.IsDefined(vertical))
            return EditResult.Fail(EditErrorCode.InvalidImageLayout,
                "floating layout contains an unknown enum value", anchorId);
        if (layout.HorizontalRelativeFrom == ImageHorizontalReference.Unknown
            || layout.VerticalRelativeFrom == ImageVerticalReference.Unknown
            || layout.HorizontalAlignment == ImageHorizontalAlignment.Unknown
            || layout.VerticalAlignment == ImageVerticalAlignment.Unknown
            || layout.WrapMode == ImageWrapMode.Unknown
            || layout.WrapSide == ImageWrapSide.Unknown
            || layout.RawHorizontalReference is not null || layout.RawVerticalReference is not null
            || layout.RawHorizontalPosition is not null || layout.RawVerticalPosition is not null
            || layout.RawWrapMode is not null || layout.RawWrapSide is not null
            || layout.RawRelativeSizeHorizontal is not null || layout.RawRelativeSizeVertical is not null
            || layout.RawFlagTokens is not null)
            return EditResult.Fail(EditErrorCode.InvalidImageLayout,
                "report-only raw floating layout tokens cannot be written", anchorId);
        if (layout.WrapMode is ImageWrapMode.Tight or ImageWrapMode.Through)
        {
            if (layout.WrapPolygon is { } polygon && polygon.Points.Count < 3)
                return EditResult.Fail(EditErrorCode.InvalidImageLayout,
                    "a wrap polygon needs a start vertex and at least two line segments", anchorId);
        }
        else if (layout.WrapPolygon is not null)
            return EditResult.Fail(EditErrorCode.InvalidImageLayout,
                "wrapPolygon applies only to tight and through wrap", anchorId);
        if ((layout.HorizontalOffsetEmu is null) == (layout.HorizontalAlignment is null)
            || (layout.VerticalOffsetEmu is null) == (layout.VerticalAlignment is null))
            return EditResult.Fail(EditErrorCode.InvalidImageLayout,
                "each floating position axis requires exactly one offset or alignment", anchorId);
        long max = checked((long)(MaxImageRenderedPoints * EmusPerPoint));
        if (layout.HorizontalOffsetEmu is { } x && Math.Abs((decimal)x) > max
            || layout.VerticalOffsetEmu is { } y && Math.Abs((decimal)y) > max)
            return EditResult.Fail(EditErrorCode.InvalidImageLayout,
                "floating position offset exceeds the runtime EMU limit", anchorId);
        if (layout.DistanceTopEmu < 0 || layout.DistanceBottomEmu < 0
            || layout.DistanceLeftEmu < 0 || layout.DistanceRightEmu < 0
            || layout.DistanceTopEmu > max || layout.DistanceBottomEmu > max
            || layout.DistanceLeftEmu > max || layout.DistanceRightEmu > max)
            return EditResult.Fail(EditErrorCode.InvalidImageLayout,
                "wrap distances must be non-negative and within the runtime EMU limit", anchorId);
        return null;
    }

    /// <summary>A validated tight/through layout without vertices means the picture rectangle;
    /// resolving that here keeps the no-op comparison and the writer looking at the same value.</summary>
    private static FloatingImageLayout NormalizeWrapPolygon(FloatingImageLayout layout) =>
        layout.WrapMode is ImageWrapMode.Tight or ImageWrapMode.Through && layout.WrapPolygon is null
            ? layout with { WrapPolygon = ImageWrapPolygon.Rectangle }
            : layout;

    private static string? ValidateImageInsertionBoundary(XElement paragraph, int offset)
    {
        if (paragraph.Descendants().Any(element => element.Name == W.fldChar || element.Name == W.instrText))
            return "images cannot be inserted into a paragraph containing a complex field";
        int consumed = 0;
        foreach (var child in paragraph.Elements().Where(IsInlineChild))
        {
            int length = string.Concat(child.DescendantsAndSelf(W.t).Select(text => (string)text)).Length;
            if (consumed < offset && offset < consumed + length && child.Name != W.r)
                return "image insertion boundary is inside an unsupported inline container";
            consumed += length;
        }
        return null;
    }

    private static int ImageOffset(XElement paragraph, XElement image)
    {
        int offset = 0;
        foreach (var run in InlineRuns(paragraph))
        {
            if (ReferenceEquals(run, image) || image.AncestorsAndSelf().Contains(run)
                || XNode.DocumentOrderComparer.Compare(run, image) >= 0) break;
            offset += RunText(run).Length;
        }
        return offset;
    }

    private static void InsertInlineElementAtOffset(XElement paragraph, int offset, XElement element)
    {
        SplitRunsAtOffset(paragraph, offset);
        SplitInlineContainersAtOffset(paragraph, offset);
        var map = RunTextMap.Build(paragraph);
        var right = map.Segments.FirstOrDefault(segment => segment.StartOffsetInBlock >= offset).Run;
        if (right is not null)
        {
            var boundary = right.AncestorsAndSelf().First(node => ReferenceEquals(node.Parent, paragraph));
            boundary.AddBeforeSelf(element);
        }
        else paragraph.Add(element);
    }

    private uint NextDocumentPropertyId()
    {
        var used = OwnedPartRelationships.StoryParts(_doc!)
            .SelectMany(owner => owner.Part.GetXDocument().Descendants(WP.docPr))
            .Select(element => uint.TryParse((string?)element.Attribute("id"),
                NumberStyles.None, CultureInfo.InvariantCulture, out var id) ? id : 0)
            .Where(id => id != 0).ToHashSet();
        for (uint id = 1; id < uint.MaxValue; id++) if (!used.Contains(id)) return id;
        throw new InvalidOperationException("no globally available wp:docPr id remains");
    }

    private static XElement BuildImageDrawing(string relationshipId, uint docPrId,
        long widthEmu, long heightEmu, string? altText, string? title,
        ImagePlacement placement, FloatingImageLayout? layout)
    {
        var docPr = new XElement(WP.docPr,
            new XAttribute("id", docPrId),
            new XAttribute("name", $"Picture {docPrId}"));
        if (altText is not null) docPr.Add(new XAttribute("descr", altText));
        if (title is not null) docPr.Add(new XAttribute("title", title));
        var cNvPr = new XElement(Pic.cNvPr,
            new XAttribute("id", docPrId),
            new XAttribute("name", $"Picture {docPrId}"));
        if (altText is not null) cNvPr.Add(new XAttribute("descr", altText));
        if (title is not null) cNvPr.Add(new XAttribute("title", title));
        var graphic = new XElement(A.graphic,
            new XElement(A.graphicData,
                new XAttribute("uri", PictureGraphicDataUri),
                new XElement(Pic._pic,
                    new XElement(Pic.nvPicPr,
                        cNvPr,
                        new XElement(Pic.cNvPicPr,
                            new XElement(A.picLocks, new XAttribute("noChangeAspect", 1)))),
                    new XElement(Pic.blipFill,
                        new XElement(A.blip, new XAttribute(ImageR + "embed", relationshipId)),
                        new XElement(A.stretch, new XElement(A.fillRect))),
                    new XElement(Pic.spPr,
                        new XElement(A.xfrm,
                            new XElement(A.off, new XAttribute("x", 0), new XAttribute("y", 0)),
                            new XElement(A.ext, new XAttribute("cx", widthEmu), new XAttribute("cy", heightEmu))),
                        new XElement(A.prstGeom, new XAttribute("prst", "rect"), new XElement(A.avLst))))));
        XElement container;
        if (placement == ImagePlacement.Inline)
        {
            container = new XElement(WP.inline,
                new XAttribute("distT", 0), new XAttribute("distB", 0),
                new XAttribute("distL", 0), new XAttribute("distR", 0),
                new XElement(WP.extent, new XAttribute("cx", widthEmu), new XAttribute("cy", heightEmu)),
                new XElement(WP.effectExtent, new XAttribute("l", 0), new XAttribute("t", 0),
                    new XAttribute("r", 0), new XAttribute("b", 0)),
                docPr,
                new XElement(WP.cNvGraphicFramePr,
                    new XElement(A.graphicFrameLocks, new XAttribute("noChangeAspect", 1))),
                graphic);
        }
        else
        {
            layout ??= new FloatingImageLayout();
            container = new XElement(WP.anchor,
                new XElement(WP.simplePos, new XAttribute("x", 0), new XAttribute("y", 0)),
                BuildHorizontalPosition(layout), BuildVerticalPosition(layout),
                new XElement(WP.extent, new XAttribute("cx", widthEmu), new XAttribute("cy", heightEmu)),
                new XElement(WP.effectExtent, new XAttribute("l", 0), new XAttribute("t", 0),
                    new XAttribute("r", 0), new XAttribute("b", 0)),
                BuildWrap(layout), docPr,
                new XElement(WP.cNvGraphicFramePr,
                    new XElement(A.graphicFrameLocks, new XAttribute("noChangeAspect", 1))),
                graphic);
            ApplyFloatingAttributes(container, layout);
        }
        return new XElement(W.drawing, container);
    }

    private static void SetDrawingExtents(XElement container, long widthEmu, long heightEmu)
    {
        var extent = container.Element(WP.extent)
            ?? throw new InvalidDataException("drawing has no wp:extent");
        extent.SetAttributeValue("cx", widthEmu);
        extent.SetAttributeValue("cy", heightEmu);
        var transformExtent = container.Descendants(A.xfrm).FirstOrDefault()?.Element(A.ext)
            ?? throw new InvalidDataException("picture has no a:xfrm/a:ext");
        transformExtent.SetAttributeValue("cx", widthEmu);
        transformExtent.SetAttributeValue("cy", heightEmu);
    }

    private static void ApplyFloatingLayout(XElement anchor, FloatingImageLayout layout)
    {
        var positionH = anchor.Element(WP.positionH)
            ?? throw new InvalidDataException("floating image has no horizontal position");
        var positionV = anchor.Element(WP.positionV)
            ?? throw new InvalidDataException("floating image has no vertical position");
        positionH.ReplaceWith(BuildHorizontalPosition(layout));
        positionV.ReplaceWith(BuildVerticalPosition(layout));
        var wrap = anchor.Elements().FirstOrDefault(element => IsWrapElement(element.Name))
            ?? throw new InvalidDataException("floating image has no wrap element");
        wrap.ReplaceWith(BuildWrap(layout));
        ApplyFloatingAttributes(anchor, layout);
    }

    private static bool IsWrapElement(XName name) =>
        name == WP.wrapNone || name == WP.wrapSquare || name == WP.wrapTight
        || name == WP.wrapThrough || name == WP.wrapTopAndBottom;

    private static void ApplyFloatingAttributes(XElement anchor, FloatingImageLayout layout)
    {
        anchor.SetAttributeValue("distT", layout.DistanceTopEmu);
        anchor.SetAttributeValue("distB", layout.DistanceBottomEmu);
        anchor.SetAttributeValue("distL", layout.DistanceLeftEmu);
        anchor.SetAttributeValue("distR", layout.DistanceRightEmu);
        anchor.SetAttributeValue("simplePos", 0);
        anchor.SetAttributeValue("relativeHeight", layout.RelativeHeight);
        anchor.SetAttributeValue("behindDoc", BoolToken(layout.BehindDocument));
        anchor.SetAttributeValue("locked", BoolToken(layout.Locked));
        anchor.SetAttributeValue("layoutInCell", BoolToken(layout.LayoutInCell));
        anchor.SetAttributeValue("allowOverlap", BoolToken(layout.AllowOverlap));
    }

    private static XElement BuildHorizontalPosition(FloatingImageLayout layout)
    {
        var position = new XElement(WP.positionH,
            new XAttribute("relativeFrom", HorizontalReferenceToken(layout.HorizontalRelativeFrom)));
        if (layout.HorizontalAlignment is { } alignment)
            position.Add(new XElement(WP.align, HorizontalAlignmentToken(alignment)));
        else position.Add(new XElement(WP.posOffset, layout.HorizontalOffsetEmu!.Value));
        return position;
    }

    private static XElement BuildVerticalPosition(FloatingImageLayout layout)
    {
        var position = new XElement(WP.positionV,
            new XAttribute("relativeFrom", VerticalReferenceToken(layout.VerticalRelativeFrom)));
        if (layout.VerticalAlignment is { } alignment)
            position.Add(new XElement(WP.align, VerticalAlignmentToken(alignment)));
        else position.Add(new XElement(WP.posOffset, layout.VerticalOffsetEmu!.Value));
        return position;
    }

    private static XElement BuildWrap(FloatingImageLayout layout) => layout.WrapMode switch
    {
        ImageWrapMode.None => new XElement(WP.wrapNone),
        ImageWrapMode.Square => new XElement(WP.wrapSquare,
            new XAttribute("wrapText", WrapSideToken(layout.WrapSide))),
        ImageWrapMode.TopAndBottom => new XElement(WP.wrapTopAndBottom),
        _ => new XElement(layout.WrapMode == ImageWrapMode.Tight ? WP.wrapTight : WP.wrapThrough,
            new XAttribute("wrapText", WrapSideToken(layout.WrapSide)),
            BuildWrapPolygon(layout.WrapPolygon ?? ImageWrapPolygon.Rectangle)),
    };

    private static XElement BuildWrapPolygon(ImageWrapPolygon polygon)
    {
        var element = new XElement(WP.wrapPolygon, new XAttribute("edited", BoolToken(polygon.Edited)));
        for (int i = 0; i < polygon.Points.Count; i++)
            element.Add(new XElement(i == 0 ? WP.start : WP.lineTo,
                new XAttribute("x", polygon.Points[i].X), new XAttribute("y", polygon.Points[i].Y)));
        return element;
    }

    private static bool TryReadWrapPolygon(XElement wrap, out ImageWrapPolygon? polygon)
    {
        polygon = null;
        if (!HasOnlyAttributes(wrap, "wrapText") || wrap.Nodes().Any(node => node is not XElement)) return false;
        var outline = wrap.Elements().SingleOrDefault();
        if (outline is null || outline.Name != WP.wrapPolygon || !HasOnlyAttributes(outline, "edited")
            || !TryBoolAttribute(outline, "edited", false, out var edited)
            || outline.Nodes().Any(node => node is not XElement)) return false;
        var vertices = outline.Elements().ToList();
        if (vertices.Count < 3 || vertices[0].Name != WP.start
            || vertices.Skip(1).Any(vertex => vertex.Name != WP.lineTo)) return false;
        var points = new List<ImageWrapPoint>(vertices.Count);
        foreach (var vertex in vertices)
        {
            if (!HasOnlyAttributes(vertex, "x", "y") || vertex.Nodes().Any()
                || !long.TryParse((string?)vertex.Attribute("x"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var x)
                || !long.TryParse((string?)vertex.Attribute("y"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var y))
                return false;
            points.Add(new ImageWrapPoint(x, y));
        }
        polygon = new ImageWrapPolygon(points, edited);
        return true;
    }

    private static bool TryReadFloatingLayout(XElement anchor,
        out FloatingImageLayout? layout, out string? unsupportedReason)
    {
        layout = null;
        unsupportedReason = null;
        bool mutable = true;
        var rawTokens = new Dictionary<string, string>(StringComparer.Ordinal);
        string? rawRelativeSizeHorizontal = null;
        string? rawRelativeSizeVertical = null;
        if (!TryBoolAttribute(anchor, "simplePos", false, out var usesSimplePosition))
        {
            rawTokens["simplePos"] = (string)anchor.Attribute("simplePos")!;
            mutable = false;
            unsupportedReason = "floating layout has a malformed simplePos flag";
        }
        else if (usesSimplePosition)
        {
            rawTokens["simplePos"] = (string?)anchor.Attribute("simplePos") ?? "1";
            mutable = false;
            unsupportedReason = "floating layout uses the read-only simplePos coordinate system";
        }
        var sizeRelH = anchor.Element(WP14.sizeRelH);
        var sizeRelV = anchor.Element(WP14.sizeRelV);
        if (sizeRelH is not null)
        {
            rawRelativeSizeHorizontal = sizeRelH.ToString(SaveOptions.DisableFormatting);
            mutable = false;
            unsupportedReason ??= "relative percentage sizing is read-only";
        }
        if (sizeRelV is not null)
        {
            rawRelativeSizeVertical = sizeRelV.ToString(SaveOptions.DisableFormatting);
            mutable = false;
            unsupportedReason ??= "relative percentage sizing is read-only";
        }
        var positionH = anchor.Element(WP.positionH);
        var positionV = anchor.Element(WP.positionV);
        var wrapElements = anchor.Elements().Where(element => IsWrapElement(element.Name)).ToList();
        if (positionH is null || positionV is null || wrapElements.Count != 1)
        {
            unsupportedReason = "floating image lacks one canonical position/wrap layout";
            return false;
        }
        var horizontalReferenceToken = (string?)positionH.Attribute("relativeFrom");
        var verticalReferenceToken = (string?)positionV.Attribute("relativeFrom");
        string? rawHorizontalReference = null;
        string? rawVerticalReference = null;
        if (!TryParseHorizontalReference(horizontalReferenceToken, out var horizontalRef))
        {
            horizontalRef = ImageHorizontalReference.Unknown;
            rawHorizontalReference = horizontalReferenceToken;
            mutable = false;
            unsupportedReason = "floating image uses an unsupported horizontal position reference";
        }
        if (!TryParseVerticalReference(verticalReferenceToken, out var verticalRef))
        {
            verticalRef = ImageVerticalReference.Unknown;
            rawVerticalReference = verticalReferenceToken;
            mutable = false;
            unsupportedReason ??= "floating image uses an unsupported vertical position reference";
        }
        string? rawHorizontalPosition = null;
        string? rawVerticalPosition = null;
        if (!TryReadHorizontalPosition(positionH, out var horizontalOffset, out var horizontalAlignment))
        {
            horizontalOffset = null;
            horizontalAlignment = null;
            rawHorizontalPosition = positionH.ToString(SaveOptions.DisableFormatting);
            mutable = false;
            unsupportedReason ??= "floating horizontal position is not one supported offset or alignment";
        }
        if (!TryReadVerticalPosition(positionV, out var verticalOffset, out var verticalAlignment))
        {
            verticalOffset = null;
            verticalAlignment = null;
            rawVerticalPosition = positionV.ToString(SaveOptions.DisableFormatting);
            mutable = false;
            unsupportedReason ??= "floating vertical position is not one supported offset or alignment";
        }
        var wrap = wrapElements[0];
        ImageWrapMode wrapMode;
        ImageWrapSide wrapSide = ImageWrapSide.BothSides;
        ImageWrapPolygon? wrapPolygon = null;
        string? rawWrapMode = null;
        string? rawWrapSide = null;
        if (wrap.Name == WP.wrapNone)
        {
            wrapMode = ImageWrapMode.None;
            if (!HasOnlyAttributes(wrap) || wrap.Nodes().Any())
            {
                rawWrapMode = wrap.ToString(SaveOptions.DisableFormatting);
                mutable = false;
                unsupportedReason ??= "floating wrap contains unmodeled attributes or children";
            }
        }
        else
        {
            wrapMode = wrap.Name == WP.wrapSquare ? ImageWrapMode.Square
                : wrap.Name == WP.wrapTight ? ImageWrapMode.Tight
                : wrap.Name == WP.wrapThrough ? ImageWrapMode.Through
                : ImageWrapMode.TopAndBottom;
            var wrapSideToken = (string?)wrap.Attribute("wrapText") ?? "bothSides";
            if (!TryParseWrapSide(wrapSideToken, out wrapSide))
            {
                wrapSide = ImageWrapSide.Unknown;
                rawWrapSide = wrapSideToken;
                mutable = false;
                unsupportedReason ??= "floating image uses an unsupported wrap side";
            }
            // Each wrap form is modeled exactly: square carries only its side, top-and-bottom
            // nothing, and tight/through their side plus a plain start/lineTo outline. Distances or
            // effect extents on the wrap element itself are unmodeled and keep the layout read-only.
            string? shapeReason = wrapMode switch
            {
                ImageWrapMode.Square => HasOnlyAttributes(wrap, "wrapText") && !wrap.Nodes().Any()
                    ? null : "floating wrap contains unmodeled attributes or children",
                ImageWrapMode.TopAndBottom => HasOnlyAttributes(wrap) && !wrap.Nodes().Any()
                    ? null : "floating wrap contains unmodeled attributes or children",
                _ => TryReadWrapPolygon(wrap, out wrapPolygon)
                    ? null : "floating wrap outline is not a plain start/lineTo polygon",
            };
            if (shapeReason is not null)
            {
                rawWrapMode = wrap.ToString(SaveOptions.DisableFormatting);
                mutable = false;
                unsupportedReason ??= shapeReason;
            }
        }
        if (!TryLongAttribute(anchor, "distT", 0, out var distT))
        {
            rawTokens["distT"] = (string)anchor.Attribute("distT")!; mutable = false; distT = 0;
        }
        if (!TryLongAttribute(anchor, "distB", 0, out var distB))
        { rawTokens["distB"] = (string)anchor.Attribute("distB")!; mutable = false; distB = 0; }
        if (!TryLongAttribute(anchor, "distL", 0, out var distL))
        { rawTokens["distL"] = (string)anchor.Attribute("distL")!; mutable = false; distL = 0; }
        if (!TryLongAttribute(anchor, "distR", 0, out var distR))
        { rawTokens["distR"] = (string)anchor.Attribute("distR")!; mutable = false; distR = 0; }
        if (!TryUIntAttribute(anchor, "relativeHeight", 0, out var relativeHeight))
        { rawTokens["relativeHeight"] = (string)anchor.Attribute("relativeHeight")!; mutable = false; relativeHeight = 0; }
        if (!TryBoolAttribute(anchor, "behindDoc", false, out var behindDocument))
        { rawTokens["behindDoc"] = (string)anchor.Attribute("behindDoc")!; mutable = false; behindDocument = false; }
        if (!TryBoolAttribute(anchor, "locked", false, out var locked))
        { rawTokens["locked"] = (string)anchor.Attribute("locked")!; mutable = false; locked = false; }
        if (!TryBoolAttribute(anchor, "layoutInCell", true, out var layoutInCell))
        { rawTokens["layoutInCell"] = (string)anchor.Attribute("layoutInCell")!; mutable = false; layoutInCell = true; }
        if (!TryBoolAttribute(anchor, "allowOverlap", true, out var allowOverlap))
        { rawTokens["allowOverlap"] = (string)anchor.Attribute("allowOverlap")!; mutable = false; allowOverlap = true; }
        if (rawTokens.Count != 0)
            unsupportedReason ??= "floating layout contains malformed numeric or boolean attributes";
        layout = new FloatingImageLayout
        {
            HorizontalRelativeFrom = horizontalRef,
            HorizontalOffsetEmu = horizontalOffset,
            HorizontalAlignment = horizontalAlignment,
            VerticalRelativeFrom = verticalRef,
            VerticalOffsetEmu = verticalOffset,
            VerticalAlignment = verticalAlignment,
            WrapMode = wrapMode,
            WrapSide = wrapSide,
            WrapPolygon = wrapPolygon,
            DistanceTopEmu = distT,
            DistanceBottomEmu = distB,
            DistanceLeftEmu = distL,
            DistanceRightEmu = distR,
            RelativeHeight = relativeHeight,
            BehindDocument = behindDocument,
            Locked = locked,
            LayoutInCell = layoutInCell,
            AllowOverlap = allowOverlap,
            RawHorizontalReference = rawHorizontalReference,
            RawVerticalReference = rawVerticalReference,
            RawHorizontalPosition = rawHorizontalPosition,
            RawVerticalPosition = rawVerticalPosition,
            RawWrapMode = rawWrapMode,
            RawWrapSide = rawWrapSide,
            RawRelativeSizeHorizontal = rawRelativeSizeHorizontal,
            RawRelativeSizeVertical = rawRelativeSizeVertical,
            RawFlagTokens = rawTokens.Count == 0 ? null : rawTokens,
        };
        if (mutable && ValidateFloatingLayout(layout) is { } error)
        {
            unsupportedReason = error.Error?.Message;
            mutable = false;
        }
        return mutable;
    }

    private static bool TryReadHorizontalPosition(XElement position,
        out long? offset, out ImageHorizontalAlignment? alignment)
    {
        offset = null;
        alignment = null;
        if (position.Attribute("relativeFrom") is null
            || !HasOnlyAttributes(position, "relativeFrom")
            || position.Elements().Count() != 1) return false;
        var offsetElements = position.Elements(WP.posOffset).ToList();
        var alignElements = position.Elements(WP.align).ToList();
        if (offsetElements.Count + alignElements.Count != 1) return false;
        var offsetElement = offsetElements.SingleOrDefault();
        var alignElement = alignElements.SingleOrDefault();
        var valueElement = offsetElement ?? alignElement!;
        if (!HasOnlyAttributes(valueElement)
            || valueElement.Nodes().Any(node => node is not XText)) return false;
        if (offsetElement is not null)
            return long.TryParse(offsetElement.Value, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var value) && Assign(value, out offset);
        return TryParseHorizontalAlignment(alignElement!.Value, out alignment);
    }

    private static bool TryReadVerticalPosition(XElement position,
        out long? offset, out ImageVerticalAlignment? alignment)
    {
        offset = null;
        alignment = null;
        if (position.Attribute("relativeFrom") is null
            || !HasOnlyAttributes(position, "relativeFrom")
            || position.Elements().Count() != 1) return false;
        var offsetElements = position.Elements(WP.posOffset).ToList();
        var alignElements = position.Elements(WP.align).ToList();
        if (offsetElements.Count + alignElements.Count != 1) return false;
        var offsetElement = offsetElements.SingleOrDefault();
        var alignElement = alignElements.SingleOrDefault();
        var valueElement = offsetElement ?? alignElement!;
        if (!HasOnlyAttributes(valueElement)
            || valueElement.Nodes().Any(node => node is not XText)) return false;
        if (offsetElement is not null)
            return long.TryParse(offsetElement.Value, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out var value) && Assign(value, out offset);
        return TryParseVerticalAlignment(alignElement!.Value, out alignment);
    }

    private static bool HasOnlyAttributes(XElement element, params XName[] allowed) =>
        element.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration
                && attribute.Name != PtOpenXml.Unid)
            .All(attribute => allowed.Contains(attribute.Name));

    private static bool Assign(long value, out long? target) { target = value; return true; }

    private static bool TryLongAttribute(XElement element, XName name, long fallback, out long value)
    {
        var token = (string?)element.Attribute(name);
        if (token is null) { value = fallback; return true; }
        return long.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
    }

    private static bool TryUIntAttribute(XElement element, XName name, uint fallback, out uint value)
    {
        var token = (string?)element.Attribute(name);
        if (token is null) { value = fallback; return true; }
        return uint.TryParse(token, NumberStyles.None, CultureInfo.InvariantCulture, out value);
    }

    private static bool TryBoolAttribute(XElement element, XName name, bool fallback, out bool value)
    {
        var token = (string?)element.Attribute(name);
        if (token is null) { value = fallback; return true; }
        switch (token)
        {
            case "1": case "true": case "on": value = true; return true;
            case "0": case "false": case "off": value = false; return true;
            default: value = fallback; return false;
        }
    }

    private static int BoolToken(bool value) => value ? 1 : 0;

    private static string HorizontalReferenceToken(ImageHorizontalReference value) => value switch
    {
        ImageHorizontalReference.Page => "page",
        ImageHorizontalReference.Margin => "margin",
        ImageHorizontalReference.Column => "column",
        ImageHorizontalReference.Character => "character",
        _ => throw new ArgumentOutOfRangeException(nameof(value)),
    };

    private static string VerticalReferenceToken(ImageVerticalReference value) => value switch
    {
        ImageVerticalReference.Page => "page",
        ImageVerticalReference.Margin => "margin",
        ImageVerticalReference.Paragraph => "paragraph",
        ImageVerticalReference.Line => "line",
        _ => throw new ArgumentOutOfRangeException(nameof(value)),
    };

    private static string HorizontalAlignmentToken(ImageHorizontalAlignment value) => value switch
    {
        ImageHorizontalAlignment.Left => "left",
        ImageHorizontalAlignment.Center => "center",
        ImageHorizontalAlignment.Right => "right",
        ImageHorizontalAlignment.Inside => "inside",
        ImageHorizontalAlignment.Outside => "outside",
        _ => throw new ArgumentOutOfRangeException(nameof(value)),
    };

    private static string VerticalAlignmentToken(ImageVerticalAlignment value) => value switch
    {
        ImageVerticalAlignment.Top => "top",
        ImageVerticalAlignment.Center => "center",
        ImageVerticalAlignment.Bottom => "bottom",
        ImageVerticalAlignment.Inside => "inside",
        ImageVerticalAlignment.Outside => "outside",
        _ => throw new ArgumentOutOfRangeException(nameof(value)),
    };

    private static string WrapSideToken(ImageWrapSide value) => value switch
    {
        ImageWrapSide.BothSides => "bothSides",
        ImageWrapSide.Left => "left",
        ImageWrapSide.Right => "right",
        ImageWrapSide.Largest => "largest",
        _ => throw new ArgumentOutOfRangeException(nameof(value)),
    };

    private static bool TryParseHorizontalReference(string? token, out ImageHorizontalReference value) =>
        EnumTry(token, out value, ("page", ImageHorizontalReference.Page),
            ("margin", ImageHorizontalReference.Margin), ("column", ImageHorizontalReference.Column),
            ("character", ImageHorizontalReference.Character));

    private static bool TryParseVerticalReference(string? token, out ImageVerticalReference value) =>
        EnumTry(token, out value, ("page", ImageVerticalReference.Page),
            ("margin", ImageVerticalReference.Margin), ("paragraph", ImageVerticalReference.Paragraph),
            ("line", ImageVerticalReference.Line));

    private static bool TryParseHorizontalAlignment(string? token, out ImageHorizontalAlignment? value)
    {
        if (EnumTry(token, out ImageHorizontalAlignment parsed,
            ("left", ImageHorizontalAlignment.Left), ("center", ImageHorizontalAlignment.Center),
            ("right", ImageHorizontalAlignment.Right), ("inside", ImageHorizontalAlignment.Inside),
            ("outside", ImageHorizontalAlignment.Outside))) { value = parsed; return true; }
        value = null;
        return false;
    }

    private static bool TryParseVerticalAlignment(string? token, out ImageVerticalAlignment? value)
    {
        if (EnumTry(token, out ImageVerticalAlignment parsed,
            ("top", ImageVerticalAlignment.Top), ("center", ImageVerticalAlignment.Center),
            ("bottom", ImageVerticalAlignment.Bottom), ("inside", ImageVerticalAlignment.Inside),
            ("outside", ImageVerticalAlignment.Outside))) { value = parsed; return true; }
        value = null;
        return false;
    }

    private static bool TryParseWrapSide(string? token, out ImageWrapSide value) =>
        EnumTry(token, out value, ("bothSides", ImageWrapSide.BothSides),
            ("left", ImageWrapSide.Left), ("right", ImageWrapSide.Right),
            ("largest", ImageWrapSide.Largest));

    private static bool EnumTry<T>(string? token, out T value, params (string Token, T Value)[] values)
        where T : struct
    {
        foreach (var candidate in values)
        {
            if (!string.Equals(token, candidate.Token, StringComparison.Ordinal)) continue;
            value = candidate.Value;
            return true;
        }
        value = default;
        return false;
    }

    private EditResult ImageMutationSuccess(ImageCandidate candidate, string imageId)
    {
        var paragraph = candidate.Outer.Ancestors(W.p).FirstOrDefault();
        var anchor = paragraph is null ? null : AnchorForElement(paragraph);
        return new EditResult { Success = true, ImageId = imageId,
            Modified = anchor is null ? Array.Empty<Anchor>() : new[] { anchor.Value } };
    }
}
