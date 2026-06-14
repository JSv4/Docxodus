#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Options for <see cref="HtmlConversionOps"/>. Mirrors the parameter set of the
/// WASM <c>DocumentConverter.ConvertDocxToHtmlComplete</c> shell so every surface
/// renders identically. Integer-coded modes match the existing WASM wire contract:
/// CommentRenderMode -1=disabled,0=Endnote,1=Inline,2=Margin;
/// PaginationMode 0=None,1=Paginated; AnnotationLabelMode 0=Above,1=Inline,2=Tooltip,3=None.
/// </summary>
internal sealed class HtmlConversionOptions
{
    public string PageTitle { get; init; } = "Document";
    public string CssClassPrefix { get; init; } = "docx-";
    public bool FabricateCssClasses { get; init; } = true;
    public string AdditionalCss { get; init; } = "";
    public int CommentRenderMode { get; init; } = -1;
    public string CommentCssClassPrefix { get; init; } = "comment-";
    public int PaginationMode { get; init; }
    public double PaginationScale { get; init; } = 1.0;
    public string PaginationCssClassPrefix { get; init; } = "page-";
    public bool RenderAnnotations { get; init; }
    public int AnnotationLabelMode { get; init; }
    public string AnnotationCssClassPrefix { get; init; } = "annot-";
    public bool RenderFootnotesAndEndnotes { get; init; }
    public bool RenderHeadersAndFooters { get; init; }
    public bool RenderTrackedChanges { get; init; }
    public bool ShowDeletedContent { get; init; } = true;
    public bool RenderMoveOperations { get; init; } = true;
    public bool RenderUnsupportedContentPlaceholders { get; init; }
    public string? DocumentLanguage { get; init; }

    /// <summary>
    /// When true, assign deterministic content-addressable Unids and stamp
    /// block-level HTML elements with <c>data-anchor</c> so the editor can address
    /// blocks in the DOM. Anchors match the markdown projector / DocxSession.
    /// </summary>
    public bool StampAnchors { get; init; }
}

/// <summary>
/// Single owner of the DOCX-bytes + <see cref="HtmlConversionOptions"/> →
/// HTML-string mapping. Both the WASM <c>DocumentConverter</c> bridge and the
/// stdio Python host route through here, so render behavior lives in one place.
/// Throws on invalid input; callers serialize errors at their boundary.
/// </summary>
internal static class HtmlConversionOps
{
    /// <summary>Render raw DOCX bytes to a self-contained HTML string.</summary>
    public static string ConvertToHtml(byte[] docxBytes, HtmlConversionOptions options)
    {
        if (docxBytes == null || docxBytes.Length == 0)
            throw new ArgumentException("No document data provided", nameof(docxBytes));
        ArgumentNullException.ThrowIfNull(options);

        // Writable stream required: WmlToHtmlConverter runs RevisionAccepter internally.
        using var memoryStream = new MemoryStream();
        memoryStream.Write(docxBytes, 0, docxBytes.Length);
        memoryStream.Position = 0;
        using var wordDoc = WordprocessingDocument.Open(memoryStream, true);

        if (options.StampAnchors)
        {
            // Deterministic, content-addressable Unids — identical to the markdown
            // projector / DocxSession, so editor anchors line up across surfaces.
            UnidHelper.AssignToAllElementsDeterministic(wordDoc.MainDocumentPart!.GetXDocument().Root!);
        }

        var renderComments = options.CommentRenderMode >= 0;

        var settings = new WmlToHtmlConverterSettings
        {
            PageTitle = options.PageTitle,
            CssClassPrefix = options.CssClassPrefix,
            FabricateCssClasses = options.FabricateCssClasses,
            AdditionalCss = options.AdditionalCss,
            GeneralCss = "body { font-family: Arial, sans-serif; margin: 20px; } " +
                         "span { white-space: pre-wrap; }",
            RenderComments = renderComments,
            CommentRenderMode = renderComments
                ? (CommentRenderMode)options.CommentRenderMode
                : CommentRenderMode.EndnoteStyle,
            CommentCssClassPrefix = options.CommentCssClassPrefix,
            IncludeCommentMetadata = true,
            RenderPagination = (PaginationMode)options.PaginationMode,
            PaginationScale = options.PaginationScale > 0 ? options.PaginationScale : 1.0,
            PaginationCssClassPrefix = options.PaginationCssClassPrefix,
            RenderAnnotations = options.RenderAnnotations,
            AnnotationLabelMode = (AnnotationLabelMode)options.AnnotationLabelMode,
            AnnotationCssClassPrefix = options.AnnotationCssClassPrefix,
            IncludeAnnotationMetadata = true,
            RenderFootnotesAndEndnotes = options.RenderFootnotesAndEndnotes,
            RenderHeadersAndFooters = options.RenderHeadersAndFooters,
            RenderTrackedChanges = options.RenderTrackedChanges,
            ShowDeletedContent = options.ShowDeletedContent,
            RenderMoveOperations = options.RenderMoveOperations,
            IncludeRevisionMetadata = true,
            RenderUnsupportedContentPlaceholders = options.RenderUnsupportedContentPlaceholders,
            UnsupportedContentCssClassPrefix = "unsupported-",
            IncludeUnsupportedContentMetadata = true,
            DocumentLanguage = options.DocumentLanguage,
            StampAnchors = options.StampAnchors,
            // Embed images as base64 data URIs — no SkiaSharp needed (WASM-safe).
            ImageHandler = CreateBase64ImageHandler(),
        };

        var htmlElement = WmlToHtmlConverter.ConvertToHtml(wordDoc, settings);
        return htmlElement.ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>Render a live session's current (possibly edited) state to HTML.</summary>
    public static string ConvertToHtml(DocxSession session, HtmlConversionOptions options)
    {
        if (session == null) throw new ArgumentNullException(nameof(session));
        return ConvertToHtml(session.Save(), options);
    }

    /// <summary>Render the session registered under <paramref name="handle"/> to HTML.</summary>
    public static string ConvertToHtml(int handle, HtmlConversionOptions options) =>
        ConvertToHtml(SessionRegistry.Get(handle), options);

    /// <summary>
    /// Render a single block (addressed by a <c>kind:scope:unid</c> anchor) to faithful
    /// HTML. Builds a throwaway document that copies the source's styles/numbering/theme
    /// parts and contains just the one block, then runs the standard converter. The full
    /// document render is the faithfulness oracle — this must match the corresponding
    /// <c>data-anchor</c> element from a full render. Known limits: a list item loses
    /// numbering continuation, and an inline image loses its (uncopied) image part.
    /// </summary>
    public static string RenderBlockHtml(byte[] docxBytes, string anchorId, HtmlConversionOptions options)
    {
        if (docxBytes == null || docxBytes.Length == 0)
            throw new ArgumentException("No document data provided", nameof(docxBytes));
        if (string.IsNullOrWhiteSpace(anchorId))
            throw new ArgumentException("No anchor id provided", nameof(anchorId));
        ArgumentNullException.ThrowIfNull(options);

        using var sourceStream = new MemoryStream();
        sourceStream.Write(docxBytes, 0, docxBytes.Length);
        sourceStream.Position = 0;
        using var sourceDoc = WordprocessingDocument.Open(sourceStream, true);

        // Assign deterministic Unids with the SAME call the full render uses, so the
        // anchor the editor saw in data-anchor resolves here by construction. The anchor
        // id is kind:scope:unid; only the unid tail is the durable handle.
        var sourceRoot = sourceDoc.MainDocumentPart!.GetXDocument().Root!;
        UnidHelper.AssignToAllElementsDeterministic(sourceRoot);

        var unid = anchorId.Substring(anchorId.LastIndexOf(':') + 1);
        var blockElement = sourceRoot.DescendantsAndSelf()
            .FirstOrDefault(e => (string?)e.Attribute(PtOpenXml.Unid) == unid)
            ?? throw new ArgumentException($"anchor not found: {anchorId}", nameof(anchorId));

        // Build a throwaway doc: copied formatting parts + just this block.
        using var blockStream = new MemoryStream();
        using (var blockDoc = WordprocessingDocument.Create(
                   blockStream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = blockDoc.AddMainDocumentPart();
            CopyPartXml(sourceDoc, blockDoc, p => p.StyleDefinitionsPart);
            CopyPartXml(sourceDoc, blockDoc, p => p.StylesWithEffectsPart);
            CopyPartXml(sourceDoc, blockDoc, p => p.NumberingDefinitionsPart);
            CopyPartXml(sourceDoc, blockDoc, p => p.ThemePart);
            CopyPartXml(sourceDoc, blockDoc, p => p.FontTablePart);
            CopyPartXml(sourceDoc, blockDoc, p => p.DocumentSettingsPart);
            // The converter requires a DocumentSettingsPart (CalculateSpanWidthForTabs
            // reads w:defaultTabStop with no null check). Ensure one exists.
            if (blockDoc.MainDocumentPart!.DocumentSettingsPart is null)
            {
                blockDoc.MainDocumentPart.AddNewPart<DocumentSettingsPart>()
                    .PutXDocument(new XDocument(
                        new XElement(W.settings, new XAttribute(XNamespace.Xmlns + "w", W.w))));
            }
            main.PutXDocument(new XDocument(
                new XElement(W.document,
                    new XAttribute(XNamespace.Xmlns + "w", W.w),
                    new XAttribute(XNamespace.Xmlns + "r", R.r),
                    new XElement(W.body, new XElement(blockElement)))));
        }
        blockStream.Position = 0;
        using var renderDoc = WordprocessingDocument.Open(blockStream, true);

        var settings = new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = options.FabricateCssClasses,
            CssClassPrefix = options.CssClassPrefix,
            StampAnchors = true,
        };
        var htmlElement = WmlToHtmlConverter.ConvertToHtml(renderDoc, settings);

        // Return the rendered block element (located by its stamped data-anchor),
        // not the full <html> wrapper.
        XElement? inner = null;
        if (unid != null)
            inner = htmlElement.Descendants().FirstOrDefault(e => (string?)e.Attribute("data-anchor") == unid);
        if (inner == null)
        {
            var body = htmlElement.Descendants().FirstOrDefault(e => e.Name.LocalName == "body");
            inner = body?.Elements().FirstOrDefault() ?? htmlElement;
        }
        return inner.ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>Clone a whole formatting part (styles/numbering/theme/font) from src to dst.</summary>
    private static void CopyPartXml<TPart>(WordprocessingDocument src, WordprocessingDocument dst,
        Func<MainDocumentPart, TPart?> get) where TPart : OpenXmlPart, IFixedContentTypePart
    {
        var srcPart = get(src.MainDocumentPart!);
        if (srcPart is null) return;
        var srcRoot = srcPart.GetXDocument().Root;
        if (srcRoot is null) return;
        var dstPart = dst.MainDocumentPart!.AddNewPart<TPart>();
        dstPart.PutXDocument(new XDocument(new XElement(srcRoot)));
    }

    private static Func<ImageInfo, XElement> CreateBase64ImageHandler()
    {
        return imageInfo =>
        {
            if (imageInfo.ImageBytes == null || imageInfo.ImageBytes.Length == 0)
                return null!;

            var mimeType = imageInfo.ContentType ?? "image/png";
            var base64 = Convert.ToBase64String(imageInfo.ImageBytes);
            var dataUri = $"data:{mimeType};base64,{base64}";

            var imgElement = new XElement(XhtmlNoNamespace.img,
                new XAttribute("src", dataUri));

            if (imageInfo.ImgStyleAttribute != null)
                imgElement.Add(imageInfo.ImgStyleAttribute);

            if (!string.IsNullOrEmpty(imageInfo.AltText))
                imgElement.Add(new XAttribute("alt", imageInfo.AltText));

            return imgElement;
        };
    }
}
