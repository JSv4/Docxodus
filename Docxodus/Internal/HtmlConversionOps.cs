#nullable enable

using System;
using System.IO;
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
