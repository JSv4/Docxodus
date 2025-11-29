using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using System.Text.Json;
using System.Xml.Linq;
using Docxodus;
using DocumentFormat.OpenXml.Packaging;

namespace DocxodusWasm;

/// <summary>
/// JSExport methods for DOCX to HTML conversion.
/// These methods are callable from JavaScript.
/// </summary>
[SupportedOSPlatform("browser")]
public partial class DocumentConverter
{
    /// <summary>
    /// Convert a DOCX file to HTML with default settings.
    /// </summary>
    /// <param name="docxBytes">The DOCX file as a byte array (from JavaScript Uint8Array)</param>
    /// <returns>HTML string or JSON error object</returns>
    [JSExport]
    public static string ConvertDocxToHtml(byte[] docxBytes)
    {
        return ConvertDocxToHtmlWithOptions(
            docxBytes,
            pageTitle: "Document",
            cssPrefix: "docx-",
            fabricateClasses: true,
            additionalCss: "",
            commentRenderMode: -1,  // -1 = don't render comments
            commentCssClassPrefix: "comment-"
        );
    }

    /// <summary>
    /// Convert a DOCX file to HTML with custom settings.
    /// </summary>
    /// <param name="docxBytes">The DOCX file as a byte array</param>
    /// <param name="pageTitle">Title for the HTML document</param>
    /// <param name="cssPrefix">Prefix for generated CSS class names</param>
    /// <param name="fabricateClasses">Whether to generate CSS classes</param>
    /// <param name="additionalCss">Additional CSS to include</param>
    /// <param name="commentRenderMode">Comment render mode: -1=disabled, 0=EndnoteStyle, 1=Inline, 2=Margin</param>
    /// <param name="commentCssClassPrefix">CSS class prefix for comments (default: "comment-")</param>
    /// <returns>HTML string or JSON error object</returns>
    [JSExport]
    public static string ConvertDocxToHtmlWithOptions(
        byte[] docxBytes,
        string pageTitle,
        string cssPrefix,
        bool fabricateClasses,
        string additionalCss,
        int commentRenderMode,
        string commentCssClassPrefix)
    {
        // Delegate to the pagination-aware version with pagination disabled
        return ConvertDocxToHtmlWithPagination(
            docxBytes,
            pageTitle,
            cssPrefix,
            fabricateClasses,
            additionalCss,
            commentRenderMode,
            commentCssClassPrefix,
            paginationMode: 0,  // None
            paginationScale: 1.0,
            paginationCssClassPrefix: "page-"
        );
    }

    /// <summary>
    /// Convert a DOCX file to HTML with pagination support.
    /// </summary>
    /// <param name="docxBytes">The DOCX file as a byte array</param>
    /// <param name="pageTitle">Title for the HTML document</param>
    /// <param name="cssPrefix">Prefix for generated CSS class names</param>
    /// <param name="fabricateClasses">Whether to generate CSS classes</param>
    /// <param name="additionalCss">Additional CSS to include</param>
    /// <param name="commentRenderMode">Comment render mode: -1=disabled, 0=EndnoteStyle, 1=Inline, 2=Margin</param>
    /// <param name="commentCssClassPrefix">CSS class prefix for comments (default: "comment-")</param>
    /// <param name="paginationMode">Pagination mode: 0=None, 1=Paginated</param>
    /// <param name="paginationScale">Scale factor for page rendering (1.0 = 100%)</param>
    /// <param name="paginationCssClassPrefix">CSS class prefix for pagination elements (default: "page-")</param>
    /// <returns>HTML string or JSON error object</returns>
    [JSExport]
    public static string ConvertDocxToHtmlWithPagination(
        byte[] docxBytes,
        string pageTitle,
        string cssPrefix,
        bool fabricateClasses,
        string additionalCss,
        int commentRenderMode,
        string commentCssClassPrefix,
        int paginationMode,
        double paginationScale,
        string paginationCssClassPrefix)
    {
        if (docxBytes == null || docxBytes.Length == 0)
        {
            return SerializeError("No document data provided");
        }

        try
        {
            // Must use writable stream - WmlToHtmlConverter calls RevisionAccepter internally
            using var memoryStream = new MemoryStream();
            memoryStream.Write(docxBytes, 0, docxBytes.Length);
            memoryStream.Position = 0;
            using var wordDoc = WordprocessingDocument.Open(memoryStream, true);

            var renderComments = commentRenderMode >= 0;

            var settings = new WmlToHtmlConverterSettings
            {
                PageTitle = pageTitle ?? "Document",
                CssClassPrefix = cssPrefix ?? "docx-",
                FabricateCssClasses = fabricateClasses,
                AdditionalCss = additionalCss ?? "",
                GeneralCss = "body { font-family: Arial, sans-serif; margin: 20px; } " +
                             "span { white-space: pre-wrap; }",
                RenderComments = renderComments,
                CommentRenderMode = renderComments ? (CommentRenderMode)commentRenderMode : CommentRenderMode.EndnoteStyle,
                CommentCssClassPrefix = commentCssClassPrefix ?? "comment-",
                IncludeCommentMetadata = true,
                RenderPagination = (PaginationMode)paginationMode,
                PaginationScale = paginationScale > 0 ? paginationScale : 1.0,
                PaginationCssClassPrefix = paginationCssClassPrefix ?? "page-"
            };

            var htmlElement = WmlToHtmlConverter.ConvertToHtml(wordDoc, settings);
            return htmlElement.ToString(SaveOptions.DisableFormatting);
        }
        catch (Exception ex)
        {
            return SerializeError(ex.Message, ex.GetType().Name, ex.StackTrace);
        }
    }

    /// <summary>
    /// Get library version information.
    /// </summary>
    [JSExport]
    public static string GetVersion()
    {
        var info = new VersionInfo
        {
            Library = "Docxodus WASM",
            DotnetVersion = Environment.Version.ToString(),
            Platform = "browser-wasm"
        };
        return JsonSerializer.Serialize(info, DocxodusJsonContext.Default.VersionInfo);
    }

    internal static string SerializeError(string error, string? type = null, string? stackTrace = null)
    {
        var response = new ErrorResponse
        {
            Error = error,
            Type = type,
            StackTrace = stackTrace
        };
        return JsonSerializer.Serialize(response, DocxodusJsonContext.Default.ErrorResponse);
    }
}
