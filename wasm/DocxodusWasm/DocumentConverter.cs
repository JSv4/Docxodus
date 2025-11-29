using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using System.Text.Json;
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
            additionalCss: ""
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
    /// <returns>HTML string or JSON error object</returns>
    [JSExport]
    public static string ConvertDocxToHtmlWithOptions(
        byte[] docxBytes,
        string pageTitle,
        string cssPrefix,
        bool fabricateClasses,
        string additionalCss)
    {
        return ConvertDocxToHtmlAdvanced(
            docxBytes, pageTitle, cssPrefix, fabricateClasses, additionalCss,
            renderComments: false, commentRenderMode: 0, commentCssClassPrefix: "comment-", includeCommentMetadata: true
        );
    }

    /// <summary>
    /// Convert a DOCX file to HTML with advanced settings including comment rendering.
    /// </summary>
    /// <param name="docxBytes">The DOCX file as a byte array</param>
    /// <param name="pageTitle">Title for the HTML document</param>
    /// <param name="cssPrefix">Prefix for generated CSS class names</param>
    /// <param name="fabricateClasses">Whether to generate CSS classes</param>
    /// <param name="additionalCss">Additional CSS to include</param>
    /// <param name="renderComments">Whether to render document comments</param>
    /// <param name="commentRenderMode">Comment render mode: 0=EndnoteStyle, 1=Inline, 2=Margin</param>
    /// <param name="commentCssClassPrefix">CSS class prefix for comments (default: "comment-")</param>
    /// <param name="includeCommentMetadata">Include author/date metadata in comment HTML</param>
    /// <returns>HTML string or JSON error object</returns>
    [JSExport]
    public static string ConvertDocxToHtmlAdvanced(
        byte[] docxBytes,
        string pageTitle,
        string cssPrefix,
        bool fabricateClasses,
        string additionalCss,
        bool renderComments,
        int commentRenderMode,
        string commentCssClassPrefix,
        bool includeCommentMetadata)
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

            var settings = new WmlToHtmlConverterSettings
            {
                PageTitle = pageTitle ?? "Document",
                CssClassPrefix = cssPrefix ?? "docx-",
                FabricateCssClasses = fabricateClasses,
                AdditionalCss = additionalCss ?? "",
                GeneralCss = "body { font-family: Arial, sans-serif; margin: 20px; } " +
                             "span { white-space: pre-wrap; }",
                RenderComments = renderComments,
                CommentRenderMode = (CommentRenderMode)commentRenderMode,
                CommentCssClassPrefix = commentCssClassPrefix ?? "comment-",
                IncludeCommentMetadata = includeCommentMetadata
            };

            var htmlElement = WmlToHtmlConverter.ConvertToHtml(wordDoc, settings);
            return htmlElement.ToString();
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
