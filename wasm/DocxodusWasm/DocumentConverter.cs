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
        // Delegate to full version with annotations disabled
        return ConvertDocxToHtmlFull(
            docxBytes,
            pageTitle,
            cssPrefix,
            fabricateClasses,
            additionalCss,
            commentRenderMode,
            commentCssClassPrefix,
            paginationMode,
            paginationScale,
            paginationCssClassPrefix,
            renderAnnotations: false,
            annotationLabelMode: 0,
            annotationCssClassPrefix: "annot-"
        );
    }

    /// <summary>
    /// Convert a DOCX file to HTML with full options including annotations.
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
    /// <param name="renderAnnotations">Whether to render custom annotations</param>
    /// <param name="annotationLabelMode">Annotation label mode: 0=Above, 1=Inline, 2=Tooltip, 3=None</param>
    /// <param name="annotationCssClassPrefix">CSS class prefix for annotations (default: "annot-")</param>
    /// <returns>HTML string or JSON error object</returns>
    [JSExport]
    public static string ConvertDocxToHtmlFull(
        byte[] docxBytes,
        string pageTitle,
        string cssPrefix,
        bool fabricateClasses,
        string additionalCss,
        int commentRenderMode,
        string commentCssClassPrefix,
        int paginationMode,
        double paginationScale,
        string paginationCssClassPrefix,
        bool renderAnnotations,
        int annotationLabelMode,
        string annotationCssClassPrefix)
    {
        // Delegate to complete version with new options disabled for backward compatibility
        return ConvertDocxToHtmlComplete(
            docxBytes,
            pageTitle,
            cssPrefix,
            fabricateClasses,
            additionalCss,
            commentRenderMode,
            commentCssClassPrefix,
            paginationMode,
            paginationScale,
            paginationCssClassPrefix,
            renderAnnotations,
            annotationLabelMode,
            annotationCssClassPrefix,
            renderFootnotesAndEndnotes: false,
            renderHeadersAndFooters: false,
            renderTrackedChanges: false,
            showDeletedContent: true,
            renderMoveOperations: true
        );
    }

    /// <summary>
    /// Convert a DOCX file to HTML with all available options.
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
    /// <param name="renderAnnotations">Whether to render custom annotations</param>
    /// <param name="annotationLabelMode">Annotation label mode: 0=Above, 1=Inline, 2=Tooltip, 3=None</param>
    /// <param name="annotationCssClassPrefix">CSS class prefix for annotations (default: "annot-")</param>
    /// <param name="renderFootnotesAndEndnotes">Whether to render footnotes and endnotes sections</param>
    /// <param name="renderHeadersAndFooters">Whether to render document headers and footers</param>
    /// <param name="renderTrackedChanges">Whether to render tracked changes (insertions/deletions)</param>
    /// <param name="showDeletedContent">Whether to show deleted content with strikethrough (only when renderTrackedChanges=true)</param>
    /// <param name="renderMoveOperations">Whether to distinguish move operations from insert/delete (only when renderTrackedChanges=true)</param>
    /// <returns>HTML string or JSON error object</returns>
    [JSExport]
    public static string ConvertDocxToHtmlComplete(
        byte[] docxBytes,
        string pageTitle,
        string cssPrefix,
        bool fabricateClasses,
        string additionalCss,
        int commentRenderMode,
        string commentCssClassPrefix,
        int paginationMode,
        double paginationScale,
        string paginationCssClassPrefix,
        bool renderAnnotations,
        int annotationLabelMode,
        string annotationCssClassPrefix,
        bool renderFootnotesAndEndnotes,
        bool renderHeadersAndFooters,
        bool renderTrackedChanges,
        bool showDeletedContent,
        bool renderMoveOperations)
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
                PaginationCssClassPrefix = paginationCssClassPrefix ?? "page-",
                RenderAnnotations = renderAnnotations,
                AnnotationLabelMode = (AnnotationLabelMode)annotationLabelMode,
                AnnotationCssClassPrefix = annotationCssClassPrefix ?? "annot-",
                IncludeAnnotationMetadata = true,
                RenderFootnotesAndEndnotes = renderFootnotesAndEndnotes,
                RenderHeadersAndFooters = renderHeadersAndFooters,
                RenderTrackedChanges = renderTrackedChanges,
                ShowDeletedContent = showDeletedContent,
                RenderMoveOperations = renderMoveOperations,
                IncludeRevisionMetadata = true
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
    /// Get all annotations from a document.
    /// </summary>
    /// <param name="docxBytes">The DOCX file as a byte array</param>
    /// <returns>JSON response with annotations array or error</returns>
    [JSExport]
    public static string GetAnnotations(byte[] docxBytes)
    {
        if (docxBytes == null || docxBytes.Length == 0)
        {
            return SerializeError("No document data provided");
        }

        try
        {
            var wmlDoc = new WmlDocument("document.docx", docxBytes);
            var annotations = AnnotationManager.GetAnnotations(wmlDoc);

            var response = new AnnotationsResponse
            {
                Annotations = annotations.Select(a => new AnnotationInfo
                {
                    Id = a.Id,
                    LabelId = a.LabelId,
                    Label = a.Label,
                    Color = a.Color,
                    Author = a.Author,
                    Created = a.Created?.ToString("o"),
                    BookmarkName = a.BookmarkName,
                    StartPage = a.StartPage,
                    EndPage = a.EndPage,
                    AnnotatedText = a.AnnotatedText,
                    Metadata = a.Metadata?.Count > 0 ? a.Metadata : null
                }).ToArray()
            };

            return JsonSerializer.Serialize(response, DocxodusJsonContext.Default.AnnotationsResponse);
        }
        catch (Exception ex)
        {
            return SerializeError(ex.Message, ex.GetType().Name, ex.StackTrace);
        }
    }

    /// <summary>
    /// Add an annotation to a document.
    /// </summary>
    /// <param name="docxBytes">The DOCX file as a byte array</param>
    /// <param name="requestJson">JSON request with annotation details</param>
    /// <returns>JSON response with modified document bytes and annotation info</returns>
    [JSExport]
    public static string AddAnnotation(byte[] docxBytes, string requestJson)
    {
        if (docxBytes == null || docxBytes.Length == 0)
        {
            return SerializeError("No document data provided");
        }

        try
        {
            var request = JsonSerializer.Deserialize(requestJson, DocxodusJsonContext.Default.AddAnnotationRequest);
            if (request == null)
            {
                return SerializeError("Invalid request JSON");
            }

            var wmlDoc = new WmlDocument("document.docx", docxBytes);

            var annotation = new DocumentAnnotation(request.Id, request.LabelId, request.Label, request.Color)
            {
                Author = request.Author
            };

            if (request.Metadata != null)
            {
                foreach (var (key, value) in request.Metadata)
                {
                    annotation.Metadata[key] = value;
                }
            }

            AnnotationRange range;
            if (!string.IsNullOrEmpty(request.SearchText))
            {
                range = AnnotationRange.FromSearch(request.SearchText, request.Occurrence);
            }
            else if (request.StartParagraphIndex.HasValue && request.EndParagraphIndex.HasValue)
            {
                range = AnnotationRange.FromParagraphs(request.StartParagraphIndex.Value, request.EndParagraphIndex.Value);
            }
            else
            {
                return SerializeError("Request must specify either SearchText or paragraph indices");
            }

            var resultDoc = AnnotationManager.AddAnnotation(wmlDoc, annotation, range);

            // Get the added annotation to return its details
            var addedAnnotation = AnnotationManager.GetAnnotation(resultDoc, request.Id);

            var response = new AddAnnotationBase64Response
            {
                Success = true,
                DocumentBytes = Convert.ToBase64String(resultDoc.DocumentByteArray),
                Annotation = addedAnnotation != null ? new AnnotationInfo
                {
                    Id = addedAnnotation.Id,
                    LabelId = addedAnnotation.LabelId,
                    Label = addedAnnotation.Label,
                    Color = addedAnnotation.Color,
                    Author = addedAnnotation.Author,
                    Created = addedAnnotation.Created?.ToString("o"),
                    BookmarkName = addedAnnotation.BookmarkName,
                    AnnotatedText = addedAnnotation.AnnotatedText
                } : null
            };
            return JsonSerializer.Serialize(response, DocxodusJsonContext.Default.AddAnnotationBase64Response);
        }
        catch (Exception ex)
        {
            return SerializeError(ex.Message, ex.GetType().Name, ex.StackTrace);
        }
    }

    /// <summary>
    /// Remove an annotation from a document.
    /// </summary>
    /// <param name="docxBytes">The DOCX file as a byte array</param>
    /// <param name="annotationId">The ID of the annotation to remove</param>
    /// <returns>Base64-encoded modified document bytes or JSON error</returns>
    [JSExport]
    public static string RemoveAnnotation(byte[] docxBytes, string annotationId)
    {
        if (docxBytes == null || docxBytes.Length == 0)
        {
            return SerializeError("No document data provided");
        }

        if (string.IsNullOrEmpty(annotationId))
        {
            return SerializeError("Annotation ID is required");
        }

        try
        {
            var wmlDoc = new WmlDocument("document.docx", docxBytes);
            var resultDoc = AnnotationManager.RemoveAnnotation(wmlDoc, annotationId);

            var response = new RemoveAnnotationResponse
            {
                Success = true,
                DocumentBytes = Convert.ToBase64String(resultDoc.DocumentByteArray)
            };
            return JsonSerializer.Serialize(response, DocxodusJsonContext.Default.RemoveAnnotationResponse);
        }
        catch (Exception ex)
        {
            return SerializeError(ex.Message, ex.GetType().Name, ex.StackTrace);
        }
    }

    /// <summary>
    /// Check if a document has any annotations.
    /// </summary>
    /// <param name="docxBytes">The DOCX file as a byte array</param>
    /// <returns>JSON with HasAnnotations boolean</returns>
    [JSExport]
    public static string HasAnnotations(byte[] docxBytes)
    {
        if (docxBytes == null || docxBytes.Length == 0)
        {
            return SerializeError("No document data provided");
        }

        try
        {
            var wmlDoc = new WmlDocument("document.docx", docxBytes);
            var hasAnnotations = AnnotationManager.HasAnnotations(wmlDoc);

            var response = new HasAnnotationsResponse { HasAnnotations = hasAnnotations };
            return JsonSerializer.Serialize(response, DocxodusJsonContext.Default.HasAnnotationsResponse);
        }
        catch (Exception ex)
        {
            return SerializeError(ex.Message, ex.GetType().Name, ex.StackTrace);
        }
    }

    /// <summary>
    /// Get the document structure for element-based annotation targeting.
    /// </summary>
    /// <param name="docxBytes">The DOCX file as a byte array</param>
    /// <returns>JSON response with document structure tree</returns>
    [JSExport]
    public static string GetDocumentStructure(byte[] docxBytes)
    {
        if (docxBytes == null || docxBytes.Length == 0)
        {
            return SerializeError("No document data provided");
        }

        try
        {
            var wmlDoc = new WmlDocument("document.docx", docxBytes);
            var structure = AnnotationManager.GetDocumentStructure(wmlDoc);

            var response = new DocumentStructureResponse
            {
                Root = ConvertElement(structure.Root),
                ElementsById = structure.ElementsById.ToDictionary(
                    kvp => kvp.Key,
                    kvp => ConvertElementShallow(kvp.Value)),
                TableColumns = structure.TableColumns.ToDictionary(
                    kvp => kvp.Key,
                    kvp => new TableColumnInfoDto
                    {
                        TableId = kvp.Value.TableId,
                        ColumnIndex = kvp.Value.ColumnIndex,
                        CellIds = kvp.Value.CellIds.ToArray(),
                        RowCount = kvp.Value.RowCount
                    })
            };

            return JsonSerializer.Serialize(response, DocxodusJsonContext.Default.DocumentStructureResponse);
        }
        catch (Exception ex)
        {
            return SerializeError(ex.Message, ex.GetType().Name, ex.StackTrace);
        }
    }

    /// <summary>
    /// Add an annotation using flexible targeting (element ID, indices, or text search).
    /// </summary>
    /// <param name="docxBytes">The DOCX file as a byte array</param>
    /// <param name="requestJson">JSON request with annotation and target details</param>
    /// <returns>JSON response with modified document bytes and annotation info</returns>
    [JSExport]
    public static string AddAnnotationWithTarget(byte[] docxBytes, string requestJson)
    {
        if (docxBytes == null || docxBytes.Length == 0)
        {
            return SerializeError("No document data provided");
        }

        try
        {
            var request = JsonSerializer.Deserialize(requestJson, DocxodusJsonContext.Default.AddAnnotationWithTargetRequest);
            if (request == null)
            {
                return SerializeError("Invalid request JSON");
            }

            var wmlDoc = new WmlDocument("document.docx", docxBytes);

            var annotation = new DocumentAnnotation(request.Id, request.LabelId, request.Label, request.Color)
            {
                Author = request.Author
            };

            if (request.Metadata != null)
            {
                foreach (var (key, value) in request.Metadata)
                {
                    annotation.Metadata[key] = value;
                }
            }

            // Build AnnotationTarget from request
            var target = new AnnotationTarget
            {
                ElementId = request.ElementId,
                SearchText = request.SearchText,
                Occurrence = request.Occurrence,
                ParagraphIndex = request.ParagraphIndex,
                RunIndex = request.RunIndex,
                TableIndex = request.TableIndex,
                RowIndex = request.RowIndex,
                CellIndex = request.CellIndex,
                ColumnIndex = request.ColumnIndex
            };

            // Parse element type if provided
            if (!string.IsNullOrEmpty(request.ElementType))
            {
                if (Enum.TryParse<DocumentElementType>(request.ElementType, true, out var elementType))
                {
                    target.ElementType = elementType;
                }
                else
                {
                    return SerializeError($"Invalid element type: {request.ElementType}");
                }
            }

            // Handle range end for paragraph ranges
            if (request.RangeEndParagraphIndex.HasValue)
            {
                target.RangeEnd = new AnnotationTarget
                {
                    ParagraphIndex = request.RangeEndParagraphIndex.Value
                };
            }

            var resultDoc = AnnotationManager.AddAnnotation(wmlDoc, annotation, target);

            // Get the added annotation to return its details
            var addedAnnotation = AnnotationManager.GetAnnotation(resultDoc, request.Id);

            var response = new AddAnnotationResponse
            {
                DocumentBytes = resultDoc.DocumentByteArray,
                Annotation = addedAnnotation != null ? new AnnotationInfo
                {
                    Id = addedAnnotation.Id,
                    LabelId = addedAnnotation.LabelId,
                    Label = addedAnnotation.Label,
                    Color = addedAnnotation.Color,
                    Author = addedAnnotation.Author,
                    Created = addedAnnotation.Created?.ToString("o"),
                    BookmarkName = addedAnnotation.BookmarkName,
                    AnnotatedText = addedAnnotation.AnnotatedText,
                    Metadata = addedAnnotation.Metadata
                } : null
            };

            return JsonSerializer.Serialize(response, DocxodusJsonContext.Default.AddAnnotationResponse);
        }
        catch (Exception ex)
        {
            return SerializeError(ex.Message, ex.GetType().Name, ex.StackTrace);
        }
    }

    private static DocumentElementInfo ConvertElement(DocumentElement element)
    {
        return new DocumentElementInfo
        {
            Id = element.Id,
            Type = element.Type.ToString(),
            TextPreview = element.TextPreview,
            Index = element.Index,
            RowIndex = element.RowIndex,
            ColumnIndex = element.ColumnIndex,
            RowSpan = element.RowSpan,
            ColumnSpan = element.ColumnSpan,
            Children = element.Children.Select(ConvertElement).ToArray()
        };
    }

    private static DocumentElementInfo ConvertElementShallow(DocumentElement element)
    {
        // For the lookup dictionary, we don't include children to avoid duplication
        return new DocumentElementInfo
        {
            Id = element.Id,
            Type = element.Type.ToString(),
            TextPreview = element.TextPreview,
            Index = element.Index,
            RowIndex = element.RowIndex,
            ColumnIndex = element.ColumnIndex,
            RowSpan = element.RowSpan,
            ColumnSpan = element.ColumnSpan,
            Children = Array.Empty<DocumentElementInfo>()
        };
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
