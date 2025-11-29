using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using System.Text.Json;
using Docxodus;
using DocumentFormat.OpenXml.Packaging;

namespace DocxodusWasm;

/// <summary>
/// JSExport methods for DOCX document comparison (redlining).
/// These methods are callable from JavaScript.
/// </summary>
[SupportedOSPlatform("browser")]
public partial class DocumentComparer
{
    /// <summary>
    /// Compare two DOCX documents and return the result as a redlined DOCX (byte array).
    /// </summary>
    /// <param name="originalBytes">The original DOCX file as a byte array</param>
    /// <param name="modifiedBytes">The modified DOCX file as a byte array</param>
    /// <param name="authorName">Author name for tracked changes</param>
    /// <returns>Redlined DOCX as byte array, or empty array on error</returns>
    [JSExport]
    public static byte[] CompareDocuments(
        byte[] originalBytes,
        byte[] modifiedBytes,
        string authorName)
    {
        if (originalBytes == null || originalBytes.Length == 0 ||
            modifiedBytes == null || modifiedBytes.Length == 0)
        {
            Console.WriteLine("Error: Missing document data");
            return Array.Empty<byte>();
        }

        try
        {
            var original = new WmlDocument("original.docx", originalBytes);
            var modified = new WmlDocument("modified.docx", modifiedBytes);

            var settings = new WmlComparerSettings
            {
                AuthorForRevisions = authorName ?? "Docxodus",
                DateTimeForRevisions = DateTime.UtcNow.ToString("o"),
                DetailThreshold = 0.15
            };

            var result = WmlComparer.Compare(original, modified, settings);
            return result.DocumentByteArray;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Comparison error: {ex.GetType().Name}: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
            return Array.Empty<byte>();
        }
    }

    /// <summary>
    /// Compare two DOCX documents and return the result as HTML.
    /// Uses default settings with tracked changes visible.
    /// </summary>
    /// <param name="originalBytes">The original DOCX file as a byte array</param>
    /// <param name="modifiedBytes">The modified DOCX file as a byte array</param>
    /// <param name="authorName">Author name for tracked changes</param>
    /// <returns>HTML string with redlined content, or JSON error object</returns>
    [JSExport]
    public static string CompareDocumentsToHtml(
        byte[] originalBytes,
        byte[] modifiedBytes,
        string authorName)
    {
        // Default: show tracked changes visually
        return CompareDocumentsToHtmlWithOptions(originalBytes, modifiedBytes, authorName, renderTrackedChanges: true);
    }

    /// <summary>
    /// Compare two DOCX documents and return the result as HTML with options.
    /// </summary>
    /// <param name="originalBytes">The original DOCX file as a byte array</param>
    /// <param name="modifiedBytes">The modified DOCX file as a byte array</param>
    /// <param name="authorName">Author name for tracked changes</param>
    /// <param name="renderTrackedChanges">If true, show insertions/deletions visually. If false, accept all changes (clean output).</param>
    /// <returns>HTML string, or JSON error object</returns>
    [JSExport]
    public static string CompareDocumentsToHtmlWithOptions(
        byte[] originalBytes,
        byte[] modifiedBytes,
        string authorName,
        bool renderTrackedChanges)
    {
        if (originalBytes == null || originalBytes.Length == 0 ||
            modifiedBytes == null || modifiedBytes.Length == 0)
        {
            return DocumentConverter.SerializeError("Missing document data");
        }

        try
        {
            var original = new WmlDocument("original.docx", originalBytes);
            var modified = new WmlDocument("modified.docx", modifiedBytes);

            var comparerSettings = new WmlComparerSettings
            {
                AuthorForRevisions = authorName ?? "Docxodus",
                DateTimeForRevisions = DateTime.UtcNow.ToString("o"),
                DetailThreshold = 0.15
            };

            var result = WmlComparer.Compare(original, modified, comparerSettings);

            // Convert the redlined document to HTML
            // Must use writable stream - WmlToHtmlConverter may call RevisionAccepter internally
            using var memoryStream = new MemoryStream();
            memoryStream.Write(result.DocumentByteArray, 0, result.DocumentByteArray.Length);
            memoryStream.Position = 0;
            using var wordDoc = WordprocessingDocument.Open(memoryStream, true);

            var htmlSettings = new WmlToHtmlConverterSettings
            {
                PageTitle = "Document Comparison",
                CssClassPrefix = "redline-",
                FabricateCssClasses = true,
                RenderTrackedChanges = renderTrackedChanges,
                IncludeRevisionMetadata = renderTrackedChanges,
                ShowDeletedContent = true,
                RenderMoveOperations = true,
            };

            // Add author color if rendering tracked changes
            if (renderTrackedChanges)
            {
                htmlSettings.AuthorColors = new Dictionary<string, string>
                {
                    { authorName ?? "Docxodus", "#007bff" }
                };
            }

            var htmlElement = WmlToHtmlConverter.ConvertToHtml(wordDoc, htmlSettings);
            return htmlElement.ToString();
        }
        catch (Exception ex)
        {
            return DocumentConverter.SerializeError(ex.Message, ex.GetType().Name, ex.StackTrace);
        }
    }

    /// <summary>
    /// Get revisions from a compared document as JSON.
    /// Uses default move detection settings.
    /// </summary>
    /// <param name="comparedDocBytes">A document that has been through comparison (has tracked changes)</param>
    /// <returns>JSON array of revisions, or JSON error object</returns>
    [JSExport]
    public static string GetRevisionsJson(byte[] comparedDocBytes)
    {
        return GetRevisionsJsonWithOptions(comparedDocBytes, true, 0.8, 3, false);
    }

    /// <summary>
    /// Get revisions from a compared document as JSON with configurable move detection.
    /// </summary>
    /// <param name="comparedDocBytes">A document that has been through comparison (has tracked changes)</param>
    /// <param name="detectMoves">Whether to detect and mark moved content (default: true)</param>
    /// <param name="moveSimilarityThreshold">Jaccard similarity threshold 0.0-1.0 (default: 0.8)</param>
    /// <param name="moveMinimumWordCount">Minimum word count for move detection (default: 3)</param>
    /// <param name="caseInsensitive">Whether similarity matching ignores case (default: false)</param>
    /// <returns>JSON array of revisions, or JSON error object</returns>
    [JSExport]
    public static string GetRevisionsJsonWithOptions(
        byte[] comparedDocBytes,
        bool detectMoves,
        double moveSimilarityThreshold,
        int moveMinimumWordCount,
        bool caseInsensitive)
    {
        if (comparedDocBytes == null || comparedDocBytes.Length == 0)
        {
            return DocumentConverter.SerializeError("No document data provided");
        }

        try
        {
            var doc = new WmlDocument("compared.docx", comparedDocBytes);
            var settings = new WmlComparerSettings
            {
                DetectMoves = detectMoves,
                MoveSimilarityThreshold = moveSimilarityThreshold,
                MoveMinimumWordCount = moveMinimumWordCount,
                CaseInsensitive = caseInsensitive
            };
            var revisions = WmlComparer.GetRevisions(doc, settings);

            var response = new RevisionsResponse
            {
                Revisions = revisions.Select(r => new RevisionInfo
                {
                    Author = r.Author ?? "",
                    Date = r.Date ?? "",
                    RevisionType = r.RevisionType.ToString(),
                    Text = r.Text ?? "",
                    MoveGroupId = r.MoveGroupId,
                    IsMoveSource = r.IsMoveSource,
                    FormatChange = r.FormatChange != null ? new FormatChangeInfo
                    {
                        OldProperties = r.FormatChange.OldProperties,
                        NewProperties = r.FormatChange.NewProperties,
                        ChangedPropertyNames = r.FormatChange.ChangedPropertyNames
                    } : null
                }).ToArray()
            };

            return JsonSerializer.Serialize(response, DocxodusJsonContext.Default.RevisionsResponse);
        }
        catch (Exception ex)
        {
            return DocumentConverter.SerializeError(ex.Message, ex.GetType().Name);
        }
    }

    /// <summary>
    /// Compare documents with detailed options.
    /// </summary>
    /// <param name="originalBytes">The original DOCX file</param>
    /// <param name="modifiedBytes">The modified DOCX file</param>
    /// <param name="authorName">Author name for tracked changes</param>
    /// <param name="detailThreshold">Detail threshold (0.0 to 1.0, default 0.15)</param>
    /// <param name="caseInsensitive">Whether comparison is case-insensitive</param>
    /// <returns>Redlined DOCX as byte array</returns>
    [JSExport]
    public static byte[] CompareDocumentsWithOptions(
        byte[] originalBytes,
        byte[] modifiedBytes,
        string authorName,
        double detailThreshold,
        bool caseInsensitive)
    {
        if (originalBytes == null || originalBytes.Length == 0 ||
            modifiedBytes == null || modifiedBytes.Length == 0)
        {
            return Array.Empty<byte>();
        }

        try
        {
            var original = new WmlDocument("original.docx", originalBytes);
            var modified = new WmlDocument("modified.docx", modifiedBytes);

            var settings = new WmlComparerSettings
            {
                AuthorForRevisions = authorName ?? "Docxodus",
                DateTimeForRevisions = DateTime.UtcNow.ToString("o"),
                DetailThreshold = detailThreshold,
                CaseInsensitive = caseInsensitive
            };

            var result = WmlComparer.Compare(original, modified, settings);
            return result.DocumentByteArray;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Comparison error: {ex.Message}");
            return Array.Empty<byte>();
        }
    }
}
