using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using System.Text.Json;
using Docxodus;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxodusWasm;

/// <summary>
/// JSExport methods for DOCX document comparison (redlining).
/// These methods are callable from JavaScript.
/// </summary>
[SupportedOSPlatform("browser")]
public partial class DocumentComparer
{
    /// <summary>
    /// Force the comparison code path fully hot.
    ///
    /// Creating the WASM runtime does not exercise the comparison engine, so
    /// the first real <see cref="CompareDocuments"/> pays a one-time warmup
    /// cost — comparison-assembly initialization plus JIT of the diff/XML
    /// stack — on top of the actual diff work (~2x the steady-state latency).
    ///
    /// <para>This method runs a complete comparison against two tiny seed
    /// documents constructed in-memory, exercising the exact code path
    /// <see cref="CompareDocuments"/> uses (<see cref="WmlComparer.Compare"/>).
    /// That resolves and JIT-compiles everything the engine touches, so a
    /// subsequent real comparison runs at steady-state speed and triggers no
    /// further <c>.wasm</c> fetches.</para>
    ///
    /// <para>Idempotent and self-contained: no caller IO, no seed fixtures to
    /// ship. Safe to call repeatedly — the warmup work is only paid once.
    /// Returns <c>"ok"</c> on success or a JSON error object; warmup is
    /// best-effort, so even the error path has already warmed the engine.</para>
    /// </summary>
    /// <returns><c>"ok"</c> on success, or a JSON error object.</returns>
    [JSExport]
    public static string Warmup()
    {
        try
        {
            // Two minimal in-memory documents that differ by a single word, so
            // the comparer produces a real insertion/deletion and walks the
            // full LCS + markup path rather than an empty fast-exit.
            var original = new WmlDocument("warmup-original.docx", BuildSeedDocx("warmup original"));
            var modified = new WmlDocument("warmup-modified.docx", BuildSeedDocx("warmup modified"));

            var settings = new WmlComparerSettings
            {
                AuthorForRevisions = "Docxodus",
                DateTimeForRevisions = DateTime.UtcNow.ToString("o"),
                DetailThreshold = 0.15
            };

            var result = WmlComparer.Compare(original, modified, settings);

            // Touch the revision-extraction path too, since callers that warm
            // the compare path almost always read revisions next.
            _ = WmlComparer.GetRevisions(result, settings);

            return "ok";
        }
        catch (Exception ex)
        {
            // The act of calling WmlComparer.Compare above has already forced
            // the assemblies to load even if the comparison itself threw, so
            // warmup has still served its purpose. Report the failure so a
            // caller can surface it, but do not throw.
            return DocumentConverter.SerializeError(ex.Message, ex.GetType().Name);
        }
    }

    /// <summary>
    /// Build a minimal but valid DOCX package (one paragraph) in memory.
    /// Includes the parts comparison expects (styles, settings).
    /// </summary>
    private static byte[] BuildSeedDocx(string text)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(
                    new Run(
                        new Text(text) { Space = SpaceProcessingModeValues.Preserve }))));

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles();

            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings();

            mainPart.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>
    /// Compare two DOCX documents and return the result as a redlined DOCX (byte array).
    /// </summary>
    /// <param name="originalBytes">The original DOCX file as a byte array</param>
    /// <param name="modifiedBytes">The modified DOCX file as a byte array</param>
    /// <param name="authorName">Author name for tracked changes</param>
    /// <param name="engine">Comparison engine (<see cref="ComparisonEngine"/> as int; 0 = WmlComparer default, 1 = DocxDiff)</param>
    /// <returns>Redlined DOCX as byte array, or empty array on error</returns>
    [JSExport]
    public static byte[] CompareDocuments(
        byte[] originalBytes,
        byte[] modifiedBytes,
        string authorName,
        int engine)
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

            var result = DocxCompare.Compare(original, modified, (ComparisonEngine)engine, settings);
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
        // Default: show tracked changes visually, WmlComparer engine (0).
        return CompareDocumentsToHtmlWithOptions(originalBytes, modifiedBytes, authorName, renderTrackedChanges: true, engine: 0);
    }

    /// <summary>
    /// Compare two DOCX documents and return the result as HTML with options.
    /// </summary>
    /// <param name="originalBytes">The original DOCX file as a byte array</param>
    /// <param name="modifiedBytes">The modified DOCX file as a byte array</param>
    /// <param name="authorName">Author name for tracked changes</param>
    /// <param name="renderTrackedChanges">If true, show insertions/deletions visually. If false, accept all changes (clean output).</param>
    /// <param name="engine">Comparison engine (<see cref="ComparisonEngine"/> as int; 0 = WmlComparer default, 1 = DocxDiff)</param>
    /// <returns>HTML string, or JSON error object</returns>
    [JSExport]
    public static string CompareDocumentsToHtmlWithOptions(
        byte[] originalBytes,
        byte[] modifiedBytes,
        string authorName,
        bool renderTrackedChanges,
        int engine)
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

            var result = DocxCompare.Compare(original, modified, (ComparisonEngine)engine, comparerSettings);

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
    /// Compare two DOCX documents and return the result as HTML with full options.
    /// Supports all comparison settings (detailThreshold, caseInsensitive) plus HTML rendering options.
    /// </summary>
    /// <param name="originalBytes">The original DOCX file as a byte array</param>
    /// <param name="modifiedBytes">The modified DOCX file as a byte array</param>
    /// <param name="authorName">Author name for tracked changes</param>
    /// <param name="detailThreshold">Detail threshold (0.0 to 1.0, default 0.15)</param>
    /// <param name="caseInsensitive">Whether comparison is case-insensitive</param>
    /// <param name="renderTrackedChanges">If true, show insertions/deletions visually. If false, accept all changes (clean output).</param>
    /// <param name="engine">Comparison engine (<see cref="ComparisonEngine"/> as int; 0 = WmlComparer default, 1 = DocxDiff)</param>
    /// <returns>HTML string, or JSON error object</returns>
    [JSExport]
    public static string CompareDocumentsToHtmlFull(
        byte[] originalBytes,
        byte[] modifiedBytes,
        string authorName,
        double detailThreshold,
        bool caseInsensitive,
        bool renderTrackedChanges,
        int engine)
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
                DetailThreshold = detailThreshold,
                CaseInsensitive = caseInsensitive
            };

            var result = DocxCompare.Compare(original, modified, (ComparisonEngine)engine, comparerSettings);

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
    /// Compare two DOCX documents with logging enabled.
    /// Returns both the redlined document and a log of any warnings/errors encountered.
    /// </summary>
    /// <param name="originalBytes">The original DOCX file as a byte array</param>
    /// <param name="modifiedBytes">The modified DOCX file as a byte array</param>
    /// <param name="authorName">Author name for tracked changes</param>
    /// <param name="detailThreshold">Detail threshold (0.0 to 1.0, default 0.15)</param>
    /// <param name="caseInsensitive">Whether comparison is case-insensitive</param>
    /// <returns>JSON response with document bytes (base64) and log entries</returns>
    [JSExport]
    public static string CompareDocumentsWithLog(
        byte[] originalBytes,
        byte[] modifiedBytes,
        string authorName,
        double detailThreshold,
        bool caseInsensitive)
    {
        if (originalBytes == null || originalBytes.Length == 0 ||
            modifiedBytes == null || modifiedBytes.Length == 0)
        {
            return JsonSerializer.Serialize(new CompareDocumentsWithLogResponse
            {
                Success = false,
                Error = "Missing document data"
            }, DocxodusJsonContext.Default.CompareDocumentsWithLogResponse);
        }

        var log = new ComparisonLog();

        try
        {
            var original = new WmlDocument("original.docx", originalBytes);
            var modified = new WmlDocument("modified.docx", modifiedBytes);

            var settings = new WmlComparerSettings
            {
                AuthorForRevisions = authorName ?? "Docxodus",
                DateTimeForRevisions = DateTime.UtcNow.ToString("o"),
                DetailThreshold = detailThreshold,
                CaseInsensitive = caseInsensitive,
                Log = log
            };

            var result = WmlComparer.Compare(original, modified, settings);

            return JsonSerializer.Serialize(new CompareDocumentsWithLogResponse
            {
                Success = true,
                DocumentBase64 = Convert.ToBase64String(result.DocumentByteArray),
                Log = ConvertLogEntries(log),
                HasWarnings = log.HasWarnings,
                HasErrors = log.HasErrors
            }, DocxodusJsonContext.Default.CompareDocumentsWithLogResponse);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new CompareDocumentsWithLogResponse
            {
                Success = false,
                Error = $"{ex.GetType().Name}: {ex.Message}",
                Log = ConvertLogEntries(log),
                HasWarnings = log.HasWarnings,
                HasErrors = log.HasErrors
            }, DocxodusJsonContext.Default.CompareDocumentsWithLogResponse);
        }
    }

    /// <summary>
    /// Compare two DOCX documents to HTML with logging enabled.
    /// Returns both the HTML output and a log of any warnings/errors encountered.
    /// </summary>
    /// <param name="originalBytes">The original DOCX file as a byte array</param>
    /// <param name="modifiedBytes">The modified DOCX file as a byte array</param>
    /// <param name="authorName">Author name for tracked changes</param>
    /// <param name="detailThreshold">Detail threshold (0.0 to 1.0, default 0.15)</param>
    /// <param name="caseInsensitive">Whether comparison is case-insensitive</param>
    /// <param name="renderTrackedChanges">If true, show insertions/deletions visually</param>
    /// <returns>JSON response with HTML and log entries</returns>
    [JSExport]
    public static string CompareDocumentsToHtmlWithLog(
        byte[] originalBytes,
        byte[] modifiedBytes,
        string authorName,
        double detailThreshold,
        bool caseInsensitive,
        bool renderTrackedChanges)
    {
        if (originalBytes == null || originalBytes.Length == 0 ||
            modifiedBytes == null || modifiedBytes.Length == 0)
        {
            return JsonSerializer.Serialize(new CompareDocumentsToHtmlWithLogResponse
            {
                Success = false,
                Error = "Missing document data"
            }, DocxodusJsonContext.Default.CompareDocumentsToHtmlWithLogResponse);
        }

        var log = new ComparisonLog();

        try
        {
            var original = new WmlDocument("original.docx", originalBytes);
            var modified = new WmlDocument("modified.docx", modifiedBytes);

            var comparerSettings = new WmlComparerSettings
            {
                AuthorForRevisions = authorName ?? "Docxodus",
                DateTimeForRevisions = DateTime.UtcNow.ToString("o"),
                DetailThreshold = detailThreshold,
                CaseInsensitive = caseInsensitive,
                Log = log
            };

            var result = WmlComparer.Compare(original, modified, comparerSettings);

            // Convert the redlined document to HTML
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

            if (renderTrackedChanges)
            {
                htmlSettings.AuthorColors = new Dictionary<string, string>
                {
                    { authorName ?? "Docxodus", "#007bff" }
                };
            }

            var htmlElement = WmlToHtmlConverter.ConvertToHtml(wordDoc, htmlSettings);

            return JsonSerializer.Serialize(new CompareDocumentsToHtmlWithLogResponse
            {
                Success = true,
                Html = htmlElement.ToString(),
                Log = ConvertLogEntries(log),
                HasWarnings = log.HasWarnings,
                HasErrors = log.HasErrors
            }, DocxodusJsonContext.Default.CompareDocumentsToHtmlWithLogResponse);
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new CompareDocumentsToHtmlWithLogResponse
            {
                Success = false,
                Error = $"{ex.GetType().Name}: {ex.Message}",
                Log = ConvertLogEntries(log),
                HasWarnings = log.HasWarnings,
                HasErrors = log.HasErrors
            }, DocxodusJsonContext.Default.CompareDocumentsToHtmlWithLogResponse);
        }
    }

    /// <summary>
    /// Convert ComparisonLog entries to DTOs for serialization.
    /// </summary>
    private static ComparisonLogEntryDto[] ConvertLogEntries(ComparisonLog log)
    {
        return log.Entries.Select(e => new ComparisonLogEntryDto
        {
            Level = e.Level.ToString(),
            Code = e.Code,
            Message = e.Message,
            Details = e.Details,
            Location = e.Location
        }).ToArray();
    }

    /// <summary>
    /// Compare documents with detailed options.
    /// </summary>
    /// <param name="originalBytes">The original DOCX file</param>
    /// <param name="modifiedBytes">The modified DOCX file</param>
    /// <param name="authorName">Author name for tracked changes</param>
    /// <param name="detailThreshold">Detail threshold (0.0 to 1.0, default 0.15)</param>
    /// <param name="caseInsensitive">Whether comparison is case-insensitive</param>
    /// <param name="engine">Comparison engine (<see cref="ComparisonEngine"/> as int; 0 = WmlComparer default, 1 = DocxDiff)</param>
    /// <returns>Redlined DOCX as byte array</returns>
    [JSExport]
    public static byte[] CompareDocumentsWithOptions(
        byte[] originalBytes,
        byte[] modifiedBytes,
        string authorName,
        double detailThreshold,
        bool caseInsensitive,
        int engine)
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

            var result = DocxCompare.Compare(original, modified, (ComparisonEngine)engine, settings);
            return result.DocumentByteArray;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Comparison error: {ex.Message}");
            return Array.Empty<byte>();
        }
    }
}
