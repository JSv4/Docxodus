using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using System.Text.Json;
using Docxodus;
using Docxodus.Internal;
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
    /// <see cref="CompareDocuments"/> uses (<see cref="DocxCompare.Compare"/> through
    /// the DocxDiff engine).
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
            // DocxDiff produces a real insertion/deletion and walks its full
            // alignment + markup path rather than an empty fast-exit.
            var original = new WmlDocument("warmup-original.docx", BuildSeedDocx("warmup original"));
            var modified = new WmlDocument("warmup-modified.docx", BuildSeedDocx("warmup modified"));

            var settings = new DocxDiffSettings
            {
                AuthorForRevisions = "Docxodus",
                DateTimeForRevisions = DateTime.UtcNow.ToString("o"),
            };

            var result = DocxCompare.Compare(original, modified, settings);

            // Touch the revision-extraction path too, since callers that warm
            // the compare path almost always read revisions next.
            using (var warmSession = new DocxSession(result.DocumentByteArray))
                _ = warmSession.ListRevisions();

            return "ok";
        }
        catch (Exception ex)
        {
            // The act of calling DocxCompare.Compare above has already forced
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

        // The public byte-array API has no document metadata beyond the package itself. Avoid
        // marshaling an exact no-op through either comparison engine so callers receive a detached,
        // byte-for-byte copy of the package they supplied.
        if (originalBytes.AsSpan().SequenceEqual(modifiedBytes))
            return (byte[])originalBytes.Clone();

        try
        {
            var original = new WmlDocument("original.docx", originalBytes);
            var modified = new WmlDocument("modified.docx", modifiedBytes);

            var settings = new DocxDiffSettings
            {
                AuthorForRevisions = authorName ?? "Docxodus",
                DateTimeForRevisions = DateTime.UtcNow.ToString("o"),
            };

            var result = DocxCompare.Compare(original, modified, settings);
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
        // Default: show tracked changes visually through the DocxDiff engine.
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

            var settings = new DocxDiffSettings
            {
                AuthorForRevisions = authorName ?? "Docxodus",
                DateTimeForRevisions = DateTime.UtcNow.ToString("o"),
            };

            var result = DocxCompare.Compare(original, modified, settings);

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
    /// Read the tracked revisions already present in a document, as JSON.
    ///
    /// <para>Through v10 this re-derived moves by running the legacy comparer's move detection over the
    /// document, which is why it took similarity/min-word/case knobs. It now reads the document's own
    /// <c>w:moveFrom</c>/<c>w:moveTo</c> markup through <see cref="DocxSession.ListRevisions"/>: the moves
    /// in the document ARE the moves, not a re-guess from a Jaccard threshold, so those knobs are gone.
    /// The reader also surfaces revision families the legacy one never did (rows, cells, content controls,
    /// numbering, property changes) and reports each move as ONE grouped entry rather than two halves.</para>
    ///
    /// <para>The payload is <see cref="DocxSessionJson.SerializeRevisionList"/> verbatim — the SAME wire
    /// shape the session's own <c>listRevisions</c> returns — so the two paths cannot drift. That is why
    /// there is no DTO here: adding one would fork the shape.</para>
    /// </summary>
    /// <param name="comparedDocBytes">A document carrying tracked changes</param>
    /// <returns>JSON array of revisions, or JSON error object</returns>
    [JSExport]
    public static string GetRevisionsJson(byte[] comparedDocBytes)
    {
        if (comparedDocBytes == null || comparedDocBytes.Length == 0)
        {
            return DocumentConverter.SerializeError("No document data provided");
        }

        try
        {
            using var session = new DocxSession(comparedDocBytes);
            return DocxSessionJson.SerializeRevisionList(session.ListRevisions());
        }
        catch (Exception ex)
        {
            return DocumentConverter.SerializeError(ex.Message, ex.GetType().Name);
        }
    }

    /// <summary>
    /// Compare two DOCX documents and return the result as HTML with full options.
    /// Supports the comparison settings that survive (caseInsensitive) plus HTML rendering options.
    /// </summary>
    /// <param name="originalBytes">The original DOCX file as a byte array</param>
    /// <param name="modifiedBytes">The modified DOCX file as a byte array</param>
    /// <param name="authorName">Author name for tracked changes</param>
    /// <param name="caseInsensitive">Whether comparison is case-insensitive</param>
    /// <param name="renderTrackedChanges">If true, show insertions/deletions visually. If false, accept all changes (clean output).</param>
    /// <returns>HTML string, or JSON error object</returns>
    [JSExport]
    public static string CompareDocumentsToHtmlFull(
        byte[] originalBytes,
        byte[] modifiedBytes,
        string authorName,
        bool caseInsensitive,
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

            var settings = new DocxDiffSettings
            {
                AuthorForRevisions = authorName ?? "Docxodus",
                DateTimeForRevisions = DateTime.UtcNow.ToString("o"),
                CaseInsensitive = caseInsensitive,
            };

            var result = DocxCompare.Compare(original, modified, settings);

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
    /// Compare documents with detailed options.
    /// </summary>
    /// <param name="originalBytes">The original DOCX file</param>
    /// <param name="modifiedBytes">The modified DOCX file</param>
    /// <param name="authorName">Author name for tracked changes</param>
    /// <param name="caseInsensitive">Whether comparison is case-insensitive</param>
    /// <returns>Redlined DOCX as byte array</returns>
    [JSExport]
    public static byte[] CompareDocumentsWithOptions(
        byte[] originalBytes,
        byte[] modifiedBytes,
        string authorName,
        bool caseInsensitive)
    {
        if (originalBytes == null || originalBytes.Length == 0 ||
            modifiedBytes == null || modifiedBytes.Length == 0)
        {
            return Array.Empty<byte>();
        }

        if (originalBytes.AsSpan().SequenceEqual(modifiedBytes))
            return (byte[])originalBytes.Clone();

        try
        {
            var original = new WmlDocument("original.docx", originalBytes);
            var modified = new WmlDocument("modified.docx", modifiedBytes);

            var settings = new DocxDiffSettings
            {
                AuthorForRevisions = authorName ?? "Docxodus",
                DateTimeForRevisions = DateTime.UtcNow.ToString("o"),
                CaseInsensitive = caseInsensitive,
            };

            var result = DocxCompare.Compare(original, modified, settings);
            return result.DocumentByteArray;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Comparison error: {ex.Message}");
            return Array.Empty<byte>();
        }
    }
}
