using System;
using System.IO;
using System.Runtime.Versioning;
using Docxodus;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxodusWasm;

/// <summary>
/// The comparison engine's one-time warm-up, behind <see cref="DocumentComparer.Warmup"/>.
///
/// <para>The first comparison a module instance runs executes the engine's whole cold path
/// (assembly resolution, type loads, static constructors, first entry into every method the
/// diff and render stages touch), much of it on the interpreter, so it costs a few hundred
/// milliseconds more than the ones after it. A seed comparison of two one-paragraph in-memory
/// documents walks the same path while allocating almost nothing, so a caller that runs it
/// first — the npm worker's <c>prepare()</c> — keeps its first real comparison at steady-state
/// latency. It is a latency tool only: the first-comparison hang it was once an invariant
/// against (issues #695, #696) was the interpreter's precise stack marking, which the build now
/// turns off (issue #811, <c>DocxodusWasm.csproj</c>).</para>
/// </summary>
[SupportedOSPlatform("browser")]
internal static class ComparisonEngine
{
    private static bool warmed;

    /// <summary>
    /// Run the seed comparison once per module instance. Subsequent calls return immediately.
    /// </summary>
    /// <returns><c>"ok"</c> on success, or a JSON error object.</returns>
    /// <remarks>
    /// Best-effort: a warm-up that throws has still forced the assemblies to load and the cold
    /// path to run, so it is latched even on failure. The failure is reported rather than
    /// thrown, because warming is never a precondition of the caller's real work.
    /// </remarks>
    internal static string EnsureWarm()
    {
        if (warmed)
            return "ok";
        warmed = true;

        try
        {
            // Two minimal in-memory documents that differ by a single word, so DocxDiff produces
            // a real insertion/deletion and walks its full alignment + markup path rather than an
            // empty fast-exit.
            var original = new WmlDocument("warmup-original.docx", BuildSeedDocx("warmup original"));
            var modified = new WmlDocument("warmup-modified.docx", BuildSeedDocx("warmup modified"));

            var settings = new DocxDiffSettings
            {
                AuthorForRevisions = "Docxodus",
                DateTimeForRevisions = DateTime.UtcNow.ToString("o"),
            };

            var result = DocxCompare.Compare(original, modified, settings);

            // Touch the revision-extraction path too, since callers that warm the compare path
            // almost always read revisions next.
            using (var warmSession = new DocxSession(result.DocumentByteArray))
                _ = warmSession.ListRevisions();

            return "ok";
        }
        catch (Exception ex)
        {
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
}
