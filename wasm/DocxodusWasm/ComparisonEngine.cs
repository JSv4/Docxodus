#nullable enable

using System;
using System.IO;
using System.Runtime.Versioning;
using Docxodus;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxodusWasm;

/// <summary>
/// The comparison engine's one-time warm-up, and the invariant every JSExport that runs a
/// comparison holds: <see cref="EnsureWarm"/> before the first real comparison of a module
/// instance (issues #695, #696).
///
/// <para><b>Why this is mandatory in the browser and meaningless natively.</b> The first
/// comparison a module instance runs is not merely slower than the ones after it — it executes
/// the engine's whole cold path (assembly resolution, type loads, static constructors, and the
/// first-time entry into every method the diff and render stages touch), and a large part of
/// that cold path runs on the Mono interpreter rather than in AOT-compiled code. Mono's GC pins
/// conservatively from the interpreter stack, so while that cold path is on the stack, every
/// nursery collection pays a scan proportional to it. A comparison also allocates heavily — a
/// package's worth of XML — so it triggers many collections. When those two coincide on a heap
/// that some earlier operation has already filled (opening a <c>DocxSession</c> to read
/// revisions, say), the collections stop being incidental and the comparison ceases to make
/// meaningful progress: measured at over ten minutes for a pair that takes 25 ms natively and
/// ~500 ms warm in the browser.</para>
///
/// <para>The seed comparison below is immune to that collapse because it is tiny — two
/// one-paragraph in-memory documents allocate almost nothing — while still executing the same
/// cold path. Paying it first therefore leaves the caller's real comparison warm, whatever its
/// size. This holds where tuning the GC does not: the collapse is not monotone in nursery size
/// (measured: 4m, 6m and 12m all collapse while 8m and 16m do not), so no nursery setting is a
/// fix, whereas warming first survives every one of those configurations.</para>
///
/// <para><b>Not a substitute for <c>prepare()</c>.</b> The npm worker still calls
/// <see cref="DocumentComparer.Warmup"/> up front to move this cost off the critical path.
/// The difference is that a caller who does not can no longer land on the collapse: the cost is
/// paid on the first comparison instead of before it, and it is paid once either way.</para>
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
    /// path to run, so it is latched even on failure — a caller must not pay a failing warm-up
    /// on every comparison. The failure is reported rather than thrown so a comparison entry
    /// point can ignore it and proceed to the caller's real work.
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
