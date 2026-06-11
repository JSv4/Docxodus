#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir;

/// <summary>
/// Equivalence harness for the IR markdown emitter (M1.4 Task 1). Drives the shipped
/// <see cref="WmlToMarkdownConverter"/> (the ORACLE) and the IR path
/// (<see cref="IrReader.Read"/> + <see cref="IrMarkdownEmitter.Emit"/>) over the whole
/// <c>TestFiles/</c> corpus, compares markdown strings and anchor indexes, writes per-fixture diffs
/// for controller triage, and asserts byte-equality on a curated must-pass list of body-simple
/// fixtures. The corpus stat is informational until Task 3 drives it to closure.
/// </summary>
public class IrMarkdownEquivalenceTests
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");

    private static readonly string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private readonly ITestOutputHelper _output;

    public IrMarkdownEquivalenceTests(ITestOutputHelper output) => _output = output;

    // --- revision-view policy (M1.4-T3 sub-task 2) -------------------------------------------------
    //
    // Equivalence is asserted over revision-FREE inputs by design. The oracle projects a document
    // AS-IS (its EmitInlineRuns accepts revisions inline: ins → plain text, del → dropped), while the
    // IR's RevisionView defaults to Accept (RevisionProcessor.AcceptRevisions rewrites the package
    // structurally before the read). For documents carrying tracked changes those two acceptance
    // paths can diverge (moved content, paragraph-mark revisions, etc.). v1 IR deliberately has no
    // as-is projection of tracked-changes docs (spec §5.1); as-is projection is a v2 IR item
    // (spec §11). To compare like-for-like we pre-accept revisions ONCE here and feed the SAME
    // accepted bytes to BOTH the oracle and the IR path — so the oracle sees an already-accepted
    // document (its inline acceptance becomes a no-op) and the IR's Accept re-pass is idempotent.

    /// <summary>
    /// Return input bytes suitable for an apples-to-apples comparison: if <paramref name="file"/>
    /// carries tracked-change markup, return a copy with revisions accepted once
    /// (<see cref="RevisionProcessor.AcceptRevisions(WmlDocument)"/>); otherwise return the file's
    /// bytes verbatim. The returned <see cref="WmlDocument"/> is a fresh instance each call so the
    /// oracle (which mutates bytes to persist Unids) and the IR path never share mutable state.
    /// </summary>
    private static WmlDocument PrepareInput(FileInfo file)
    {
        var doc = new WmlDocument(file.FullName);
        return HasRevisionMarkup(doc) ? RevisionProcessor.AcceptRevisions(doc) : doc;
    }

    private static readonly XName[] RevisionElementNames =
    {
        // The same set IrReader.HasRevisionMarkup uses: the tracked-change markup whose presence
        // means the two acceptance paths could diverge.
        XName.Get("ins", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        XName.Get("del", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        XName.Get("moveFrom", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        XName.Get("moveTo", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        XName.Get("rPrChange", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        XName.Get("pPrChange", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        XName.Get("tblPrChange", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        XName.Get("trPrChange", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        XName.Get("tcPrChange", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        XName.Get("sectPrChange", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        XName.Get("numberingChange", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
    };

    private static bool HasRevisionMarkup(WmlDocument doc)
    {
        try
        {
            using var stream = new OpenXmlMemoryStreamDocument(doc);
            using var wdoc = stream.GetWordprocessingDocument();
            var names = new HashSet<XName>(RevisionElementNames);
            foreach (var part in EnumerableScopeRoots(wdoc))
                if (part is not null && part.DescendantsAndSelf().Any(e => names.Contains(e.Name)))
                    return true;
            return false;
        }
        catch
        {
            return false;
        }
    }

    private static IEnumerable<System.Xml.Linq.XElement?> EnumerableScopeRoots(WordprocessingDocument wdoc)
    {
        var main = wdoc.MainDocumentPart;
        if (main is null) yield break;
        yield return main.GetXDocument().Root;
        foreach (var h in main.HeaderParts) yield return h.GetXDocument().Root;
        foreach (var f in main.FooterParts) yield return f.GetXDocument().Root;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart.GetXDocument().Root;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart.GetXDocument().Root;
    }

    /// <summary>
    /// The curated set of genuinely body-simple fixtures (plain paragraphs / headings / bulleted
    /// lists, no tables/images/multipart/numbered-counter content) whose IR-emitted markdown AND
    /// anchor index must be byte-equal to the oracle's. Verified by inspecting the oracle output and
    /// the corpus equality report. This list grows per task as more of the projection is ported.
    /// </summary>
    public static IEnumerable<object[]> MustPassFixtures()
    {
        foreach (var name in MustPassNames)
            yield return new object[] { name };
    }

    // Populated empirically from the corpus equality report (see MarkdownEquivalence_CorpusReport).
    // Every entry is a fixture whose body is simple enough that the Task-1 emitter reaches byte
    // equality with the oracle on both markdown and the (AutoNumberPrefix-excluded) anchor index.
    private static readonly string[] MustPassNames =
    {
        "CA001-Plain.docx",        // plain paragraphs — the baseline shape
        "CZ002-Multi-Paragraphs.docx", // several plain paragraphs + anchors
        "HC023-Hyperlink.docx",    // [text](url) hyperlink rendering
        "HC024-Tabs-01.docx",      // w:tab → 4 spaces
        "HC039-Bold.docx",         // **bold** delimiter
        "HC035-Strike-Through.docx", // ~~strike~~ delimiter
        // M1.4-T2: bulleted lists (review follow-up), tables, and images.
        "HC010-Test-05.docx",      // bulleted list items (·-format) → "-" markers + 2-space indent
        "CA005-Table.docx",        // simple table → GFM pipe table + tbl/tr/tc index entries
        "CA014-Complex-Table.docx", // 8x9 table over the cell cap → opaque ```table rows/cols block
        "HC042-Image-Png.docx",    // inline image: oracle emits no image markup; IR matches (no img line)
        // A clean body-only in-pPr sectPr corpus fixture does not exist (every TestFiles sectPr
        // fixture is also multipart or revision-tainted — T3 territory), so the {#sec:…} + thematic
        // break is pinned programmatically in IrMarkdownRuleTests.Rule_InlineSectionBreak instead.

        // --- M1.4-T3: numbering counters, multipart scopes, note refs, revision-accepted inputs ---
        "CA003-Numbered-List.docx",           // numbered (decimal) list markers via the resolved-marker fact
        "LIR001-en-US-ordinal.docx",          // ordinal numbering + Heading-with-numPr trailing-blank rule
        "LIR003-en-US-upperLetter.docx",      // upperLetter counter markers
        "DB012-Lists-With-Different-Numberings.docx", // multiple numId counters in one document
        "CA008-Footnote-Reference.docx",      // body [^fn-…] note ref + # Footnotes definition section
        "DB007-Notes.docx",                   // footnotes scope: labels from note Unid, definition lines
        "RC007-Endnotes-After.docx",          // # Endnotes section
        "DB002-Sections-With-Headers.docx",   // # Headers / ## hdrN multipart structure + dividers
        "DA236-Page-Num-in-Footer.docx",      // # Footers / ## ftrN multipart structure
        "WC020-FootNote-After-1.docx",        // footnote ref + section, revision-accepted by the harness
        // Per-rule pins for the resolved-marker display rules, fldSimple drop, inline-SDT drop,
        // block-level-SDT skip, and tab grouping live in IrMarkdownRuleTests (added this task).
    };

    // --- corpus report (informational; asserts only the must-pass list + totality) ---------------

    [Fact]
    [Trait("Category", "Corpus")]
    public void MarkdownEquivalence_CorpusReport()
    {
        var files = TestFilesDir.GetFiles("*.docx", SearchOption.AllDirectories)
            .OrderBy(f => f.FullName, StringComparer.Ordinal)
            .ToList();

        var artifactsDir = ArtifactsDir();
        if (Directory.Exists(artifactsDir)) Directory.Delete(artifactsDir, recursive: true);
        Directory.CreateDirectory(artifactsDir);

        int equal = 0, different = 0, skipped = 0, threw = 0;
        var equalNames = new List<string>();
        var emitterFailures = new List<string>();

        foreach (var file in files)
        {
            if (!CanOpen(file)) { skipped++; continue; }

            // Revision-view policy (see PrepareInput): pre-accept revisions once for fixtures that
            // carry tracked changes, and feed the SAME accepted bytes to both paths.
            WmlDocument prepared;
            try
            {
                prepared = PrepareInput(file);
            }
            catch (Exception ex)
            {
                _output.WriteLine($"[prepare-skip] {file.Name}: {ex.GetType().Name}");
                skipped++;
                continue;
            }

            string oracleMd;
            IReadOnlyDictionary<string, AnchorTarget> oracleIndex;
            try
            {
                // The oracle mutates the document bytes (persists Unids) — run it on a copy.
                var oracleDoc = new WmlDocument(prepared);
                var projection = WmlToMarkdownConverter.Convert(oracleDoc, new WmlToMarkdownConverterSettings());
                oracleMd = projection.Markdown;
                oracleIndex = projection.AnchorIndex;
            }
            catch (Exception ex)
            {
                // The oracle itself rejecting a fixture is out of scope for this harness.
                _output.WriteLine($"[oracle-skip] {file.Name}: {ex.GetType().Name}");
                skipped++;
                continue;
            }

            string irMd;
            IReadOnlyDictionary<string, AnchorTarget> irIndex;
            try
            {
                var ir = IrReader.Read(new WmlDocument(prepared));
                var result = IrMarkdownEmitter.Emit(ir, new WmlToMarkdownConverterSettings());
                irMd = result.Markdown;
                irIndex = result.AnchorIndex;
            }
            catch (Exception ex)
            {
                // Totality violation: the emitter must never throw. Record and continue so the
                // report is complete, but fail the test at the end.
                threw++;
                emitterFailures.Add($"{file.Name}: {ex.GetType().Name}: {ex.Message}");
                continue;
            }

            var mdEqual = string.Equals(oracleMd, irMd, StringComparison.Ordinal);
            var indexEqual = BodyIndexEqual(oracleIndex, irIndex, out var indexDiff);

            if (mdEqual && indexEqual)
            {
                equal++;
                equalNames.Add(file.Name);
            }
            else
            {
                different++;
                WriteDiff(artifactsDir, file.Name, oracleMd, irMd, indexDiff);
            }
        }

        _output.WriteLine($"Corpus markdown equivalence: {equal} equal / {equal + different} comparable " +
                          $"({skipped} skipped, {threw} emitter-threw) of {files.Count} *.docx.");
        _output.WriteLine("Equal fixtures:");
        foreach (var n in equalNames.OrderBy(n => n, StringComparer.Ordinal))
            _output.WriteLine("  " + n);

        Assert.True(threw == 0,
            $"IrMarkdownEmitter.Emit threw on {threw} fixture(s) (totality violation):" +
            Environment.NewLine + string.Join(Environment.NewLine, emitterFailures));
        Assert.True(equal > 0, "No fixtures reached markdown equivalence — harness or emitter regression.");
    }

    // --- must-pass byte equality ------------------------------------------

    [Theory]
    [MemberData(nameof(MustPassFixtures))]
    public void MarkdownEquivalence_MustPassFixtures(string fixtureName)
    {
        var file = TestFilesDir.GetFiles(fixtureName, SearchOption.AllDirectories)
            .OrderBy(f => f.FullName, StringComparer.Ordinal)
            .First();

        var prepared = PrepareInput(file);

        var oracleDoc = new WmlDocument(prepared);
        var projection = WmlToMarkdownConverter.Convert(oracleDoc, new WmlToMarkdownConverterSettings());

        var ir = IrReader.Read(new WmlDocument(prepared));
        var result = IrMarkdownEmitter.Emit(ir, new WmlToMarkdownConverterSettings());

        Assert.Equal(projection.Markdown, result.Markdown);
        Assert.True(BodyIndexEqual(projection.AnchorIndex, result.AnchorIndex, out var diff),
            $"Anchor-index mismatch for {fixtureName}:{Environment.NewLine}{diff}");
    }

    // --- index comparison -------------------------------------------------

    /// <summary>
    /// Compare the oracle and IR anchor indexes restricted to BODY entries (Task 1 scope). For each
    /// body anchor the oracle produced, the IR must produce an entry with the same Anchor.Id/Kind/
    /// Scope/Unid, identical PartUri, and identical TextPreview. AutoNumberPrefix is EXCLUDED from
    /// the comparison — the IR counter walk lands in M1.4-T3 (see emitter TODO). Returns a
    /// human-readable diff in <paramref name="diff"/> on mismatch.
    /// </summary>
    private static bool BodyIndexEqual(
        IReadOnlyDictionary<string, AnchorTarget> oracle,
        IReadOnlyDictionary<string, AnchorTarget> ir,
        out string diff)
    {
        var sb = new StringBuilder();
        var oracleBody = oracle.Where(kv => kv.Value.Anchor.Scope == "body")
            .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.Ordinal);

        foreach (var (key, oTarget) in oracleBody.OrderBy(kv => kv.Key, StringComparer.Ordinal))
        {
            if (!ir.TryGetValue(key, out var iTarget))
            {
                sb.AppendLine($"  missing in IR: {key}");
                continue;
            }
            if (oTarget.Anchor.Id != iTarget.Anchor.Id
                || oTarget.Anchor.Kind != iTarget.Anchor.Kind
                || oTarget.Anchor.Scope != iTarget.Anchor.Scope
                || oTarget.Anchor.Unid != iTarget.Anchor.Unid)
                sb.AppendLine($"  anchor mismatch: {key} oracle={oTarget.Anchor} ir={iTarget.Anchor}");
            if (oTarget.PartUri != iTarget.PartUri)
                sb.AppendLine($"  partUri mismatch: {key} oracle={oTarget.PartUri} ir={iTarget.PartUri}");
            if (oTarget.TextPreview != iTarget.TextPreview)
                sb.AppendLine($"  textPreview mismatch: {key} oracle='{oTarget.TextPreview}' ir='{iTarget.TextPreview}'");
            // AutoNumberPrefix now resolved on the IR (the reader captures the live-package marker),
            // so it participates in the comparison (M1.4-T3 removed the exclusion).
            if (oTarget.AutoNumberPrefix != iTarget.AutoNumberPrefix)
                sb.AppendLine($"  autoNumberPrefix mismatch: {key} oracle='{oTarget.AutoNumberPrefix}' ir='{iTarget.AutoNumberPrefix}'");
        }

        // Body anchors the IR produced that the oracle did not.
        foreach (var key in ir.Keys.Where(k => ir[k].Anchor.Scope == "body"))
            if (!oracleBody.ContainsKey(key))
                sb.AppendLine($"  extra in IR: {key}");

        diff = sb.ToString();
        return diff.Length == 0;
    }

    // --- diff artifact ----------------------------------------------------

    private static void WriteDiff(string dir, string fixtureName, string oracleMd, string irMd, string indexDiff)
    {
        const int maxLines = 60;
        var sb = new StringBuilder();
        sb.AppendLine($"# Equivalence diff: {fixtureName}");
        sb.AppendLine();
        if (indexDiff.Length > 0)
        {
            sb.AppendLine("## Anchor-index diff (body, AutoNumberPrefix excluded)");
            sb.AppendLine(indexDiff);
        }
        sb.AppendLine("## Markdown unified diff (first " + maxLines + " differing lines)");
        var oLines = oracleMd.Replace("\r\n", "\n").Split('\n');
        var iLines = irMd.Replace("\r\n", "\n").Split('\n');
        int shown = 0;
        int max = Math.Max(oLines.Length, iLines.Length);
        for (int i = 0; i < max && shown < maxLines; i++)
        {
            var o = i < oLines.Length ? oLines[i] : "<EOF>";
            var n = i < iLines.Length ? iLines[i] : "<EOF>";
            if (!string.Equals(o, n, StringComparison.Ordinal))
            {
                sb.AppendLine($"@@ line {i + 1}");
                sb.AppendLine($"- {o}");
                sb.AppendLine($"+ {n}");
                shown++;
            }
        }
        if (shown == 0) sb.AppendLine("(markdown equal; index-only diff)");

        var safe = fixtureName.Replace(Path.DirectorySeparatorChar, '_').Replace(Path.AltDirectorySeparatorChar, '_');
        File.WriteAllText(Path.Combine(dir, safe + ".diff"), sb.ToString());
    }

    private static string ArtifactsDir([CallerFilePath] string thisFile = "") =>
        Path.Combine(Path.GetDirectoryName(thisFile)!, "EquivalenceArtifacts");

    private static bool CanOpen(FileInfo file)
    {
        try
        {
            using var fs = file.OpenRead();
            using var _ = WordprocessingDocument.Open(fs, false);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
