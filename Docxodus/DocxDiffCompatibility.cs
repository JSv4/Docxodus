#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

/// <summary>How thoroughly DocxDiff has been verified on a document feature class.</summary>
public enum DocxDiffCoverage
{
    /// <summary>No dedicated fidelity coverage — treat DocxDiff output as unverified for this construct.</summary>
    Untested,
    /// <summary>Some coverage, but a documented gap or risk remains.</summary>
    Partial,
    /// <summary>A dedicated fidelity campaign verified this construct; no warning is raised.</summary>
    Covered,
}

/// <summary>One feature class DocxDiff may or may not handle faithfully, with its current coverage status.</summary>
public sealed record DocxDiffFeature(
    string Id,
    string DisplayName,
    DocxDiffCoverage Coverage,
    string Note,
    string? CampaignRef = null);

/// <summary>A catalog feature found PRESENT in a document while NOT <see cref="DocxDiffCoverage.Covered"/>.</summary>
public sealed record DocxDiffCompatibilityFinding(DocxDiffFeature Feature, int Occurrences);

/// <summary>The result of inspecting a document against the DocxDiff feature catalog.</summary>
public sealed record DocxDiffCompatibilityReport(
    IReadOnlyList<DocxDiffCompatibilityFinding> Warnings,
    IReadOnlyList<DocxDiffFeature> CoveredPresent)
{
    /// <summary>True iff at least one under-tested feature is present in the inspected document.</summary>
    public bool HasWarnings => Warnings.Count > 0;

    /// <summary>A human-readable multi-line summary for logging — one line per warning.</summary>
    public string Summarize()
    {
        if (Warnings.Count == 0)
            return "DocxDiff compatibility: no under-tested features detected.";
        var sb = new StringBuilder();
        sb.AppendLine($"DocxDiff compatibility: {Warnings.Count} under-tested feature(s) detected:");
        foreach (var w in Warnings)
            sb.AppendLine($"  [{w.Feature.Coverage}] {w.Feature.DisplayName} (x{w.Occurrences}) - {w.Feature.Note}");
        return sb.ToString().TrimEnd();
    }
}

/// <summary>Thrown by a <see cref="DocxDiff"/> entry point when an input document contains an under-tested
/// feature and <see cref="DocxDiffSettings.ThrowOnCompatibilityWarning"/> is set. The <see cref="Report"/>
/// carries the structured warnings; the message is its summary.</summary>
public sealed class DocxDiffCompatibilityException : DocxodusException
{
    public DocxDiffCompatibilityReport Report { get; }

    public DocxDiffCompatibilityException(DocxDiffCompatibilityReport report)
        : base((report ?? throw new ArgumentNullException(nameof(report))).Summarize()) => Report = report;
}

/// <summary>
/// A rudimentary document-feature classifier for DocxDiff. <see cref="Inspect(byte[])"/> scans a DOCX for
/// constructs DocxDiff has not had a fidelity campaign for and returns them as warnings; <see cref="Catalog"/>
/// is the same status-tagged feature list independent of any document — the campaign roadmap. The classifier
/// does NOT change <see cref="DocxDiff"/>; a caller runs it as a pre-flight and decides whether to log, gate,
/// or throw. As a campaign lands, flip the catalog entry's <see cref="DocxDiffCoverage"/> to
/// <see cref="DocxDiffCoverage.Covered"/> and set its CampaignRef.
/// </summary>
public static class DocxDiffCompatibility
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";

    private sealed record Entry(DocxDiffFeature Feature, Func<IReadOnlyList<XElement>, int> Detect);

    // ---- detection helpers ----------------------------------------------------------------------

    /// <summary>Total count of elements across all content roots whose name is one of <paramref name="names"/>.</summary>
    private static int Count(IReadOnlyList<XElement> roots, params XName[] names)
    {
        var set = new HashSet<XName>(names);
        return roots.Sum(r => r.DescendantsAndSelf().Count(e => set.Contains(e.Name)));
    }

    /// <summary>Every field instruction string in the document: each <c>w:fldSimple/@w:instr</c> plus each
    /// <c>w:instrText</c> run value (split-run instructions are matched by keyword regex, so this is sufficient).</summary>
    private static IEnumerable<string> FieldInstructions(IReadOnlyList<XElement> roots)
    {
        foreach (var r in roots)
        {
            foreach (var f in r.DescendantsAndSelf(W + "fldSimple"))
                if ((string?)f.Attribute(W + "instr") is { } i)
                    yield return i;
            foreach (var it in r.DescendantsAndSelf(W + "instrText"))
                yield return it.Value;
        }
    }

    private static bool AnyInstr(IReadOnlyList<XElement> roots, string pattern) =>
        FieldInstructions(roots).Any(i => Regex.IsMatch(i, pattern, RegexOptions.IgnoreCase));

    private static int CountInstr(IReadOnlyList<XElement> roots, string pattern) =>
        FieldInstructions(roots).Count(i => Regex.IsMatch(i, pattern, RegexOptions.IgnoreCase));

    private static bool AnyHyperlinkAnchor(IReadOnlyList<XElement> roots) =>
        roots.Any(r => r.DescendantsAndSelf(W + "hyperlink").Any(h => h.Attribute(W + "anchor") != null));

    // ---- the catalog (feature + detector) -------------------------------------------------------

    private static readonly Entry[] _entries =
    {
        new(new("bookmarksCrossRefs", "Bookmarks & internal cross-references", DocxDiffCoverage.Covered,
                "Bookmark/cross-reference fidelity campaign.", "5be37a1"),
            roots =>
            {
                int bk = Count(roots, W + "bookmarkStart");
                bool refs = AnyInstr(roots, @"\b(REF|PAGEREF|NOTEREF)\b") || AnyHyperlinkAnchor(roots);
                return bk > 0 && refs ? bk : 0;
            }),
        new(new("footnotesEndnotes", "Footnotes & endnotes", DocxDiffCoverage.Covered,
                "Footnote/endnote + note-in-note campaigns.", "#239/#240"),
            roots => Count(roots, W + "footnoteReference", W + "endnoteReference")),
        new(new("comments", "Comments", DocxDiffCoverage.Covered,
                "Comment fidelity campaign: fine per-word markup on edited commented paragraphs + "
                + "id/range/reference/definition + threaded reply integrity.", "#243"),
            roots => Count(roots, W + "commentReference")),
        new(new("complexFields", "Complex fields (TOC/SEQ/INDEX/...)", DocxDiffCoverage.Partial,
                "REF/PAGEREF are solid; other field types are uncampaigned."),
            roots => CountInstr(roots, @"\b(TOC|INDEX|SEQ|HYPERLINK|STYLEREF|MERGEFIELD|DOCPROPERTY|INCLUDETEXT)\b")),
        new(new("contentControls", "Content controls (w:sdt)", DocxDiffCoverage.Partial,
                "Only an inline-SDT insertion scenario is tested; nested/block/data-bound uncampaigned."),
            roots => Count(roots, W + "sdt")),
        new(new("math", "Office Math (OMML)", DocxDiffCoverage.Partial,
                "Round-trips in a single fixture; no fidelity campaign."),
            roots => Count(roots, M + "oMath")),
        new(new("drawingml", "DrawingML (images/charts/SmartArt)", DocxDiffCoverage.Partial,
                "Inline-image media import is tested; charts/SmartArt/anchored drawings uncampaigned."),
            roots => Count(roots, W + "drawing")),
        new(new("textboxes", "Textboxes", DocxDiffCoverage.Untested,
                "IrTextbox modeling exists but has had no campaign."),
            roots => Count(roots, W + "txbxContent", V + "textbox")),
        new(new("rtlComplexScript", "RTL / complex script", DocxDiffCoverage.Untested,
                "No RTL/complex-script coverage."),
            roots => Count(roots, W + "bidi", W + "rtl", W + "cs")),
        new(new("oleEmbeddedObjects", "OLE / embedded objects", DocxDiffCoverage.Untested,
                "No coverage."),
            roots => Count(roots, W + "object", O + "OLEObject")),
        new(new("revisionsInInput", "Pre-existing tracked changes in input", DocxDiffCoverage.Covered,
                "Characterized + pinned: the engine diffs the ACCEPTED VIEW of each input (rule N13), so the "
                + "produced body carries only THIS diff's revisions. Pre-existing markup in carried-over parts "
                + "(headers/footers, unchanged footnotes/endnotes, styles, comments) is passed through verbatim "
                + "unless DocxDiffSettings.PreAcceptInputRevisions is set, which accepts BOTH inputs first so the "
                + "result is revision-free except for this diff and round-trips against the accepted view in body/"
                + "header/footer/note/style scopes. The flag cleans exactly the parts RevisionProcessor."
                + "AcceptRevisions processes; carried-over parts it does NOT process — the comments part and the "
                + "glossary (building-blocks) document part — keep their pre-existing revisions. "
                + "Two honest costs of accept-all: it flattens prior authorship and change boundaries (you "
                + "lose who edited what and where), and 'accept all' is itself a policy that overrides any change a "
                + "prior reviewer had left unaccepted (effectively rejected). See ir_diff_engine.md + "
                + "ooxml_corner_cases.md.", "PreAcceptInputRevisions"),
            roots => Count(roots, W + "ins", W + "del", W + "moveFrom", W + "moveTo")),
    };

    /// <summary>The full feature catalog — the campaign roadmap, independent of any document. Filter to
    /// <see cref="DocxDiffCoverage.Untested"/> then <see cref="DocxDiffCoverage.Partial"/> for what to tackle next.</summary>
    public static IReadOnlyList<DocxDiffFeature> Catalog { get; } = _entries.Select(e => e.Feature).ToList();

    // ---- inspection -----------------------------------------------------------------------------

    /// <summary>Inspect a <see cref="WmlDocument"/> for under-tested features.</summary>
    public static DocxDiffCompatibilityReport Inspect(WmlDocument doc) =>
        Inspect((doc ?? throw new ArgumentNullException(nameof(doc))).DocumentByteArray);

    /// <summary>Inspect raw DOCX bytes for under-tested features. Throws if the bytes are not a readable OOXML
    /// package (corrupt, encrypted, or non-DOCX) — the same input <see cref="DocxDiff.Compare(WmlDocument,
    /// WmlDocument, DocxDiffSettings?)"/> would itself reject; returns an EMPTY report for a package with no main
    /// document part.</summary>
    public static DocxDiffCompatibilityReport Inspect(byte[] docxBytes)
    {
        if (docxBytes == null) throw new ArgumentNullException(nameof(docxBytes));
        var roots = ContentRoots(docxBytes);
        var warnings = new List<DocxDiffCompatibilityFinding>();
        var covered = new List<DocxDiffFeature>();
        foreach (var e in _entries)
        {
            int n = e.Detect(roots);
            if (n <= 0)
                continue;
            if (e.Feature.Coverage == DocxDiffCoverage.Covered)
                covered.Add(e.Feature);
            else
                warnings.Add(new DocxDiffCompatibilityFinding(e.Feature, n));
        }
        return new DocxDiffCompatibilityReport(warnings, covered);
    }

    /// <summary>The XML roots of every diff-relevant content part: main body, headers, footers, footnotes,
    /// endnotes. Each is loaded fully (detached) so it survives the package being closed.</summary>
    private static IReadOnlyList<XElement> ContentRoots(byte[] bytes)
    {
        var roots = new List<XElement>();
        using var ms = new MemoryStream(bytes, writable: false);
        using var wdoc = WordprocessingDocument.Open(ms, false);
        var main = wdoc.MainDocumentPart;
        if (main == null)
            return roots;

        void Add(OpenXmlPart? part)
        {
            if (part == null) return;
            using var s = part.GetStream(FileMode.Open, FileAccess.Read);
            var x = XDocument.Load(s);
            if (x.Root != null)
                roots.Add(x.Root);
        }

        Add(main);
        foreach (var h in main.HeaderParts) Add(h);
        foreach (var f in main.FooterParts) Add(f);
        Add(main.FootnotesPart);
        Add(main.EndnotesPart);
        return roots;
    }
}
