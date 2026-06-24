#nullable enable
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;

namespace DiffHarness;

/// <summary>
/// Runs <see cref="DocxDiff"/> on a (left, right) pair and emits all artifacts needed to compare our
/// output against LibreOffice's: the tracked-changes document (<c>ours.docx</c>), the consumer revision
/// list in both granularities, the edit-script JSON, and a round-trip verification report
/// (<c>accept(ours) == right</c> / <c>reject(ours) == left</c> at the text level).
/// </summary>
internal static class DiffRunner
{
    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true };

    public static RoundTripReport Run(string leftPath, string rightPath, string outDir, string author)
    {
        Directory.CreateDirectory(outDir);
        var leftBytes = File.ReadAllBytes(leftPath);
        var rightBytes = File.ReadAllBytes(rightPath);
        var left = new WmlDocument(leftPath, leftBytes);
        var right = new WmlDocument(rightPath, rightBytes);

        var settings = new DocxDiffSettings { AuthorForRevisions = author };

        // 1) Tracked-changes document.
        var ours = DocxDiff.Compare(left, right, settings);
        File.WriteAllBytes(Path.Combine(outDir, "ours.docx"), ours.DocumentByteArray);

        // 2) Revisions in both granularities (fine = engine-native; compat = WmlComparer-comparable).
        var revsFine = DocxDiff.GetRevisions(left, right, settings);
        var compatSettings = new DocxDiffSettings
        {
            AuthorForRevisions = author,
            RevisionGranularity = DocxDiffRevisionGranularity.WmlComparerCompatible,
        };
        var revsCompat = DocxDiff.GetRevisions(left, right, compatSettings);
        File.WriteAllText(Path.Combine(outDir, "revisions.fine.json"),
            JsonSerializer.Serialize(revsFine.Select(RevDto.From), Json));
        File.WriteAllText(Path.Combine(outDir, "revisions.compat.json"),
            JsonSerializer.Serialize(revsCompat.Select(RevDto.From), Json));

        // 3) Edit script as data.
        File.WriteAllText(Path.Combine(outDir, "editscript.json"),
            DocxDiff.GetEditScriptJson(left, right, settings));

        // 4) Round-trip verification (body & notes exact; header/footer dedup-set + bloat metric).
        var leftText = TextExtractor.Extract(leftBytes);
        var rightText = TextExtractor.Extract(rightBytes);
        var accept = TextExtractor.Extract(RevisionAccepter.AcceptRevisions(ours).DocumentByteArray);
        var reject = TextExtractor.Extract(RevisionProcessor.RejectRevisions(ours).DocumentByteArray);

        // 5) Footnote/endnote STRUCTURE: the Compare output's note ids must be unique and every body note
        //    reference must resolve to exactly one definition — the invariant the duplicate-id / dropped-ref
        //    corruption violated (schema-invalid output that LibreOffice silently dropped a note from).
        var note = CheckNoteStructure(ours.DocumentByteArray);

        var report = new RoundTripReport(
            AcceptBodyEqualsRight: accept.BodyEquals(rightText),
            RejectBodyEqualsLeft: reject.BodyEquals(leftText),
            AcceptNotesEqualRight: accept.NotesEqual(rightText),
            RejectNotesEqualLeft: reject.NotesEqual(leftText),
            AcceptHdrFtrSetEqualsRight: accept.HeaderFooterSetEquals(rightText),
            RejectHdrFtrSetEqualsLeft: reject.HeaderFooterSetEquals(leftText),
            RevisionCountFine: revsFine.Count,
            RevisionCountCompat: revsCompat.Count,
            // bloat: header/footer part count of the produced doc vs the originals
            HdrFtrPartsOurs: accept.HeaderFooterPartCount,
            HdrFtrPartsOriginal: rightText.HeaderFooterPartCount,
            AcceptBodyFirstDiff: accept.BodyEquals(rightText) ? null : FirstDiff(accept.Body, rightText.Body),
            RejectBodyFirstDiff: reject.BodyEquals(leftText) ? null : FirstDiff(reject.Body, leftText.Body),
            OursFootnoteIdsUnique: note.FootnoteIdsUnique,
            OursEndnoteIdsUnique: note.EndnoteIdsUnique,
            OursFootnoteRefsAllResolve: note.FootnoteRefsAllResolve,
            OursEndnoteRefsAllResolve: note.EndnoteRefsAllResolve);
        File.WriteAllText(Path.Combine(outDir, "roundtrip.json"),
            JsonSerializer.Serialize(report, Json));
        return report;
    }

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>Footnote/endnote STRUCTURE check on a produced document: definition ids are unique and every
    /// body reference resolves to exactly one definition. Mirrors <c>DocxDiffScenarioTests</c>'s in-process
    /// invariant so the LibreOffice-corpus run flags the same corruption headlessly.</summary>
    private static NoteStructure CheckNoteStructure(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart;
        if (main == null)
            return new NoteStructure(true, true, true, true);

        var body = main.GetXDocument().Root?.Element(W + "body");

        (bool unique, bool resolve) Check(OpenXmlPart? part, string defName, string refName)
        {
            var defIds = part?.GetXDocument().Root?.Elements(W + defName)
                             .Select(e => (string?)e.Attribute(W + "id"))
                             .Where(id => id != null).Select(id => id!).ToList()
                         ?? new List<string>();
            bool unique = defIds.Count == defIds.Distinct().Count();
            var counts = defIds.GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());
            bool resolve = body == null || body.Descendants(W + refName)
                .Select(r => (string?)r.Attribute(W + "id"))
                .All(id => id != null && counts.TryGetValue(id, out var n) && n == 1);
            return (unique, resolve);
        }

        var (fnU, fnR) = Check(main.FootnotesPart, "footnote", "footnoteReference");
        var (enU, enR) = Check(main.EndnotesPart, "endnote", "endnoteReference");
        return new NoteStructure(fnU, enU, fnR, enR);
    }

    private readonly record struct NoteStructure(
        bool FootnoteIdsUnique, bool EndnoteIdsUnique,
        bool FootnoteRefsAllResolve, bool EndnoteRefsAllResolve);

    /// <summary>Return a short human-readable description of the first divergence, or null if equal.</summary>
    private static string? FirstDiff(string a, string b)
    {
        if (a == b) return null;
        int n = Math.Min(a.Length, b.Length);
        int i = 0;
        while (i < n && a[i] == b[i]) i++;
        string Win(string s, int at) =>
            s.Substring(Math.Max(0, at - 40), Math.Min(s.Length, at + 40) - Math.Max(0, at - 40))
             .Replace("\n", "\\n");
        return $"at char {i} (lenA={a.Length}, lenB={b.Length}): " +
               $"A=…{Win(a, i)}…  |  B=…{Win(b, i)}…";
    }
}

internal sealed record RoundTripReport(
    bool AcceptBodyEqualsRight,
    bool RejectBodyEqualsLeft,
    bool AcceptNotesEqualRight,
    bool RejectNotesEqualLeft,
    bool AcceptHdrFtrSetEqualsRight,
    bool RejectHdrFtrSetEqualsLeft,
    int RevisionCountFine,
    int RevisionCountCompat,
    int HdrFtrPartsOurs,
    int HdrFtrPartsOriginal,
    string? AcceptBodyFirstDiff,
    string? RejectBodyFirstDiff,
    bool OursFootnoteIdsUnique = true,
    bool OursEndnoteIdsUnique = true,
    bool OursFootnoteRefsAllResolve = true,
    bool OursEndnoteRefsAllResolve = true)
{
    /// <summary>Content-level round-trip success: body + notes + header/footer (dedup) all match.</summary>
    public bool ContentClean =>
        AcceptBodyEqualsRight && RejectBodyEqualsLeft &&
        AcceptNotesEqualRight && RejectNotesEqualLeft &&
        AcceptHdrFtrSetEqualsRight && RejectHdrFtrSetEqualsLeft;

    /// <summary>Footnote/endnote structural soundness of the produced document: unique definition ids and
    /// every body note reference resolving to exactly one definition (the duplicate-id / dropped-ref guard).</summary>
    public bool NoteStructureClean =>
        OursFootnoteIdsUnique && OursEndnoteIdsUnique &&
        OursFootnoteRefsAllResolve && OursEndnoteRefsAllResolve;
}

internal sealed record RevDto(
    string Type, string Text, string? LeftAnchor, string? RightAnchor,
    int? MoveGroupId, bool? IsMoveSource, string[]? ChangedFormat)
{
    public static RevDto From(DocxDiffRevision r) => new(
        r.Type.ToString(), r.Text, r.LeftAnchor, r.RightAnchor,
        r.MoveGroupId, r.IsMoveSource,
        r.FormatChange?.ChangedPropertyNames.ToArray());
}
