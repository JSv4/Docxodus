#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests;

/// <summary>
/// Headline real-document verification of bookmark / internal cross-reference fidelity, exercising the dense
/// bookmark-backed cross-references of the NVCA model contracts (COI: ~392 bookmarks / 82 REF fields; SPA:
/// ~192 bookmarks / 68 REF fields). For each contract a variant is produced by editing many bookmark-BEARING
/// and bookmark-REFERENCING paragraphs, then <see cref="DocxDiff.Compare"/> must hold the same structural
/// invariants as the synthetic corpus AND emit a schema-valid document LibreOffice can load and field-refresh.
/// <para>The fixtures live at the paths the campaign was scoped against; the test no-ops (passes) when they are
/// absent so the suite stays portable. When present it also writes the Compare outputs to the scratchpad so the
/// LibreOffice oracle can verify them out of process.</para>
/// </summary>
public class DocxDiffBookmarkRealDocTests
{
    private readonly ITestOutputHelper _out;
    public DocxDiffBookmarkRealDocTests(ITestOutputHelper o) => _out = o;

    public static IEnumerable<object[]> Contracts() => new[]
    {
        new object[] { "/home/jman/Downloads/NVCA-Model-COI-10-1-2025.docx", "coi" },
        new object[] { "/home/jman/Downloads/NVCA-Model-SPA-10-28-2025-1.docx", "spa" },
    };

    private const string ScratchDir =
        "/tmp/claude-1000/-home-jman-Code-Docxodus/393f24c6-24d8-495e-bd80-9afe2474b645/scratchpad/nvca";

    [Theory]
    [MemberData(nameof(Contracts))]
    public void NvcaContract_BookmarkAndCrossReferenceFidelity(string path, string tag)
    {
        if (!File.Exists(path))
        {
            _out.WriteLine($"[skip] fixture absent: {path}");
            return; // portable: no-op when the NVCA fixtures are not on this machine
        }

        var leftBytes = File.ReadAllBytes(path);
        var left = new WmlDocument(path, leftBytes);
        var rightBytes = EditBookmarkAndReferenceParagraphs(leftBytes);
        var right = new WmlDocument(path, rightBytes);

        var result = DocxDiff.Compare(left, right);

        // emit artifacts for the LibreOffice oracle
        Directory.CreateDirectory(ScratchDir);
        File.WriteAllBytes(Path.Combine(ScratchDir, $"{tag}-left.docx"), leftBytes);
        File.WriteAllBytes(Path.Combine(ScratchDir, $"{tag}-right.docx"), rightBytes);
        File.WriteAllBytes(Path.Combine(ScratchDir, $"{tag}-compare.docx"), result.DocumentByteArray);

        // (1) schema validity — no NEW errors vs the inputs. Keys are PATH-INSENSITIVE (id + description, not
        // XPath): the diff legitimately relocates runs, so a PRE-EXISTING source quirk (the NVCA COI carries 65
        // <w:w w:val="0"> character-scale runs the strict validator rejects but Word tolerates) reappears at a
        // shifted run-path and must NOT count as new — it auto-excludes via baseErrors (the source carries the
        // identical error type). The duplicate-bookmark-id and numbering-order issues this campaign FIXED are
        // caught here; any new bookmark/field error would surface.
        var baseErrors = SchemaErrors(leftBytes).Concat(SchemaErrors(rightBytes)).ToHashSet();
        var newErrors = SchemaErrors(result.DocumentByteArray).Where(e => !baseErrors.Contains(e)).ToList();
        Assert.True(newErrors.Count == 0,
            $"[{tag}] Compare introduced {newErrors.Count} new schema error(s): {string.Join(" | ", newErrors.Take(6))}");

        // (2) bookmark-structure soundness of the Compare output: unique ids + 1:1 pairing
        var (startIds, endIds) = BookmarkIds(result.DocumentByteArray);
        Assert.True(startIds.Count == startIds.Distinct().Count(),
            $"[{tag}] duplicate bookmarkStart id(s): {Dup(startIds)}");
        Assert.True(endIds.Count == endIds.Distinct().Count(),
            $"[{tag}] duplicate bookmarkEnd id(s): {Dup(endIds)}");
        Assert.True(startIds.OrderBy(x => x).SequenceEqual(endIds.OrderBy(x => x)),
            $"[{tag}] bookmarkStart/End ids not 1:1 paired (starts:{startIds.Count} ends:{endIds.Count})");

        // (3) bookmark-structure round-trip: accept ≡ right, reject ≡ left (names + reference resolution)
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);

        // body-TEXT round-trip too (accept ≡ right, reject ≡ left): guards the noBreakHyphen/softHyphen/sym
        // char-accounting class of bug a bookmark/field-dense edit surfaces.
        Assert.Equal(BodyText(rightBytes), BodyText(accepted.DocumentByteArray));
        Assert.Equal(BodyText(leftBytes), BodyText(rejected.DocumentByteArray));
        _out.WriteLine($"[{tag}] bookmarks={startIds.Count} refs={Projection(rightBytes).RefCount} " +
            $"accept-dangling={Projection(AsBytes(accepted)).Dangling} reject-dangling={Projection(AsBytes(rejected)).Dangling}");
        File.WriteAllText(Path.Combine(ScratchDir, $"{tag}-proj-right.txt"), Projection(rightBytes).Signature);
        File.WriteAllText(Path.Combine(ScratchDir, $"{tag}-proj-accept.txt"), Projection(AsBytes(accepted)).Signature);
        File.WriteAllText(Path.Combine(ScratchDir, $"{tag}-proj-left.txt"), Projection(leftBytes).Signature);
        File.WriteAllText(Path.Combine(ScratchDir, $"{tag}-proj-reject.txt"), Projection(AsBytes(rejected)).Signature);
        Assert.Equal(Projection(rightBytes).Signature, Projection(AsBytes(accepted)).Signature);
        Assert.Equal(Projection(leftBytes).Signature, Projection(AsBytes(rejected)).Signature);

        // every internal reference in the Compare output resolves to a present bookmark name (no dangling)
        Assert.Equal(0, Projection(result.DocumentByteArray).Dangling);
    }

    // ---- variant construction -------------------------------------------------------------------

    /// <summary>Edit many bookmark-BEARING and bookmark-REFERENCING body paragraphs: append a sentinel word to a
    /// text run of each, producing a real per-paragraph diff over the exact paragraphs that carry bookmark markers
    /// and REF/PAGEREF field instructions.</summary>
    private static byte[] EditBookmarkAndReferenceParagraphs(byte[] bytes)
    {
        using var ms = new MemoryStream();
        ms.Write(bytes, 0, bytes.Length);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            int bk = 0, rf = 0;
            foreach (var p in body.Descendants<Paragraph>())
            {
                bool carriesBookmark = p.Descendants<BookmarkStart>().Any();
                bool carriesRef = p.Descendants<FieldCode>().Any(fc =>
                    Regex.IsMatch(fc.Text ?? "", @"\b(REF|PAGEREF|NOTEREF)\b"));
                if (!carriesBookmark && !carriesRef) continue;
                // edit the first visible text run of the paragraph (skip field-internal display runs is fine —
                // any real text change makes a diff). Cap the count so the run stays fast but representative.
                var t = p.Descendants<Text>().FirstOrDefault(x => x.Text.Trim().Length > 3);
                if (t == null) continue;
                if (carriesBookmark && bk >= 40 && carriesRef && rf >= 40) continue;
                t.Text = t.Text + " (amended)";
                if (carriesBookmark) bk++;
                if (carriesRef) rf++;
            }
        }
        return ms.ToArray();
    }

    // ---- observers ------------------------------------------------------------------------------

    private static byte[] AsBytes(WmlDocument d) => d.DocumentByteArray;

    private static string BodyText(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var w = WordprocessingDocument.Open(ms, false);
        var body = w.MainDocumentPart?.Document?.Body;
        return body is null ? "" : string.Concat(body.Descendants<Text>().Select(t => t.Text));
    }

    private static (List<int> Starts, List<int> Ends) BookmarkIds(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var w = WordprocessingDocument.Open(ms, false);
        var body = w.MainDocumentPart?.Document?.Body;
        var s = new List<int>();
        var e = new List<int>();
        if (body != null)
        {
            foreach (var b in body.Descendants().Where(x => x.LocalName == "bookmarkStart"))
                if (int.TryParse(Attr(b, "id"), out var v)) s.Add(v);
            foreach (var b in body.Descendants().Where(x => x.LocalName == "bookmarkEnd"))
                if (int.TryParse(Attr(b, "id"), out var v)) e.Add(v);
        }
        return (s, e);
    }

    private record Proj(string Signature, int RefCount, int Dangling);

    /// <summary>Bookmark/reference structural projection: sorted bookmark name set + each internal reference's
    /// (kind,target,resolves) triple. Equal projections ⇒ equal bookmark/cross-reference structure.</summary>
    private static Proj Projection(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var w = WordprocessingDocument.Open(ms, false);
        var body = w.MainDocumentPart?.Document?.Body;
        if (body == null) return new Proj("", 0, 0);

        var names = body.Descendants().Where(e => e.LocalName == "bookmarkStart")
            .Select(e => Attr(e, "name")).Where(n => n.Length > 0)
            .Distinct().OrderBy(n => n, StringComparer.Ordinal).ToList();
        var nameSet = new HashSet<string>(names);

        var refs = InternalRefs(body).ToList();
        int dangling = refs.Count(r => !nameSet.Contains(r.Target));
        var refSig = refs.Select(r => $"{r.Kind}:{r.Target}:{(nameSet.Contains(r.Target) ? "ok" : "X")}")
            .OrderBy(s => s, StringComparer.Ordinal);
        return new Proj($"N[{string.Join(",", names)}]R[{string.Join(",", refSig)}]", refs.Count, dangling);
    }

    private record Ref(string Kind, string Target);

    private static IEnumerable<Ref> InternalRefs(OpenXmlElement body)
    {
        var list = new List<Ref>();
        foreach (var h in body.Descendants<Hyperlink>())
            if (!string.IsNullOrEmpty(h.Anchor?.Value)) list.Add(new Ref("anchor", h.Anchor!.Value!));
        foreach (var f in body.Descendants<SimpleField>())
            if (Parse(f.Instruction?.Value ?? "") is { } t) list.Add(t);
        var instr = "";
        foreach (var e in body.Descendants().Where(x => x.LocalName is "fldChar" or "instrText" or "delInstrText"))
        {
            if (e.LocalName == "fldChar")
            {
                var ty = Attr(e, "fldCharType");
                if (ty == "begin") instr = "";
                else if (ty is "separate" or "end") { if (Parse(instr) is { } t) list.Add(t); instr = ""; }
            }
            else instr += e.InnerText;
        }
        return list;
    }

    private static Ref? Parse(string instr)
    {
        var toks = Regex.Matches(instr.Trim(), "\"[^\"]*\"|\\S+").Select(m => m.Value).ToList();
        if (toks.Count == 0) return null;
        var kw = toks[0].ToUpperInvariant();
        if (kw is "REF" or "PAGEREF" or "NOTEREF") return toks.Count >= 2 ? new Ref(kw, toks[1].Trim('"')) : null;
        if (kw == "HYPERLINK")
            for (int i = 1; i < toks.Count - 1; i++)
                if (toks[i] == "\\l") return new Ref("HYPERLINK\\l", toks[i + 1].Trim('"'));
        return null;
    }

    private static string Attr(OpenXmlElement e, string n) =>
        e.GetAttributes().FirstOrDefault(a => a.LocalName == n).Value ?? "";

    private static string Dup(List<int> ids) =>
        string.Join(",", ids.GroupBy(x => x).Where(g => g.Count() > 1).Select(g => g.Key));

    private static HashSet<string> SchemaErrors(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var w = WordprocessingDocument.Open(ms, false);
        var v = new OpenXmlValidator(FileFormatVersions.Office2019);
        // PATH-INSENSITIVE key (id + description, no XPath): the diff relocates runs, so a pre-existing source
        // error must match across input↔output by TYPE, not position. A genuinely new error (e.g. a duplicate
        // bookmark id, whose description names the colliding value) still differs from anything in the source.
        return v.Validate(w).Select(e => $"{e.Id}: {e.Description}").ToHashSet();
    }
}
