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
using DocxodusDiffParityFixtures;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Bookmark / internal cross-reference structural integrity for <see cref="DocxDiff.Compare"/>, the
/// bookmark counterpart to <c>DocxDiffScenarioTests.Scenario_PreservesFootnoteStructure</c>. For every
/// shape in <see cref="DocxDiffBookmarkFixtures"/> the Compare output must be:
/// <list type="number">
/// <item><b>Schema valid</b> — no NEW OpenXml errors vs the inputs (catches the duplicate-bookmark-id
/// <c>Sem_UniqueAttributeValue</c>).</item>
/// <item><b>Bookmark-structurally sound in the intermediate</b> — bookmark ids are UNIQUE and every
/// <c>w:bookmarkStart</c> pairs 1:1 with a <c>w:bookmarkEnd</c> of the same id.</item>
/// <item><b>Round-trip at the bookmark-structure level</b> — the bookmark/reference projection of
/// <c>accept</c> equals RIGHT's and of <c>reject</c> equals LEFT's: names preserved, and every internal
/// reference (hyperlink <c>w:anchor</c>, and <c>REF</c>/<c>PAGEREF</c>/<c>NOTEREF</c>/<c>HYPERLINK \l</c>
/// field instructions in <c>w:instrText</c>/<c>w:fldSimple</c>) resolves exactly as it does in the source.
/// This is a structural assertion, never a text multiset.</item>
/// </list>
/// </summary>
public class DocxDiffBookmarkStructureTests
{
    public static IEnumerable<object[]> AllScenarios() =>
        DocxDiffBookmarkFixtures.Names().Select(n => new object[] { n });

    [Theory]
    [MemberData(nameof(AllScenarios))]
    public void Bookmark_Compare_IsSchemaValid(string scenario)
    {
        var (left, right) = DocxDiffBookmarkFixtures.Build(scenario);
        var result = DocxDiff.Compare(left, right);

        var baseErrors = SchemaErrors(left).Concat(SchemaErrors(right)).ToHashSet();
        var newErrors = SchemaErrors(result).Where(e => !baseErrors.Contains(e)).ToList();
        Assert.True(newErrors.Count == 0,
            $"[{scenario}] Compare introduced {newErrors.Count} new schema error(s): {string.Join(" | ", newErrors.Take(5))}");
    }

    [Theory]
    [MemberData(nameof(AllScenarios))]
    public void Bookmark_Compare_HasUniquePairedBookmarks(string scenario)
    {
        var (left, right) = DocxDiffBookmarkFixtures.Build(scenario);
        var result = DocxDiff.Compare(left, right);

        var (startIds, endIds, _) = BookmarkMarkers(result);

        // unique ids on each side
        Assert.True(startIds.Count == startIds.Distinct().Count(),
            $"[{scenario}] duplicate bookmarkStart id(s): [{string.Join(",", startIds)}]");
        Assert.True(endIds.Count == endIds.Distinct().Count(),
            $"[{scenario}] duplicate bookmarkEnd id(s): [{string.Join(",", endIds)}]");

        // 1:1 pairing: the set of start ids equals the set of end ids
        var onlyStart = startIds.Except(endIds).ToList();
        var onlyEnd = endIds.Except(startIds).ToList();
        Assert.True(onlyStart.Count == 0, $"[{scenario}] bookmarkStart id(s) without matching End: [{string.Join(",", onlyStart)}]");
        Assert.True(onlyEnd.Count == 0, $"[{scenario}] bookmarkEnd id(s) without matching Start: [{string.Join(",", onlyEnd)}]");
    }

    [Theory]
    [MemberData(nameof(AllScenarios))]
    public void Bookmark_RoundTrips_AtStructureLevel(string scenario)
    {
        var (left, right) = DocxDiffBookmarkFixtures.Build(scenario);
        var result = DocxDiff.Compare(left, right);

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);

        // body text round-trip (the WmlComparer output contract)
        Assert.Equal(BodyText(right), BodyText(accepted));
        Assert.Equal(BodyText(left), BodyText(rejected));

        // bookmark-structure round-trip: accept ≡ right, reject ≡ left
        Assert.Equal(Projection(right), Projection(accepted));
        Assert.Equal(Projection(left), Projection(rejected));
    }

    [Fact]
    public void Bookmark_IdentityDiff_IsCleanNoOp()
    {
        foreach (var name in DocxDiffBookmarkFixtures.Names())
        {
            var (left, _) = DocxDiffBookmarkFixtures.Build(name);
            // left vs left: no spurious bookmark churn, schema clean.
            var result = DocxDiff.Compare(left, left);
            Assert.Empty(SchemaErrors(result).Where(e => !SchemaErrors(left).Contains(e)));
            Assert.Equal(Projection(left), Projection(RevisionProcessor.AcceptRevisions(result)));
            Assert.Equal(Projection(left), Projection(RevisionProcessor.RejectRevisions(result)));
        }
    }

    // ---- bookmark/reference projection ----------------------------------------------------------

    /// <summary>
    /// A renderer-independent structural snapshot of a document's bookmark layer: the sorted set of
    /// bookmark NAMES plus, for each internal reference (document order), the (kind, target, resolves)
    /// triple. Two documents with the same projection have the same bookmark/cross-reference structure
    /// regardless of ids, run splitting, or revision markup.
    /// </summary>
    private static string Projection(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var body = w.MainDocumentPart?.Document?.Body;
        if (body is null) return "names=[]; refs=[]";

        var names = body.Descendants().Where(e => e.LocalName == "bookmarkStart")
            .Select(e => AttrValue(e, "name"))
            .Where(n => n.Length > 0 && !n.StartsWith("_GoBack", StringComparison.Ordinal))
            .OrderBy(n => n, StringComparer.Ordinal).ToList();
        var nameSet = new HashSet<string>(names);

        var refs = InternalReferences(body)
            .Select(r => $"{r.Kind}:{r.Target}:{(nameSet.Contains(r.Target) ? "ok" : "dangling")}")
            .OrderBy(s => s, StringComparer.Ordinal).ToList();

        return $"names=[{string.Join(",", names)}]; refs=[{string.Join(",", refs)}]";
    }

    private record Ref(string Kind, string Target);

    /// <summary>Internal references in a body: hyperlink <c>w:anchor</c> + REF/PAGEREF/NOTEREF/HYPERLINK \l
    /// instructions from <c>w:fldSimple</c> and <c>w:instrText</c> field runs.</summary>
    private static IEnumerable<Ref> InternalReferences(OpenXmlElement body)
    {
        var result = new List<Ref>();

        // internal hyperlinks (have an anchor, no external r:id)
        foreach (var h in body.Descendants<Hyperlink>())
        {
            var anchor = h.Anchor?.Value;
            if (!string.IsNullOrEmpty(anchor))
                result.Add(new Ref("anchor", anchor!));
        }

        // fldSimple instructions
        foreach (var f in body.Descendants<SimpleField>())
        {
            var t = ParseRefTarget(f.Instruction?.Value ?? "");
            if (t is not null) result.Add(t);
        }

        // fldChar fields: concatenate consecutive instrText/delInstrText into one instruction string.
        // (In accept/reject docs fields are clean; in the intermediate this is a best-effort scan.)
        var instr = "";
        foreach (var e in body.Descendants().Where(x =>
                     x.LocalName is "fldChar" or "instrText" or "delInstrText"))
        {
            if (e.LocalName == "fldChar")
            {
                var type = AttrValue(e, "fldCharType");
                if (type == "begin") instr = "";
                else if (type is "separate" or "end")
                {
                    var t = ParseRefTarget(instr);
                    if (t is not null) result.Add(t);
                    instr = "";
                }
            }
            else
            {
                instr += e.InnerText;
            }
        }
        return result;
    }

    /// <summary>Parse a field instruction for a bookmark-targeting reference. Returns null if it is not
    /// a REF/PAGEREF/NOTEREF or HYPERLINK \l field.</summary>
    private static Ref? ParseRefTarget(string instr)
    {
        var toks = Tokenize(instr);
        if (toks.Count == 0) return null;
        var kw = toks[0].ToUpperInvariant();
        if (kw is "REF" or "PAGEREF" or "NOTEREF")
            return toks.Count >= 2 ? new Ref(kw, Unquote(toks[1])) : null;
        if (kw == "HYPERLINK")
        {
            // HYPERLINK \l "bookmark"  → internal anchor
            for (int i = 1; i < toks.Count - 1; i++)
                if (toks[i] == "\\l")
                    return new Ref("HYPERLINK\\l", Unquote(toks[i + 1]));
        }
        return null;
    }

    private static List<string> Tokenize(string s)
    {
        // split on whitespace, keeping quoted runs together
        var matches = Regex.Matches(s.Trim(), "\"[^\"]*\"|\\S+");
        return matches.Select(m => m.Value).ToList();
    }

    private static string Unquote(string s) => s.Trim('"');

    // ---- low-level observers --------------------------------------------------------------------

    private static (List<int> Starts, List<int> Ends, List<string> Names) BookmarkMarkers(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var body = w.MainDocumentPart?.Document?.Body;
        var starts = new List<int>();
        var ends = new List<int>();
        var names = new List<string>();
        if (body != null)
        {
            foreach (var e in body.Descendants().Where(x => x.LocalName == "bookmarkStart"))
            {
                if (int.TryParse(AttrValue(e, "id"), out var id)) starts.Add(id);
                var n = AttrValue(e, "name");
                if (n.Length > 0) names.Add(n);
            }
            foreach (var e in body.Descendants().Where(x => x.LocalName == "bookmarkEnd"))
                if (int.TryParse(AttrValue(e, "id"), out var id)) ends.Add(id);
        }
        return (starts, ends, names);
    }

    private static string AttrValue(OpenXmlElement e, string localName) =>
        e.GetAttributes().FirstOrDefault(a => a.LocalName == localName).Value ?? "";

    private static string BodyText(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var body = w.MainDocumentPart?.Document?.Body;
        return body is null ? "" : string.Concat(body.Descendants<Text>().Select(t => t.Text));
    }

    private static HashSet<string> SchemaErrors(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        return validator.Validate(w).Select(e => $"{e.Id}@{e.Path?.XPath}: {e.Description}").ToHashSet();
    }
}
