#nullable enable
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using DocxodusDiffParityFixtures;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// LibreOffice-free codification of the <c>tools/diffharness</c> verification campaign. For every edit
/// type × feature in <see cref="DocxDiffScenarioFixtures"/>, asserts <see cref="DocxDiff.Compare"/>'s
/// universal invariants — the durable correctness checks that need no external renderer:
/// <list type="number">
/// <item><b>Round-trip.</b> <c>AcceptRevisions(Compare) ≡ right</c> and <c>RejectRevisions(Compare) ≡ left</c>
/// at the body + note-store text level (the WmlComparer output contract). Header/footer scopes are
/// deliberately not diffed, so they are not part of this assertion — exactly as the oracle behaves.</item>
/// <item><b>No header/footer part duplication.</b> The produced document carries the SAME number of
/// header/footer parts as the base (regression guard for the F1 fix, generalized across every edit).</item>
/// <item><b>Schema validity.</b> The produced document introduces no NEW OpenXml schema errors vs the base.</item>
/// </list>
/// The campaign's narrow regression tests live in <see cref="DocxDiffTests"/>
/// (<c>GetRevisions_TableColumnChange…</c>, <c>Compare_InsertedContentControl…</c>) and
/// <c>IrMarkupRendererTests</c> (<c>Render_does_not_duplicate_header_parts…</c>); this is the broad matrix.
/// </summary>
public class DocxDiffScenarioTests
{
    public static IEnumerable<object[]> AllScenarios() =>
        DocxDiffScenarioFixtures.Names().Select(n => new object[] { n });

    // ---- universal invariants (hold for EVERY scenario) ----------------------------------------

    [Theory]
    [MemberData(nameof(AllScenarios))]
    public void Scenario_RoundTrips_AndDoesNotBloatHeaderFooterParts(string scenario)
    {
        var (left, right) = DocxDiffScenarioFixtures.Build(scenario);

        var result = DocxDiff.Compare(left, right);

        // Round-trip: accept ⇒ right, reject ⇒ left (body + note-store content).
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyText(right), BodyText(accepted));
        Assert.Equal(BodyText(left), BodyText(rejected));
        Assert.Equal(NoteTexts(right), NoteTexts(accepted));
        Assert.Equal(NoteTexts(left), NoteTexts(rejected));

        // No header/footer part duplication (the F1 fix, generalized).
        Assert.Equal(HeaderFooterPartCount(left), HeaderFooterPartCount(result));
    }

    // ---- schema validity (EVERY scenario — the footnote-referencing edits are now clean too) ----

    [Theory]
    [MemberData(nameof(AllScenarios))]
    public void Scenario_ProducesSchemaValidOutput(string scenario)
    {
        var (left, right) = DocxDiffScenarioFixtures.Build(scenario);
        var result = DocxDiff.Compare(left, right);

        var baseErrors = SchemaErrors(left);
        var newErrors = SchemaErrors(result).Where(e => !baseErrors.Contains(e)).ToList();
        Assert.True(newErrors.Count == 0,
            $"Compare introduced {newErrors.Count} new schema error(s): {string.Join(" | ", newErrors.Take(5))}");
    }

    // ---- footnote-structure integrity (oracle layer 2: id ↔ reference ↔ text) -------------------

    /// <summary>
    /// Footnote-structure round-trip — the structural counterpart to the text round-trip above. For EVERY
    /// scenario, <see cref="DocxDiff.Compare"/>'s output must be footnote-structurally sound:
    /// <list type="number">
    /// <item><b>Unique definition ids.</b> No two real <c>w:footnote</c> definitions share a <c>@w:id</c>.</item>
    /// <item><b>Every body reference resolves.</b> Each body <c>w:footnoteReference @w:id</c> names exactly one
    /// definition.</item>
    /// <item><b>Accept ≡ right / reject ≡ left at the id→ref→text level.</b> Resolving each body reference (in
    /// document order) to its definition text yields the RIGHT document's referenced-footnote sequence on accept
    /// and the LEFT's on reject — not just the order-insensitive multiset the text round-trip checks.</item>
    /// </list>
    /// This is the invariant the duplicate-id / dropped-reference corruption violated.
    /// </summary>
    [Theory]
    [MemberData(nameof(AllScenarios))]
    public void Scenario_PreservesFootnoteStructure(string scenario)
    {
        var (left, right) = DocxDiffScenarioFixtures.Build(scenario);
        var result = DocxDiff.Compare(left, right);

        // (1) unique definition ids in the Compare output.
        var ids = FootnoteIds(result);
        Assert.True(ids.Count == ids.Distinct().Count(),
            $"duplicate footnote id(s) in Compare output: [{string.Join(",", ids)}]");

        // (2) every body footnote reference resolves to exactly one definition.
        var unresolved = UnresolvedReferenceIds(result);
        Assert.True(unresolved.Count == 0,
            $"body footnote reference(s) resolve to ≠1 definition: [{string.Join(",", unresolved)}]");

        // (3) id → ref → text integrity: accept ≡ right, reject ≡ left in body-reference order.
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(ReferencedFootnoteTexts(right), ReferencedFootnoteTexts(accepted));
        Assert.Equal(ReferencedFootnoteTexts(left), ReferencedFootnoteTexts(rejected));
    }

    /// <summary>Sanity: the synthetic base is itself schema-valid and an identity diff is a clean no-op.</summary>
    [Fact]
    public void Base_IsSchemaValid_And_IdentityDiffRoundTrips()
    {
        var doc = DocxDiffScenarioFixtures.BaseDoc();
        Assert.Empty(SchemaErrors(doc));

        var result = DocxDiff.Compare(doc, doc);
        Assert.Equal(BodyText(doc), BodyText(RevisionProcessor.AcceptRevisions(result)));
        Assert.Equal(BodyText(doc), BodyText(RevisionProcessor.RejectRevisions(result)));
        Assert.Equal(HeaderFooterPartCount(doc), HeaderFooterPartCount(result));
        Assert.Empty(DocxDiff.GetRevisions(doc, doc));
        Assert.Empty(SchemaErrors(result));
    }

    // ---- observation helpers (no external renderer) --------------------------------------------

    private static string BodyText(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var body = w.MainDocumentPart?.Document?.Body;
        return body is null ? "" : string.Concat(body.Descendants<Text>().Select(t => t.Text));
    }

    /// <summary>Sorted multiset of per-footnote texts — robust to the renumbering Compare (and the oracle) do.</summary>
    private static List<string> NoteTexts(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var footnotes = w.MainDocumentPart?.FootnotesPart?.Footnotes;
        if (footnotes is null) return new List<string>();
        return footnotes.Elements<Footnote>()
            .Select(f => string.Concat(f.Descendants<Text>().Select(t => t.Text)))
            .Where(t => t.Length > 0)
            .OrderBy(t => t, System.StringComparer.Ordinal)
            .ToList();
    }

    private static List<long> FootnoteIds(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var footnotes = w.MainDocumentPart?.FootnotesPart?.Footnotes;
        return footnotes is null
            ? new List<long>()
            : footnotes.Elements<Footnote>().Where(f => f.Id is not null).Select(f => f.Id!.Value).ToList();
    }

    /// <summary>Body footnote-reference ids (document order) that do NOT resolve to exactly one definition.</summary>
    private static List<string> UnresolvedReferenceIds(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        var defIds = (main.FootnotesPart?.Footnotes?.Elements<Footnote>()
                         .Where(f => f.Id is not null)
                         .Select(f => f.Id!.Value) ?? Enumerable.Empty<long>())
                     .ToList();
        var defCount = defIds.GroupBy(x => x).ToDictionary(g => g.Key, g => g.Count());
        var bad = new List<string>();
        foreach (var r in main.Document?.Body?.Descendants<FootnoteReference>() ?? Enumerable.Empty<FootnoteReference>())
        {
            var id = r.Id?.Value;
            if (id is null || !defCount.TryGetValue(id.Value, out var n) || n != 1)
                bad.Add(id?.ToString() ?? "(null)");
        }
        return bad;
    }

    /// <summary>For each body footnote reference (document order), the text of the footnote definition it
    /// resolves to — the id ↔ reference ↔ text projection. Unresolvable references surface as "(unresolved)" so
    /// a dropped reference or dangling id changes the sequence rather than silently matching.</summary>
    private static List<string> ReferencedFootnoteTexts(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        var defText = new Dictionary<long, string>();
        foreach (var f in main.FootnotesPart?.Footnotes?.Elements<Footnote>() ?? Enumerable.Empty<Footnote>())
            if (f.Id is not null)
                defText[f.Id.Value] = string.Concat(f.Descendants<Text>().Select(t => t.Text));
        var seq = new List<string>();
        foreach (var r in main.Document?.Body?.Descendants<FootnoteReference>() ?? Enumerable.Empty<FootnoteReference>())
        {
            var id = r.Id?.Value;
            seq.Add(id is not null && defText.TryGetValue(id.Value, out var t) ? t : "(unresolved)");
        }
        return seq;
    }

    private static int HeaderFooterPartCount(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var main = w.MainDocumentPart!;
        return main.HeaderParts.Count() + main.FooterParts.Count();
    }

    private static HashSet<string> SchemaErrors(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var w = WordprocessingDocument.Open(ms, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
        return validator.Validate(w)
            .Select(e => $"{e.Id}@{e.Path?.XPath}: {e.Description}")
            .ToHashSet();
    }
}
