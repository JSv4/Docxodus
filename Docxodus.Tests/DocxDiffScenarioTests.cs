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
    /// <summary>
    /// Scenarios that edit a FOOTNOTE-REFERENCING paragraph currently expose an unfixed engine bug: the
    /// produced footnotes part gets a DUPLICATE footnote id (and a body footnote reference is dropped),
    /// so the output is schema-invalid — even though the body + note TEXT round-trips. The WmlComparer
    /// oracle handles the same inputs cleanly, and the corruption reproduces on the real NVCA contract
    /// (e.g. body-replace-word → footnote id 1 duplicated). Tracked by
    /// <see cref="FootnoteReferencingParagraphEdit_CorruptsFootnoteIds_KnownBug"/>; these scenarios are
    /// excluded from the schema-validity theory until the note-path fix lands (then move them back).
    /// </summary>
    private static readonly HashSet<string> KnownFootnoteIdBug = new()
    {
        "body-replace-word", "body-delete-paragraph", "body-split-paragraph",
        "whole-paragraph-replace", "multi-edit",
    };

    public static IEnumerable<object[]> AllScenarios() =>
        DocxDiffScenarioFixtures.Names().Select(n => new object[] { n });

    public static IEnumerable<object[]> SchemaCleanScenarios() =>
        DocxDiffScenarioFixtures.Names().Where(n => !KnownFootnoteIdBug.Contains(n)).Select(n => new object[] { n });

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

    // ---- schema validity (every scenario EXCEPT the known footnote-id bug) ----------------------

    [Theory]
    [MemberData(nameof(SchemaCleanScenarios))]
    public void Scenario_ProducesSchemaValidOutput(string scenario)
    {
        var (left, right) = DocxDiffScenarioFixtures.Build(scenario);
        var result = DocxDiff.Compare(left, right);

        var baseErrors = SchemaErrors(left);
        var newErrors = SchemaErrors(result).Where(e => !baseErrors.Contains(e)).ToList();
        Assert.True(newErrors.Count == 0,
            $"Compare introduced {newErrors.Count} new schema error(s): {string.Join(" | ", newErrors.Take(5))}");
    }

    /// <summary>
    /// Characterizes the KNOWN footnote-id bug (see <see cref="KnownFootnoteIdBug"/>): editing a
    /// footnote-referencing paragraph currently produces a duplicate footnote id. This test asserts the
    /// bug is STILL present, so it stays green while tracked — when the note-path fix lands, this test
    /// FAILS, which is the signal to delete it and fold these scenarios back into
    /// <see cref="Scenario_ProducesSchemaValidOutput"/>. The oracle (WmlComparer) is clean on these inputs.
    /// </summary>
    [Fact]
    public void FootnoteReferencingParagraphEdit_CorruptsFootnoteIds_KnownBug()
    {
        var stillBuggy = new List<string>();
        foreach (var scenario in KnownFootnoteIdBug)
        {
            var (left, right) = DocxDiffScenarioFixtures.Build(scenario);
            var ids = FootnoteIds(DocxDiff.Compare(left, right));
            if (ids.Count != ids.Distinct().Count())
                stillBuggy.Add(scenario);
        }
        Assert.Equal(KnownFootnoteIdBug.OrderBy(x => x), stillBuggy.OrderBy(x => x));
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
