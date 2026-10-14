#nullable enable

using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Headless CI guard for the byte-in / byte-out accept/reject surface added to
/// <see cref="DocxDiffOps"/> — the primitive both the WASM/npm and stdio/docx-scalpel clients route through to
/// verify a redline's round-trip contract. This is the .NET-level oracle behind the client round-trip tests
/// (<c>npm/tests/docx-diff.spec.ts</c>, <c>python/tests/test_docx_diff.py</c>): if the Ops surface itself were
/// wrong, every client would be too. Asserts the actual contract — accept(compare(left,right)) ≡ right and
/// reject ≡ left at the body-text level — not the shape of the result.
/// </summary>
public class DocxDiffOpsRoundTripTests
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");
    private static byte[] Wc(string name) => File.ReadAllBytes(Path.Combine(TestFilesDir.FullName, "WC", name));

    private static string BodyText(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var w = WordprocessingDocument.Open(ms, false);
        var body = w.MainDocumentPart?.Document?.Body;
        return body is null ? "" : string.Concat(body.Descendants<Text>().Select(t => t.Text));
    }

    [Theory]
    [InlineData("WC001-Digits.docx", "WC001-Digits-Mod.docx")]
    [InlineData("WC004-Large.docx", "WC004-Large-Mod.docx")]
    public void AcceptRejectRoundTrip_MaterializesRightAndLeft(string leftName, string rightName)
    {
        var left = Wc(leftName);
        var right = Wc(rightName);

        var redline = DocxDiffOps.Compare(left, right, null);
        var accepted = DocxDiffOps.AcceptRevisions(redline);
        var rejected = DocxDiffOps.RejectRevisions(redline);

        Assert.NotEqual(BodyText(left), BodyText(right));     // the pair genuinely differs
        Assert.Equal(BodyText(right), BodyText(accepted));    // accept ≡ right
        Assert.Equal(BodyText(left), BodyText(rejected));     // reject ≡ left
    }

    [Fact]
    public void AcceptOrReject_EmptyInput_Throws()
    {
        Assert.Throws<ArgumentException>(() => DocxDiffOps.AcceptRevisions(Array.Empty<byte>()));
        Assert.Throws<ArgumentException>(() => DocxDiffOps.RejectRevisions(Array.Empty<byte>()));
    }

    // ─── CompareProducts — one memoized pass, every product (issue #594) ─────

    [Fact]
    public void CompareProducts_EachProductMatchesItsStandaloneOp()
    {
        var left = Wc("WC001-Digits.docx");
        var right = Wc("WC001-Digits-Mod.docx");

        var products = DocxDiffOps.CompareProducts(
            left, right, null,
            redline: true, revisions: true, editScript: true, semanticChanges: true);

        Assert.Equal(DocxDiffOps.Compare(left, right, null), products.RedlineBytes);
        Assert.Equal(DocxDiffOps.GetRevisionsJson(left, right, null), products.RevisionsJson);
        Assert.Equal(DocxDiffOps.GetEditScriptJson(left, right, null), products.EditScriptJson);
        Assert.Equal(
            DocxDiffOps.GetSemanticChangesJson(left, right, null),
            products.SemanticChangesJson);
    }

    [Fact]
    public void CompareProducts_UnrequestedProductsAreNull()
    {
        var left = Wc("WC001-Digits.docx");
        var right = Wc("WC001-Digits-Mod.docx");

        var products = DocxDiffOps.CompareProducts(
            left, right, null,
            redline: false, revisions: true, editScript: false, semanticChanges: false);

        Assert.Null(products.RedlineBytes);
        Assert.NotNull(products.RevisionsJson);
        Assert.Null(products.EditScriptJson);
        Assert.Null(products.SemanticChangesJson);
    }

    [Fact]
    public void CompareProductsJson_EnvelopeCarriesTheStandaloneWireShapes()
    {
        var left = Wc("WC001-Digits.docx");
        var right = Wc("WC001-Digits-Mod.docx");

        var envelope = System.Text.Json.JsonDocument.Parse(
            DocxDiffOps.CompareProductsJson(left, right, null, null)).RootElement;

        Assert.Equal(
            DocxDiffOps.Compare(left, right, null),
            Convert.FromBase64String(envelope.GetProperty("redlineB64").GetString()!));

        var standaloneRevisions = System.Text.Json.JsonDocument
            .Parse(DocxDiffOps.GetRevisionsJson(left, right, null))
            .RootElement.GetProperty("revisions");
        Assert.Equal(
            standaloneRevisions.GetRawText(),
            envelope.GetProperty("revisions").GetRawText());

        Assert.True(envelope.TryGetProperty("editScript", out var script)
            && script.ValueKind == System.Text.Json.JsonValueKind.Object);
        Assert.True(envelope.TryGetProperty("semanticChanges", out var semantic)
            && semantic.ValueKind == System.Text.Json.JsonValueKind.Object);
    }

    [Fact]
    public void CompareProductsJson_SelectionAndValidation()
    {
        var left = Wc("WC001-Digits.docx");
        var right = Wc("WC001-Digits-Mod.docx");

        var envelope = System.Text.Json.JsonDocument.Parse(
            DocxDiffOps.CompareProductsJson(left, right, null, "[\"revisions\"]")).RootElement;
        Assert.True(envelope.TryGetProperty("revisions", out _));
        Assert.False(envelope.TryGetProperty("redlineB64", out _));
        Assert.False(envelope.TryGetProperty("editScript", out _));
        Assert.False(envelope.TryGetProperty("semanticChanges", out _));

        Assert.Throws<ArgumentException>(() =>
            DocxDiffOps.CompareProductsJson(left, right, null, "[\"typo\"]"));
        Assert.Throws<ArgumentException>(() =>
            DocxDiffOps.CompareProductsJson(left, right, null, "[]"));
        Assert.Throws<ArgumentException>(() =>
            DocxDiffOps.CompareProductsJson(left, right, null, "{\"redline\":true}"));
    }

    // ─── One baseline, many candidates (issue #617) ──────────────────────

    [Fact]
    public void CompareBatch_EachResultMatchesTheSinglePairEnvelope()
    {
        var baseline = Wc("WC001-Digits.docx");
        var candidates = new[] { Wc("WC001-Digits-Mod.docx"), Wc("WC002-DeleteAtEnd.docx") };
        var candidatesJson = System.Text.Json.JsonSerializer.Serialize(
            candidates.Select((bytes, i) => new { name = $"c{i}", docB64 = Convert.ToBase64String(bytes) }));

        var batch = System.Text.Json.JsonDocument.Parse(
            DocxDiffOps.CompareBatchJson(baseline, candidatesJson, null, null));
        var results = batch.RootElement.GetProperty("results").EnumerateArray().ToList();

        Assert.Equal(candidates.Length, results.Count);
        for (var i = 0; i < candidates.Length; i++)
        {
            Assert.Equal($"c{i}", results[i].GetProperty("name").GetString());
            // Every product byte-for-byte what the single-pair envelope carries, name aside.
            var single = System.Text.Json.JsonDocument.Parse(
                DocxDiffOps.CompareProductsJson(baseline, candidates[i], null, null));
            foreach (var key in new[] { "redlineB64", "revisions", "editScript", "semanticChanges" })
                Assert.Equal(
                    single.RootElement.GetProperty(key).GetRawText(),
                    results[i].GetProperty(key).GetRawText());
        }
    }

    /// <summary>One malformed candidate must not cost the caller the rest of the batch.</summary>
    [Fact]
    public void CompareBatch_AFailingCandidateCarriesItsErrorAndTheRestStillCompare()
    {
        var baseline = Wc("WC001-Digits.docx");
        var candidatesJson = System.Text.Json.JsonSerializer.Serialize(new[]
        {
            new { name = "good", docB64 = Convert.ToBase64String(Wc("WC001-Digits-Mod.docx")) },
            new { name = "junk", docB64 = Convert.ToBase64String(new byte[] { 1, 2, 3, 4 }) },
            new { name = "also-good", docB64 = Convert.ToBase64String(Wc("WC002-DeleteAtEnd.docx")) },
        });

        var results = System.Text.Json.JsonDocument
            .Parse(DocxDiffOps.CompareBatchJson(baseline, candidatesJson, null, "[\"revisions\"]"))
            .RootElement.GetProperty("results").EnumerateArray().ToList();

        Assert.Equal(3, results.Count);
        Assert.True(results[0].TryGetProperty("revisions", out _));
        Assert.True(results[1].TryGetProperty("error", out _));
        Assert.False(results[1].TryGetProperty("revisions", out _));
        Assert.True(results[2].TryGetProperty("revisions", out _));
    }

    [Fact]
    public void CompareBatch_RejectsACandidateWithoutBytesRatherThanSkippingIt()
    {
        var baseline = Wc("WC001-Digits.docx");
        Assert.Throws<ArgumentException>(() =>
            DocxDiffOps.CompareBatchJson(baseline, "[{\"name\":\"nameless\"}]", null, null));
    }

    [Fact]
    public void ConsolidateProducts_EachProductMatchesItsStandaloneOp()
    {
        var baseDoc = Wc("WC001-Digits.docx");
        var reviewersJson = System.Text.Json.JsonSerializer.Serialize(new[]
        {
            new { author = "A", docB64 = Convert.ToBase64String(Wc("WC001-Digits-Mod.docx")) },
            new { author = "B", docB64 = Convert.ToBase64String(Wc("WC002-DeleteAtEnd.docx")) },
        });

        var products = DocxDiffOps.ConsolidateProducts(
            baseDoc, reviewersJson, null,
            redline: true, revisions: true, editScript: true, conflicts: true);

        Assert.Equal(DocxDiffOps.Consolidate(baseDoc, reviewersJson, null), products.RedlineBytes);
        Assert.Equal(
            DocxDiffOps.GetConsolidatedRevisionsJson(baseDoc, reviewersJson, null),
            products.RevisionsJson);
        Assert.Equal(
            DocxDiffOps.GetConsolidatedEditScriptJson(baseDoc, reviewersJson, null),
            products.EditScriptJson);
        Assert.Equal(DocxDiffOps.GetConflictsJson(baseDoc, reviewersJson, null), products.ConflictsJson);
    }
}
