#nullable enable
using System;
using System.Text.Json;
using Docxodus.Internal;
using Xunit;
namespace Docxodus.Tests;
public class DocxDiffOpsConsolidateTests
{
    private static byte[] Doc(params string[] paras) => Docxodus.Tests.Ir.Diff.Docs.Para(paras).DocumentByteArray;

    [Theory]
    [InlineData(null, false, "Crates hold nine dozen extra samples.")]
    [InlineData("""{"wordSeparators":".;!?"}""", true, "Crates hold six dozen samples.")]
    [InlineData("""{"wordSeparators":".;!?","conflictResolution":1}""", true, "Crates hold nine dozen samples.")]
    public void Word_boundaries_determine_whether_two_edits_compete(
        string? settings, bool expectConflict, string acceptedText)
    {
        const string original = "Crates hold six dozen samples.";
        var bytes = Doc(original);
        var reviewers = JsonSerializer.Serialize(new[]
        {
            new { author = "Ash", docB64 = Convert.ToBase64String(Doc("Crates hold nine dozen samples.")) },
            new { author = "Elm", docB64 = Convert.ToBase64String(Doc("Crates hold six dozen extra samples.")) },
        });
        using var report = JsonDocument.Parse(DocxDiffOps.GetConflictsJson(bytes, reviewers, settings));
        Assert.Equal(expectConflict, report.RootElement.GetProperty("conflicts").GetArrayLength() > 0);
        var output = new WmlDocument("review.docx", DocxDiffOps.Consolidate(bytes, reviewers, settings));
        Assert.Equal(acceptedText, Ir.Diff.Docs.PlainText(RevisionProcessor.AcceptRevisions(output)));
        Assert.Equal(original, Ir.Diff.Docs.PlainText(RevisionProcessor.RejectRevisions(output)));
    }

    [Fact]
    public void Consolidate_round_trips_through_wire()
    {
        var b = Doc("alpha one", "gamma three");
        var r1 = Convert.ToBase64String(Doc("alpha one EDITED", "gamma three"));
        var r2 = Convert.ToBase64String(Doc("alpha one", "gamma three EDITED"));
        var reviewers = $"[{{\"author\":\"Bob\",\"docB64\":\"{r1}\"}},{{\"author\":\"Fred\",\"docB64\":\"{r2}\"}}]";
        var merged = DocxDiffOps.Consolidate(b, reviewers, null);
        Assert.True(merged.Length > 0);
    }

    [Fact]
    public void GetConflictsJson_is_valid_json_with_conflicts_array()
    {
        var b = Doc("the quick brown fox");
        var r1 = Convert.ToBase64String(Doc("the SLOW brown fox"));
        var r2 = Convert.ToBase64String(Doc("the FAST brown fox"));
        var reviewers = $"[{{\"author\":\"Bob\",\"docB64\":\"{r1}\"}},{{\"author\":\"Fred\",\"docB64\":\"{r2}\"}}]";
        var json = DocxDiffOps.GetConflictsJson(b, reviewers, null);
        using var doc = JsonDocument.Parse(json);
        Assert.True(doc.RootElement.TryGetProperty("conflicts", out var arr));
        Assert.True(arr.GetArrayLength() >= 1);
    }

    [Fact]
    public void ParseSettings_reads_culture_so_it_is_reachable_off_dotnet()
    {
        // Regression (engine audit): DocxDiffSettings.Culture is a real .NET setting (consumed by the
        // tokenizer for case folding) but ParseSettings never read a "culture" key, so it was unreachable
        // from every non-.NET layer (WASM/npm/python). The wire parser must surface it.
        var s = DocxDiffOps.ParseSettings("{\"caseInsensitive\":true,\"culture\":\"tr-TR\"}");
        Assert.Equal("tr-TR", s.Culture?.Name);
    }

    [Fact]
    public void ConflictResolution_setting_parsed_from_wire()
    {
        var b = Doc("the quick brown fox");
        var r1 = Convert.ToBase64String(Doc("the SLOW brown fox"));
        var r2 = Convert.ToBase64String(Doc("the FAST brown fox"));
        var reviewers = $"[{{\"author\":\"Bob\",\"docB64\":\"{r1}\"}},{{\"author\":\"Fred\",\"docB64\":\"{r2}\"}}]";
        // StackAll = 2 -> accepted body should contain both edits
        var merged = DocxDiffOps.Consolidate(b, reviewers, "{\"conflictResolution\":2}");
        var doc = new Docxodus.WmlDocument("m.docx", merged);
        var accepted = Docxodus.Tests.Ir.Diff.Docs.PlainText(Docxodus.RevisionAccepter.AcceptRevisions(doc));
        Assert.Contains("SLOW", accepted);
        Assert.Contains("FAST", accepted);
    }
}
