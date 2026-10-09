using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using Docxodus;
using Docxodus.Internal;
using Docxodus.McpServer;
using Xunit;
using static Docxodus.Tests.Ir.Diff.RevisionsInInputFixtures;

namespace Docxodus.Tests;

/// <summary>
/// N-way consolidate compares the ACCEPTED view of the base and of every reviewer copy, as two-way
/// comparison does (issue #1020). The engine already diffs the accepted body, so a plain fixture
/// cannot tell the policies apart: the discriminating input is a base whose header and footnote carry
/// pending tracked changes by a prior author ("ScopePrior"). Those parts are carried over from the base
/// into the consolidated redline, so without the pre-accept that markup leaks under its original author.
/// </summary>
public class ConsolidateFrontDoorTests
{
    private const string PriorAuthor = "ScopePrior";

    private static WmlDocument Version(string body) => MultiScopeRevisionDoc(
        body, "BodyPrior", "ins-x", "del-x", PriorAuthor, "hdr-prior", "fn-prior");

    private static WmlDocument Base() => Version("Body alpha");

    private static DocxDiffReviewer[] Reviewers() => new[]
    {
        new DocxDiffReviewer { Author = "Reviewer A", Document = Version("Body gamma") },
        new DocxDiffReviewer { Author = "Reviewer B", Document = Version("Body alpha delta") },
    };

    private static string ReviewersJson() => JsonSerializer.Serialize(Reviewers().Select(r => new
    {
        author = r.Author,
        docB64 = Convert.ToBase64String(r.Document.DocumentByteArray),
    }));

    private static void AssertNoPriorMarkup(WmlDocument consolidated)
    {
        var authors = RevisionAuthorsAllScopes(consolidated);
        Assert.DoesNotContain(PriorAuthor, authors);
        Assert.DoesNotContain("BodyPrior", authors);
        // The consolidation still did its work: both reviewers' edits are attributed.
        Assert.Contains("Reviewer A", authors);
    }

    [Fact]
    public void FrontDoorConsolidate_DoesNotLeakTheBasesHeaderAndFootnoteRevisions()
    {
        AssertNoPriorMarkup(DocxCompare.Consolidate(Base(), Reviewers()));
    }

    [Fact]
    public void FrontDoorConsolidate_LeavesTheCallersSettingsUntouched()
    {
        var settings = new DocxDiffConsolidateSettings { ConflictResolution = ConflictResolution.StackAll };

        DocxCompare.Consolidate(Base(), Reviewers(), settings);

        Assert.False(settings.Diff.PreAcceptInputRevisions);
        Assert.Equal(ConflictResolution.StackAll, settings.ConflictResolution);
    }

    [Fact]
    public void FacadeConsolidate_DoesNotLeakTheBasesHeaderAndFootnoteRevisions()
    {
        var bytes = DocxDiffOps.Consolidate(Base().DocumentByteArray, ReviewersJson(), null);

        AssertNoPriorMarkup(new WmlDocument("consolidated.docx", bytes));
    }

    [Fact]
    public void FacadeConsolidateProducts_DoesNotLeakTheBasesHeaderAndFootnoteRevisions()
    {
        var products = DocxDiffOps.ConsolidateProducts(Base().DocumentByteArray, ReviewersJson(), null,
            redline: true, revisions: false, editScript: false, conflicts: false);

        AssertNoPriorMarkup(new WmlDocument("consolidated.docx", products.RedlineBytes!));
    }

    [Fact]
    public void FacadeConsolidate_AppliesThePolicyEvenWhenTheWireAsksForTheRawEngine()
    {
        var bytes = DocxDiffOps.Consolidate(
            Base().DocumentByteArray, ReviewersJson(), "{\"preAcceptInputRevisions\":false}");

        AssertNoPriorMarkup(new WmlDocument("consolidated.docx", bytes));
    }

    [Fact]
    public void FacadeConsolidate_PreserveInputRevisionsStillTurnsThePreAcceptOff()
    {
        var bytes = DocxDiffOps.Consolidate(
            Base().DocumentByteArray, ReviewersJson(), "{\"preserveInputRevisions\":true}");

        Assert.Contains(PriorAuthor, RevisionAuthorsAllScopes(new WmlDocument("consolidated.docx", bytes)));
    }

    [Fact]
    public void StdioHostConsolidate_DoesNotLeakTheBasesHeaderAndFootnoteRevisions()
    {
        var args = JsonSerializer.SerializeToElement(new
        {
            baseB64 = Convert.ToBase64String(Base().DocumentByteArray),
            reviewers = JsonSerializer.Deserialize<JsonElement>(ReviewersJson()),
        });

        using var result = JsonDocument.Parse(Docxodus.PyHost.Dispatcher.Dispatch("docx_diff_consolidate", args));

        var bytes = Convert.FromBase64String(result.RootElement.GetProperty("docxB64").GetString()!);
        AssertNoPriorMarkup(new WmlDocument("consolidated.docx", bytes));
    }

    [Fact]
    public void McpConsolidate_DoesNotLeakTheBasesHeaderAndFootnoteRevisions()
    {
        var root = Path.Combine(Path.GetTempPath(), $"docxodus-consolidate-front-door-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var store = new SessionStore(new LocalFileDocumentStore(root));
        try
        {
            File.WriteAllBytes(Path.Combine(root, "base.docx"), Base().DocumentByteArray);
            var reviewers = Reviewers();
            File.WriteAllBytes(Path.Combine(root, "a.docx"), reviewers[0].Document.DocumentByteArray);
            File.WriteAllBytes(Path.Combine(root, "b.docx"), reviewers[1].Document.DocumentByteArray);

            using var args = JsonDocument.Parse(
                """{"baselinePath":"base.docx","revisedPaths":["a.docx","b.docx"],"authors":["Reviewer A","Reviewer B"],"outputPath":"out.docx"}""");
            var response = Dispatcher.Call(store, "docxodus_compare", args.RootElement.Clone());

            AssertNoPriorMarkup(new WmlDocument("out.docx", File.ReadAllBytes(Path.Combine(root, "out.docx"))));
            using var summary = JsonDocument.Parse(response);
            Assert.False(summary.RootElement.GetProperty("revisions").GetProperty("byAuthor").TryGetProperty(PriorAuthor, out _));
        }
        finally
        {
            store.CloseAll();
            Directory.Delete(root, recursive: true);
        }
    }
}
