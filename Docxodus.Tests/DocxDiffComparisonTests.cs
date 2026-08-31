#nullable enable

using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="DocxDiff.CreateComparison"/> (issue #594): one alignment pass serving every data
/// product. The contract is (1) every product is identical to what the stateless static
/// produces on the same inputs/settings, and (2) repeated product calls are served from the
/// memoized pass, not recomputed.
/// </summary>
public class DocxDiffComparisonTests
{
    private static WmlDocument Doc(params string[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                paragraphs.Select(text => new Paragraph(new Run(new Text(text))))));
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()));
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("test.docx", stream.ToArray());
    }

    private static (WmlDocument Left, WmlDocument Right) Pair() => (
        Doc("The quick brown fox.", "Second paragraph here.", "Closing paragraph."),
        Doc("The quick red fox.", "Second paragraph here.", "A fresh closing paragraph."));

    [Fact]
    public void EveryProduct_MatchesItsStatelessStatic()
    {
        var (left, right) = Pair();
        var settings = new DocxDiffSettings { AuthorForRevisions = "Daisy" };
        var comparison = DocxDiff.CreateComparison(left, right, settings);

        Assert.Equal(
            DocxDiff.Compare(left, right, settings).DocumentByteArray,
            comparison.ToRedline().DocumentByteArray);

        var staticRevisions = DocxDiff.GetRevisions(left, right, settings);
        var memoRevisions = comparison.GetRevisions();
        Assert.Equal(staticRevisions.Count, memoRevisions.Count);
        Assert.Equal(
            staticRevisions.Select(r => r.ToString()),
            memoRevisions.Select(r => r.ToString()));

        Assert.Equal(
            DocxDiff.GetEditScriptJson(left, right, settings),
            comparison.GetEditScriptJson());
    }

    [Fact]
    public void SemanticProducts_MatchTheStatics_WithTheComparisonSettingsFlowing()
    {
        var (left, right) = Pair();
        var settings = new DocxDiffSettings { AuthorForRevisions = "Daisy" };
        var comparison = DocxDiff.CreateComparison(left, right, settings);

        // With no explicit options, the comparison's own settings flow into the semantic pass —
        // the same JSON the static produces when handed those settings explicitly.
        var options = new SemanticDiffOptions { DiffSettings = settings };
        Assert.Equal(
            DocxDiff.GetSemanticChangesJson(left, right, options),
            comparison.GetSemanticChangesJson());

        // Explicit options are honored verbatim.
        Assert.Equal(
            DocxDiff.GetSemanticChangesJson(left, right, null),
            comparison.GetSemanticChangesJson(new SemanticDiffOptions()));
    }

    [Fact]
    public void RepeatedCalls_ServeTheMemoizedProduct()
    {
        var (left, right) = Pair();
        var comparison = DocxDiff.CreateComparison(left, right);

        Assert.Same(comparison.GetRevisions(), comparison.GetRevisions());
        Assert.Same(comparison.GetEditScriptJson(), comparison.GetEditScriptJson());
        Assert.Same(comparison.GetSemanticChangesJson(), comparison.GetSemanticChangesJson());
        Assert.Same(
            comparison.ToRedline().DocumentByteArray,
            comparison.ToRedline().DocumentByteArray);
    }

    [Fact]
    public void IdenticalInputs_ShortCircuitLikeTheStatics()
    {
        var doc = Doc("Unchanged text.");
        var same = new WmlDocument("copy.docx", doc.DocumentByteArray);
        var comparison = DocxDiff.CreateComparison(doc, same);

        Assert.Equal(doc.DocumentByteArray, comparison.ToRedline().DocumentByteArray);
        Assert.Empty(comparison.GetRevisions());
    }

    [Fact]
    public void NullInputs_AreRejectedAtCreation()
    {
        var doc = Doc("x");
        Assert.Throws<System.ArgumentNullException>(() => DocxDiff.CreateComparison(null!, doc));
        Assert.Throws<System.ArgumentNullException>(() => DocxDiff.CreateComparison(doc, null!));
    }

    // ─── Reusable snapshots (issue #617) ─────────────────────────────────

    [Fact]
    public void SnapshotComparison_ProducesTheSameProductsAsTheDocumentComparison()
    {
        var (left, right) = Pair();
        var expected = DocxDiff.CreateComparison(left, right);
        var actual = DocxDiff.CreateComparison(
            DocxDiff.CreateSnapshot(left), DocxDiff.CreateSnapshot(right));

        Assert.Equal(expected.GetEditScriptJson(), actual.GetEditScriptJson());
        Assert.Equal(
            expected.GetRevisions().Select(r => $"{r.Type}|{r.Text}"),
            actual.GetRevisions().Select(r => $"{r.Type}|{r.Text}"));
        Assert.Equal(
            expected.ToRedline().DocumentByteArray, actual.ToRedline().DocumentByteArray);
    }

    /// <summary>
    /// The point of the type: one baseline, many counterparties, one read of the baseline. The
    /// products must be exactly the fan-out of independent comparisons.
    /// </summary>
    [Fact]
    public void OneBaselineSnapshot_ServesEveryCandidateWithUnchangedProducts()
    {
        var baseline = Doc("Clause one.", "Clause two.", "Clause three.");
        var candidates = new[]
        {
            Doc("Clause one, amended.", "Clause two.", "Clause three."),
            Doc("Clause one.", "Clause two, amended.", "Clause three."),
            Doc("Clause one.", "Clause two.", "Clause three, amended."),
        };

        var snapshot = DocxDiff.CreateSnapshot(baseline);
        Assert.False(snapshot.IsMaterialized);

        foreach (var candidate in candidates)
        {
            var viaSnapshot = DocxDiff.CreateComparison(snapshot, DocxDiff.CreateSnapshot(candidate));
            Assert.Equal(
                DocxDiff.GetEditScriptJson(baseline, candidate), viaSnapshot.GetEditScriptJson());
            Assert.Equal(
                DocxDiff.Compare(baseline, candidate).DocumentByteArray,
                viaSnapshot.ToRedline().DocumentByteArray);
        }

        // …and the baseline was read once, not once per candidate.
        Assert.True(snapshot.IsMaterialized);
    }

    /// <summary>
    /// A snapshot is read under one input-revision policy and is only valid for comparisons that
    /// share it. Silently reusing one read the other way would compare a different view of the
    /// document than the caller asked for, so the mismatch is refused on both sides.
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SnapshotReadUnderADifferentInputRevisionPolicy_IsRefused(bool snapshotAccepts)
    {
        var (left, right) = Pair();
        var snapshotSettings = new DocxDiffSettings { PreAcceptInputRevisions = snapshotAccepts };
        var comparisonSettings = new DocxDiffSettings { PreAcceptInputRevisions = !snapshotAccepts };

        var mismatched = DocxDiff.CreateSnapshot(left, snapshotSettings);
        var matching = DocxDiff.CreateSnapshot(right, comparisonSettings);

        Assert.Equal(snapshotAccepts, mismatched.InputRevisionsAccepted);
        Assert.Contains("input-revision policy",
            Assert.Throws<ArgumentException>(
                () => DocxDiff.CreateComparison(mismatched, matching, comparisonSettings)).Message);
        Assert.Contains("input-revision policy",
            Assert.Throws<ArgumentException>(
                () => DocxDiff.CreateComparison(matching, mismatched, comparisonSettings)).Message);
    }

    /// <summary>`PreserveInputRevisions` overrides the accept-flatten, so a snapshot created under
    /// it reads the same way as one created with neither flag — and is interchangeable with it.</summary>
    [Fact]
    public void PreserveInputRevisions_ReadsLikeTheDefaultAndIsInterchangeable()
    {
        var (left, right) = Pair();
        var preserve = new DocxDiffSettings
        {
            PreAcceptInputRevisions = true,
            PreserveInputRevisions = true,
        };

        var snapshot = DocxDiff.CreateSnapshot(left, preserve);
        Assert.False(snapshot.InputRevisionsAccepted);
        Assert.NotNull(DocxDiff.CreateComparison(snapshot, DocxDiff.CreateSnapshot(right)));
    }

    /// <summary>A reused snapshot must not make a caller's compatibility subscription go quiet: the
    /// pre-flight is a property of the comparison, not of the read, so it fires on every comparison
    /// the snapshot takes part in.</summary>
    [Fact]
    public void ReusedSnapshot_StillPreflightsEveryComparison()
    {
        var (left, right) = Pair();
        var baseline = DocxDiff.CreateSnapshot(left);
        var calls = 0;
        var settings = new DocxDiffSettings { OnCompatibilityWarning = _ => calls++ };

        _ = DocxDiff.CreateComparison(baseline, DocxDiff.CreateSnapshot(right), settings).GetRevisions();
        var firstRound = calls;
        _ = DocxDiff.CreateComparison(baseline, DocxDiff.CreateSnapshot(right), settings).GetRevisions();

        // Whatever the document warns about, the second comparison reports it as often as the first.
        Assert.Equal(firstRound * 2, calls);
    }

    // ─── Memoized consolidation (issue #617) ─────────────────────────────

    [Fact]
    public void ConsolidationProducts_MatchTheStatelessStatics()
    {
        var baseDoc = Doc("Clause one.", "Clause two.");
        var reviewers = new[]
        {
            new DocxDiffReviewer { Author = "A", Document = Doc("Clause one, per A.", "Clause two.") },
            new DocxDiffReviewer { Author = "B", Document = Doc("Clause one.", "Clause two, per B.") },
        };

        var consolidation = DocxDiff.CreateConsolidation(baseDoc, reviewers);

        Assert.Equal(
            DocxDiff.GetConsolidatedEditScriptJson(baseDoc, reviewers),
            consolidation.GetConsolidatedEditScriptJson());
        Assert.Equal(
            DocxDiff.GetConflicts(baseDoc, reviewers).Select(c => c.Id),
            consolidation.GetConflicts().Select(c => c.Id));
        Assert.Equal(
            DocxDiff.GetConsolidatedRevisions(baseDoc, reviewers).Select(r => $"{r.Author}|{r.Text}"),
            consolidation.GetConsolidatedRevisions().Select(r => $"{r.Author}|{r.Text}"));
        Assert.Equal(
            DocxDiff.Consolidate(baseDoc, reviewers).DocumentByteArray,
            consolidation.Consolidate().DocumentByteArray);
    }

    /// <summary>
    /// The gap this closes: only <c>Consolidate</c> ran the compatibility pre-flight, so a caller
    /// who set <c>ThrowOnCompatibilityWarning</c> and asked for conflicts, consolidated revisions or
    /// the consolidated edit script was silently never told. All four now run the same gate, which
    /// is the N-way half of what #622 fixed on the pairwise side.
    /// </summary>
    [Fact]
    public void EveryConsolidateEntryPoint_HonorsTheCompatibilitySubscription()
    {
        // A real fixture the catalog warns about, so the subscription has something to report.
        var warns = new WmlDocument(
            Path.Combine("../../../../TestFiles", "CU011-Chart-Embedded-Xlsx-03.docx"));
        Assert.True(DocxDiff.InspectCompatibility(warns).HasWarnings, "fixture no longer warns");
        var reviewers = new[] { new DocxDiffReviewer { Author = "A", Document = warns } };

        var entryPoints = new (string Name, Func<DocxDiffConsolidateSettings, object?> Run)[]
        {
            ("Consolidate", s => DocxDiff.Consolidate(warns, reviewers, s)),
            ("GetConflicts", s => DocxDiff.GetConflicts(warns, reviewers, s)),
            ("GetConsolidatedRevisions", s => DocxDiff.GetConsolidatedRevisions(warns, reviewers, s)),
            ("GetConsolidatedEditScriptJson", s => DocxDiff.GetConsolidatedEditScriptJson(warns, reviewers, s)),
        };

        foreach (var (name, run) in entryPoints)
        {
            var reports = 0;
            run(new DocxDiffConsolidateSettings
            {
                Diff = new DocxDiffSettings { OnCompatibilityWarning = _ => reports++ },
            });
            Assert.True(reports > 0, $"{name} never reported a compatibility warning");

            Assert.Throws<DocxDiffCompatibilityException>(() => run(new DocxDiffConsolidateSettings
            {
                Diff = new DocxDiffSettings { ThrowOnCompatibilityWarning = true },
            }));
        }
    }
}
