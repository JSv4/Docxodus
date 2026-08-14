#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Xml.Linq;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Memory-bounding tests for the undo ring (UR0xx) and the estimator it budgets against.
///
/// <para>A depth cap alone never bounded memory: each undo step is a deep clone of every
/// snapshot-scoped part, so one step costs whatever the DOCUMENT costs. These tests pin the
/// second bound — evict oldest until under budget, redo before undo, never surrender the last
/// undo step — and the reporting that makes the eviction visible rather than silent.</para>
/// </summary>
public class UndoRingMemoryTests
{
    // ─── Ring mechanics (cost supplied directly, no documents involved) ───

    /// <summary>Each entry costs 10 bytes; a 100-byte budget holds ten.</summary>
    private static UndoRing<string> Ring(int depth, long budget) =>
        new(depth, budget, static _ => 10);

    [Fact]
    public void UR001_DepthBoundStillApplies()
    {
        var ring = Ring(depth: 3, budget: 0);
        for (int i = 0; i < 10; i++) ring.RecordPreOp($"s{i}");

        Assert.Equal(3, ring.UndoCount);
        Assert.False(ring.EvictedForMemory); // depth, not memory, did this
    }

    [Fact]
    public void UR002_BudgetEvictsOldestBeyondTheByteCeiling()
    {
        // Depth would allow 100; the budget allows 5 entries of 10 bytes.
        var ring = Ring(depth: 100, budget: 50);
        for (int i = 0; i < 20; i++) ring.RecordPreOp($"s{i}");

        Assert.Equal(5, ring.UndoCount);
        Assert.Equal(50, ring.RetainedBytes);
        Assert.True(ring.EvictedForMemory);

        // The SURVIVORS are the most recent five — history is trimmed from the old end.
        var kept = Enumerable.Range(0, 5).Select(_ => ring.PopForUndo().snapshot).ToArray();
        Assert.Equal(new[] { "s19", "s18", "s17", "s16", "s15" }, kept);
    }

    /// <summary>A snapshot larger than the entire budget must still leave one step of undo.
    /// Silently having no undo at all is worst on exactly the large documents that trip this.</summary>
    [Fact]
    public void UR003_LastUndoEntryIsNeverEvicted()
    {
        var ring = new UndoRing<string>(capacity: 10, budgetBytes: 5, costOf: static _ => 1000);
        ring.RecordPreOp("only");
        ring.RecordPreOp("newest");

        Assert.Equal(1, ring.UndoCount);
        Assert.True(ring.EvictedForMemory);
        Assert.Equal("newest", ring.PopForUndo().snapshot);
    }

    /// <summary>Redo is speculative and is surrendered before the user's own undo history.</summary>
    [Fact]
    public void UR004_RedoIsEvictedBeforeUndo()
    {
        var ring = Ring(depth: 100, budget: 30);
        ring.RecordPreOp("u1");
        ring.RecordPreOp("u2");

        // Undo twice, building redo entries, then push past the budget.
        ring.PopForUndo();
        ring.RecordForRedo("r1");
        ring.RecordForRedo("r2");
        ring.RecordForRedo("r3");

        Assert.True(ring.RetainedBytes <= 30);
        Assert.True(ring.UndoCount >= 1, "undo history must outlive speculative redo entries");
    }

    [Fact]
    public void UR005_BudgetOfZeroDisablesMeasurementEntirely()
    {
        var ring = new UndoRing<string>(capacity: 100, budgetBytes: 0,
            costOf: static _ => throw new System.InvalidOperationException(
                "cost selector must not run when the budget is disabled"));

        for (int i = 0; i < 50; i++) ring.RecordPreOp($"s{i}");

        Assert.Equal(50, ring.UndoCount);
        Assert.Equal(0, ring.RetainedBytes);
        Assert.False(ring.EvictedForMemory);
    }

    [Fact]
    public void UR006_PopAndClearKeepTheByteAccountingHonest()
    {
        var ring = Ring(depth: 100, budget: 1000);
        for (int i = 0; i < 8; i++) ring.RecordPreOp($"s{i}");
        Assert.Equal(80, ring.RetainedBytes);

        ring.PopForUndo();
        ring.PopForUndo();
        Assert.Equal(60, ring.RetainedBytes);

        ring.Clear();
        Assert.Equal(0, ring.RetainedBytes);
        Assert.Equal(0, ring.UndoCount);
    }

    // ─── Estimator ───────────────────────────────────────────────────────

    /// <summary>The estimate must grow with the document, and must exceed serialized length —
    /// budgeting against serialized size is the under-count this estimator exists to avoid.</summary>
    [Fact]
    public void UR010_EstimateGrowsWithDocumentAndExceedsSerializedLength()
    {
        XDocument Build(int paragraphs)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            return new XDocument(new XElement(w + "document",
                new XElement(w + "body",
                    Enumerable.Range(0, paragraphs).Select(i =>
                        new XElement(w + "p",
                            new XAttribute(w + "rsidR", $"00{i:D6}"),
                            new XElement(w + "r", new XElement(w + "t", $"paragraph number {i}")))))));
        }

        var small = XmlMemoryEstimator.Estimate(Build(10));
        var large = XmlMemoryEstimator.Estimate(Build(1000));

        Assert.True(small > 0);
        Assert.True(large > small * 50, $"estimate should scale with size: {small} -> {large}");

        var serializedChars = Build(1000).ToString(SaveOptions.DisableFormatting).Length;
        Assert.True(large > serializedChars,
            $"estimate {large} must exceed serialized length {serializedChars} — a live tree costs " +
            "more than its markup, which is the whole reason serialized size is not the measure");
    }

    [Fact]
    public void UR011_EstimateOfAnEmptyDocumentIsSmallButNonZero()
    {
        var doc = new XDocument(new XElement("root"));
        var estimate = XmlMemoryEstimator.Estimate(doc);
        Assert.InRange(estimate, 1, 1024);
    }

    // ─── Session integration ─────────────────────────────────────────────

    private static string FirstBodyParagraph(DocxSession s) =>
        s.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;

    /// <summary>The default session measures its snapshots and reports the retained total.</summary>
    [Fact]
    public void UR020_SessionReportsRetainedUndoMemory()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(s);

        Assert.Equal(0, s.UndoCount);
        Assert.Equal(0, s.UndoMemoryBytes);

        Assert.True(s.ReplaceText(anchor, "one").Success);
        Assert.Equal(1, s.UndoCount);
        Assert.True(s.UndoMemoryBytes > 0, "a recorded snapshot must contribute measured bytes");
        Assert.False(s.UndoHistoryTrimmedForMemory);
    }

    /// <summary>A budget small enough to bind trims history and says so — the behavior an editor
    /// needs in order to explain why undo stops short of the configured depth.</summary>
    [Fact]
    public void UR021_TightBudgetTrimsHistoryAndReportsIt()
    {
        using var s = new DocxSession(
            DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { UndoDepth = 50, UndoMemoryBudgetBytes = 1 });

        var anchor = FirstBodyParagraph(s);
        for (int i = 0; i < 5; i++)
            Assert.True(s.ReplaceText(anchor, $"edit {i}").Success);

        Assert.True(s.UndoHistoryTrimmedForMemory);
        Assert.Equal(1, s.UndoCount);          // the guaranteed last step
        Assert.True(s.Undo());                 // and it still works
        Assert.False(s.Undo());
    }

    /// <summary>Opting out restores the historical depth-only behavior exactly, including not
    /// paying for measurement.</summary>
    [Fact]
    public void UR022_ZeroBudgetRestoresDepthOnlyBounding()
    {
        using var s = new DocxSession(
            DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { UndoDepth = 4, UndoMemoryBudgetBytes = 0 });

        var anchor = FirstBodyParagraph(s);
        for (int i = 0; i < 10; i++)
            Assert.True(s.ReplaceText(anchor, $"edit {i}").Success);

        Assert.Equal(4, s.UndoCount);
        Assert.Equal(0, s.UndoMemoryBytes);    // nothing measured
        Assert.False(s.UndoHistoryTrimmedForMemory);
    }

    /// <summary>The default depth is the documented one, and the wire default agrees with it —
    /// the JSON surface used to hardcode 50 independently of the .NET default.</summary>
    [Fact]
    public void UR023_WireDefaultsTrackTheDotNetDefaults()
    {
        var dotnet = new DocxSessionSettings();
        Assert.Equal(20, dotnet.UndoDepth);
        Assert.Equal(128L * 1024 * 1024, dotnet.UndoMemoryBudgetBytes);

        var fromEmptyJson = DocxSessionJson.ParseSettings("{}");
        Assert.Equal(dotnet.UndoDepth, fromEmptyJson.UndoDepth);
        Assert.Equal(dotnet.UndoMemoryBudgetBytes, fromEmptyJson.UndoMemoryBudgetBytes);

        var explicitJson = DocxSessionJson.ParseSettings(
            "{\"undoDepth\":7,\"undoMemoryBudgetBytes\":1234567}");
        Assert.Equal(7, explicitJson.UndoDepth);
        Assert.Equal(1234567L, explicitJson.UndoMemoryBudgetBytes);
    }
}
