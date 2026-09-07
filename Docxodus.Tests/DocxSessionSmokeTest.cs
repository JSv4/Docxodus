#nullable enable

// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests;

/// <summary>
/// End-to-end smoke test that exercises the DocxSession API on a real legal-style
/// fixture and verifies edits round-trip through Save/reopen. This is the
/// "shake the box and see what falls out" test the agentic editing flow will rely on.
/// </summary>
public class DocxSessionSmokeTest
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");
    private readonly ITestOutputHelper _output;

    public DocxSessionSmokeTest(ITestOutputHelper output) => _output = output;

    /// <summary>
    /// Sorts AnchorTargets by their first appearance in a given projection markdown,
    /// so callers see document order rather than anchor-id hex order.
    /// </summary>
    private static int InitialOrderIndex(string markdown, string anchorId)
    {
        var idx = markdown.IndexOf("{#" + anchorId + "}", StringComparison.Ordinal);
        return idx < 0 ? int.MaxValue : idx;
    }

    [Fact]
    public void DS999_AgenticWorkflowOnRealDocument()
    {
        // ── Arrange ──────────────────────────────────────────────────────
        var path = Path.Combine(TestFilesDir.FullName, "HC001-5DayTourPlanTemplate.docx");
        var bytes = File.ReadAllBytes(path);
        _output.WriteLine($"Loaded fixture: {path} ({bytes.Length} bytes)");

        using var session = new DocxSession(bytes);

        // ── Read: project to markdown ─────────────────────────────────────
        var initial = session.Project();
        Assert.NotEmpty(initial.Markdown);
        Assert.True(initial.AnchorIndex.Count > 0);
        _output.WriteLine($"Initial projection: {initial.Markdown.Length} chars, " +
                          $"{initial.AnchorIndex.Count} anchors");

        var firstParagraph = initial.AnchorIndex.Values
            .Where(t => t.Anchor.Kind is "p" or "h" or "li" && t.Anchor.Scope == "body")
            .OrderBy(t => InitialOrderIndex(initial.Markdown, t.Anchor.Id))
            .First();
        _output.WriteLine($"First addressable block: {firstParagraph.Anchor.Id}");

        // ── Mutation 1: ReplaceText (Tier A) ──────────────────────────────
        var r1 = session.ReplaceText(firstParagraph.Anchor.Id,
            "**SMOKETESTMARKER1:** Agentically replaced opening paragraph.");
        Assert.True(r1.Success, r1.Error?.Message);
        Assert.NotNull(r1.Patch);
        _output.WriteLine($"[1] ReplaceText OK; modified={r1.Modified.Count}");

        // ── Mutation 2: InsertParagraph After (Tier B) ────────────────────
        var r2 = session.InsertParagraph(firstParagraph.Anchor.Id, Position.After,
            "## Inserted Heading\n\nAgentically inserted body paragraph below the heading.");
        Assert.True(r2.Success, r2.Error?.Message);
        Assert.Equal(2, r2.Created.Count);
        var newHeading = r2.Created[0];
        Assert.Equal("h", newHeading.Kind);
        _output.WriteLine($"[2] InsertParagraph OK; created heading {newHeading.Id}");

        // ── Mutation 3: SplitParagraph (Tier B) ───────────────────────────
        // Pick a body paragraph with ≥10 chars that ISN'T the marker'd one
        var splittable = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body")
            .Where(t => t.Anchor.Id != firstParagraph.Anchor.Id)
            .FirstOrDefault(t => session.GetAnchorInfo(t.Anchor.Id) is { TextPreview.Length: > 10 });
        if (splittable is not null)
        {
            var r3 = session.SplitParagraph(splittable.Anchor.Id, 5);
            Assert.True(r3.Success, r3.Error?.Message);
            Assert.Single(r3.Created);
            _output.WriteLine($"[3] SplitParagraph OK at offset 5");
        }

        // ── Mutation 4: ApplyFormat span (Tier C) ─────────────────────────
        // Apply to the marker'd block (still addressable by firstParagraph.Anchor.Id)
        var someParagraph = session.Project().AnchorIndex.Values
            .FirstOrDefault(t => t.Anchor.Id == firstParagraph.Anchor.Id);
        if (someParagraph is not null)
        {
            var r4 = session.ApplyFormat(someParagraph.Anchor.Id,
                new CharSpan(0, 3), new FormatOp { Bold = true });
            Assert.True(r4.Success, r4.Error?.Message);
            _output.WriteLine($"[4] ApplyFormat span bold OK");
        }

        // ── Mutation 5: Raw.GetXml → mutate → Raw.ReplaceXml (escape hatch) ──
        // Use the marker'd block so the round-trip preserves both marker and injection
        var rawXml = session.Raw.GetXml(firstParagraph.Anchor.Id);
        Assert.Contains("w:p", rawXml);
        var modifiedRaw = rawXml.Replace("</w:p>",
            "<w:r><w:t xml:space=\"preserve\"> RAWINJECTED</w:t></w:r></w:p>");
        var r5 = session.Raw.ReplaceXml(firstParagraph.Anchor.Id, modifiedRaw);
        Assert.True(r5.Success, r5.Error?.Message);
        _output.WriteLine($"[5] Raw.ReplaceXml OK");

        var afterEdits = session.Project();
        _output.WriteLine("──── afterEdits.Markdown ────");
        _output.WriteLine(afterEdits.Markdown);
        _output.WriteLine("──── /afterEdits ────");
        Assert.Contains("SMOKETESTMARKER1", afterEdits.Markdown);
        Assert.Contains("Inserted Heading", afterEdits.Markdown);
        Assert.Contains("RAWINJECTED", afterEdits.Markdown);
        _output.WriteLine($"After edits projection: {afterEdits.Markdown.Length} chars, " +
                          $"{afterEdits.AnchorIndex.Count} anchors");

        // ── Save + reopen round-trip ──────────────────────────────────────
        var saved = session.Save();
        _output.WriteLine($"Saved bytes: {saved.Length}");
        Assert.NotEmpty(saved);

        using var reopened = new DocxSession(saved);
        var reprojected = reopened.Project();
        Assert.Contains("SMOKETESTMARKER1", reprojected.Markdown);
        Assert.Contains("Inserted Heading", reprojected.Markdown);
        Assert.Contains("RAWINJECTED", reprojected.Markdown);
        _output.WriteLine($"Reopened projection: {reprojected.Markdown.Length} chars");

        // ── Verify saved bytes are still a valid DOCX per the SDK ─────────
        using var ms = new MemoryStream(saved);
        using var verify = WordprocessingDocument.Open(ms, isEditable: false);
        Assert.NotNull(verify.MainDocumentPart);
        _output.WriteLine("Saved DOCX opens cleanly via the SDK.");

        // ── Undo all the way back ─────────────────────────────────────────
        int undoCount = 0;
        while (session.Undo()) undoCount++;
        _output.WriteLine($"Undid {undoCount} ops");
        var finalProj = session.Project();
        Assert.DoesNotContain("SMOKETESTMARKER1", finalProj.Markdown);
        Assert.DoesNotContain("RAWINJECTED", finalProj.Markdown);
        _output.WriteLine("Post-undo projection no longer contains markers — undo chain works.");
    }
}
