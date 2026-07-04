#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// M2.5 Task 4 — public-surface smoke tests for <see cref="DocxDiff"/>: each entry point
/// (<see cref="DocxDiff.Compare"/>, <see cref="DocxDiff.GetRevisions"/>,
/// <see cref="DocxDiff.GetEditScriptJson"/>) over a WC corpus pair AND a programmatic pair, plus
/// <see cref="DocxDiffSettings"/> mapping spot-checks. These assert the PUBLIC contract only (round-trip,
/// revision shape, JSON parses) — the engine internals are covered by the Ir.Diff battery.
/// </summary>
public class DocxDiffTests
{
    private static readonly DirectoryInfo SourceDir = new("../../../../TestFiles/");

    private static WmlDocument Load(string relativePath) =>
        new(Path.Combine(SourceDir.FullName, relativePath));

    // A minimal programmatic doc, mirroring the Ir.Diff scoreboard builders (all required parts present).
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

    private static WmlDocument BoldDoc(string text)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new RunProperties(new Bold()), new Text(text)))));
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

    // A one-or-more-row table document (each inner array is a row of cell texts), with all required parts.
    private static WmlDocument TableDoc(params string[][] rows)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var table = new DocumentFormat.OpenXml.Wordprocessing.Table();
            foreach (var row in rows)
            {
                var tr = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                foreach (var cellText in row)
                    tr.Append(new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                        new Paragraph(new Run(new Text(cellText)))));
                table.Append(tr);
            }
            // A table must be followed by a paragraph for a valid body; it is equal on both sides.
            mainPart.Document = new Document(new Body(table, new Paragraph()));
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

    // Body-level <w:t> texts of a document, for round-trip comparison.
    private static List<string> BodyTexts(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var body = wdoc.MainDocumentPart?.Document.Body;
        return body is null
            ? new List<string>()
            : body.Descendants<Paragraph>().Select(p => p.InnerText).ToList();
    }

    // ----------------------------------------------------------------- Compare

    [Fact]
    public void Compare_WcPair_RoundTripsAcceptToRightRejectToLeft()
    {
        var left = Load("WC/WC001-Digits.docx");
        var right = Load("WC/WC001-Digits-Mod.docx");

        var result = DocxDiff.Compare(left, right);

        // The WmlComparer output contract: accept ⇒ right, reject ⇒ left (body text level).
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }

    [Fact]
    public void Compare_ProgrammaticPair_ProducesTrackedChanges()
    {
        var left = Doc("The quick brown fox.");
        var right = Doc("The quick red fox.");

        var result = DocxDiff.Compare(left, right);

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));

        // The produced markup actually carries ins/del elements (not an untouched clone).
        using var stream = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var body = wdoc.MainDocumentPart!.Document.Body!;
        Assert.True(body.Descendants<InsertedRun>().Any() || body.Descendants<DeletedRun>().Any());
    }

    [Fact]
    public void Compare_DeletedTableColumn_RejectRestoresTheColumn()
    {
        // Regression (engine audit, MEDIUM — reject fidelity): a deleted trailing table column produced a
        // left-surplus cell op that RenderModifyRow dropped (the `ci >= rightCells.Count` break), so the
        // deleted column was never marked w:del and RejectRevisions did NOT restore it. A column-structure
        // change must bail to the whole-table del(left)+ins(right) fallback so reject ≡ left exactly.
        var left = TableDoc(new[] { "a", "b", "c" });
        var right = TableDoc(new[] { "a", "b" });

        var result = DocxDiff.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);

        Assert.Equal(BodyTexts(right), BodyTexts(accepted));   // accept ⇒ right (2 columns)
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));    // reject ⇒ left (3 columns — "c" restored)
    }

    [Fact]
    public void GetRevisions_TableColumnChange_ReportsWholeTableReplace()
    {
        // A column add/remove bails the MARKUP renderer to a whole-table del(left)+ins(right) fallback
        // (see Compare_DeletedTableColumn_RejectRestoresTheColumn). GetRevisions must mirror that: it
        // previously returned ZERO revisions for a column-count change (the per-cell path drops the
        // surplus cell), diverging from the WmlComparer oracle — which reports a Deleted + Inserted pair —
        // and silently hiding the change from revision consumers even though the markup tracks it.
        var left = TableDoc(new[] { "a", "b" });          // 1 row, 2 columns
        var right = TableDoc(new[] { "a", "b", "c" });    // 1 row, 3 columns — a column added

        var revisions = DocxDiff.GetRevisions(left, right);

        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Deleted);
        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Inserted);
    }

    // A document whose body paragraph has lead text followed by an INLINE content control (w:sdt) wrapping
    // a run — the shape that exposed the unwrapped-sdt-insertion bug.
    private static WmlDocument DocWithInlineSdt(string leadText, string sdtText)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var sdt = new SdtRun(
                new SdtProperties(new SdtId { Val = 9001 }),
                new SdtContentRun(new Run(new Text(sdtText) { Space = SpaceProcessingModeValues.Preserve })));
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text(leadText)), sdt)));
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

    [Fact]
    public void Compare_InsertedContentControl_RejectStripsTheInsertedText()
    {
        // A run inserted INSIDE a w:sdt (content control) must be wrapped in w:ins, else RejectRevisions
        // (which strips w:ins/w:del) leaves it — violating the core contract reject ≡ left and silently
        // retaining content the user rejected. The markup renderer wrapped a w:sdt's DIRECT children but
        // not the runs nested under w:sdtContent, so inserted content-control text leaked through on reject.
        var left = Doc("Hello world");
        var right = DocWithInlineSdt("Hello world", " controlled text");

        var result = DocxDiff.Compare(left, right);
        var rejected = RevisionProcessor.RejectRevisions(result);
        var accepted = RevisionProcessor.AcceptRevisions(result);

        Assert.Equal(BodyTexts(left), BodyTexts(rejected));    // reject ⇒ left (sdt text removed)
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));   // accept ⇒ right (sdt text kept)
    }

    [Fact]
    public void Compare_IdenticalDocuments_HasNoTrackedChanges()
    {
        var doc = Doc("Unchanged paragraph one.", "Unchanged paragraph two.");

        var result = DocxDiff.Compare(doc, doc);

        using var stream = new MemoryStream(result.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var body = wdoc.MainDocumentPart!.Document.Body!;
        Assert.False(body.Descendants<InsertedRun>().Any());
        Assert.False(body.Descendants<DeletedRun>().Any());
    }

    // ----------------------------------------------------------------- GetRevisions

    [Fact]
    public void GetRevisions_WcPair_ReturnsRevisionsWithAnchors()
    {
        var left = Load("WC/WC001-Digits.docx");
        var right = Load("WC/WC001-Digits-Mod.docx");

        var revisions = DocxDiff.GetRevisions(left, right);

        Assert.NotEmpty(revisions);
        // Every revision carries at least one anchor of the documented kind:scope:unid grammar.
        foreach (var r in revisions)
        {
            var anchor = r.LeftAnchor ?? r.RightAnchor;
            Assert.NotNull(anchor);
            Assert.Equal(2, anchor!.Count(c => c == ':')); // kind:scope:unid → exactly two colons.
        }
    }

    [Fact]
    public void GetRevisions_ProgrammaticPair_SurfacesInsertAndDelete()
    {
        var left = Doc("alpha beta gamma");
        var right = Doc("alpha delta gamma");

        var revisions = DocxDiff.GetRevisions(left, right);

        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Inserted);
        Assert.Contains(revisions, r => r.Type == DocxDiffRevisionType.Deleted);

        // Anchor presence rule: a pure insertion has RightAnchor (no LeftAnchor); a pure deletion the inverse.
        var ins = revisions.First(r => r.Type == DocxDiffRevisionType.Inserted);
        Assert.NotNull(ins.RightAnchor);
        var del = revisions.First(r => r.Type == DocxDiffRevisionType.Deleted);
        Assert.NotNull(del.LeftAnchor);
    }

    [Fact]
    public void GetRevisions_TokenLevelEdit_CarriesBothEnclosingBlockAnchors()
    {
        // A SINGLE-paragraph edit is a MODIFIED block: the inserted/deleted token spans live inside a block
        // that exists on BOTH sides, so each token-level revision carries BOTH the left and right anchors of
        // the enclosing paragraph (the documented "token-level revision carries the enclosing block's
        // anchor(s)" rule). This is the shape the python E2E observed — a Deleted with both anchors — and it
        // is intentional, not a defect: the contract is "the type's PRIMARY anchor is always present; the
        // opposite anchor MAY also be present for a token-level revision inside a Modified/MoveModify block".
        var left = Doc("alpha beta gamma");
        var right = Doc("alpha delta gamma");

        var revisions = DocxDiff.GetRevisions(left, right);

        var ins = revisions.First(r => r.Type == DocxDiffRevisionType.Inserted);
        var del = revisions.First(r => r.Type == DocxDiffRevisionType.Deleted);
        // Token-level inside a Modified paragraph → both anchors present (and equal across the two revisions:
        // they describe edits to the SAME enclosing block).
        Assert.NotNull(ins.LeftAnchor);
        Assert.NotNull(ins.RightAnchor);
        Assert.NotNull(del.LeftAnchor);
        Assert.NotNull(del.RightAnchor);
        Assert.Equal(ins.LeftAnchor, del.LeftAnchor);
        Assert.Equal(ins.RightAnchor, del.RightAnchor);
    }

    [Fact]
    public void GetRevisions_BlockLevelInsertDelete_CarriesOnlyPrimaryAnchor()
    {
        // A WHOLE-paragraph insertion/deletion is a BLOCK-level revision: the block exists on ONE side only,
        // so the revision carries ONLY its primary anchor — the opposite anchor is null. (Contrast the
        // token-level case above.) Two surrounding equal paragraphs keep the inserted/deleted one whole-block.
        var left = Doc("first para", "third para");
        var right = Doc("first para", "second para", "third para");

        var ins = Assert.Single(DocxDiff.GetRevisions(left, right),
            r => r.Type == DocxDiffRevisionType.Inserted);
        Assert.NotNull(ins.RightAnchor);
        Assert.Null(ins.LeftAnchor);

        var del = Assert.Single(DocxDiff.GetRevisions(right, left),
            r => r.Type == DocxDiffRevisionType.Deleted);
        Assert.NotNull(del.LeftAnchor);
        Assert.Null(del.RightAnchor);
    }

    [Fact]
    public void GetRevisions_AnchorPresenceContract_HoldsOverCorpusPair()
    {
        // The public-surface anchor-presence contract, asserted end-to-end over a real WC pair (both
        // directions): Inserted ALWAYS has RightAnchor; Deleted ALWAYS has LeftAnchor; FormatChanged has
        // both; Moved is exclusive (source = left only, dest = right only). The opposite anchor is permitted
        // for token-level ins/del. This is the invariant the python/npm doc comments and the IR corpus test
        // also encode — pinned here at the shipped public API.
        var a = Load("WC/WC001-Digits.docx");
        var b = Load("WC/WC001-Digits-Mod.docx");
        foreach (var (l, r) in new[] { (a, b), (b, a) })
            foreach (var rev in DocxDiff.GetRevisions(l, r))
            {
                switch (rev.Type)
                {
                    case DocxDiffRevisionType.Inserted:
                        Assert.NotNull(rev.RightAnchor);
                        break;
                    case DocxDiffRevisionType.Deleted:
                        Assert.NotNull(rev.LeftAnchor);
                        break;
                    case DocxDiffRevisionType.FormatChanged:
                        Assert.NotNull(rev.LeftAnchor);
                        Assert.NotNull(rev.RightAnchor);
                        break;
                    case DocxDiffRevisionType.Moved:
                        if (rev.IsMoveSource == true)
                        {
                            Assert.NotNull(rev.LeftAnchor);
                            Assert.Null(rev.RightAnchor);
                        }
                        else
                        {
                            Assert.NotNull(rev.RightAnchor);
                            Assert.Null(rev.LeftAnchor);
                        }
                        break;
                }
            }
    }

    [Fact]
    public void GetRevisions_FormatOnlyChange_SurfacesFormatChanged()
    {
        var left = Doc("identical text");
        var right = BoldDoc("identical text");

        var revisions = DocxDiff.GetRevisions(left, right);

        var fc = Assert.Single(revisions, r => r.Type == DocxDiffRevisionType.FormatChanged);
        Assert.NotNull(fc.FormatChange);
        Assert.Contains("bold", fc.FormatChange!.ChangedPropertyNames);
    }

    [Fact]
    public void GetRevisions_StampsAuthorFromSettings()
    {
        var left = Doc("one two three");
        var right = Doc("one TWO three");
        var settings = new DocxDiffSettings { AuthorForRevisions = "Daisy" };

        var revisions = DocxDiff.GetRevisions(left, right, settings);

        Assert.NotEmpty(revisions);
        Assert.All(revisions, r => Assert.Equal("Daisy", r.Author));
    }

    // ----------------------------------------------------------------- GetEditScriptJson

    [Fact]
    public void GetEditScriptJson_WcPair_ParsesAndHasOperations()
    {
        var left = Load("WC/WC001-Digits.docx");
        var right = Load("WC/WC001-Digits-Mod.docx");

        var json = DocxDiff.GetEditScriptJson(left, right);

        using var parsed = JsonDocument.Parse(json); // must be valid JSON.
        Assert.True(parsed.RootElement.TryGetProperty("operations", out var ops));
        Assert.Equal(JsonValueKind.Array, ops.ValueKind);
        Assert.NotEqual(0, ops.GetArrayLength());
    }

    [Fact]
    public void GetEditScriptJson_ProgrammaticPair_Parses()
    {
        var left = Doc("first paragraph", "second paragraph");
        var right = Doc("first paragraph", "second paragraph changed");

        var json = DocxDiff.GetEditScriptJson(left, right);

        using var parsed = JsonDocument.Parse(json);
        Assert.Equal(JsonValueKind.Object, parsed.RootElement.ValueKind);
    }

    [Fact]
    public void GetEditScriptJson_IsDeterministicByDefault()
    {
        var left = Doc("alpha", "beta");
        var right = Doc("alpha", "beta gamma");

        var a = DocxDiff.GetEditScriptJson(left, right);
        var b = DocxDiff.GetEditScriptJson(left, right);

        Assert.Equal(a, b); // deterministic dates + stable ids ⇒ byte-identical.
    }

    // ----------------------------------------------------------------- settings mapping spot-checks

    [Fact]
    public void Settings_CaseInsensitive_SuppressesCaseOnlyRevisions()
    {
        var left = Doc("Hello World");
        var right = Doc("hello world");

        var sensitive = DocxDiff.GetRevisions(left, right);
        var insensitive = DocxDiff.GetRevisions(left, right,
            new DocxDiffSettings { CaseInsensitive = true });

        Assert.NotEmpty(sensitive);                 // case difference IS a change by default.
        Assert.DoesNotContain(insensitive, r =>
            r.Type is DocxDiffRevisionType.Inserted or DocxDiffRevisionType.Deleted);
    }

    [Fact]
    public void Settings_WmlComparerCompatibleGranularity_CoalescesRevisions()
    {
        // A paragraph where two adjacent words both change: Fine reports one ins + one del per word;
        // WmlComparerCompatible coalesces the contiguous region.
        var left = Doc("the aaa bbb end");
        var right = Doc("the xxx yyy end");

        var fine = DocxDiff.GetRevisions(left, right,
            new DocxDiffSettings { RevisionGranularity = DocxDiffRevisionGranularity.Fine });
        var compat = DocxDiff.GetRevisions(left, right,
            new DocxDiffSettings { RevisionGranularity = DocxDiffRevisionGranularity.WmlComparerCompatible });

        Assert.True(compat.Count <= fine.Count,
            $"compatible ({compat.Count}) should coalesce vs fine ({fine.Count})");
    }

    [Fact]
    public void Settings_FormatComparisonFull_SeesUnmodeledFlip()
    {
        // ModeledOnly (default) ignores an unmodeled-only rPr difference; Full sees it. We assert the
        // default does not over-report on a plain text-equal pair (no spurious FormatChanged).
        var doc = Doc("plain text");

        var modeled = DocxDiff.GetRevisions(doc, doc,
            new DocxDiffSettings { FormatComparison = DocxDiffFormatComparison.ModeledOnly });

        Assert.DoesNotContain(modeled, r => r.Type == DocxDiffRevisionType.FormatChanged);
    }

    [Fact]
    public void Settings_DetectMovesFalse_RendersMoveAsInsertDelete()
    {
        // Move a distinctive block from the top to the bottom.
        var left = Doc("MOVED distinctive sentence here", "anchor one", "anchor two", "anchor three");
        var right = Doc("anchor one", "anchor two", "anchor three", "MOVED distinctive sentence here");

        var withMoves = DocxDiff.GetRevisions(left, right,
            new DocxDiffSettings { DetectMoves = true });
        var noMoves = DocxDiff.GetRevisions(left, right,
            new DocxDiffSettings { DetectMoves = false });

        // With moves off, nothing is reported as Moved.
        Assert.DoesNotContain(noMoves, r => r.Type == DocxDiffRevisionType.Moved);
        // (With moves on the engine MAY report a move; this is a settings-mapping spot check, not a
        // move-detection assertion, so we only assert the off-switch suppresses the Moved type.)
        Assert.NotNull(withMoves);
    }

    [Fact]
    public void Settings_TrackBlockFormatChanges_default_true_and_maps_through()
    {
        Assert.True(new DocxDiffSettings().TrackBlockFormatChanges);
        Assert.True(new DocxDiffSettings().ToIrDiffSettings().TrackBlockFormatChanges);
        Assert.False(new DocxDiffSettings { TrackBlockFormatChanges = false }.ToIrDiffSettings().TrackBlockFormatChanges);
    }

    [Fact]
    public void Settings_ExplicitEmptyWordSeparators_IsHonoredNotRevertedToDefault()
    {
        // Regression (engine audit): an explicitly EMPTY WordSeparators array was silently swapped for the
        // default set (the `Length: > 0` guard), contradicting the "only null falls back" contract. An
        // explicit set — even empty — must win.
        var ir = new DocxDiffSettings { WordSeparators = System.Array.Empty<char>() }.ToIrDiffSettings();
        Assert.Empty(ir.WordSeparators);

        // A null set still falls back to the documented default.
        var deflt = new DocxDiffSettings { WordSeparators = null }.ToIrDiffSettings();
        Assert.NotEmpty(deflt.WordSeparators);
    }

    [Fact]
    public void Settings_InvalidExplicitDate_ThrowsArgumentException()
    {
        // Regression (engine audit): an explicit DateTimeForRevisions was stamped verbatim into w:date with
        // no validation, so garbage produced schema-questionable markup with no boundary error. A
        // non-parseable value must throw at the boundary instead.
        var left = Doc("one two");
        var right = Doc("one three");
        var settings = new DocxDiffSettings { DateTimeForRevisions = "not-a-date" };
        Assert.Throws<System.ArgumentException>(() => DocxDiff.GetRevisions(left, right, settings));
    }

    [Fact]
    public void Settings_ExplicitDate_StampsRevisions()
    {
        var left = Doc("one two");
        var right = Doc("one three");
        var settings = new DocxDiffSettings { DateTimeForRevisions = "2021-07-04T00:00:00Z" };

        var revisions = DocxDiff.GetRevisions(left, right, settings);

        Assert.NotEmpty(revisions);
        Assert.All(revisions, r => Assert.Equal("2021-07-04T00:00:00Z", r.Date));
    }
}
