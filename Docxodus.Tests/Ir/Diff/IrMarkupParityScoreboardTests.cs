#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.4 Task 4 — the MARKUP parity scoreboard. The companion to <see cref="IrParityScoreboardTests"/> (which
/// scores the <c>GetRevisions</c>-count rows): this scores the MARKUP_BLOCKED rows of the WmlComparer suite —
/// the cases whose original assertion rides on PRODUCED OOXML (native <c>w:moveFrom</c>/<c>w:moveTo</c>,
/// <c>w:rPrChange</c>, accept/reject round-trip, schema validity, numbering continuity, no-throw on body-level
/// markers, thread-safety) — re-expressed against <see cref="IrMarkupRenderer"/>'s output and scored PASS /
/// DEVIATION / FAIL with the same soft-assert ratchet.
///
/// <para><b>Why a second scoreboard.</b> The GetRevisions scoreboard reads revisions straight off the edit
/// script (no document is produced). These rows instead PRODUCE a tracked-revisions document via
/// <see cref="IrMarkupRenderer.Render"/> and assert on it — the exact surface M2.4 Task 4 delivers. Together
/// the two scoreboards span the full WmlComparer parity bar (the program directive's 218).</para>
/// </summary>
[Trait("Category", "Parity")]
public class IrMarkupParityScoreboardTests
{
    private readonly ITestOutputHelper _out;
    public IrMarkupParityScoreboardTests(ITestOutputHelper output) => _out = output;

    private static readonly XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // ------------------------------------------------------------------ render helper

    /// <summary>Build the edit script over two docs (Accept-view IRs, the adapter's read options) and render
    /// the native tracked-revisions markup document.</summary>
    private static WmlDocument RenderMarkup(WmlDocument left, WmlDocument right, IrDiffSettings? settings = null)
    {
        settings ??= new IrDiffSettings();
        var ro = new IrReaderOptions { RetainSources = false, RevisionView = Docxodus.Ir.RevisionView.Accept };
        var irLeft = IrReader.Read(left, ro);
        var irRight = IrReader.Read(right, ro);
        var script = IrEditScriptBuilder.Build(irLeft, irRight, settings);
        return IrMarkupRenderer.Render(script, left, right, settings);
    }

    private static XElement BodyOf(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        return wd.MainDocumentPart!.GetXDocument().Root!.Element(w + "body")!;
    }

    // ------------------------------------------------------------------ the scoreboard run

    [Fact]
    public void Markup_parity_scoreboard_over_markup_blocked_WmlComparer_cases()
    {
        var board = new Scoreboard(DocumentedDeviations);

        ScoreMoveMarkupCases(board);
        ScoreFormatChangeMarkupCases(board);
        ScoreLegalNumberingCases(board);
        ScoreBodyLevelCases(board);
        ScoreParallelRaceCase(board);

        board.Report(_out);

        // The markup floor: every ported MARKUP_BLOCKED case lands PASS (markup correct) or DEVIATION
        // (adjudicated reader/aligner-level difference — see DocumentedDeviations), never an undocumented FAIL.
        // 34 = 16 native-move-markup + 6 w:rPrChange + 5 legal-numbering + 6 body-level + 1 parallel-race
        // assertions ported from WmlComparerMoveDetection/FormatChange/LegalNumbering/BodyLevel/ParallelRace.
        // Combined with the 179 GetRevisions-scoreboard rows (IrParityScoreboardTests), the two scoreboards span
        // the WmlComparer parity surface the M2.4 gate measures (see the ## M2.4 Outcome gate report).
        const int MarkupParityFloor = 34; // M2.4 Task 4 — the ported MARKUP_BLOCKED set, all PASS-or-deviation.
        Assert.True(board.Total > 0, "Markup scoreboard scored no cases.");
        Assert.Equal(board.Total, board.Pass + board.Deviation + board.Fail);
        Assert.True(board.Pass + board.Deviation >= MarkupParityFloor,
            $"MARKUP PARITY REGRESSION: {board.Pass} PASS + {board.Deviation} DEVIATION = " +
            $"{board.Pass + board.Deviation} < floor {MarkupParityFloor}. Undocumented FAILs: " +
            string.Join(", ", board.FailingIds));
    }

    private static readonly IReadOnlyDictionary<string, string> DocumentedDeviations =
        new Dictionary<string, string>
        {
            // No markup-level deviations: every ported markup assertion holds against IrMarkupRenderer's output.
            // (The 6 reader/aligner/rId-remap deviations are corpus accept/reject ROUND-TRIP pairs, catalogued in
            // IrMarkupRendererTests.Task4BlockedPairs — they are not GetRevisions/markup-shape assertions and so
            // are not ported here; the markup SHAPES the programmatic fixtures assert all pass.)
        };

    // ------------------------------------------------------------------ move markup (16)

    private void ScoreMoveMarkupCases(Scoreboard board)
    {
        var swap2L = new[] { "First paragraph with enough words for move detection.",
                             "Second paragraph with sufficient content here today." };
        var swap2R = new[] { "Second paragraph with sufficient content here today.",
                             "First paragraph with enough words for move detection." };
        var moveSettings = new IrDiffSettings { MoveSimilarityThreshold = 0.8, MoveMinimumTokenCount = 3 };

        board.Score("MoveMarkup-ContainsMoveFrom", "D", () =>
            Assert.True(MoveBody(swap2L, swap2R, moveSettings).Descendants(w + "moveFrom").Any(), "has moveFrom"));

        board.Score("MoveMarkup-ContainsMoveTo", "D", () =>
            Assert.True(MoveBody(swap2L, swap2R, moveSettings).Descendants(w + "moveTo").Any(), "has moveTo"));

        board.Score("MoveMarkup-ContainsRangeMarkers", "D", () =>
        {
            var b = MoveBody(swap2L, swap2R, moveSettings);
            Assert.True(b.Descendants(w + "moveFromRangeStart").Any());
            Assert.True(b.Descendants(w + "moveFromRangeEnd").Any());
            Assert.True(b.Descendants(w + "moveToRangeStart").Any());
            Assert.True(b.Descendants(w + "moveToRangeEnd").Any());
            Assert.Equal(b.Descendants(w + "moveFromRangeStart").Count(), b.Descendants(w + "moveFromRangeEnd").Count());
            Assert.Equal(b.Descendants(w + "moveToRangeStart").Count(), b.Descendants(w + "moveToRangeEnd").Count());
        });

        board.Score("MoveMarkup-LinkPairsViaName", "D", () =>
        {
            var b = MoveBody(swap2L, swap2R, moveSettings);
            var from = b.Descendants(w + "moveFromRangeStart").Select(e => (string?)e.Attribute(w + "name")).ToHashSet();
            var to = b.Descendants(w + "moveToRangeStart").Select(e => (string?)e.Attribute(w + "name")).ToHashSet();
            Assert.NotEmpty(from);
            Assert.True(from.SetEquals(to), "from/to names pair");
        });

        board.Score("MoveMarkup-WhenDisabledNoMoveElements", "D", () =>
        {
            var b = MoveBody(swap2L, swap2R, new IrDiffSettings { RenderMoves = false });
            Assert.Empty(b.Descendants(w + "moveFrom"));
            Assert.Empty(b.Descendants(w + "moveTo"));
            Assert.True(b.Descendants(w + "ins").Any() || b.Descendants(w + "del").Any());
        });

        board.Score("MoveMarkup-RequiredAttributes", "D", () =>
        {
            var b = MoveBody(swap2L, swap2R, moveSettings with { AuthorForRevisions = "TestAuthor" });
            foreach (var e in b.Descendants(w + "moveFrom").Concat(b.Descendants(w + "moveTo")))
            {
                Assert.NotNull(e.Attribute(w + "id"));
                Assert.Equal("TestAuthor", (string?)e.Attribute(w + "author"));
                Assert.NotNull(e.Attribute(w + "date"));
            }
        });

        board.Score("MoveMarkup-RangeIdsProperlyPaired", "D", () =>
        {
            var b = MoveBody(swap2L, swap2R, moveSettings);
            var fromStart = b.Descendants(w + "moveFromRangeStart").Select(e => (string?)e.Attribute(w + "id")).ToHashSet();
            var fromEnd = b.Descendants(w + "moveFromRangeEnd").Select(e => (string?)e.Attribute(w + "id")).ToHashSet();
            var toStart = b.Descendants(w + "moveToRangeStart").Select(e => (string?)e.Attribute(w + "id")).ToHashSet();
            var toEnd = b.Descendants(w + "moveToRangeEnd").Select(e => (string?)e.Attribute(w + "id")).ToHashSet();
            Assert.True(fromStart.SetEquals(fromEnd), "moveFrom range ids pair");
            Assert.True(toStart.SetEquals(toEnd), "moveTo range ids pair");
        });

        board.Score("SimplifyMoveMarkup-ConvertMoveFromToDel", "D", () =>
        {
            var b = MoveBody(swap2L, swap2R, moveSettings with { SimplifyMoveMarkup = true });
            Assert.Empty(b.Descendants(w + "moveFrom"));
            Assert.Empty(b.Descendants(w + "moveTo"));
            Assert.True(b.Descendants(w + "del").Any());
        });

        board.Score("SimplifyMoveMarkup-ConvertMoveToToIns", "D", () =>
            Assert.True(MoveBody(swap2L, swap2R, moveSettings with { SimplifyMoveMarkup = true })
                .Descendants(w + "ins").Any()));

        board.Score("SimplifyMoveMarkup-RemoveRangeMarkers", "D", () =>
        {
            var b = MoveBody(swap2L, swap2R, moveSettings with { SimplifyMoveMarkup = true });
            Assert.Empty(b.Descendants(w + "moveFromRangeStart"));
            Assert.Empty(b.Descendants(w + "moveFromRangeEnd"));
            Assert.Empty(b.Descendants(w + "moveToRangeStart"));
            Assert.Empty(b.Descendants(w + "moveToRangeEnd"));
        });

        board.Score("SimplifyMoveMarkup-PreserveAttributes", "D", () =>
        {
            var b = MoveBody(swap2L, swap2R, moveSettings with { SimplifyMoveMarkup = true, AuthorForRevisions = "TestAuthor" });
            foreach (var e in b.Descendants(w + "del").Concat(b.Descendants(w + "ins")))
            {
                Assert.NotNull(e.Attribute(w + "author"));
                Assert.NotNull(e.Attribute(w + "id"));
            }
        });

        board.Score("SimplifyMoveMarkup-WhenFalsePreserveMoveElements", "D", () =>
        {
            var b = MoveBody(swap2L, swap2R, moveSettings with { SimplifyMoveMarkup = false });
            Assert.True(b.Descendants(w + "moveFrom").Any());
            Assert.True(b.Descendants(w + "moveTo").Any());
        });

        // Mixed-change uniqueness: moves + an edit + a delete; non-range ins/del/moveFrom/moveTo ids unique.
        board.Score("MoveMarkup-MixedChangesUniqueIds", "D", () =>
        {
            var l = new[] { "Alpha paragraph with enough words for move detection here.",
                            "Beta paragraph that will be modified by an edit later.",
                            "Gamma paragraph that stays unchanged across the revision." };
            var r = new[] { "Gamma paragraph that stays unchanged across the revision.",
                            "Alpha paragraph with enough words for move detection here.",
                            "Beta paragraph that will be CHANGED by an edit later." };
            var b = MoveBody(l, r, moveSettings);
            var ids = b.Descendants().Where(e => e.Name == w + "ins" || e.Name == w + "del" ||
                                                 e.Name == w + "moveFrom" || e.Name == w + "moveTo")
                .Select(e => (string?)e.Attribute(w + "id")).Where(s => s != null).ToList();
            Assert.Equal(ids.Count, ids.Distinct().Count());
        });

        board.Score("MoveMarkup-MoveNamesFollowConvention", "D", () =>
        {
            var b = MoveBody(swap2L, swap2R, moveSettings);
            var names = b.Descendants(w + "moveFromRangeStart").Select(e => (string?)e.Attribute(w + "name")).ToList();
            Assert.NotEmpty(names);
            Assert.All(names, n => Assert.StartsWith("move", n));
            Assert.DoesNotContain("", names);
        });

        board.Score("MoveMarkup-RecognizedByWmlComparerGetRevisions", "D", () =>
        {
            // THE ORACLE: WmlComparer.GetRevisions over OUR output sees Moved revisions.
            var rendered = RenderMarkup(MoveDoc(swap2L), MoveDoc(swap2R), moveSettings);
            var moved = WmlComparer.GetRevisions(rendered, new WmlComparerSettings())
                .Count(r => r.RevisionType == WmlComparer.WmlComparerRevisionType.Moved);
            Assert.True(moved >= 2, $"GetRevisions sees {moved} Moved");
        });

        board.Score("MoveMarkup-ValidSchema", "D", () =>
            Assert.Equal(0, SchemaErrorCount(RenderMarkup(MoveDoc(swap2L), MoveDoc(swap2R), moveSettings))));
    }

    private static XElement MoveBody(string[] left, string[] right, IrDiffSettings settings) =>
        BodyOf(RenderMarkup(MoveDoc(left), MoveDoc(right), settings));

    // ------------------------------------------------------------------ format change markup (6)

    private void ScoreFormatChangeMarkupCases(Scoreboard board)
    {
        board.Score("FormatChange-AddBold-rPrChange", "E", () =>
            Assert.NotEmpty(FmtBody(Para("sample text here"), BoldPara("sample text here")).Descendants(w + "rPrChange")));

        board.Score("FormatChange-RemoveBold-rPrChange", "E", () =>
            Assert.NotEmpty(FmtBody(BoldPara("sample text here"), Para("sample text here")).Descendants(w + "rPrChange")));

        board.Score("FormatChange-BoldToItalic-rPrChange", "E", () =>
            Assert.NotEmpty(FmtBody(BoldPara("sample text here"), ItalicPara("sample text here")).Descendants(w + "rPrChange")));

        board.Score("FormatChange-AddMultiple-rPrChange", "E", () =>
            Assert.NotEmpty(FmtBody(Para("sample text here"), BoldItalicPara("sample text here")).Descendants(w + "rPrChange")));

        board.Score("FormatChange-RequiredAttributes", "E", () =>
        {
            var rendered = RenderMarkup(Para("sample text here"), BoldPara("sample text here"),
                new IrDiffSettings { AuthorForRevisions = "Test Author" });
            var changes = BodyOf(rendered).Descendants(w + "rPrChange").ToList();
            Assert.NotEmpty(changes);
            foreach (var c in changes)
            {
                Assert.NotNull(c.Attribute(w + "id"));
                Assert.Equal("Test Author", (string?)c.Attribute(w + "author"));
                Assert.NotNull(c.Attribute(w + "date"));
            }
        });

        board.Score("FormatChange-ContainsOldProperties", "E", () =>
        {
            // Bold → plain: the rPrChange's inner rPr must carry the OLD bold property.
            var change = FmtBody(BoldPara("sample text here"), Para("sample text here"))
                .Descendants(w + "rPrChange").First();
            var oldRpr = change.Element(w + "rPr");
            Assert.NotNull(oldRpr);
            Assert.NotNull(oldRpr!.Element(w + "b"));
        });
    }

    private static XElement FmtBody(WmlDocument left, WmlDocument right) =>
        BodyOf(RenderMarkup(left, right, new IrDiffSettings()));

    // ------------------------------------------------------------------ legal numbering (5)

    private void ScoreLegalNumberingCases(Scoreboard board)
    {
        board.Score("LegalNum-001-OriginalHasLegal-Preserves", "G", () =>
        {
            var rendered = RenderMarkup(NumberedList("First item", legal: true), NumberedList("First item modified", legal: false));
            Assert.True(HasLegalNumbering(rendered), "compared preserves isLgl");
            Assert.Equal(0, SchemaErrorCount(rendered));
        });

        board.Score("LegalNum-002-RevisedHasLegal-Preserves", "G", () =>
        {
            var rendered = RenderMarkup(NumberedList("First item", legal: false), NumberedList("First item modified", legal: true));
            Assert.True(HasLegalNumbering(rendered), "compared preserves isLgl");
            Assert.Equal(0, SchemaErrorCount(rendered));
        });

        board.Score("LegalNum-003-BothHaveLegal-Preserves", "G", () =>
        {
            var rendered = RenderMarkup(NumberedList("First item", legal: true), NumberedList("First item modified", legal: true));
            Assert.True(HasLegalNumbering(rendered), "compared preserves isLgl");
            Assert.Equal(0, SchemaErrorCount(rendered));
        });

        board.Score("LegalNum-004-MultiLevel-Preserves", "G", () =>
        {
            var rendered = RenderMarkup(
                MultiLevelNumbering(new[] { "Section 1", "Subsection 1.1", "Sub-subsection 1.1.1" }, legal: false),
                MultiLevelNumbering(new[] { "Section 1", "Subsection 1.1", "Sub-subsection 1.1.1 modified" }, legal: true));
            Assert.True(HasLegalNumbering(rendered), "compared preserves isLgl");
            Assert.Equal(0, SchemaErrorCount(rendered));
        });

        board.Score("LegalNum-005-DifferentNumIds-Merges", "G", () =>
        {
            var rendered = RenderMarkup(NumberedList("First item", legal: false, numId: 1),
                                        NumberedList("First item modified", legal: true, numId: 2));
            Assert.True(AbstractNumCount(rendered) >= 1, "numbering carried");
            Assert.Equal(0, SchemaErrorCount(rendered));
        });
    }

    // ------------------------------------------------------------------ body-level elements (6)

    private void ScoreBodyLevelCases(Scoreboard board)
    {
        var bmStart = new XElement(w + "bookmarkStart", new XAttribute(w + "id", "1"), new XAttribute(w + "name", "MARK1"));
        var bmEnd = new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "1"));
        var permStart = new XElement(w + "permStart", new XAttribute(w + "id", "1"), new XAttribute(w + "edGrp", "everyone"));
        var permEnd = new XElement(w + "permEnd", new XAttribute(w + "id", "1"));
        var proof1 = new XElement(w + "proofErr", new XAttribute(w + "type", "spellStart"));
        var proof2 = new XElement(w + "proofErr", new XAttribute(w + "type", "spellEnd"));

        board.Score("BodyBookmark-001-NoThrow", "BL", () =>
        {
            var before = BodyLevelDoc(bmStart, bmEnd, "First paragraph.", "Second paragraph.");
            var after = BodyLevelDoc(null, null, "First paragraph.", "Second paragraph modified.");
            var rendered = RenderMarkup(before, after);
            Assert.NotNull(rendered.DocumentByteArray);
        });

        board.Score("BodyBookmark-002-ReverseDirection", "BL", () =>
        {
            var before = BodyLevelDoc(null, null, "First paragraph.", "Second paragraph.");
            var after = BodyLevelDoc(bmStart, bmEnd, "First paragraph.", "Second paragraph modified.");
            var rendered = RenderMarkup(before, after);
            Assert.NotNull(rendered.DocumentByteArray);
        });

        board.Score("BodyPerm-001-NoThrow", "BL", () =>
        {
            var before = BodyLevelDoc(permStart, permEnd, "Para A.", "Para B.");
            var after = BodyLevelDoc(null, null, "Para A.", "Para B modified.");
            Assert.NotNull(RenderMarkup(before, after).DocumentByteArray);
        });

        board.Score("BodyProofErr-001-NoThrow", "BL", () =>
        {
            var before = BodyLevelDoc(proof1, proof2, "Para A.", "Para B.");
            var after = BodyLevelDoc(null, null, "Para A.", "Para B modified.");
            Assert.NotNull(RenderMarkup(before, after).DocumentByteArray);
        });

        board.Score("BodyBookmark-003-BothSides", "BL", () =>
        {
            var before = BodyLevelDoc(bmStart, bmEnd, "Para A.", "Para B.");
            var after = BodyLevelDoc(
                new XElement(w + "bookmarkStart", new XAttribute(w + "id", "2"), new XAttribute(w + "name", "MARK2")),
                new XElement(w + "bookmarkEnd", new XAttribute(w + "id", "2")), "Para A modified.", "Para C.");
            Assert.NotNull(RenderMarkup(before, after).DocumentByteArray);
        });

        board.Score("BodyLevel-RenderedSchemaValid", "BL", () =>
        {
            var before = BodyLevelDoc(bmStart, bmEnd, "First paragraph.", "Second paragraph.");
            var after = BodyLevelDoc(null, null, "First paragraph.", "Second paragraph modified.");
            // No NEW schema errors over the inputs' baseline.
            int baseline = Math.Max(SchemaErrorCount(before), SchemaErrorCount(after));
            Assert.True(SchemaErrorCount(RenderMarkup(before, after)) <= baseline, "no new schema errors");
        });
    }

    // ------------------------------------------------------------------ parallel race (1)

    private void ScoreParallelRaceCase(Scoreboard board)
    {
        board.Score("ParallelRace-NoStatics-ThreadSafe", "PR", () =>
        {
            // The IR markup renderer must be thread-safe BY CONSTRUCTION (no mutable statics): 16 concurrent
            // renders of distinct inputs must all succeed and each round-trip independently.
            var tasks = Enumerable.Range(0, 16).Select(salt => Task.Run(() =>
            {
                var left = MoveDoc(new[] { $"Race paragraph {salt} alpha bravo charlie delta echo.",
                                           $"Race paragraph {salt} foxtrot golf hotel india juliet." });
                var right = MoveDoc(new[] { $"Race paragraph {salt} foxtrot golf hotel india juliet.",
                                            $"Race paragraph {salt} alpha bravo CHANGED charlie delta echo." });
                var rendered = RenderMarkup(left, right, new IrDiffSettings { MoveSimilarityThreshold = 0.6 });
                // Independently round-trip each: accept ≡ right, reject ≡ left (by body text).
                var acc = BodyTextHashes(RevisionProcessor.AcceptRevisions(rendered));
                var rej = BodyTextHashes(RevisionProcessor.RejectRevisions(rendered));
                return acc.SequenceEqual(BodyTextHashes(right)) && rej.SequenceEqual(BodyTextHashes(left));
            })).ToArray();
            Task.WaitAll(tasks);
            Assert.All(tasks, t => Assert.True(t.Result, "each concurrent render round-trips"));
        });
    }

    private static List<string> BodyTextHashes(WmlDocument doc)
    {
        var ir = IrReader.Read(doc, new IrReaderOptions { RetainSources = false, RevisionView = Docxodus.Ir.RevisionView.Accept });
        var blocks = ir.Body.Blocks.ToList();
        if (blocks.Count > 0 && blocks[^1] is IrSectionBreak)
            blocks.RemoveAt(blocks.Count - 1);
        return blocks.OfType<IrParagraph>().Select(p => p.ContentHash.ToHex()).ToList();
    }

    // ------------------------------------------------------------------ fixture builders

    private static WmlDocument MoveDoc(string[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                paragraphs.Select(t => new Paragraph(new Run(new Text(t))))));
            AddDefaults(mainPart);
            doc.Save();
        }
        return new WmlDocument("move.docx", stream.ToArray());
    }

    private static WmlDocument Para(string text) => RunDoc(text, null);
    private static WmlDocument BoldPara(string text) => RunDoc(text, new RunProperties(new Bold()));
    private static WmlDocument ItalicPara(string text) => RunDoc(text, new RunProperties(new Italic()));
    private static WmlDocument BoldItalicPara(string text) => RunDoc(text, new RunProperties(new Bold(), new Italic()));

    private static WmlDocument RunDoc(string text, RunProperties? rPr)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var run = rPr != null ? new Run(rPr, new Text(text)) : new Run(new Text(text));
            mainPart.Document = new Document(new Body(new Paragraph(run)));
            AddDefaults(mainPart);
            doc.Save();
        }
        return new WmlDocument("fmt.docx", stream.ToArray());
    }

    private static WmlDocument NumberedList(string text, bool legal, int numId = 1)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(new Paragraph(
                new ParagraphProperties(new NumberingProperties(
                    new NumberingLevelReference { Val = 0 }, new NumberingId { Val = numId })),
                new Run(new Text(text)))));
            AddDefaults(mainPart);
            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            var level = new Level(
                new StartNumberingValue { Val = 1 },
                new NumberingFormat { Val = NumberFormatValues.Decimal },
                new LevelText { Val = "%1." },
                new LevelJustification { Val = LevelJustificationValues.Left }) { LevelIndex = 0 };
            if (legal)
                level.InsertAfter(new IsLegalNumberingStyle(), level.GetFirstChild<NumberingFormat>());
            var abstractNum = new AbstractNum(level) { AbstractNumberId = 1 };
            var numInstance = new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = numId };
            numberingPart.Numbering = new Numbering(abstractNum, numInstance);
            doc.Save();
        }
        return new WmlDocument("num.docx", stream.ToArray());
    }

    private static WmlDocument MultiLevelNumbering(string[] texts, bool legal, int numId = 1)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body();
            for (int i = 0; i < texts.Length; i++)
                body.AppendChild(new Paragraph(
                    new ParagraphProperties(new NumberingProperties(
                        new NumberingLevelReference { Val = i }, new NumberingId { Val = numId })),
                    new Run(new Text(texts[i]))));
            mainPart.Document = new Document(body);
            AddDefaults(mainPart);
            var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            var abstractNum = new AbstractNum { AbstractNumberId = 1 };
            for (int i = 0; i < texts.Length; i++)
            {
                var level = new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = $"%{i + 1}." },
                    new LevelJustification { Val = LevelJustificationValues.Left }) { LevelIndex = i };
                if (legal)
                    level.InsertAfter(new IsLegalNumberingStyle(), level.GetFirstChild<NumberingFormat>());
                abstractNum.AppendChild(level);
            }
            var numInstance = new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = numId };
            numberingPart.Numbering = new Numbering(abstractNum, numInstance);
            doc.Save();
        }
        return new WmlDocument("mlnum.docx", stream.ToArray());
    }

    /// <summary>Build a doc whose body carries optional body-level marker elements between the first and
    /// remaining paragraphs (mirrors WmlComparerBodyLevelElementsTests.BuildDocxWithBodyLevelElement).</summary>
    private static WmlDocument BodyLevelDoc(XElement? marker1, XElement? marker2, params string[] paraTexts)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            AddDefaults(mainPart);
            var bodyXml = new System.Text.StringBuilder("<w:body xmlns:w=\"" + w + "\">");
            for (int i = 0; i < paraTexts.Length; i++)
            {
                bodyXml.Append($"<w:p><w:r><w:t>{System.Security.SecurityElement.Escape(paraTexts[i])}</w:t></w:r></w:p>");
                if (i == 0 && marker1 != null)
                {
                    bodyXml.Append(marker1);
                    if (marker2 != null) bodyXml.Append(marker2);
                }
            }
            bodyXml.Append("</w:body>");
            var docXml = new XElement(w + "document", new XAttribute(XNamespace.Xmlns + "w", w.NamespaceName),
                XElement.Parse(bodyXml.ToString()));
            mainPart.GetXDocument().Add(docXml);   // replace placeholder content
            mainPart.PutXDocument();
            doc.Save();
        }
        return new WmlDocument("bodylevel.docx", stream.ToArray());
    }

    private static void AddDefaults(MainDocumentPart mainPart)
    {
        if (mainPart.StyleDefinitionsPart == null)
        {
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()));
        }
        if (mainPart.DocumentSettingsPart == null)
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
    }

    // ------------------------------------------------------------------ inspection helpers

    private static bool HasLegalNumbering(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        return wd.MainDocumentPart!.NumberingDefinitionsPart?.GetXDocument().Descendants(w + "isLgl").Any() ?? false;
    }

    private static int AbstractNumCount(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        return wd.MainDocumentPart!.NumberingDefinitionsPart?.GetXDocument().Descendants(w + "abstractNum").Count() ?? 0;
    }

    private static int SchemaErrorCount(WmlDocument doc)
    {
        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
        using var wd = WordprocessingDocument.Open(ms, false);
        var validator = new OpenXmlValidator();
        return validator.Validate(wd).Count(e =>
            e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema &&
            !OxPt.WcTests.ExpectedErrors.Contains(e.Description));
    }

    // ------------------------------------------------------------------ soft-assert scoreboard

    private sealed class SoftAssertException : Exception
    {
        public SoftAssertException(string message) : base(message) { }
    }

    private enum RowState { Pass, Deviation, Fail }

    private sealed class Scoreboard
    {
        private readonly List<(string Id, string Category, RowState State, string Detail)> _rows = new();
        private readonly IReadOnlyDictionary<string, string> _deviations;

        public Scoreboard(IReadOnlyDictionary<string, string> deviations) => _deviations = deviations;

        public int Pass { get; private set; }
        public int Deviation { get; private set; }
        public int Fail { get; private set; }
        public int Total => _rows.Count;
        public IEnumerable<string> FailingIds => _rows.Where(r => r.State == RowState.Fail).Select(r => r.Id);

        public void Score(string id, string category, Action body)
        {
            string? failDetail = null;
            try { body(); }
            catch (Exception ex) { failDetail = $"{ex.GetType().Name}: {ex.Message}"; }

            if (failDetail is null)
            {
                if (_deviations.ContainsKey(id))
                {
                    _rows.Add((id, category, RowState.Fail,
                        "STALE DEVIATION: this case now PASSES — remove it from DocumentedDeviations."));
                    Fail++;
                }
                else { _rows.Add((id, category, RowState.Pass, "")); Pass++; }
                return;
            }
            if (_deviations.TryGetValue(id, out var reason))
            {
                _rows.Add((id, category, RowState.Deviation, $"{failDetail}  —  {reason}"));
                Deviation++;
            }
            else { _rows.Add((id, category, RowState.Fail, failDetail)); Fail++; }
        }

        public void Report(ITestOutputHelper o)
        {
            o.WriteLine("===== IR MARKUP PARITY SCOREBOARD (MARKUP_BLOCKED cases) =====");
            o.WriteLine($"Total: {Total}   PASS: {Pass}   DEVIATION: {Deviation}   FAIL: {Fail}   " +
                        $"({100.0 * (Pass + Deviation) / Math.Max(1, Total):F1}% pass-or-deviation)");
            foreach (var g in _rows.GroupBy(r => r.Category).OrderBy(g => g.Key))
                o.WriteLine($"  [{g.Key,-3}] {g.Count(r => r.State == RowState.Pass)} pass + " +
                            $"{g.Count(r => r.State == RowState.Deviation)} deviation / {g.Count()}");
            o.WriteLine("FAILING (must be empty for the floor):");
            foreach (var r in _rows.Where(r => r.State == RowState.Fail))
                o.WriteLine($"  FAIL  {r.Id,-48} {r.Detail}");
            o.WriteLine("PASSING:");
            foreach (var r in _rows.Where(r => r.State == RowState.Pass))
                o.WriteLine($"  PASS  {r.Id}");
        }
    }
}
