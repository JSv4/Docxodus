#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tests for the editor features built to draft an SEC Form S-1 cover page through
/// <see cref="DocxSession"/>: run font size, paragraph borders / horizontal rules,
/// table insertion, and the blank-document factory. Test IDs use the DS2xx range.
/// </summary>
public class DocxSessionS1FeaturesTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string FirstBodyParagraph(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;

    private static XElement DocumentXml(byte[] docxBytes)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.GetXDocument().Root!;
    }

    // ─── F1: font size ──────────────────────────────────────────────────

    [Fact]
    public void DS201_FontSize_SetsSzAndSzCsInHalfPoints()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var r = session.ApplyFormat(anchor, null, new FormatOp { FontSizePts = 20 });
        Assert.True(r.Success, r.Error?.Message);

        var root = DocumentXml(session.Save());
        var run = root.Descendants(W + "r").First(x => x.Value.Length > 0);
        Assert.Equal("40", (string?)run.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val"));
        Assert.Equal("40", (string?)run.Element(W + "rPr")?.Element(W + "szCs")?.Attribute(W + "val"));
    }

    [Fact]
    public void DS201b_FontSize_FractionalRoundsToHalfPoint()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        session.ApplyFormat(anchor, null, new FormatOp { FontSizePts = 7.5 });

        var root = DocumentXml(session.Save());
        var run = root.Descendants(W + "r").First(x => x.Value.Length > 0);
        Assert.Equal("15", (string?)run.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val"));
    }

    [Fact]
    public void DS202_FontSize_ZeroClearsExplicitSize()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        session.ApplyFormat(anchor, null, new FormatOp { FontSizePts = 18 });
        session.ApplyFormat(anchor, null, new FormatOp { FontSizePts = 0 });

        var root = DocumentXml(session.Save());
        Assert.Empty(root.Descendants(W + "sz"));
    }

    // ─── F1b: font family ───────────────────────────────────────────────

    [Fact]
    public void DS220_FontFamily_SetsRFontsAsciiHAnsiCs()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var r = session.ApplyFormat(anchor, null, new FormatOp { FontFamily = "Times New Roman" });
        Assert.True(r.Success, r.Error?.Message);

        var root = DocumentXml(session.Save());
        var run = root.Descendants(W + "r").First(x => x.Value.Length > 0);
        var rFonts = run.Element(W + "rPr")?.Element(W + "rFonts");
        Assert.NotNull(rFonts);
        Assert.Equal("Times New Roman", (string?)rFonts!.Attribute(W + "ascii"));
        Assert.Equal("Times New Roman", (string?)rFonts.Attribute(W + "hAnsi"));
        Assert.Equal("Times New Roman", (string?)rFonts.Attribute(W + "cs"));
    }

    [Fact]
    public void DS221_FontFamily_InsertedInSchemaOrder_AndValidates()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        // Apply size + bold FIRST so a naive append would put w:rFonts after w:sz (out of order).
        session.ApplyFormat(anchor, null, new FormatOp { Bold = true, FontSizePts = 14 });
        session.ApplyFormat(anchor, null, new FormatOp { FontFamily = "Georgia" });

        var bytes = session.Save();
        var root = DocumentXml(bytes);
        var rPr = root.Descendants(W + "r").First(x => x.Value.Length > 0).Element(W + "rPr")!;
        // w:rFonts is the first EG_RPrBase child after an optional w:rStyle (none here),
        // so it must be the first rPr child — before w:b / w:sz.
        Assert.Equal(W + "rFonts", rPr.Elements().First().Name);

        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(doc).ToList();
        Assert.Empty(errors);
    }

    [Fact]
    public void DS222_FontFamily_EmptyStringClearsRFonts()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        session.ApplyFormat(anchor, null, new FormatOp { FontFamily = "Arial" });
        session.ApplyFormat(anchor, null, new FormatOp { FontFamily = "" });

        var root = DocumentXml(session.Save());
        var run = root.Descendants(W + "r").First(x => x.Value.Length > 0);
        Assert.Null(run.Element(W + "rPr")?.Element(W + "rFonts"));
    }

    // ─── F1c: Enter inherits run formatting ─────────────────────────────

    [Fact]
    public void DS230_SplitAtEnd_NewParagraphInheritsRunFormatting()
    {
        // Drafting a uniformly-formatted filing: format a whole paragraph bold + Times + 16pt
        // (direct run formatting, as the editor's ribbon applies it), press Enter at the end, and
        // keep typing. The new line must continue in the SAME formatting (matches Word) — not reset
        // to the document default.
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        session.ApplyFormat(anchor, null, new FormatOp { Bold = true, FontFamily = "Times New Roman", FontSizePts = 16 });

        var split = session.SplitParagraph(anchor, "First paragraph.".Length); // Enter at end
        Assert.True(split.Success, split.Error?.Message);
        var newAnchor = split.Created!.Single().Id;

        var typed = session.ReplaceText(newAnchor, "continued");
        Assert.True(typed.Success, typed.Error?.Message);

        var root = DocumentXml(session.Save());
        var run = root.Descendants(W + "r").First(r => r.Value == "continued");
        var rPr = run.Element(W + "rPr");
        Assert.NotNull(rPr);
        Assert.NotNull(rPr!.Element(W + "b"));                                                  // bold carried
        Assert.Equal("Times New Roman", (string?)rPr.Element(W + "rFonts")?.Attribute(W + "ascii")); // font carried
        Assert.Equal("32", (string?)rPr.Element(W + "sz")?.Attribute(W + "val"));               // 16pt carried
    }

    [Fact]
    public void DS231_SplitInheritance_ProducesValidOoxml()
    {
        // The carried paragraph-mark rPr must be inserted in schema order so the document validates.
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        session.ApplyFormat(anchor, null, new FormatOp { Bold = true, FontFamily = "Georgia", FontSizePts = 18 });
        var split = session.SplitParagraph(anchor, "First paragraph.".Length);
        session.ReplaceText(split.Created!.Single().Id, "more");

        var bytes = session.Save();
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(doc).ToList();
        Assert.Empty(errors);
    }

    // ─── F1d: right tab stop ────────────────────────────────────────────

    [Fact]
    public void DS240_InsertTab_Right_AddsRightTabStopAndTabRun()
    {
        // The "As filed… / Registration No." filing row: one paragraph, left text + a right-aligned
        // tab stop at the margin + a tab + (later) right text — no two-column table needed.
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var r = session.InsertTab(anchor, "First paragraph.".Length, TabStopAlignment.Right);
        Assert.True(r.Success, r.Error?.Message);

        var root = DocumentXml(session.Save());
        var para = root.Descendants(W + "p").First(p => p.Value.Contains("First paragraph."));

        // A right tab STOP on the paragraph (w:pPr/w:tabs/w:tab with val=right, a positive pos).
        var stop = para.Element(W + "pPr")?.Element(W + "tabs")?.Element(W + "tab");
        Assert.NotNull(stop);
        Assert.Equal("right", (string?)stop!.Attribute(W + "val"));
        Assert.True(int.Parse((string)stop.Attribute(W + "pos")!) > 0);

        // A tab RUN in the content (a w:tab whose parent is a w:r), after the text.
        var tabRun = para.Descendants(W + "tab").Where(t => t.Parent!.Name == W + "r").ToList();
        Assert.Single(tabRun);
    }

    [Fact]
    public void DS241_InsertTab_ProducesValidOoxml_AndRoundTrips()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        session.InsertTab(anchor, "First paragraph.".Length, TabStopAlignment.Right);

        var bytes = session.Save();
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(doc).ToList();
        Assert.Empty(errors);

        // Round-trips: reopen sees the tab stop survive.
        using var session2 = new DocxSession(bytes);
        var root = DocumentXml(session2.Save());
        Assert.Contains(root.Descendants(W + "tab"),
            t => t.Parent!.Name == W + "tabs" && (string?)t.Attribute(W + "val") == "right");
    }

    [Fact]
    public void DS242_InsertTab_Twice_DoesNotDuplicateTheStop()
    {
        // Idempotent stop: applying a right tab on the same paragraph again must not stack a second
        // identical stop (a tab RUN is added each time — that's content — but the STOP is de-duped).
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        session.InsertTab(anchor, "First paragraph.".Length, TabStopAlignment.Right);
        session.InsertTab(anchor, "First paragraph.".Length, TabStopAlignment.Right);

        var root = DocumentXml(session.Save());
        var para = root.Descendants(W + "p").First(p => p.Value.Contains("First paragraph."));
        var stops = para.Element(W + "pPr")!.Element(W + "tabs")!.Elements(W + "tab")
            .Where(t => (string?)t.Attribute(W + "val") == "right").ToList();
        Assert.Single(stops);
    }

    // ─── F2: paragraph borders ──────────────────────────────────────────

    [Fact]
    public void DS203_BottomBorder_EmitsPBdrBottom()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var r = session.SetParagraphFormat(anchor, new ParagraphFormatOp
        {
            BottomBorder = new ParagraphBorderEdge { Style = "single", Size = 18, Color = "000000" },
        });
        Assert.True(r.Success, r.Error?.Message);

        var root = DocumentXml(session.Save());
        var bottom = root.Descendants(W + "pBdr").Elements(W + "bottom").Single();
        Assert.Equal("single", (string?)bottom.Attribute(W + "val"));
        Assert.Equal("18", (string?)bottom.Attribute(W + "sz"));
        Assert.Equal("000000", (string?)bottom.Attribute(W + "color"));
    }

    [Fact]
    public void DS204_InsertHorizontalRule_AddsEmptyBorderedParagraph()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        int before = DocumentXml(session.Save()).Descendants(W + "p").Count();

        var r = session.InsertHorizontalRule(anchor, Position.After);
        Assert.True(r.Success, r.Error?.Message);
        Assert.NotEmpty(r.Created);

        var root = DocumentXml(session.Save());
        Assert.Equal(before + 1, root.Descendants(W + "p").Count());
        // The new paragraph has a bottom border and no run text.
        var rule = root.Descendants(W + "p").Single(p => p.Element(W + "pPr")?.Element(W + "pBdr") is not null);
        Assert.NotNull(rule.Element(W + "pPr")!.Element(W + "pBdr")!.Element(W + "bottom"));
        Assert.Equal(string.Empty, rule.Value);
    }

    // ─── F3: table insertion ────────────────────────────────────────────

    [Fact]
    public void DS205_InsertTable_BuildsGridWithSeededContentAndAlignment()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var r = session.InsertTable(anchor, Position.After, rows: 2, cols: 3, new TableInsertOptions
        {
            Borderless = true,
            CellAlignment = ParagraphAlignment.Center,
            CellContents = new[] { "Texas", "7370", "01-0627671", "(State)", "(SIC)", "(IRS)" },
        });
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal(6, r.Created.Count); // one cell-paragraph anchor per cell

        var root = DocumentXml(session.Save());
        var tbl = root.Descendants(W + "tbl").Single();
        Assert.Equal(2, tbl.Elements(W + "tr").Count());
        Assert.All(tbl.Elements(W + "tr"), tr => Assert.Equal(3, tr.Elements(W + "tc").Count()));
        Assert.Equal(3, tbl.Element(W + "tblGrid")!.Elements(W + "gridCol").Count());

        // Seeded content + centered cell paragraphs.
        Assert.Contains("Texas", tbl.Value);
        Assert.Contains("01-0627671", tbl.Value);
        var firstCellP = tbl.Element(W + "tr")!.Element(W + "tc")!.Element(W + "p")!;
        Assert.Equal("center", (string?)firstCellP.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val"));

        // Borderless => explicit "none" table borders.
        var borders = tbl.Element(W + "tblPr")!.Element(W + "tblBorders")!;
        Assert.All(borders.Elements(), e => Assert.Equal("none", (string?)e.Attribute(W + "val")));
    }

    [Fact]
    public void DS214_InsertTable_ColumnWidths_LandInGridAndCells()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        // A wide-left / narrow-right 2-column split (the S-1 filing-header row).
        var r = session.InsertTable(anchor, Position.After, 1, 2, new TableInsertOptions
        {
            Borderless = true,
            ColumnWidths = new[] { 7000, 2576 },
        });
        Assert.True(r.Success, r.Error?.Message);

        var tbl = DocumentXml(session.Save()).Descendants(W + "tbl").Single();

        // w:tblGrid carries the explicit column widths in order.
        var gridCols = tbl.Element(W + "tblGrid")!.Elements(W + "gridCol").ToList();
        Assert.Equal(2, gridCols.Count);
        Assert.Equal("7000", (string?)gridCols[0].Attribute(W + "w"));
        Assert.Equal("2576", (string?)gridCols[1].Attribute(W + "w"));

        // Each cell's w:tcW matches its column.
        var cells = tbl.Element(W + "tr")!.Elements(W + "tc").ToList();
        Assert.Equal("7000", (string?)cells[0].Element(W + "tcPr")?.Element(W + "tcW")?.Attribute(W + "w"));
        Assert.Equal("2576", (string?)cells[1].Element(W + "tcPr")?.Element(W + "tcW")?.Attribute(W + "w"));
    }

    [Fact]
    public void DS215_InsertTable_ColumnWidths_WrongCount_IsRejected()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        // 3 widths for a 2-column table is a caller error — fail loudly, don't silently equalize.
        var r = session.InsertTable(anchor, Position.After, 1, 2, new TableInsertOptions
        {
            ColumnWidths = new[] { 1000, 2000, 3000 },
        });
        Assert.False(r.Success);
    }

    [Fact]
    public void DS206_InsertTable_BorderedByDefault()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        var r = session.InsertTable(anchor, Position.After, 1, 2);
        Assert.True(r.Success, r.Error?.Message);

        var tbl = DocumentXml(session.Save()).Descendants(W + "tbl").Single();
        var borders = tbl.Element(W + "tblPr")!.Element(W + "tblBorders")!;
        Assert.All(borders.Elements(), e => Assert.Equal("single", (string?)e.Attribute(W + "val")));
    }

    [Fact]
    public void DS207_InsertTable_CreatedCellsAreFillableByAnchor()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        var r = session.InsertTable(anchor, Position.After, 1, 2, new TableInsertOptions
        {
            CellContents = new[] { "left", "right" },
        });
        Assert.True(r.Success, r.Error?.Message);

        // Each created cell anchor is addressable for a subsequent edit.
        var cellAnchor = r.Created.First();
        var fill = session.ReplaceText(cellAnchor.Id, "FILLED");
        Assert.True(fill.Success, fill.Error?.Message);
        Assert.Contains("FILLED", DocumentXml(session.Save()).Descendants(W + "tbl").Single().Value);
    }

    // ─── F0: blank document ─────────────────────────────────────────────

    [Fact]
    public void DS208_CreateBlankDocx_OpensAsEditableSession()
    {
        var bytes = DocxSession.CreateBlankDocxBytes();
        using var session = new DocxSession(bytes);
        var proj = session.Project();
        Assert.True(proj.AnchorIndex.Count >= 1);

        // A blank doc must support the basic drafting ops without error: typing into the
        // single (empty) body paragraph is the entry point for drafting from scratch.
        var anchor = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;
        var typed = session.ReplaceText(anchor, "Hello S-1");
        Assert.True(typed.Success, typed.Error?.Message);
        // Verify against saved OOXML text (the markdown projection escapes the hyphen).
        Assert.Contains("Hello S-1", DocumentXml(session.Save()).Descendants(W + "body").Single().Value);
    }

    // ─── End-to-end: build an S-1-style page and validate the OOXML schema ──

    [Fact]
    public void DS210_DraftS1CoverPage_ProducesSchemaValidOoxml()
    {
        // Mirror the browser smoke test: a blank doc, then every new feature exercised —
        // font size, paragraph border / horizontal rule, and borderless + bordered tables.
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var sentinel = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p").Anchor.Id;

        // 2-col header table (borderless), justified left/right + 8pt.
        var hdr = session.InsertTable(sentinel, Position.Before, 1, 2, new TableInsertOptions
        {
            Borderless = true,
            CellContents = new[] { "As filed on May 20, 2026", "Registration No. 333-" },
        });
        session.SetParagraphFormat(hdr.Created[0].Id, new ParagraphFormatOp { Alignment = ParagraphAlignment.Left });
        session.SetParagraphFormat(hdr.Created[1].Id, new ParagraphFormatOp { Alignment = ParagraphAlignment.Right });
        session.ApplyFormat(hdr.Created[0].Id, null, new FormatOp { FontSizePts = 8 });

        // Heavy rule.
        session.InsertHorizontalRule(sentinel, Position.Before, new ParagraphBorderEdge { Style = "single", Size = 24, Color = "000000" });

        // Big bold centered title.
        var title = session.InsertParagraph(sentinel, Position.Before, "FORM S-1").Created[0].Id;
        session.ApplyFormat(title, null, new FormatOp { Bold = true, FontSizePts = 22 });
        session.SetParagraphFormat(title, new ParagraphFormatOp { Alignment = ParagraphAlignment.Center });

        session.InsertHorizontalRule(sentinel, Position.Before);

        // Bordered 2×3 value/label table.
        session.InsertTable(sentinel, Position.Before, 2, 3, new TableInsertOptions
        {
            CellAlignment = ParagraphAlignment.Center,
            CellContents = new[] { "Texas", "7370", "01-0627671", "(State)", "(SIC)", "(IRS)" },
        });

        var bytes = session.Save();

        using var ms = new MemoryStream(bytes);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
        var errors = validator.Validate(wDoc)
            .Select(e => $"{e.Path?.XPath}: {e.Description}")
            .ToList();
        Assert.True(errors.Count == 0, "OOXML schema errors:\n" + string.Join("\n", errors));
    }

    [Fact]
    public void DS209_CreateBlankDocx_HasNormalStyleAndOpensInWord()
    {
        var bytes = DocxSession.CreateBlankDocxBytes();
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);

        // A default "Normal" paragraph style + a section + a body paragraph exist.
        var stylesRoot = doc.MainDocumentPart!.StyleDefinitionsPart!.GetXDocument().Root!;
        Assert.Contains(stylesRoot.Elements(W + "style"),
            s => (string?)s.Attribute(W + "styleId") == "Normal");
        var body = doc.MainDocumentPart.GetXDocument().Root!.Element(W + "body")!;
        Assert.NotNull(body.Element(W + "sectPr"));
        Assert.NotEmpty(body.Elements(W + "p"));
    }

    // ─── F-fix: line-break fidelity (Shift+Enter → w:br, not a raw newline) ──

    [Fact]
    public void DS211_HardLineBreak_BecomesWbr_NotLiteralNewline()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        // GFM hard break: two trailing spaces + newline within ONE paragraph.
        var r = session.ReplaceText(anchor, "Line one  \nLine two");
        Assert.True(r.Success, r.Error?.Message);

        var root = DocumentXml(session.Save());

        // It stays ONE paragraph (a line break, not a paragraph split)...
        var matching = root.Descendants(W + "p")
            .Where(p => p.Value.Contains("Line one") || p.Value.Contains("Line two"))
            .ToList();
        Assert.Single(matching);
        var para = matching[0];

        // ...containing a real w:br...
        Assert.NotEmpty(para.Descendants(W + "br"));

        // ...and NO w:t carries a literal newline (the Word-infidelity we are fixing).
        Assert.DoesNotContain(para.Descendants(W + "t"), t => t.Value.Contains('\n'));
    }

    [Fact]
    public void DS212_HardLineBreak_RoundTripsThroughProjection()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        session.ReplaceText(anchor, "Line one  \nLine two");

        // Re-open the saved bytes and project: the hard break survives as the
        // canonical GFM "  \n" (symmetric with WmlToMarkdownConverter's w:br output).
        using var session2 = new DocxSession(session.Save());
        Assert.Contains("Line one  \nLine two", session2.Project().Markdown);
    }

    [Fact]
    public void DS213_BlankLineSeparatorStillSplitsParagraphs_NotABreak()
    {
        // Guard: a BLANK line (paragraph separator) must NOT become a w:br — only an
        // intra-paragraph single newline does. InsertParagraph accepts multi-block md.
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var r = session.InsertParagraph(anchor, Position.After, "Alpha\n\nBeta");
        Assert.True(r.Success, r.Error?.Message);

        var root = DocumentXml(session.Save());
        Assert.Contains(root.Descendants(W + "p"), p => p.Value == "Alpha");
        Assert.Contains(root.Descendants(W + "p"), p => p.Value == "Beta");
        Assert.DoesNotContain(root.Descendants(W + "p"),
            p => (p.Value.Contains("Alpha") || p.Value.Contains("Beta")) && p.Descendants(W + "br").Any());
    }

    // ─── Finding 1a: HR border must not propagate on Enter-split of an empty rule ───────

    [Fact]
    public void DS216_SplitEmptyBorderedParagraph_DoesNotInheritBorder()
    {
        // Reproduces the S-1 smoke-test footgun: pressing Enter inside an HR (an empty
        // bottom-bordered paragraph) must NOT give the new paragraph a border.
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var anchor = FirstBodyParagraph(session);

        var hr = session.InsertHorizontalRule(anchor, Position.After);
        Assert.True(hr.Success, hr.Error?.Message);
        var hrAnchor = hr.Created[0].Id;

        var split = session.SplitParagraph(hrAnchor, 0);
        Assert.True(split.Success, split.Error?.Message);

        var root = DocumentXml(session.Save());
        // Exactly one paragraph carries a border (the original rule), not two — the new
        // empty paragraph below the rule is borderless so the user can type body text there.
        Assert.Equal(1, root.Descendants(W + "p").Count(p => p.Element(W + "pPr")?.Element(W + "pBdr") is not null));
    }

    [Fact]
    public void DS217_SplitBorderedParagraphWithText_KeepsBorderOnNewParagraph()
    {
        // Guard: only the EMPTY-source (rule) case drops the border. A bordered paragraph
        // that has text (a boxed/underlined block) still splits with the border on both halves
        // — this protects against an over-broad "always strip on split" implementation.
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var anchor = FirstBodyParagraph(session);

        var ins = session.InsertParagraph(anchor, Position.After, "Boxed heading text");
        var boxAnchor = ins.Created[0].Id;
        var fmt = session.SetParagraphFormat(boxAnchor, new ParagraphFormatOp
        {
            BottomBorder = new ParagraphBorderEdge { Style = "single", Size = 12, Color = "auto" },
        });
        Assert.True(fmt.Success, fmt.Error?.Message);

        var split = session.SplitParagraph(boxAnchor, 5);
        Assert.True(split.Success, split.Error?.Message);

        var root = DocumentXml(session.Save());
        // Both halves of the split text paragraph keep the bottom border.
        Assert.Equal(2, root.Descendants(W + "p").Count(p => p.Element(W + "pPr")?.Element(W + "pBdr") is not null));
    }

    // ─── Finding 2: a table at body end must be followed by a paragraph ─────────────────

    [Fact]
    public void DS218_InsertTableAtBodyEnd_AppendsTrailingParagraph()
    {
        // Reproduces "cannot append after a trailing table": inserting after the last block
        // left the table as the final body element (</w:tbl></w:sectPr>). Word's convention
        // (and an editable surface) needs a paragraph after the table.
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var anchor = FirstBodyParagraph(session);

        var r = session.InsertTable(anchor, Position.After, 1, 2);
        Assert.True(r.Success, r.Error?.Message);

        var body = DocumentXml(session.Save()).Element(W + "body")!;
        var tbl = body.Elements(W + "tbl").Single();
        var next = tbl.ElementsAfterSelf().FirstOrDefault();
        Assert.NotNull(next);                       // not the last element anymore
        Assert.Equal(W + "p", next!.Name);          // and what follows is a paragraph
        Assert.Equal(W + "sectPr", body.Elements().Last().Name); // sectPr still closes the body
    }

    [Fact]
    public void DS219_InsertTableFollowedByParagraph_DoesNotAddExtraTrailing()
    {
        // Guard against double-insert: when a paragraph already follows the table, no extra one.
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var anchor = FirstBodyParagraph(session);

        var r = session.InsertTable(anchor, Position.Before, 1, 2);
        Assert.True(r.Success, r.Error?.Message);

        var body = DocumentXml(session.Save()).Element(W + "body")!;
        // [tbl, originalParagraph, sectPr] — exactly one direct body paragraph, not two.
        Assert.Equal(1, body.Elements(W + "p").Count());
    }

    // ─── F5: table cells inherit the document font ──────────────────────

    [Fact]
    public void DS223_InsertTable_CellFontFamily_StampsSeededRunFonts()
    {
        // A table inserted into a Times document should have Times cells, not the
        // blank-doc docDefaults (Calibri). Seeded content runs carry the font directly.
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var r = session.InsertTable(anchor, Position.After, 1, 2, new TableInsertOptions
        {
            Borderless = true,
            CellFontFamily = "Times New Roman",
            CellContents = new[] { "Texas", "7370" },
        });
        Assert.True(r.Success, r.Error?.Message);

        var tbl = DocumentXml(session.Save()).Descendants(W + "tbl").Single();
        foreach (var run in tbl.Descendants(W + "r").Where(x => x.Value.Length > 0))
        {
            var ascii = (string?)run.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "ascii");
            Assert.Equal("Times New Roman", ascii);
        }
    }

    [Fact]
    public void DS224_InsertTable_CellFontFamily_StampsEmptyCellMarkFont()
    {
        // Empty cells (the editor's grid-picker flow) carry the font on the paragraph-mark
        // run properties so later typing inherits it.
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var r = session.InsertTable(anchor, Position.After, 2, 2, new TableInsertOptions
        {
            Borderless = true,
            CellFontFamily = "Times New Roman",
        });
        Assert.True(r.Success, r.Error?.Message);

        var tbl = DocumentXml(session.Save()).Descendants(W + "tbl").Single();
        foreach (var cellP in tbl.Descendants(W + "tc").Select(tc => tc.Element(W + "p")!))
        {
            var markAscii = (string?)cellP.Element(W + "pPr")?.Element(W + "rPr")?
                .Element(W + "rFonts")?.Attribute(W + "ascii");
            Assert.Equal("Times New Roman", markAscii);
        }
    }

    [Fact]
    public void DS225_TypingIntoEmptyFontCell_InheritsTheMarkFont()
    {
        // The critical typed-later path: the editor commits text into an empty cell via
        // ReplaceText, which rebuilds runs. The new run must inherit the cell's mark font.
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var r = session.InsertTable(anchor, Position.After, 1, 2, new TableInsertOptions
        {
            Borderless = true,
            CellFontFamily = "Times New Roman",
        });
        Assert.True(r.Success, r.Error?.Message);
        var cellAnchor = r.Created[0].Id;

        var typed = session.ReplaceText(cellAnchor, "Texas");
        Assert.True(typed.Success, typed.Error?.Message);

        var tbl = DocumentXml(session.Save()).Descendants(W + "tbl").Single();
        var run = tbl.Descendants(W + "r").First(x => x.Value == "Texas");
        var ascii = (string?)run.Element(W + "rPr")?.Element(W + "rFonts")?.Attribute(W + "ascii");
        Assert.Equal("Times New Roman", ascii);
    }

    [Fact]
    public void DS226_InsertTable_CellFontFamily_Validates()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var r = session.InsertTable(anchor, Position.After, 2, 2, new TableInsertOptions
        {
            Borderless = true,
            CellFontFamily = "Times New Roman",
            CellContents = new[] { "Texas", "7370" },
        });
        Assert.True(r.Success, r.Error?.Message);

        var bytes = session.Save();
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(doc).ToList();
        Assert.Empty(errors);
    }

    [Fact]
    public void DS227_InsertTable_NoCellFontFamily_LeavesCellsFontless()
    {
        // Default (no CellFontFamily) preserves prior behavior — cells inherit docDefaults.
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var r = session.InsertTable(anchor, Position.After, 1, 2, new TableInsertOptions
        {
            Borderless = true,
            CellContents = new[] { "Texas", "7370" },
        });
        Assert.True(r.Success, r.Error?.Message);

        var tbl = DocumentXml(session.Save()).Descendants(W + "tbl").Single();
        var run = tbl.Descendants(W + "r").First(x => x.Value == "Texas");
        Assert.Null(run.Element(W + "rPr")?.Element(W + "rFonts"));
    }

    // ─── F-indent: hanging / first-line indent ──────────────────────────

    private static XElement? FirstInd(byte[] docxBytes) =>
        DocumentXml(docxBytes).Descendants(W + "p").First().Element(W + "pPr")?.Element(W + "ind");

    [Fact]
    public void DS243_SetParagraphFormat_LeftIndent_SetsWIndLeft()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var r = session.SetParagraphFormat(anchor, new ParagraphFormatOp { LeftIndent = 1440 });
        Assert.True(r.Success, r.Error?.Message);

        var ind = FirstInd(session.Save());
        Assert.Equal("1440", (string?)ind?.Attribute(W + "left"));
    }

    [Fact]
    public void DS244_SetParagraphFormat_FirstLineIndentPositive_SetsFirstLine_RemovesHanging()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        session.SetParagraphFormat(anchor, new ParagraphFormatOp { LeftIndent = 720, FirstLineIndent = -360 });
        var r = session.SetParagraphFormat(anchor, new ParagraphFormatOp { FirstLineIndent = 360 });
        Assert.True(r.Success, r.Error?.Message);

        var ind = FirstInd(session.Save());
        Assert.Equal("360", (string?)ind?.Attribute(W + "firstLine"));
        Assert.Null((string?)ind?.Attribute(W + "hanging"));
    }

    [Fact]
    public void DS245_SetParagraphFormat_FirstLineIndentNegative_SetsHanging_RemovesFirstLine()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        session.SetParagraphFormat(anchor, new ParagraphFormatOp { FirstLineIndent = 200 });
        var r = session.SetParagraphFormat(anchor, new ParagraphFormatOp { LeftIndent = 720, FirstLineIndent = -360 });
        Assert.True(r.Success, r.Error?.Message);

        var ind = FirstInd(session.Save());
        Assert.Equal("360", (string?)ind?.Attribute(W + "hanging"));
        Assert.Equal("720", (string?)ind?.Attribute(W + "left"));
        Assert.Null((string?)ind?.Attribute(W + "firstLine"));
    }

    [Fact]
    public void DS246_SetParagraphFormat_FirstLineIndentZero_ClearsBoth_KeepsLeft()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        session.SetParagraphFormat(anchor, new ParagraphFormatOp { LeftIndent = 720, FirstLineIndent = -360 });
        var r = session.SetParagraphFormat(anchor, new ParagraphFormatOp { FirstLineIndent = 0 });
        Assert.True(r.Success, r.Error?.Message);

        var ind = FirstInd(session.Save());
        Assert.Null((string?)ind?.Attribute(W + "hanging"));
        Assert.Null((string?)ind?.Attribute(W + "firstLine"));
        Assert.Equal("720", (string?)ind?.Attribute(W + "left"));
    }

    [Fact]
    public void DS247_ParseParagraphFormatOp_ReadsLeftAndFirstLineIndent()
    {
        var op = Docxodus.Internal.DocxSessionJson.ParseParagraphFormatOp(
            "{\"leftIndent\":720,\"firstLineIndent\":-360}");
        Assert.Equal(720, op.LeftIndent);
        Assert.Equal(-360, op.FirstLineIndent);
    }

    // ─── F-numbering: configurable multi-level numbering ─────────────────

    [Fact]
    public void DS248_EnsureMultilevel_BuildsAbstractNum_AndDedupsSameScheme()
    {
        var bytes = DocxSession.CreateBlankDocxBytes();
        using var stream = new MemoryStream(); // expandable — adding a numbering part grows the package
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        using var doc = WordprocessingDocument.Open(stream, true);
        var levels = new System.Collections.Generic.List<NumberingLevel>
        {
            new() { Format = NumberFormat.Decimal, LevelText = "%1." },
            new() { Format = NumberFormat.LowerLetter, LevelText = "(%2)" },
        };
        int n1 = Docxodus.Internal.NumberingFactory.EnsureMultilevel(doc, levels, restart: false);
        int n2 = Docxodus.Internal.NumberingFactory.EnsureMultilevel(doc, levels, restart: false);
        Assert.Equal(n1, n2); // same scheme → same numId (one continuous sequence)

        int n3 = Docxodus.Internal.NumberingFactory.EnsureMultilevel(doc, levels, restart: true);
        Assert.NotEqual(n1, n3); // restart → fresh numId

        var root = doc.MainDocumentPart!.NumberingDefinitionsPart!.GetXDocument().Root!;
        var abs = root.Elements(W + "abstractNum").Single(a =>
            a.Elements(W + "lvl").Any(l => (string?)l.Element(W + "lvlText")?.Attribute(W + "val") == "(%2)"));
        Assert.Equal("lowerLetter",
            (string?)abs.Elements(W + "lvl").First(l => (string?)l.Attribute(W + "ilvl") == "1")
                .Element(W + "numFmt")?.Attribute(W + "val"));
    }

    [Fact]
    public void DS249_ApplyMultilevelNumbering_SetsNumPr()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        var levels = new System.Collections.Generic.List<NumberingLevel>
        {
            new() { Format = NumberFormat.Decimal, LevelText = "%1." },
            new() { Format = NumberFormat.Decimal, LevelText = "%1.%2" },
            new() { Format = NumberFormat.LowerLetter, LevelText = "(%3)" },
        };
        var res = session.ApplyMultilevelNumbering(anchor, levels, level: 0);
        Assert.True(res.Success, res.Error?.Message);

        var numPr = DocumentXml(session.Save()).Descendants(W + "p").First()
            .Element(W + "pPr")?.Element(W + "numPr");
        Assert.NotNull(numPr);
        Assert.Equal("0", (string?)numPr!.Element(W + "ilvl")?.Attribute(W + "val"));
    }

    [Fact]
    public void DS250_ApplyMultilevelNumbering_ThenSetListLevel_Promotes()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        var levels = new System.Collections.Generic.List<NumberingLevel>
        {
            new() { Format = NumberFormat.Decimal, LevelText = "%1." },
            new() { Format = NumberFormat.LowerLetter, LevelText = "(%2)" },
        };
        var applied = session.ApplyMultilevelNumbering(anchor, levels, level: 0);
        Assert.True(applied.Success, applied.Error?.Message);

        var bumped = session.SetListLevel(applied.Modified[0].Id, 1);
        Assert.True(bumped.Success, bumped.Error?.Message);

        var ilvl = DocumentXml(session.Save()).Descendants(W + "p").First()
            .Element(W + "pPr")?.Element(W + "numPr")?.Element(W + "ilvl")?.Attribute(W + "val");
        Assert.Equal("1", (string?)ilvl);
    }

    [Fact]
    public void DS251_ApplyMultilevelNumbering_RoundTrips()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        var levels = new System.Collections.Generic.List<NumberingLevel>
        {
            new() { Format = NumberFormat.LowerLetter, LevelText = "(%1)" },
        };
        session.ApplyMultilevelNumbering(anchor, levels, 0);

        using var reopened = new DocxSession(session.Save());
        var numPr = DocumentXml(reopened.Save()).Descendants(W + "p").First()
            .Element(W + "pPr")?.Element(W + "numPr");
        Assert.NotNull(numPr);
    }
}
