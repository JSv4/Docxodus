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
}
