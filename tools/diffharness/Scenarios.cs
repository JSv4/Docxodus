#nullable enable
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DiffHarness;

/// <summary>
/// The edit × feature scenario catalog. Each scenario applies ONE isolated, content-anchored edit to a
/// fresh copy of the base contract to produce its <c>right.docx</c>, so the resulting diff exercises
/// exactly one behaviour. Anchors are matched by text content (not absolute index) so the catalog is
/// resilient to document revisions. A missing anchor throws loudly — a scenario must never silently
/// produce a no-op variant.
/// </summary>
internal static class Scenarios
{
    private sealed record Scenario(
        string Id, string Feature, string EditType, string Desc, Action<WordprocessingDocument> Mutate,
        // True for edits confined to a scope DocxDiff deliberately does not diff (header/footer), matching
        // the WmlComparer oracle. For these, accept!=right on that scope is EXPECTED, not a defect — the
        // round-trip pass criterion drops the accept-side header/footer gate.
        bool UndiffedScopeOnly = false);

    public static int Generate(string basePath, string outRoot)
    {
        var baseBytes = File.ReadAllBytes(basePath);
        Directory.CreateDirectory(outRoot);

        var scenarios = Catalog();
        var manifest = new List<object>();
        foreach (var s in scenarios)
        {
            var dir = Path.Combine(outRoot, s.Id);
            Directory.CreateDirectory(dir);
            File.WriteAllBytes(Path.Combine(dir, "left.docx"), baseBytes);
            var right = Apply(baseBytes, s.Mutate);
            File.WriteAllBytes(Path.Combine(dir, "right.docx"), right);
            manifest.Add(new { id = s.Id, feature = s.Feature, editType = s.EditType, desc = s.Desc,
                                undiffedScopeOnly = s.UndiffedScopeOnly });
            Console.WriteLine($"  + {s.Id}");
        }
        File.WriteAllText(Path.Combine(outRoot, "manifest.json"),
            JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }));
        return scenarios.Count;
    }

    private static byte[] Apply(byte[] baseBytes, Action<WordprocessingDocument> mutate)
    {
        using var ms = new MemoryStream();
        ms.Write(baseBytes, 0, baseBytes.Length);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, true))  // AutoSave flushes on dispose
        {
            mutate(doc);
        }
        return ms.ToArray();
    }

    // ---- the catalog ---------------------------------------------------------------------------

    private static List<Scenario> Catalog() =>
    [
        // ===== body paragraphs: text edits =====
        new("body-replace-word", "body-para", "replace",
            "Replace a single word in a body paragraph (Purchaser -> Investor).",
            d => ReplaceInFirstText(Body(d), "Purchaser", "Investor")),

        new("body-insert-word", "body-para", "insert",
            "Insert a word into a sentence.",
            d => ReplaceInFirstText(Body(d), "Purchase and Sale of Preferred Stock",
                                            "Purchase and Sale of New Preferred Stock")),

        new("body-delete-word", "body-para", "delete",
            "Delete a word from a heading.",
            d => ReplaceInFirstText(Body(d), "Purchase and Sale of Preferred Stock",
                                            "Purchase and Sale of Stock")),

        new("body-replace-phrase", "body-para", "replace",
            "Replace a multi-word phrase within a paragraph.",
            d => ReplaceInFirstText(Body(d), "Preferred Stock Purchase Agreement",
                                            "Preferred Equity Subscription Agreement")),

        // ===== body paragraphs: structural =====
        new("body-insert-paragraph", "body-para", "insert-block",
            "Insert a brand-new paragraph after a heading.",
            d => InsertParagraphAfter(Body(d), "Purchase and Sale of Preferred Stock",
                                              "This is an inserted clause for diff testing purposes.")),

        new("body-delete-paragraph", "body-para", "delete-block",
            "Delete an entire body paragraph.",
            d => DeleteTopLevelPara(Body(d), "shall apply to each such closing unless otherwise specified")),

        new("body-move-paragraph", "body-para", "move",
            "Move a paragraph from one location to a distant one (cut & paste).",
            d => MoveTopLevelPara(Body(d),
                    moveText: "shall apply to each such closing unless otherwise specified",
                    afterAnchor: "Purchase and Sale of Preferred Stock")),

        new("body-split-paragraph", "body-para", "split",
            "Split one paragraph into two at a sentence boundary.",
            d => SplitTopLevelPara(Body(d), "THIS SERIES")),

        // ===== formatting (text unchanged) =====
        new("format-bold-run", "format", "format",
            "Bold a run without changing its text.",
            d => SetRunFormat(Body(d), "Purchase and Sale of Preferred Stock", r => r.Bold = new Bold())),

        new("format-italic-run", "format", "format",
            "Italicize a run without changing its text.",
            d => SetRunFormat(Body(d), "shall apply to each such closing", r => r.Italic = new Italic())),

        new("format-fontsize-run", "format", "format",
            "Change a run's font size.",
            d => SetRunFormat(Body(d), "Purchase and Sale of Preferred Stock",
                r => { r.FontSize = new FontSize { Val = "36" }; r.FontSizeComplexScript = new FontSizeComplexScript { Val = "36" }; })),

        new("format-color-run", "format", "format",
            "Change a run's color to red.",
            d => SetRunFormat(Body(d), "shall apply to each such closing",
                r => r.Color = new Color { Val = "FF0000" })),

        new("format-underline-run", "format", "format",
            "Underline a run (anchor chosen NOT already underlined).",
            d => SetRunFormat(Body(d), "shall apply to each such closing",
                r => r.Underline = new Underline { Val = UnderlineValues.Single })),

        // ===== styles =====
        new("style-change-paragraph", "style", "style",
            "Change a paragraph's style id.",
            d => SetParagraphStyle(Body(d), "shall apply to each such closing", "Heading2")),

        // ===== tables =====
        new("table-cell-edit", "table", "replace",
            "Edit text in a table cell.",
            d => ReplaceInFirstText(FirstTable(d), "Name and Address of Purchaser",
                                                   "Name and Address of Investor")),

        new("table-cell-insert-word", "table", "insert",
            "Insert a word in a table cell.",
            d => ReplaceInFirstText(FirstTable(d), "Total Shares", "Grand Total Shares")),

        new("table-insert-row", "table", "insert-block",
            "Append a cloned row to the first table.",
            d => InsertTableRow(FirstTable(d))),

        new("table-delete-row", "table", "delete-block",
            "Delete the last data row from the first table.",
            d => DeleteTableRow(FirstTable(d))),

        // ===== headers / footers =====
        new("header-edit", "header", "replace",
            "Edit text in a running header.",
            d => ReplaceInPart(HeaderContaining(d, "redline this against prior NVCA versions"),
                               "prior NVCA versions", "prior model versions"),
            UndiffedScopeOnly: true),

        new("footer-edit", "footer", "replace",
            "Edit text in a running footer.",
            d => ReplaceInPart(FooterContaining(d, "ACTIVE/"), "ACTIVE/", "DRAFT/"),
            UndiffedScopeOnly: true),

        // ===== footnotes =====
        new("footnote-edit", "footnote", "replace",
            "Edit text inside a footnote.",
            d => ReplaceInPart(FootnotesRoot(d), "mandatory tranches", "mandatory funding tranches")),

        // ===== round 2: tables (column ops, cell format) =====
        new("table-insert-column", "table", "insert-block",
            "Append a column (a w:tc per row + a w:gridCol) to the first table.",
            d => InsertTableColumn(FirstTable(d))),

        new("table-delete-column", "table", "delete-block",
            "Delete the last column from the first table.",
            d => DeleteTableColumn(FirstTable(d))),

        new("table-cell-format", "table", "format",
            "Bold a table cell's text without changing it.",
            d => SetRunFormat(FirstTable(d), "Total Shares", r => r.Bold = new Bold())),

        // ===== round 2: footnotes (structural) =====
        new("footnote-insert", "footnote", "insert-block",
            "Add a new footnote and reference it from a body paragraph.",
            d => InsertFootnote(d, "Purchase and Sale of Preferred Stock",
                               "Newly inserted footnote text for diff testing.")),

        new("footnote-delete", "footnote", "delete-block",
            "Delete a footnote and its body reference.",
            d => DeleteFirstFootnote(d)),

        // ===== round 2: multi-paragraph + structural body =====
        new("multi-paragraph-delete", "body-para", "delete-block",
            "Delete three consecutive body paragraphs.",
            d => DeleteConsecutiveParas(Body(d), "Purchase and Sale of Preferred Stock", 3)),

        new("multi-paragraph-insert", "body-para", "insert-block",
            "Insert three consecutive new paragraphs.",
            d => InsertConsecutiveParas(Body(d), "Purchase and Sale of Preferred Stock", 3)),

        new("whole-paragraph-replace", "body-para", "replace",
            "Replace ALL text of a paragraph with entirely new text.",
            d => ReplaceWholeParagraph(Body(d), "shall apply to each such closing",
                                       "This paragraph has been completely rewritten for diff testing.")),

        new("move-heading-block", "body-para", "move",
            "Move a (styled) heading paragraph to a distant location.",
            d => MoveTopLevelPara(Body(d),
                    moveText: "Purchase and Sale of Preferred Stock",
                    afterAnchor: "shall apply to each such closing unless otherwise specified")),

        // ===== mixed / stress =====
        new("multi-edit", "mixed", "mixed",
            "Several independent edits across body, table, and a format change.",
            d => {
                ReplaceInFirstText(Body(d), "Purchaser", "Investor");
                ReplaceInFirstText(FirstTable(d), "Total Shares", "Aggregate Shares");
                SetRunFormat(Body(d), "Purchase and Sale of Preferred Stock", r => r.Bold = new Bold());
            }),
    ];

    // ---- targeting helpers ---------------------------------------------------------------------

    private static Body Body(WordprocessingDocument d) =>
        d.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("no body");

    private static Table FirstTable(WordprocessingDocument d) =>
        Body(d).Descendants<Table>().FirstOrDefault()
        ?? throw new InvalidOperationException("no table in document");

    private static OpenXmlElement HeaderContaining(WordprocessingDocument d, string text) =>
        d.MainDocumentPart!.HeaderParts.Select(h => (OpenXmlElement?)h.Header)
            .FirstOrDefault(h => h is not null && TextOf(h).Contains(text))
        ?? throw new InvalidOperationException($"no header containing '{text}'");

    private static OpenXmlElement FooterContaining(WordprocessingDocument d, string text) =>
        d.MainDocumentPart!.FooterParts.Select(f => (OpenXmlElement?)f.Footer)
            .FirstOrDefault(f => f is not null && TextOf(f).Contains(text))
        ?? throw new InvalidOperationException($"no footer containing '{text}'");

    private static OpenXmlElement FootnotesRoot(WordprocessingDocument d) =>
        d.MainDocumentPart?.FootnotesPart?.Footnotes
        ?? throw new InvalidOperationException("no footnotes part");

    private static string TextOf(OpenXmlElement e) => string.Concat(e.Descendants<Text>().Select(t => t.Text));

    // ---- mutation primitives -------------------------------------------------------------------

    /// <summary>Replace the first <c>w:t</c> (anywhere under <paramref name="root"/>) containing
    /// <paramref name="find"/>, preserving surrounding whitespace.</summary>
    private static void ReplaceInFirstText(OpenXmlElement root, string find, string repl)
    {
        var t = root.Descendants<Text>().FirstOrDefault(x => x.Text.Contains(find))
            ?? throw new InvalidOperationException($"anchor text not found: '{find}'");
        t.Text = t.Text.Replace(find, repl);
        t.Space = SpaceProcessingModeValues.Preserve;
    }

    private static void ReplaceInPart(OpenXmlElement root, string find, string repl) =>
        ReplaceInFirstText(root, find, repl);

    private static Paragraph FindTopLevelPara(Body body, string contains) =>
        body.Elements<Paragraph>().FirstOrDefault(p => TextOf(p).Contains(contains))
        ?? throw new InvalidOperationException($"no top-level paragraph containing '{contains}'");

    private static void InsertParagraphAfter(Body body, string anchor, string newText)
    {
        var p = FindTopLevelPara(body, anchor);
        var np = new Paragraph(new Run(new Text(newText) { Space = SpaceProcessingModeValues.Preserve }));
        p.InsertAfterSelf(np);
    }

    private static void DeleteTopLevelPara(Body body, string contains) =>
        FindTopLevelPara(body, contains).Remove();

    private static void MoveTopLevelPara(Body body, string moveText, string afterAnchor)
    {
        var src = FindTopLevelPara(body, moveText);
        var anchor = FindTopLevelPara(body, afterAnchor);
        var clone = (Paragraph)src.CloneNode(true);
        src.Remove();
        anchor.InsertAfterSelf(clone);
    }

    /// <summary>Split a paragraph after its first run, moving the rest into a new following paragraph.</summary>
    private static void SplitTopLevelPara(Body body, string contains)
    {
        var p = FindTopLevelPara(body, contains);
        var runs = p.Elements<Run>().ToList();
        if (runs.Count < 2) throw new InvalidOperationException($"paragraph '{contains}' has too few runs to split");
        int half = runs.Count / 2;
        var np = new Paragraph();
        // preserve paragraph properties on the new paragraph
        if (p.ParagraphProperties is { } pPr) np.AppendChild((ParagraphProperties)pPr.CloneNode(true));
        for (int i = half; i < runs.Count; i++)
        {
            runs[i].Remove();
            np.AppendChild(runs[i]);
        }
        p.InsertAfterSelf(np);
    }

    private static void SetRunFormat(OpenXmlElement root, string anchorText, Action<RunProperties> apply)
    {
        var run = root.Descendants<Run>()
            .FirstOrDefault(r => r.Elements<Text>().Any(t => t.Text.Contains(anchorText)))
            ?? throw new InvalidOperationException($"no run containing '{anchorText}'");
        var rPr = run.RunProperties;
        if (rPr is null) { rPr = new RunProperties(); run.PrependChild(rPr); }
        apply(rPr);
    }

    private static void SetParagraphStyle(Body body, string contains, string styleId)
    {
        var p = FindTopLevelPara(body, contains);
        var pPr = p.ParagraphProperties;
        if (pPr is null) { pPr = new ParagraphProperties(); p.PrependChild(pPr); }
        pPr.ParagraphStyleId = new ParagraphStyleId { Val = styleId };
    }

    private static void InsertTableRow(Table table)
    {
        var rows = table.Elements<TableRow>().ToList();
        var template = rows[^1];
        var clone = (TableRow)template.CloneNode(true);
        // Give the new row distinctive content so the insertion is visible. Use the first cell's first
        // paragraph; synthesize a run if the cloned cell was empty.
        var firstCell = clone.Elements<TableCell>().FirstOrDefault()
            ?? throw new InvalidOperationException("cloned row has no cell");
        var firstText = firstCell.Descendants<Text>().FirstOrDefault();
        if (firstText is not null)
        {
            firstText.Text = "Inserted Row Cell";
            firstText.Space = SpaceProcessingModeValues.Preserve;
        }
        else
        {
            var para = firstCell.Elements<Paragraph>().FirstOrDefault();
            if (para is null) { para = new Paragraph(); firstCell.AppendChild(para); }
            para.AppendChild(new Run(new Text("Inserted Row Cell") { Space = SpaceProcessingModeValues.Preserve }));
        }
        template.InsertAfterSelf(clone);
    }

    private static void DeleteTableRow(Table table)
    {
        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count < 2)
            throw new InvalidOperationException("table has too few rows to delete safely");
        rows[^1].Remove();  // delete the last (data) row, keeping at least the header
    }

    // ---- round 2 primitives --------------------------------------------------------------------

    private static void InsertTableColumn(Table table)
    {
        var grid = table.Elements<TableGrid>().FirstOrDefault();
        if (grid is null) { grid = new TableGrid(); table.PrependChild(grid); }
        grid.AppendChild(new GridColumn { Width = "1500" });

        bool firstRow = true;
        foreach (var row in table.Elements<TableRow>())
        {
            var lastCell = row.Elements<TableCell>().LastOrDefault();
            var newCell = lastCell != null
                ? (TableCell)lastCell.CloneNode(true)
                : new TableCell(new Paragraph());
            if (firstRow)
            {
                var t = newCell.Descendants<Text>().FirstOrDefault();
                if (t is not null) { t.Text = "New Column"; t.Space = SpaceProcessingModeValues.Preserve; }
                else newCell.AppendChild(new Paragraph(new Run(new Text("New Column"))));
                firstRow = false;
            }
            row.AppendChild(newCell);
        }
    }

    private static void DeleteTableColumn(Table table)
    {
        table.Elements<TableGrid>().FirstOrDefault()?.Elements<GridColumn>().LastOrDefault()?.Remove();
        foreach (var row in table.Elements<TableRow>())
            row.Elements<TableCell>().LastOrDefault()?.Remove();
    }

    private static void InsertFootnote(WordprocessingDocument d, string anchorText, string footnoteText)
    {
        var main = d.MainDocumentPart!;
        var footnotes = main.FootnotesPart?.Footnotes
            ?? throw new InvalidOperationException("no footnotes part");
        long newId = footnotes.Elements<Footnote>().Select(f => f.Id?.Value ?? 0L).DefaultIfEmpty(0L).Max() + 1;
        footnotes.AppendChild(new Footnote(
            new Paragraph(new Run(new Text(footnoteText) { Space = SpaceProcessingModeValues.Preserve })))
        { Id = newId });

        var run = Body(d).Descendants<Run>()
            .FirstOrDefault(r => r.Elements<Text>().Any(t => t.Text.Contains(anchorText)))
            ?? throw new InvalidOperationException($"no run containing '{anchorText}'");
        run.InsertAfterSelf(new Run(new FootnoteReference { Id = newId }));
    }

    private static void DeleteFirstFootnote(WordprocessingDocument d)
    {
        var main = d.MainDocumentPart!;
        var fnRef = main.Document!.Body!.Descendants<FootnoteReference>().FirstOrDefault()
            ?? throw new InvalidOperationException("no footnote reference in body");
        long id = fnRef.Id?.Value ?? throw new InvalidOperationException("footnote reference has no id");
        (fnRef.Ancestors<Run>().FirstOrDefault() ?? (OpenXmlElement)fnRef).Remove();
        main.FootnotesPart?.Footnotes?.Elements<Footnote>()
            .FirstOrDefault(f => f.Id?.Value == id)?.Remove();
    }

    private static void DeleteConsecutiveParas(Body body, string anchor, int count)
    {
        var start = FindTopLevelPara(body, anchor);
        var elems = body.Elements().ToList();
        int idx = elems.IndexOf(start);
        var toRemove = elems.Skip(idx).OfType<Paragraph>().Take(count).ToList();
        foreach (var p in toRemove) p.Remove();
    }

    private static void InsertConsecutiveParas(Body body, string anchor, int count)
    {
        Paragraph last = FindTopLevelPara(body, anchor);
        for (int i = 0; i < count; i++)
        {
            var np = new Paragraph(new Run(
                new Text($"Inserted paragraph number {i + 1} for diff testing.")
                { Space = SpaceProcessingModeValues.Preserve }));
            last.InsertAfterSelf(np);
            last = np;
        }
    }

    private static void ReplaceWholeParagraph(Body body, string anchor, string newText)
    {
        var p = FindTopLevelPara(body, anchor);
        foreach (var r in p.Elements<Run>().ToList()) r.Remove();
        p.AppendChild(new Run(new Text(newText) { Space = SpaceProcessingModeValues.Preserve }));
    }
}
