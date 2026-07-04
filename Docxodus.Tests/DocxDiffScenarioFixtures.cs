#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
// Docxodus defines its own Table/TableRow/TableCell/… types that clash with the WordprocessingML ones,
// so DON'T `using Docxodus;` here — alias only the one Docxodus type this file needs.
using WmlDocument = Docxodus.WmlDocument;
using WpTable = DocumentFormat.OpenXml.Wordprocessing.Table;

// Deliberately NOT under the `Docxodus` namespace: that would make Docxodus' own Table/TableRow/TableCell
// (from XlsxTables) visible by the enclosing-namespace rule and clash with the WordprocessingML types.
namespace DocxodusDiffParityFixtures;

/// <summary>
/// A LibreOffice-free port of the <c>tools/diffharness</c> verification campaign: a synthetic but
/// feature-rich base document plus a catalog of isolated, content-anchored edits, used by
/// <see cref="DocxDiffScenarioTests"/> to assert <see cref="DocxDiff"/>'s universal invariants
/// (round-trip <c>accept ≡ right</c> / <c>reject ≡ left</c>, no header/footer part duplication, schema
/// validity) on every edit type × feature.
/// </summary>
/// <remarks>
/// The base is deliberately shaped to exercise the bugs the campaign found:
/// <list type="bullet">
/// <item>TWO sections — section 1 ends with an INNER <c>w:sectPr</c> (inside a paragraph's <c>w:pPr</c>)
/// carrying a <c>w:headerReference</c>/<c>w:footerReference</c>; section 2's body <c>w:sectPr</c> carries the
/// same. The inner section-break paragraph is an Equal block on most edits, the exact shape that made the
/// renderer duplicate the RIGHT's header/footer parts (F1).</item>
/// <item>A 2×3 table (cell edits, row/column add/remove).</item>
/// <item>A footnotes part with two referenced footnotes (notes round-trip survives footnote renumbering).</item>
/// <item>Distinctive body paragraphs (word/paragraph/format/move/split edits).</item>
/// </list>
/// </remarks>
internal static class DocxDiffScenarioFixtures
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private static byte[]? s_baseBytes;

    /// <summary>The synthetic feature-rich base document bytes (built once, reused).</summary>
    public static byte[] BaseBytes => s_baseBytes ??= BuildBase();

    /// <summary>Open the base as a fresh editable copy and apply <paramref name="mutate"/> to produce a variant.</summary>
    private static WmlDocument Apply(Action<WordprocessingDocument> mutate)
    {
        using var ms = new MemoryStream();
        ms.Write(BaseBytes, 0, BaseBytes.Length);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, true))  // AutoSave flushes on dispose
            mutate(doc);
        return new WmlDocument("variant.docx", ms.ToArray());
    }

    public static WmlDocument BaseDoc() => new("base.docx", (byte[])BaseBytes.Clone());

    /// <summary>The left (base) and right (mutated) documents for a named scenario.</summary>
    public static (WmlDocument Left, WmlDocument Right) Build(string scenario)
    {
        var mutate = Catalog().TryGetValue(scenario, out var m)
            ? m
            : throw new ArgumentException($"unknown scenario '{scenario}'", nameof(scenario));
        return (BaseDoc(), Apply(mutate));
    }

    public static IEnumerable<string> Names() => Catalog().Keys;

    // ---- the scenario catalog ------------------------------------------------------------------

    private static IReadOnlyDictionary<string, Action<WordprocessingDocument>> Catalog() => new
        Dictionary<string, Action<WordprocessingDocument>>
    {
        // body paragraphs — text edits
        ["body-replace-word"] = d => ReplaceInFirstText(Body(d), "Purchaser", "Investor"),
        ["body-insert-word"] = d => ReplaceInFirstText(Body(d), "Preferred Stock", "New Preferred Stock"),
        ["body-delete-word"] = d => ReplaceInFirstText(Body(d), "each closing", "closing"),
        ["body-replace-phrase"] = d => ReplaceInFirstText(Body(d),
            "terms and conditions", "terms, conditions and provisions"),

        // body paragraphs — structural
        ["body-insert-paragraph"] = d => InsertParagraphAfter(Body(d),
            "Purchase and Sale", "This is an inserted clause for diff testing."),
        ["body-delete-paragraph"] = d => DeleteTopLevelPara(Body(d), "shall apply to each closing"),
        ["body-move-paragraph"] = d => MoveTopLevelPara(Body(d),
            moveText: "shall apply to each closing", afterAnchor: "Purchase and Sale"),
        ["body-split-paragraph"] = d => SplitTopLevelPara(Body(d), "This Agreement is made"),
        ["whole-paragraph-replace"] = d => ReplaceWholeParagraph(Body(d), "shall apply to each closing",
            "This paragraph has been completely rewritten for diff testing purposes."),

        // formatting (text unchanged)
        ["format-bold-run"] = d => SetRunFormat(Body(d), "Purchase and Sale", r => r.Bold = new Bold()),
        ["format-italic-run"] = d => SetRunFormat(Body(d), "shall apply", r => r.Italic = new Italic()),
        ["format-size-run"] = d => SetRunFormat(Body(d), "Purchase and Sale",
            r => { r.FontSize = new FontSize { Val = "36" }; r.FontSizeComplexScript = new FontSizeComplexScript { Val = "36" }; }),
        ["format-color-run"] = d => SetRunFormat(Body(d), "shall apply", r => r.Color = new Color { Val = "FF0000" }),
        ["format-underline-run"] = d => SetRunFormat(Body(d), "shall apply",
            r => r.Underline = new Underline { Val = UnderlineValues.Single }),
        ["style-change-paragraph"] = d => SetParagraphStyle(Body(d), "shall apply", "Heading2"),

        // tables
        ["table-cell-edit"] = d => ReplaceInFirstText(FirstTable(d), "Acme Corp", "Beta LLC"),
        ["table-cell-insert-word"] = d => ReplaceInFirstText(FirstTable(d), "Name and Address", "Full Name and Address"),
        ["table-insert-row"] = d => InsertTableRow(FirstTable(d)),
        ["table-delete-row"] = d => DeleteTableRow(FirstTable(d)),
        ["table-insert-column"] = d => InsertTableColumn(FirstTable(d)),
        ["table-delete-column"] = d => DeleteTableColumn(FirstTable(d)),
        ["table-cell-format"] = d => SetRunFormat(FirstTable(d), "Shares", r => r.Bold = new Bold()),

        // footnotes
        ["footnote-edit"] = d => ReplaceInPart(FootnotesRoot(d), "mandatory tranches", "mandatory funding tranches"),
        ["footnote-insert"] = d => InsertFootnote(d, "Purchase and Sale", "Newly inserted footnote text."),
        ["footnote-delete"] = d => DeleteFirstFootnote(d),

        // headers / footers (undiffed scope — round-trips on reject; output must not bloat parts)
        ["header-edit"] = d => ReplaceInPart(HeaderContaining(d, "Confidential"), "Confidential", "Privileged"),
        ["footer-edit"] = d => ReplaceInPart(FooterContaining(d, "ACTIVE/"), "ACTIVE/", "DRAFT/"),

        // content control (F4)
        ["sdt-insert"] = d => InsertInlineSdt(Body(d), "Purchase and Sale", " (controlled insertion)"),

        // mixed
        ["multi-edit"] = d =>
        {
            ReplaceInFirstText(Body(d), "Purchaser", "Investor");
            ReplaceInFirstText(FirstTable(d), "Acme Corp", "Beta LLC");
            SetRunFormat(Body(d), "Purchase and Sale", r => r.Bold = new Bold());
        },
    };

    // ---- base document builder -----------------------------------------------------------------

    private static byte[] BuildBase()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var styles = main.AddNewPart<StyleDefinitionsPart>();
            styles.Styles = new Styles(
                new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))),
                new Style(new StyleName { Val = "Heading 2" }) { Type = StyleValues.Paragraph, StyleId = "Heading2" });
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var hdr = main.AddNewPart<HeaderPart>();
            var hdrId = main.GetIdOfPart(hdr);
            WritePartXml(hdr, $"<w:hdr xmlns:w=\"{W}\"><w:p><w:r><w:t xml:space=\"preserve\">Confidential — Preferred Stock Purchase Agreement</w:t></w:r></w:p></w:hdr>");

            var ftr = main.AddNewPart<FooterPart>();
            var ftrId = main.GetIdOfPart(ftr);
            WritePartXml(ftr, $"<w:ftr xmlns:w=\"{W}\"><w:p><w:r><w:t>ACTIVE/12345</w:t></w:r></w:p></w:ftr>");

            var fn = main.AddNewPart<FootnotesPart>();
            WritePartXml(fn,
                $"<w:footnotes xmlns:w=\"{W}\">" +
                "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>" +
                "<w:footnote w:id=\"1\"><w:p><w:r><w:t xml:space=\"preserve\">Include this provision only if there are mandatory tranches.</w:t></w:r></w:p></w:footnote>" +
                "<w:footnote w:id=\"2\"><w:p><w:r><w:t xml:space=\"preserve\">See Section 1.2 for the closing schedule.</w:t></w:r></w:p></w:footnote>" +
                "</w:footnotes>");

            string SectPr() =>
                $"<w:sectPr><w:headerReference w:type=\"default\" r:id=\"{hdrId}\"/>" +
                $"<w:footerReference w:type=\"default\" r:id=\"{ftrId}\"/>" +
                "<w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>";

            var documentXml =
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{R}\"><w:body>" +
                // section 1
                "<w:p><w:r><w:t>Series Preferred Stock Purchase Agreement</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">This Agreement is made by and among the Company and each Purchaser listed in Exhibit A hereto.</w:t></w:r>" +
                "<w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>" +
                "<w:p><w:r><w:t>Purchase and Sale of Preferred Stock</w:t></w:r></w:p>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">The terms and conditions shall apply to each closing unless otherwise specified herein.</w:t></w:r>" +
                "<w:r><w:footnoteReference w:id=\"2\"/></w:r></w:p>" +
                // INNER section break (carries header/footer references — the F1 trigger)
                $"<w:p><w:pPr>{SectPr()}</w:pPr></w:p>" +
                // section 2: a table + a paragraph
                "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
                "<w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>" +
                "<w:tr><w:tc><w:p><w:r><w:t>Name and Address</w:t></w:r></w:p></w:tc>" +
                "<w:tc><w:p><w:r><w:t>Shares</w:t></w:r></w:p></w:tc>" +
                "<w:tc><w:p><w:r><w:t>Price</w:t></w:r></w:p></w:tc></w:tr>" +
                "<w:tr><w:tc><w:p><w:r><w:t>Acme Corp</w:t></w:r></w:p></w:tc>" +
                "<w:tc><w:p><w:r><w:t>100</w:t></w:r></w:p></w:tc>" +
                "<w:tc><w:p><w:r><w:t>1000</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                "<w:p><w:r><w:t>Closing conditions and representations are set forth below.</w:t></w:r></w:p>" +
                // trailing body section break
                SectPr() +
                "</w:body></w:document>";
            WritePartXml(main, documentXml);
        }
        return ms.ToArray();
    }

    private static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
        writer.Write(xml);
    }

    // ---- targeting helpers ---------------------------------------------------------------------

    private static Body Body(WordprocessingDocument d) =>
        d.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("no body");

    private static WpTable FirstTable(WordprocessingDocument d) =>
        Body(d).Descendants<WpTable>().FirstOrDefault() ?? throw new InvalidOperationException("no table");

    private static OpenXmlElement HeaderContaining(WordprocessingDocument d, string text) =>
        d.MainDocumentPart!.HeaderParts.Select(h => (OpenXmlElement?)h.Header)
            .First(h => h is not null && TextOf(h).Contains(text))!;

    private static OpenXmlElement FooterContaining(WordprocessingDocument d, string text) =>
        d.MainDocumentPart!.FooterParts.Select(f => (OpenXmlElement?)f.Footer)
            .First(f => f is not null && TextOf(f).Contains(text))!;

    private static OpenXmlElement FootnotesRoot(WordprocessingDocument d) =>
        d.MainDocumentPart?.FootnotesPart?.Footnotes ?? throw new InvalidOperationException("no footnotes part");

    private static string TextOf(OpenXmlElement e) => string.Concat(e.Descendants<Text>().Select(t => t.Text));

    // ---- mutation primitives (ported from tools/diffharness/Scenarios.cs) ----------------------

    private static void ReplaceInFirstText(OpenXmlElement root, string find, string repl)
    {
        var t = root.Descendants<Text>().FirstOrDefault(x => x.Text.Contains(find))
            ?? throw new InvalidOperationException($"anchor text not found: '{find}'");
        t.Text = t.Text.Replace(find, repl);
        t.Space = SpaceProcessingModeValues.Preserve;
    }

    private static void ReplaceInPart(OpenXmlElement root, string find, string repl) => ReplaceInFirstText(root, find, repl);

    private static Paragraph FindTopLevelPara(Body body, string contains) =>
        body.Elements<Paragraph>().FirstOrDefault(p => TextOf(p).Contains(contains))
        ?? throw new InvalidOperationException($"no top-level paragraph containing '{contains}'");

    private static void InsertParagraphAfter(Body body, string anchor, string newText) =>
        FindTopLevelPara(body, anchor).InsertAfterSelf(
            new Paragraph(new Run(new Text(newText) { Space = SpaceProcessingModeValues.Preserve })));

    private static void DeleteTopLevelPara(Body body, string contains) => FindTopLevelPara(body, contains).Remove();

    private static void MoveTopLevelPara(Body body, string moveText, string afterAnchor)
    {
        var src = FindTopLevelPara(body, moveText);
        var anchor = FindTopLevelPara(body, afterAnchor);
        var clone = (Paragraph)src.CloneNode(true);
        src.Remove();
        anchor.InsertAfterSelf(clone);
    }

    private static void SplitTopLevelPara(Body body, string contains)
    {
        var p = FindTopLevelPara(body, contains);
        var runs = p.Elements<Run>().ToList();
        if (runs.Count < 1) throw new InvalidOperationException("paragraph has no runs to split");
        // split the single run's text into two paragraphs at a word boundary
        var text = string.Concat(runs.SelectMany(r => r.Elements<Text>()).Select(t => t.Text));
        int mid = text.IndexOf(' ', text.Length / 2);
        if (mid < 0) mid = text.Length / 2;
        foreach (var r in runs) r.Remove();
        p.AppendChild(new Run(new Text(text[..mid]) { Space = SpaceProcessingModeValues.Preserve }));
        var np = new Paragraph(new Run(new Text(text[mid..].TrimStart()) { Space = SpaceProcessingModeValues.Preserve }));
        p.InsertAfterSelf(np);
    }

    private static void ReplaceWholeParagraph(Body body, string anchor, string newText)
    {
        var p = FindTopLevelPara(body, anchor);
        foreach (var r in p.Elements<Run>().ToList()) r.Remove();
        p.AppendChild(new Run(new Text(newText) { Space = SpaceProcessingModeValues.Preserve }));
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

    private static void InsertTableRow(WpTable table)
    {
        var clone = (TableRow)table.Elements<TableRow>().Last().CloneNode(true);
        var firstText = clone.Descendants<Text>().FirstOrDefault();
        if (firstText is not null) { firstText.Text = "Inserted Row"; firstText.Space = SpaceProcessingModeValues.Preserve; }
        table.Elements<TableRow>().Last().InsertAfterSelf(clone);
    }

    private static void DeleteTableRow(WpTable table)
    {
        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count < 2) throw new InvalidOperationException("table too small to delete a row");
        rows[^1].Remove();
    }

    private static void InsertTableColumn(WpTable table)
    {
        var grid = table.Elements<TableGrid>().First();
        grid.AppendChild(new GridColumn { Width = "1500" });
        bool firstRow = true;
        foreach (var row in table.Elements<TableRow>())
        {
            var clone = (TableCell)row.Elements<TableCell>().Last().CloneNode(true);
            var t = clone.Descendants<Text>().FirstOrDefault();
            if (t is not null) { t.Text = firstRow ? "New Column" : "x"; t.Space = SpaceProcessingModeValues.Preserve; }
            row.AppendChild(clone);
            firstRow = false;
        }
    }

    private static void DeleteTableColumn(WpTable table)
    {
        table.Elements<TableGrid>().First().Elements<GridColumn>().Last().Remove();
        foreach (var row in table.Elements<TableRow>())
            row.Elements<TableCell>().Last().Remove();
    }

    private static void InsertFootnote(WordprocessingDocument d, string anchorText, string footnoteText)
    {
        var footnotes = d.MainDocumentPart!.FootnotesPart!.Footnotes!;
        long newId = footnotes.Elements<Footnote>().Select(f => f.Id?.Value ?? 0L).DefaultIfEmpty(0L).Max() + 1;
        footnotes.AppendChild(new Footnote(
            new Paragraph(new Run(new Text(footnoteText) { Space = SpaceProcessingModeValues.Preserve }))) { Id = newId });
        var run = Body(d).Descendants<Run>()
            .First(r => r.Elements<Text>().Any(t => t.Text.Contains(anchorText)));
        run.InsertAfterSelf(new Run(new FootnoteReference { Id = newId }));
    }

    private static void DeleteFirstFootnote(WordprocessingDocument d)
    {
        var main = d.MainDocumentPart!;
        var fnRef = main.Document!.Body!.Descendants<FootnoteReference>().First();
        long id = fnRef.Id!.Value;
        (fnRef.Ancestors<Run>().FirstOrDefault() ?? (OpenXmlElement)fnRef).Remove();
        main.FootnotesPart!.Footnotes!.Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == id)?.Remove();
    }

    private static void InsertInlineSdt(Body body, string anchorText, string sdtText)
    {
        var run = body.Descendants<Run>().First(r => r.Elements<Text>().Any(t => t.Text.Contains(anchorText)));
        run.InsertAfterSelf(new SdtRun(
            new SdtProperties(new SdtId { Val = 9001 }),
            new SdtContentRun(new Run(new Text(sdtText) { Space = SpaceProcessingModeValues.Preserve }))));
    }
}
