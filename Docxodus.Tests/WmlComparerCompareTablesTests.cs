#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tests for <see cref="WmlComparerSettings.CompareTables"/> — Word's "Tables" compare option.
/// When off, body-level tables take no part in the comparison and the result carries the left
/// document's tables verbatim and unmarked.
/// </summary>
public class WmlComparerCompareTablesTests
{
    private static WmlDocument Doc(params XElement[] bodyChildren)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.PutXDocument(new XDocument(
                new XElement(W.document,
                    new XAttribute(XNamespace.Xmlns + "w", W.w),
                    new XElement(W.body, bodyChildren))));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(
                new XDocument(new XElement(W.styles, new XAttribute(XNamespace.Xmlns + "w", W.w))));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(
                new XDocument(new XElement(W.settings, new XAttribute(XNamespace.Xmlns + "w", W.w))));
        }
        return new WmlDocument("test.docx", ms.ToArray());
    }

    private static XElement Para(string text) =>
        new(W.p, new XElement(W.r, new XElement(W.t, text)));

    private static XElement Table(params string[][] rows) =>
        new(W.tbl,
            new XElement(W.tblPr,
                new XElement(W.tblW, new XAttribute(W._w, "0"), new XAttribute(W.type, "auto"))),
            new XElement(W.tblGrid,
                Enumerable.Range(0, rows[0].Length)
                    .Select(_ => new XElement(W.gridCol, new XAttribute(W._w, "3000")))),
            rows.Select(cells => new XElement(W.tr,
                cells.Select(c => new XElement(W.tc, Para(c))))));

    private static WmlComparerSettings Settings(bool compareTables) =>
        new() { CompareTables = compareTables, DateTimeForRevisions = "2000-01-01T00:00:00Z" };

    private static XElement Body(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    private static int RevisionCount(WmlDocument result, WmlComparerSettings settings) =>
        WmlComparer.GetRevisions(result, settings).Count;

    /// <summary>The block sequence of the produced body: paragraph text, or "TBL" for a table.</summary>
    private static List<string> BlockOrder(WmlDocument doc) =>
        Body(doc).Elements()
            .Where(e => e.Name == W.p || e.Name == W.tbl)
            .Select(e => e.Name == W.tbl ? "TBL" : e.Value)
            .ToList();

    /// <summary>
    /// Ignoring a scope must never produce a package Word would refuse. Also asserts no marker
    /// paragraph leaked into the output, whatever shape the renderer gave it.
    /// </summary>
    private static void AssertSchemaValidAndNoPlaceholder(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);

        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(wDoc)
            .Where(e => e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema &&
                        !OxPt.WcTests.ExpectedErrors.Contains(e.Description))
            .Select(e => e.Description)
            .ToList();
        Assert.Empty(errors);

        // Match the full internal marker (prefix + 32 hex), not the bare prefix — a document may
        // legitimately contain the prefix as ordinary prose.
        var xml = wDoc.MainDocumentPart!.GetXDocument().ToString();
        Assert.DoesNotMatch(@"DocxodusIgnoredTable[0-9a-f]{32}", xml);
    }

    /// <summary>
    /// WCT001 — the reported case: a table on the left, none on the right, identical prose.
    /// With tables off this is not a change at all.
    /// </summary>
    [Fact]
    public void WCT001_TableOnlyOnTheLeft_IsIgnored()
    {
        var left = Doc(Table(new[] { "1", "2", "3" }), Para("Hans Christian Andersen"));
        var right = Doc(Para("Hans Christian Andersen"));

        var offSettings = Settings(compareTables: false);
        var ignored = WmlComparer.Compare(left, right, offSettings);

        Assert.Equal(0, RevisionCount(ignored, offSettings));
        AssertSchemaValidAndNoPlaceholder(ignored);

        var table = Assert.Single(Body(ignored).Elements(W.tbl));
        Assert.Equal(3, table.Descendants(W.tc).Count());
        Assert.Empty(table.Descendants(W.ins));
        Assert.Empty(table.Descendants(W.del));

        // Guard: with the option ON the same pair genuinely differs, so the test is not vacuous.
        var onSettings = Settings(compareTables: true);
        Assert.NotEqual(0, RevisionCount(WmlComparer.Compare(left, right, onSettings), onSettings));
    }

    /// <summary>
    /// WCT002 — a row added to a table, prose identical on both sides. The added row is the only
    /// difference, so with tables off there is nothing to report.
    /// </summary>
    [Fact]
    public void WCT002_AddedTableRow_IsIgnored()
    {
        var threeRows = new[]
        {
            new[] { "1", "2", "3" }, new[] { "4", "5", "6" }, new[] { "7", "8", "9" },
        };
        var left = Doc(Table(threeRows), Para("Hans Christian Andersen"));
        var right = Doc(Table(threeRows.Append(new[] { "A", "B", "C" }).ToArray()),
            Para("Hans Christian Andersen"));

        var offSettings = Settings(compareTables: false);
        var ignored = WmlComparer.Compare(left, right, offSettings);

        Assert.Equal(0, RevisionCount(ignored, offSettings));
        AssertSchemaValidAndNoPlaceholder(ignored);

        // The left table survives as three rows — the added row is not carried over.
        var table = Assert.Single(Body(ignored).Elements(W.tbl));
        Assert.Equal(3, table.Elements(W.tr).Count());

        var onSettings = Settings(compareTables: true);
        Assert.NotEqual(0, RevisionCount(WmlComparer.Compare(left, right, onSettings), onSettings));
    }

    /// <summary>
    /// WCT003 — ignoring tables must not silence prose. A paragraph edit alongside a table edit is
    /// still reported; only the table part is dropped.
    /// </summary>
    [Fact]
    public void WCT003_ParagraphChangeIsStillReported_WhenTablesAreIgnored()
    {
        var left = Doc(Table(new[] { "1", "2", "3" }), Para("Hans Christian Andersen"));
        var right = Doc(Table(new[] { "9", "9", "9" }), Para("Someone Else"));

        var settings = Settings(compareTables: false);
        var result = WmlComparer.Compare(left, right, settings);

        AssertSchemaValidAndNoPlaceholder(result);
        Assert.NotEmpty(WmlComparer.GetRevisions(result, settings));

        var table = Assert.Single(Body(result).Elements(W.tbl));
        Assert.Empty(table.Descendants(W.ins));
        Assert.Empty(table.Descendants(W.del));
        Assert.Equal("1", table.Descendants(W.t).First().Value);
    }

    /// <summary>
    /// WCT004 — a table-only document has no paragraph anywhere, so the marker is the body's only
    /// block; it must not survive into the result.
    /// </summary>
    [Fact]
    public void WCT004_TableOnlyDocument_KeepsTheTableAndLeavesNoMarker()
    {
        var left = Doc(Table(new[] { "1", "2" }));
        var right = Doc(Table(new[] { "8", "9" }));

        var settings = Settings(compareTables: false);
        var result = WmlComparer.Compare(left, right, settings);

        Assert.Equal(0, RevisionCount(result, settings));
        AssertSchemaValidAndNoPlaceholder(result);
        Assert.Equal(new[] { "TBL" }, BlockOrder(result));
        Assert.Equal("1", Assert.Single(Body(result).Elements(W.tbl)).Descendants(W.t).First().Value);
    }

    /// <summary>WCT005 — the default is ON, so existing callers are unaffected.</summary>
    [Fact]
    public void WCT005_DefaultIsCompareTablesOn()
    {
        Assert.True(new WmlComparerSettings().CompareTables);
    }

    /// <summary>
    /// WCT006 — a table that is NOT the first body child keeps its place. This is the case the
    /// original unid-anchored implementation got wrong (the table moved to the top of the body).
    /// </summary>
    [Fact]
    public void WCT006_TableKeepsItsPosition_WhenNotFirstBlock()
    {
        var left = Doc(Para("AAA"), Table(new[] { "1", "2" }), Para("ZZZ"));
        var right = Doc(Para("AAA"), Para("ZZZ"));

        var settings = Settings(compareTables: false);
        var result = WmlComparer.Compare(left, right, settings);

        Assert.Equal(new[] { "AAA", "TBL", "ZZZ" }, BlockOrder(result));
        Assert.Equal(0, RevisionCount(result, settings));
        AssertSchemaValidAndNoPlaceholder(result);
    }

    /// <summary>
    /// WCT007 — several tables at different positions each stay put, in order, and keep their own
    /// content (no cross-assignment).
    /// </summary>
    [Fact]
    public void WCT007_MultipleTables_KeepTheirPositionsAndContent()
    {
        // Each table needs a paragraph after it: two adjacent w:tbl elements are one table.
        var left = Doc(
            Para("AAA"), Table(new[] { "T1" }),
            Para("BBB"), Table(new[] { "T2" }),
            Para("CCC"), Table(new[] { "T3" }),
            Para("DDD"));
        var right = Doc(Para("AAA"), Para("BBB"), Para("CCC"), Para("DDD"));

        var settings = Settings(compareTables: false);
        var result = WmlComparer.Compare(left, right, settings);

        Assert.Equal(
            new[] { "AAA", "TBL", "BBB", "TBL", "CCC", "TBL", "DDD" },
            BlockOrder(result));
        Assert.Equal(
            new[] { "T1", "T2", "T3" },
            Body(result).Elements(W.tbl).Select(t => t.Descendants(W.t).First().Value).ToArray());
        Assert.Equal(0, RevisionCount(result, settings));
        AssertSchemaValidAndNoPlaceholder(result);
    }

    /// <summary>
    /// WCT008 — a table only the RIGHT document has leaves nothing behind: there is no left table to
    /// carry over, and the marker standing in for it must not leak.
    /// </summary>
    [Fact]
    public void WCT008_TableOnlyOnTheRight_LeavesNoMarker()
    {
        var left = Doc(Para("AAA"), Para("ZZZ"));
        var right = Doc(Para("AAA"), Table(new[] { "9" }), Para("ZZZ"));

        var settings = Settings(compareTables: false);
        var result = WmlComparer.Compare(left, right, settings);

        Assert.Equal(new[] { "AAA", "ZZZ" }, BlockOrder(result));
        Assert.Equal(0, RevisionCount(result, settings));
        AssertSchemaValidAndNoPlaceholder(result);
    }

    /// <summary>
    /// WCT009 — a footnote cited only from inside an ignored table must still resolve, otherwise Word
    /// asks to repair the file.
    /// </summary>
    [Fact]
    public void WCT009_FootnoteCitedOnlyInsideAnIgnoredTable_StillResolves()
    {
        var left = DocWithFootnoteInTable();
        var right = Doc(Para("AAA"));

        var settings = Settings(compareTables: false);
        var result = WmlComparer.Compare(left, right, settings);

        AssertSchemaValidAndNoPlaceholder(result);

        using var ms = new MemoryStream(result.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        var referenced = wDoc.MainDocumentPart!.GetXDocument()
            .Descendants(W.footnoteReference)
            .Select(r => (string?)r.Attribute(W.id))
            .Where(id => id != null)
            .ToList();

        Assert.NotEmpty(referenced);

        var defined = wDoc.MainDocumentPart.FootnotesPart?.GetXDocument()
            .Root!.Elements(W.footnote)
            .Select(f => (string?)f.Attribute(W.id))
            .ToHashSet() ?? new HashSet<string?>();

        Assert.All(referenced, id => Assert.Contains(id, defined));
    }

    /// <summary>
    /// WCT010 — the right document has MORE tables than the left. The left table must survive and the
    /// extra right table must not become a revision.
    /// </summary>
    [Fact]
    public void WCT010_RightHasMoreTables_LeftTableSurvives()
    {
        var left = Doc(Para("AAA"), Para("BBB"), Table(new[] { "keep" }), Para("CCC"));
        var right = Doc(
            Para("AAA"), Table(new[] { "new" }),
            Para("BBB"), Table(new[] { "keep" }), Para("CCC"));

        var settings = Settings(compareTables: false);
        var result = WmlComparer.Compare(left, right, settings);

        Assert.Equal(new[] { "AAA", "BBB", "TBL", "CCC" }, BlockOrder(result));
        Assert.Equal("keep", Assert.Single(Body(result).Elements(W.tbl)).Descendants(W.t).First().Value);
        Assert.Equal(0, RevisionCount(result, settings));
        AssertSchemaValidAndNoPlaceholder(result);
    }

    /// <summary>
    /// WCT011 — a table removed BEFORE other tables must not perturb how the prose matches: the
    /// paragraphs are identical in both documents, so nothing may be reported.
    /// </summary>
    [Fact]
    public void WCT011_TableRemovedBeforeOtherTables_DoesNotChurnProse()
    {
        var left = Doc(
            Para("AAA"), Table(new[] { "T1" }),
            Para("BBB"), Table(new[] { "T2" }), Para("CCC"));
        var right = Doc(Para("AAA"), Para("BBB"), Table(new[] { "T2" }), Para("CCC"));

        var settings = Settings(compareTables: false);
        var result = WmlComparer.Compare(left, right, settings);

        Assert.Equal(new[] { "AAA", "TBL", "BBB", "TBL", "CCC" }, BlockOrder(result));
        Assert.Equal(
            new[] { "T1", "T2" },
            Body(result).Elements(W.tbl).Select(t => t.Descendants(W.t).First().Value).ToArray());
        Assert.Equal(0, RevisionCount(result, settings));
        AssertSchemaValidAndNoPlaceholder(result);
    }

    /// <summary>
    /// WCT012 — a table moved past a paragraph is still not a change, and the prose stays quiet.
    /// </summary>
    [Fact]
    public void WCT012_MovedTable_IsIgnored()
    {
        var left = Doc(Para("AAA"), Table(new[] { "T" }), Para("BBB"), Para("CCC"));
        var right = Doc(Para("AAA"), Para("BBB"), Table(new[] { "T" }), Para("CCC"));

        var settings = Settings(compareTables: false);
        var result = WmlComparer.Compare(left, right, settings);

        // The LEFT position wins, since the left document's tables are what is carried over.
        Assert.Equal(new[] { "AAA", "TBL", "BBB", "CCC" }, BlockOrder(result));
        Assert.Equal(0, RevisionCount(result, settings));
        AssertSchemaValidAndNoPlaceholder(result);
    }

    /// <summary>
    /// WCT013 — marker-lookalike prose present on ONE side only (the pairing that could break the
    /// marker): the table must survive and no marker token may leak. The marker text carries no word
    /// separator, so the LCS treats it as one indivisible word and real prose cannot alias its prefix.
    /// </summary>
    [Fact]
    public void WCT013_MarkerLookalikeProseOnOneSide_DoesNotConsumeTheTable()
    {
        var left = Doc(Para("AAA"), Table(new[] { "T" }), Para("ZZZ"));
        var right = Doc(Para("AAA"), Para("DocxodusIgnoredTable"), Para("ZZZ"));

        var settings = Settings(compareTables: false);
        var result = WmlComparer.Compare(left, right, settings);

        // The table survives with its content, and the internal marker never leaks (the lookalike
        // prose is preserved as ordinary right-side content, which is fine — only tables are ignored).
        Assert.Equal("T", Assert.Single(Body(result).Elements(W.tbl)).Descendants(W.t).First().Value);
        AssertSchemaValidAndNoPlaceholder(result);
    }

    /// <summary>
    /// WCT014 — a right-side paragraph that carries content but NO text (a line break lives in a run
    /// with no w:t) adjacent to an ignored table must not be swept away when its marker run is removed.
    /// </summary>
    [Fact]
    public void WCT014_TextlessButRealRightParagraph_SurvivesMarkerRemoval()
    {
        var breakPara = new XElement(W.p, new XElement(W.r, new XElement(W.br)));

        var left = Doc(Para("AAA"), Table(new[] { "T" }), Para("ZZZ"));
        var right = Doc(Para("AAA"), breakPara, Para("ZZZ"));

        var settings = Settings(compareTables: false);
        var result = WmlComparer.Compare(left, right, settings);

        var body = Body(result);
        Assert.Single(body.Elements(W.tbl));
        Assert.Single(body.Descendants(W.br)); // the break was not deleted with the marker run
        AssertSchemaValidAndNoPlaceholder(result);
    }

    /// <summary>
    /// WCT015 — Consolidate routes through the same internals: the left table must survive, the prose
    /// revision must land, and no marker may leak.
    /// </summary>
    [Fact]
    public void WCT015_Consolidate_KeepsTheTableAndTheProseRevision()
    {
        var original = Doc(Para("AAA"), Table(new[] { "T" }), Para("BBB"));
        var revised = Doc(Para("AAA"), Table(new[] { "CHANGED" }), Para("BBB edited"));

        var settings = Settings(compareTables: false);
        var consolidated = WmlComparer.Consolidate(
            original,
            new List<WmlRevisedDocumentInfo>
            {
                new WmlRevisedDocumentInfo { RevisedDocument = revised, Revisor = "Rev1" },
            },
            settings);

        // Consolidate juxtaposes revisions in its own wrapper table, so do not count tables: assert the
        // ignored table's own content survived, the prose edit landed, and no marker or table edit leaked.
        var body = Body(consolidated);
        Assert.Contains(body.Descendants(W.tbl),
            t => t.Descendants(W.t).Any(x => x.Value == "T"));
        Assert.NotEmpty(body.Descendants(W.ins));
        var xml = body.ToString();
        Assert.DoesNotMatch(@"DocxodusIgnoredTable[0-9a-f]{32}", xml);
        Assert.DoesNotContain("CHANGED", xml);
    }

    /// <summary>A document whose ONLY footnote reference sits in a body-level table cell.</summary>
    private static WmlDocument DocWithFootnoteInTable()
    {
        var cell = new XElement(W.tc,
            new XElement(W.tcPr,
                new XElement(W.tcW, new XAttribute(W._w, "3000"), new XAttribute(W.type, "dxa"))),
            new XElement(W.p,
                new XElement(W.r, new XElement(W.t, "cell")),
                new XElement(W.r, new XElement(W.footnoteReference, new XAttribute(W.id, "2")))));

        var table = new XElement(W.tbl,
            new XElement(W.tblPr,
                new XElement(W.tblW, new XAttribute(W._w, "0"), new XAttribute(W.type, "auto"))),
            new XElement(W.tblGrid, new XElement(W.gridCol, new XAttribute(W._w, "3000"))),
            new XElement(W.tr, cell));

        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.PutXDocument(new XDocument(
                new XElement(W.document,
                    new XAttribute(XNamespace.Xmlns + "w", W.w),
                    new XElement(W.body, Para("AAA"), table))));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(
                new XDocument(new XElement(W.styles, new XAttribute(XNamespace.Xmlns + "w", W.w))));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(
                new XDocument(new XElement(W.settings, new XAttribute(XNamespace.Xmlns + "w", W.w))));
            main.AddNewPart<FootnotesPart>().PutXDocument(new XDocument(
                new XElement(W.footnotes,
                    new XAttribute(XNamespace.Xmlns + "w", W.w),
                    new XElement(W.footnote, new XAttribute(W.id, "-1"),
                        new XAttribute(W.type, "separator"),
                        new XElement(W.p, new XElement(W.r, new XElement(W.separator)))),
                    new XElement(W.footnote, new XAttribute(W.id, "0"),
                        new XAttribute(W.type, "continuationSeparator"),
                        new XElement(W.p, new XElement(W.r, new XElement(W.continuationSeparator)))),
                    new XElement(W.footnote, new XAttribute(W.id, "2"),
                        new XElement(W.p, new XElement(W.r, new XElement(W.t, "note text")))))));
        }
        return new WmlDocument("left.docx", ms.ToArray());
    }
}
