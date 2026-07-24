#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Docxodus;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using WTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;

namespace Docxodus.Tests;

/// <summary>
/// Replace-gap seam discipline (root-caused from the Word-oracle corpus):
/// (1) the seam TERMINATOR — the paragraph that survives accept — must carry the INS-side (right)
/// pPr as current with the left archived in <c>w:pPrChange</c>, not silently keep the left's; and
/// (2) in-gap pairing is order-preserving — content-equal empty paragraphs must not pair across
/// already-formed pairs (they were silently relocated across tables, breaking reject ≡ left).
/// </summary>
public class DocxDiffSeamDisciplineTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static WmlDocument ParaDoc(bool centered, params string[] texts)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(texts.Select(t =>
            {
                var p = new Paragraph(new Run(new Text(t)));
                if (centered)
                    p.PrependChild(new ParagraphProperties(
                        new Justification { Val = JustificationValues.Center }));
                return (OpenXmlElement)p;
            })));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults());
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("t.docx", stream.ToArray());
    }

    [Fact]
    public void SeamTerminator_CarriesEmptyPPr_WordShape()
    {
        // Middle paragraphs share tokens (ModifyBlock); the outer two are ~0.2 Jaccard, so they
        // lower to Delete+Insert (a 2x2 gap — no 1x1 force-pair) and render through the seam.
        // Word's seam shape (every decoded oracle sighting): the surviving seam paragraph carries
        // an EMPTY pPr — no style, no props, no pPrChange — so neither side's paragraph
        // formatting outlives the seam. (Word itself gives up property-level round-trip here;
        // text-level accept ≡ right / reject ≡ left still holds.)
        var left = ParaDoc(centered: true,
            "Alpha ancient prose.", "This document demonstrates the old body.", "Omega bygone prose.");
        var right = ParaDoc(centered: false,
            "Zulu contemporary words.", "This document demonstrates superscript body.", "Kappa modern words.");

        var redline = DocxDiff.Compare(left, right);

        // accept ≡ right at the property level: the right has no centered paragraphs, and the
        // LEFT's centering must not leak through the surviving seam paragraphs.
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        using (var s = new MemoryStream(accepted.DocumentByteArray))
        using (var d = WordprocessingDocument.Open(s, false))
        {
            var centeredAfterAccept = d.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>()
                .Count(p => p.ParagraphProperties?.Justification?.Val?.Value == JustificationValues.Center);
            Assert.Equal(0, centeredAfterAccept);
        }

        // The redline itself: surviving mixed ins+del (seam) paragraphs carry NO pPr content.
        using (var s = new MemoryStream(redline.DocumentByteArray))
        using (var d = WordprocessingDocument.Open(s, false))
        {
            var mixed = d.MainDocumentPart!.Document.Body!.Elements<Paragraph>()
                .Where(p => p.Descendants<InsertedRun>().Any() && p.Descendants<DeletedRun>().Any())
                .Where(p => p.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Deleted>() is null)
                .Where(p => p.ParagraphProperties?.GetFirstChild<ParagraphPropertiesChange>() is null)
                .ToList();
            Assert.NotEmpty(mixed);
            foreach (var p in mixed)
                Assert.True(p.ParagraphProperties is null || !p.ParagraphProperties.HasChildren,
                    $"seam paragraph '{p.InnerText}' must carry an empty pPr, got: {p.ParagraphProperties?.OuterXml}");
        }
    }

    [Fact]
    public void PairedParagraph_RightOnlyStyleRef_IsDroppedFromCurrentPPr()
    {
        // The left style universe has no "GrandTitle"; the right styles every paragraph with it.
        // Word expresses a PAIRED paragraph's format change within the LEFT style universe: the
        // unresolvable pStyle is dropped from the current pPr (direct props only) — the oracle's
        // output for exactly this corpus shape carries no pStyle and no imported style definition.
        static WmlDocument StyledDoc(bool styled, params string[] texts)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                main.Document = new Document(new Body(texts.Select(t =>
                {
                    var p = new Paragraph(new Run(new Text(t)));
                    if (styled)
                        p.PrependChild(new ParagraphProperties(new ParagraphStyleId { Val = "GrandTitle" }));
                    return (OpenXmlElement)p;
                })));
                var styles = new Styles(new DocDefaults());
                if (styled)
                    styles.Append(new Style(new StyleName { Val = "Grand Title" })
                    { Type = StyleValues.Paragraph, StyleId = "GrandTitle" });
                main.AddNewPart<StyleDefinitionsPart>().Styles = styles;
                main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                doc.Save();
            }
            return new WmlDocument("t.docx", stream.ToArray());
        }

        var left = StyledDoc(styled: false,
            "Alpha ancient prose.", "This document demonstrates the old body.", "Omega bygone prose.");
        var right = StyledDoc(styled: true,
            "Zulu contemporary words.", "This document demonstrates superscript body.", "Kappa modern words.");

        var redline = DocxDiff.Compare(left, right);

        using var s = new MemoryStream(redline.DocumentByteArray);
        using var d = WordprocessingDocument.Open(s, false);
        var refs = d.MainDocumentPart!.Document.Body!
            .Descendants<ParagraphStyleId>()
            .Where(id => id.Val?.Value == "GrandTitle" &&
                         id.Ancestors<ParagraphPropertiesChange>().FirstOrDefault() is null)
            .ToList();
        Assert.Empty(refs);
    }

    [Fact]
    public void DanglingParagraphStyleRefs_AreRemovedFromVerbatimAndArchivedPPr()
    {
        // The malformed sources name Title/HeadingN but define only Normal plus one unrelated,
        // valid paragraph style. The first paragraph is EqualBlock (verbatim clone); the second
        // changes formatting and therefore archives the left pPr inside pPrChange. Both paths
        // previously leaked dangling pStyle references into output, unlike Word's repair.
        static WmlDocument StyledDoc(bool right)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                Paragraph MakeParagraph(string style, string text, string? before = null)
                {
                    var spacing = new SpacingBetweenLines { After = "80", Line = "240" };
                    if (before is not null)
                        spacing.Before = before;
                    return new Paragraph(
                        new ParagraphProperties(new ParagraphStyleId { Val = style }, spacing),
                        new Run(new Text(text)));
                }

                var title = new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId { Val = "Title" },
                        new SpacingBetweenLines { Line = "276" }),
                    new Run(new Text("Shared title.")));
                main.Document = new Document(new Body(
                    title,
                    MakeParagraph(right ? "Heading3" : "Heading2", "Shared heading.", right ? "320" : "360"),
                    MakeParagraph("DefinedHeading", "Defined style.")));

                var normal = new Style(new StyleName { Val = "Normal" })
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Normal",
                    Default = true,
                };
                var defined = new Style(new StyleName { Val = "Defined Heading" })
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "DefinedHeading",
                };
                main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(), normal, defined);
                main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                doc.Save();
            }
            return new WmlDocument("t.docx", stream.ToArray());
        }

        static XElement BodyXml(WmlDocument doc)
        {
            using var stream = new MemoryStream(doc.DocumentByteArray);
            using var wdoc = WordprocessingDocument.Open(stream, false);
            return XDocument.Load(wdoc.MainDocumentPart!.GetStream()).Root!.Element(W + "body")!;
        }

        static List<string> ParagraphTexts(WmlDocument doc)
        {
            using var stream = new MemoryStream(doc.DocumentByteArray);
            using var wdoc = WordprocessingDocument.Open(stream, false);
            return wdoc.MainDocumentPart!.Document.Body!.Elements<Paragraph>()
                .Select(p => p.InnerText).ToList();
        }

        var left = StyledDoc(right: false);
        var right = StyledDoc(right: true);
        var redline = DocxDiff.Compare(left, right);
        var body = BodyXml(redline);

        // The only retained pStyle is backed by a final paragraph-style definition.
        var refs = body.Descendants(W + "pStyle")
            .Select(pStyle => (string?)pStyle.Attribute(W + "val"))
            .ToList();
        Assert.Equal(new[] { "DefinedHeading" }, refs);

        var headingPPr = body.Elements(W + "p")
            .Single(p => p.Value.Contains("Shared heading.", System.StringComparison.Ordinal))
            .Element(W + "pPr")!;
        Assert.Null(headingPPr.Element(W + "pStyle"));
        Assert.Equal("320", (string?)headingPPr.Element(W + "spacing")?.Attribute(W + "before"));
        var archivedPPr = headingPPr.Element(W + "pPrChange")!.Element(W + "pPr")!;
        Assert.Null(archivedPPr.Element(W + "pStyle"));
        Assert.Equal("360", (string?)archivedPPr.Element(W + "spacing")?.Attribute(W + "before"));

        Assert.Equal(ParagraphTexts(right), ParagraphTexts(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(ParagraphTexts(left), ParagraphTexts(RevisionProcessor.RejectRevisions(redline)));
    }

    private static WmlDocument BlockDoc(params string[] blocks)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var body = new Body();
            foreach (var b in blocks)
            {
                if (b.StartsWith("TBL:"))
                    body.Append(new WTable(
                        new TableGrid(new GridColumn()),
                        new WTableRow(new WTableCell(new Paragraph(new Run(new Text(b[4..])))))));
                else if (b == "~")
                    body.Append(new Paragraph(new ParagraphProperties(
                        new SpacingBetweenLines { Line = "276" })));
                else if (b.Length == 0)
                    body.Append(new Paragraph());
                else
                    body.Append(new Paragraph(new Run(new Text(b))));
            }
            main.Document = new Document(body);
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults());
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("t.docx", stream.ToArray());
    }

    private static List<(string Kind, string Text)> Shape(byte[] bytes)
    {
        using var s = new MemoryStream(bytes);
        using var d = WordprocessingDocument.Open(s, false);
        return d.MainDocumentPart!.Document.Body!.Elements()
            .Where(e => e is Paragraph or WTable)
            .Select(e => (e is WTable ? "tbl" : "p", e.InnerText))
            .ToList();
    }

    [Fact]
    public void EmptyDeletedParagraphs_StayInPlace_RejectReproducesLeft()
    {
        // Left: heading, TWO EMPTY spacing-pPr paragraphs, table, trailing bare empty.
        var left = BlockDoc("Support Tickets", "~", "~", "TBL:Old", "");
        // Right: three new paragraphs, a different table, bare empty, another heading + table,
        // bare empty. The bare empties after the LATER tables are the cross-gap pairing bait.
        var right = BlockDoc("Table Widths", "This document includes tables.", "Test One heading",
            "TBL:New1", "", "Test Two heading", "TBL:New2", "");

        var redline = DocxDiff.Compare(left, right);

        // (a) The two deleted empty paragraphs appear BEFORE the first table in the redline.
        var redShape = Shape(redline.DocumentByteArray);
        var firstTbl = redShape.FindIndex(b => b.Kind == "tbl");
        var emptiesBeforeTable = redShape.Take(firstTbl).Count(b => b.Kind == "p" && b.Text.Length == 0);
        Assert.True(emptiesBeforeTable >= 2,
            $"expected the 2 deleted empty paragraphs before the table, found {emptiesBeforeTable}");

        // (b) reject ≡ left — block kinds, texts AND ORDER.
        var rejected = RevisionProcessor.RejectRevisions(redline);
        Assert.Equal(Shape(left.DocumentByteArray), Shape(rejected.DocumentByteArray));
    }
}
