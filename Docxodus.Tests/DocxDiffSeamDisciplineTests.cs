#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
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
    public void SeamTerminator_AcceptCarriesRightPPr_ArchivesLeftInPPrChange()
    {
        // Middle paragraphs share tokens (ModifyBlock); the outer two are ~0.2 Jaccard, so they
        // lower to Delete+Insert (a 2x2 gap — no 1x1 force-pair) and render through the seam.
        var left = ParaDoc(centered: true,
            "Alpha ancient prose.", "This document demonstrates the old body.", "Omega bygone prose.");
        var right = ParaDoc(centered: false,
            "Zulu contemporary words.", "This document demonstrates superscript body.", "Kappa modern words.");

        var redline = DocxDiff.Compare(left, right);

        // accept ≡ right at the PROPERTY level: the right has no centered paragraphs.
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        using (var s = new MemoryStream(accepted.DocumentByteArray))
        using (var d = WordprocessingDocument.Open(s, false))
        {
            var centeredAfterAccept = d.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>()
                .Count(p => p.ParagraphProperties?.Justification?.Val?.Value == JustificationValues.Center);
            Assert.Equal(0, centeredAfterAccept);
        }

        // reject ≡ left at the property level too.
        var rejected = RevisionProcessor.RejectRevisions(redline);
        using (var s = new MemoryStream(rejected.DocumentByteArray))
        using (var d = WordprocessingDocument.Open(s, false))
        {
            var uncentered = d.MainDocumentPart!.Document.Body!
                .Elements<Paragraph>()
                .Count(p => p.ParagraphProperties?.Justification?.Val?.Value != JustificationValues.Center);
            Assert.Equal(0, uncentered);
        }

        // The redline itself: every surviving mixed ins+del paragraph carries the RIGHT pPr as
        // current and archives the LEFT in pPrChange.
        using (var s = new MemoryStream(redline.DocumentByteArray))
        using (var d = WordprocessingDocument.Open(s, false))
        {
            var mixed = d.MainDocumentPart!.Document.Body!.Elements<Paragraph>()
                .Where(p => p.Descendants<InsertedRun>().Any() && p.Descendants<DeletedRun>().Any())
                .Where(p => p.ParagraphProperties?.ParagraphMarkRunProperties?.GetFirstChild<Deleted>() is null)
                .ToList();
            Assert.NotEmpty(mixed);
            foreach (var p in mixed)
            {
                var currentCentered =
                    p.ParagraphProperties?.Justification?.Val?.Value == JustificationValues.Center;
                var archived = p.ParagraphProperties?.GetFirstChild<ParagraphPropertiesChange>() is not null;
                Assert.False(currentCentered, $"surviving paragraph '{p.InnerText}' kept the LEFT pPr");
                Assert.True(archived, $"surviving paragraph '{p.InnerText}' lost the LEFT pPr archive");
            }
        }
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
