#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using WTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using WTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Regression guard for the cross-table blank-spacer release bug. When base and next carry
/// ASYMMETRIC leading/trailing blank ("spacer") paragraphs around an edited table, a fungible blank
/// on one side gets matched to a blank on the OTHER side of the table (a cross-table pairing). The
/// crossing-resolution pass in <see cref="Docxodus.Ir.Diff.IrBlockAligner"/> used to release the
/// heavy Modified table into a whole-deleted + whole-inserted table pair (two <c>w:tbl</c>) rather
/// than demoting the fungible blank. Word emits ONE table with native per-row markup and renders the
/// blank as a separate ins/del. These tests pin that: exactly one table survives, the added row keeps
/// native <c>w:trPr/w:ins</c> row markup, and the accept ≡ right / reject ≡ left contract holds.
/// </summary>
public class IrTableCrossingReleaseTests
{
    /// <summary>
    /// Build a doc: one shared heading paragraph, <paramref name="leadingBlanks"/> empty spacer
    /// paragraphs, a single-column table with one cell paragraph per <paramref name="rowTexts"/>
    /// entry, then <paramref name="trailingBlanks"/> empty spacer paragraphs.
    /// </summary>
    private static WmlDocument HeadingBlanksTableDoc(
        string heading, int leadingBlanks, int trailingBlanks, params string[] rowTexts)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();

            var children = new List<OpenXmlElement>
            {
                new Paragraph(new Run(new Text(heading))),
            };
            for (int i = 0; i < leadingBlanks; i++)
                children.Add(new Paragraph());

            var tableChildren = new List<OpenXmlElement>
            {
                new TableProperties(),
                new TableGrid(new GridColumn()),
            };
            tableChildren.AddRange(rowTexts.Select(t =>
                (OpenXmlElement)new WTableRow(new WTableCell(new Paragraph(new Run(new Text(t)))))));
            children.Add(new WTable(tableChildren));

            for (int i = 0; i < trailingBlanks; i++)
                children.Add(new Paragraph());

            mainPart.Document = new Document(new Body(children));
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()));
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            doc.Save();
        }
        return new WmlDocument("crossing.docx", stream.ToArray());
    }

    private static int TableCount(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        return wdoc.MainDocumentPart!.Document.Body!.Descendants<WTable>().Count();
    }

    private static int InsertedRowMarkerCount(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        return wdoc.MainDocumentPart!.Document.Body!
            .Descendants<TableRowProperties>()
            .Count(trPr => trPr.Elements<Inserted>().Any());
    }

    private static List<string> BodyTexts(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var body = wdoc.MainDocumentPart?.Document.Body;
        return body is null
            ? new List<string>()
            : body.Descendants<Paragraph>().Select(p => p.InnerText).ToList();
    }

    [Fact]
    public void CrossTableBlankSpacer_KeepsOneTable_WithNativeRowMarkup()
    {
        // Base: heading, 1 leading blank, table[RowA, RowB], 1 trailing blank.
        // Next: heading, 2 leading blanks, table[RowA, RowB, RowC (added)], 1 trailing blank.
        // The asymmetric leading blanks make the left trailing blank pair to a right blank BEFORE the
        // table (a cross-table match), which used to release the whole Modified table into del+ins.
        var left = HeadingBlanksTableDoc("Shared Heading", leadingBlanks: 1, trailingBlanks: 1,
            "Row A", "Row B");
        var right = HeadingBlanksTableDoc("Shared Heading", leadingBlanks: 2, trailingBlanks: 1,
            "Row A", "Row B", "Row C");

        var result = DocxDiff.Compare(left, right);

        // Headline: ONE table (native per-row markup), NOT two (whole-deleted + whole-inserted).
        Assert.Equal(1, TableCount(result));

        // The added row keeps native w:trPr/w:ins row markup.
        Assert.True(InsertedRowMarkerCount(result) >= 1,
            "expected a native inserted-row marker (w:trPr/w:ins) for the added row");

        // Round-trip: accept ≡ right, reject ≡ left (the demoted blank is ins+del; the table returns
        // to the already-round-tripping Modified path).
        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }

    [Fact]
    public void CrossTableBlankSpacer_RowEdited_KeepsOneTable()
    {
        // Same asymmetric-blank shape, but the table edit is a cell-text change on an existing row.
        var left = HeadingBlanksTableDoc("Shared Heading", leadingBlanks: 1, trailingBlanks: 1,
            "Row A", "Row B");
        var right = HeadingBlanksTableDoc("Shared Heading", leadingBlanks: 2, trailingBlanks: 1,
            "Row A", "Row B edited");

        var result = DocxDiff.Compare(left, right);

        Assert.Equal(1, TableCount(result));

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(BodyTexts(right), BodyTexts(accepted));
        Assert.Equal(BodyTexts(left), BodyTexts(rejected));
    }
}
