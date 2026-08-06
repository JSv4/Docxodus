#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Degenerate bodies that Word's UI cannot produce but a generator can: a body with no
/// <c>w:p</c> at all, and a last paragraph carrying no <c>w:pPr</c>. Both meet
/// <c>MoveLastSectPrIntoLastParagraph</c>, which used to dereference a null paragraph.
/// </summary>
public class WmlComparerEmptyBodyTests
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
                    new XElement(W.body, bodyChildren,
                        new XElement(W.sectPr,
                            new XElement(W.pgSz,
                                new XAttribute(W._w, "12240"), new XAttribute(W.h, "15840")))))));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(
                new XDocument(new XElement(W.styles, new XAttribute(XNamespace.Xmlns + "w", W.w))));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(
                new XDocument(new XElement(W.settings, new XAttribute(XNamespace.Xmlns + "w", W.w))));
        }
        return new WmlDocument("t.docx", ms.ToArray());
    }

    /// <summary>A paragraph with NO w:pPr — the shape that exercises the pPr-creation branch.</summary>
    private static XElement Para(string text) =>
        new(W.p, new XElement(W.r, new XElement(W.t, text)));

    private static WmlComparerSettings Settings() =>
        new() { SimplifyMoveMarkup = true, DateTimeForRevisions = "2000-01-01T00:00:00Z" };

    private static XElement Body(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    /// <summary>
    /// EB001 — a body stripped of every paragraph, as either side.
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EB001_BodyWithNoParagraphs_DoesNotThrow(bool emptyOnTheLeft)
    {
        var populated = Doc(Para("alpha"), Para("bravo"));
        var empty = Doc();

        var (left, right) = emptyOnTheLeft ? (empty, populated) : (populated, empty);

        var result = WmlComparer.Compare(left, right, Settings());

        // The populated side's content is accounted for, and the section stays a body child.
        var revisions = WmlComparer.GetRevisions(result, Settings());
        var expected = emptyOnTheLeft
            ? WmlComparer.WmlComparerRevisionType.Inserted
            : WmlComparer.WmlComparerRevisionType.Deleted;
        Assert.All(revisions, r => Assert.Equal(expected, r.RevisionType));
        Assert.Contains("alpha", string.Concat(revisions.Select(r => r.Text)));
        Assert.Single(Body(result).Elements(W.sectPr));
    }

    /// <summary>
    /// EB002 — both sides paragraph-less. The two packages must not be byte-identical, or
    /// <see cref="DocxCompare"/>'s exact-no-op shortcut would return a clone without running the
    /// comparison at all; a differing sectPr keeps the pipeline engaged.
    /// </summary>
    [Fact]
    public void EB002_BothBodiesEmpty_DoNotThrow()
    {
        var left = Doc();
        var right = DocWithPageWidth("12242");
        Assert.NotEqual(left.DocumentByteArray, right.DocumentByteArray);

        var result = WmlComparer.Compare(left, right, Settings());

        Assert.Empty(WmlComparer.GetRevisions(result, Settings()));
        Assert.Empty(Body(result).Elements(W.p));
    }

    private static WmlDocument DocWithPageWidth(string width)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.PutXDocument(new XDocument(
                new XElement(W.document,
                    new XAttribute(XNamespace.Xmlns + "w", W.w),
                    new XElement(W.body,
                        new XElement(W.sectPr,
                            new XElement(W.pgSz,
                                new XAttribute(W._w, width), new XAttribute(W.h, "15840")))))));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(
                new XDocument(new XElement(W.styles, new XAttribute(XNamespace.Xmlns + "w", W.w))));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(
                new XDocument(new XElement(W.settings, new XAttribute(XNamespace.Xmlns + "w", W.w))));
        }
        return new WmlDocument("t2.docx", ms.ToArray());
    }

    /// <summary>
    /// EB003 — the helper directly: with a last paragraph that has no <c>w:pPr</c>, the section must end
    /// up inside a REAL pPr element on that paragraph. Adding the XName instead left the pPr detached, so
    /// the section went with it and the name landed as a text node.
    /// </summary>
    [Fact]
    public void EB003_SectionMovesIntoACreatedPPr()
    {
        var body = new XElement(W.body,
            new XElement(W.p, new XElement(W.r, new XElement(W.t, "alpha"))),
            new XElement(W.sectPr, new XElement(W.pgSz, new XAttribute(W._w, "12240"))));

        WmlComparer.MoveLastSectPrIntoLastParagraph(body);

        var paragraph = Assert.Single(body.Elements(W.p));
        var pPr = paragraph.Element(W.pPr);
        Assert.NotNull(pPr);
        Assert.NotNull(pPr!.Element(W.sectPr));
        Assert.Empty(body.Elements(W.sectPr));
        Assert.Empty(paragraph.Nodes().OfType<XText>());
    }

    /// <summary>
    /// EB004 — a body whose only block content is a TABLE: the early return does not fire, and the
    /// section is moved into the table's last cell paragraph rather than crashing.
    /// </summary>
    [Fact]
    public void EB004_BodyWithOnlyATable_DoesNotThrow()
    {
        var table = new XElement(W.tbl,
            new XElement(W.tblPr,
                new XElement(W.tblW, new XAttribute(W._w, "0"), new XAttribute(W.type, "auto"))),
            new XElement(W.tblGrid, new XElement(W.gridCol, new XAttribute(W._w, "3000"))),
            new XElement(W.tr, new XElement(W.tc, Para("cell"))));

        var result = WmlComparer.Compare(Doc(table), Doc(table, Para("added")), Settings());

        Assert.NotNull(result);
        Assert.Single(Body(result).Elements(W.tbl));
    }
}
