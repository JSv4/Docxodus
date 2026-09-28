#nullable enable

// Regression tests for ListItemRetriever, found while removing #nullable disable
// from the file (issue #650).

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Xml.Linq;
using Docxodus;
using Docxodus.Ir;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

public class ListItemRetrieverTests
{
    /// <summary>
    /// Minimal numbering.xml with one abstractNum defining NO levels at all, and one num
    /// (numId=1) referencing it. ListItemSourceSet/ListItemSource's own internal per-level
    /// fallback (Main, then NumStyleLink) therefore exhausts without finding anything at
    /// any level, which is exactly the precondition for ListItemInfo.Lvl's own outer
    /// fallback loop to run.
    /// </summary>
    private static XDocument BuildNumberingWithNoLevels()
    {
        return new XDocument(
            new XElement(W.numbering,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XElement(W.abstractNum,
                    new XAttribute(W.abstractNumId, 1)),
                new XElement(W.num,
                    new XAttribute(W.numId, 1),
                    new XElement(W.abstractNumId, new XAttribute(W.val, 1)))));
    }

    [Fact]
    public void ListItemInfo_Lvl_StyleOnlySource_FallbackLoopDoesNotThrow()
    {
        // Regression test: ListItemInfo.Lvl's paragraph-less fallback loop used to walk
        // FromParagraph.Lvl(i) even when FromParagraph was null (a copy-paste bug from the
        // paragraph-sourced branch above it). ListItemSource.Lvl already exhausts every
        // level 0..ilvl internally (via Main then NumStyleLink) before returning null, so
        // this outer loop only runs once that inner search has already failed everywhere —
        // exactly the case here, with an abstractNum that defines no levels at all. Before
        // the fix, entering the loop threw a NullReferenceException on the first iteration
        // instead of returning null like every other "not found" path in this file.
        var numXDoc = BuildNumberingWithNoLevels();
        var styleSource = new ListItemRetriever.ListItemSource(numXDoc, numXDoc, numId: 1);

        var listItemInfo = new ListItemRetriever.ListItemInfo
        {
            FromStyle = styleSource,
            // FromParagraph left null: this is a style-only list item.
        };

        var lvl = listItemInfo.Lvl(1);

        Assert.Null(lvl);
    }

    [Fact]
    public void RetrieveListItem_CommentParagraph_IsNotAListItem()
    {
        // Regression test for #814: the comments part is not one of the content parts that
        // list numbering initializes, so a comment paragraph stayed unannotated and
        // RetrieveListItem dereferenced null.
        using var stream = new MemoryStream(BuildListInBodyAndComment());
        using var wordDoc = WordprocessingDocument.Open(stream, false);
        var main = wordDoc.MainDocumentPart!;
        var bodyListItem = main.GetXDocument().Descendants(W.p).Single();
        var commentParagraphs = main.WordprocessingCommentsPart!.GetXDocument().Descendants(W.p).ToList();

        // Neither comment paragraph (one with w:numPr, one without) gets a marker, while the
        // same numbering definition still numbers the body paragraph.
        Assert.Equal(2, commentParagraphs.Count);
        Assert.All(commentParagraphs, p => Assert.Null(ListItemRetriever.RetrieveListItem(wordDoc, p)));
        Assert.Equal("1.", ListItemRetriever.RetrieveListItem(wordDoc, bodyListItem));
    }

    [Fact]
    public void IrRead_DocumentWithComments_RaisesNoListItemRetrieverException()
    {
        // IrReader swallowed RetrieveListItem failures, so #814 surfaced only as a first-chance
        // exception per comment paragraph on every IR read.
        var thrown = new List<Exception>();
        var testThread = Environment.CurrentManagedThreadId;
        EventHandler<FirstChanceExceptionEventArgs> record = (_, e) =>
        {
            if (Environment.CurrentManagedThreadId == testThread
                && e.Exception.TargetSite?.DeclaringType == typeof(ListItemRetriever))
                thrown.Add(e.Exception);
        };

        AppDomain.CurrentDomain.FirstChanceException += record;
        try
        {
            IrReader.Read(new WmlDocument("comments.docx", BuildListInBodyAndComment()));
        }
        finally
        {
            AppDomain.CurrentDomain.FirstChanceException -= record;
        }

        Assert.Empty(thrown);
    }

    /// <summary>
    /// A document whose body holds one decimal list item, plus a comment holding a paragraph
    /// with the same <c>w:numPr</c> and a plain paragraph.
    /// </summary>
    private static byte[] BuildListInBodyAndComment()
    {
        static Paragraph ListItem(string text) => new(
            new ParagraphProperties(new NumberingProperties(
                new NumberingLevelReference { Val = 0 },
                new NumberingId { Val = 1 })),
            new Run(new Text(text)));

        using var stream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = wordDoc.AddMainDocumentPart();
            main.Document = new Document(new Body(ListItem("Body item")));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.AddNewPart<NumberingDefinitionsPart>().Numbering = new Numbering(
                new AbstractNum(new Level(
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." },
                    new StartNumberingValue { Val = 1 }) { LevelIndex = 0 }) { AbstractNumberId = 1 },
                new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 1 });
            main.AddNewPart<WordprocessingCommentsPart>().Comments = new Comments(
                new Comment(ListItem("Comment list item"), new Paragraph(new Run(new Text("Comment text"))))
                { Id = "0", Author = "Reviewer" });
        }
        return stream.ToArray();
    }

    private const string DecimalLevel = """<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>""";
    private const string ValidList1 = """<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0">""" + DecimalLevel + """</w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>""";
    private const string Num2UsesAbstract1 = """<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>""";
    private const string Num2UsesAbstract2 = """<w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>""";

    [Theory]
    // numId 2 points at an abstractNum that doesn't exist.
    [InlineData("""<w:num w:numId="2"><w:abstractNumId w:val="99"/></w:num>""", 0, null)]
    // No w:lvl is defined at or below the paragraph's level.
    [InlineData("""<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="3">""" + DecimalLevel + "</w:lvl></w:abstractNum>" + Num2UsesAbstract2, 0, null)]
    // The paragraph's level is past the last counter slot.
    [InlineData(Num2UsesAbstract1, 10, null)]
    // The only w:lvl has no w:ilvl attribute.
    [InlineData("""<w:abstractNum w:abstractNumId="2"><w:lvl>""" + DecimalLevel + "</w:lvl></w:abstractNum>" + Num2UsesAbstract2, 0, null)]
    // A non-integer w:start counts from the default of 0.
    [InlineData("""<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="abc"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>""" + Num2UsesAbstract2, 0, "0.")]
    // A negative counter can't be written in Roman numerals, so it falls back to decimal.
    [InlineData("""<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="-1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>""" + Num2UsesAbstract2, 0, "-1.")]
    // A %0 placeholder refers to no level and stays literal.
    [InlineData("""<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%0."/></w:lvl></w:abstractNum>""" + Num2UsesAbstract2, 0, "%0.")]
    // Level 1 looks like it continues level 0, but with no level 0 there is no format to continue with.
    [InlineData("""<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/></w:lvl></w:abstractNum>""" + Num2UsesAbstract2, 1, "1.")]
    public void RetrieveListItem_MalformedNumbering_DegradesWithoutThrowing(string malformedNumbering, int ilvl, string? expectedMarker)
    {
        // Regression test for #818: each of these threw from RetrieveListItem, which callers
        // either hid behind a broad catch (IR, markdown) or let escape (HTML conversion).
        var bytes = BuildList(ValidList1 + malformedNumbering, (NumId: 2, Ilvl: ilvl), (NumId: 1, Ilvl: 0));

        // The malformed list must not stop the valid one from being numbered.
        Assert.Equal(new[] { expectedMarker, "1." }, RetrieveMarkers(bytes));
        AssertConsumersRead(bytes);
    }

    [Fact]
    public void RetrieveListItem_ListParagraphAtUndefinedLevel_LosesOnlyItsOwnMarker()
    {
        // Both paragraphs share the style|numId cache entry, so the level check must run per paragraph.
        var bytes = BuildList(ValidList1, (NumId: 1, Ilvl: 0), (NumId: 1, Ilvl: 10));

        Assert.Equal(new[] { "1.", null }, RetrieveMarkers(bytes));
    }

    [Fact]
    public void RetrieveListItem_EmptyNumberingPart_IsNotAListItem()
    {
        var bytes = BuildList(numbering: null, (NumId: 1, Ilvl: 0));

        Assert.Equal(new string?[] { null }, RetrieveMarkers(bytes));
        AssertConsumersRead(bytes);
    }

    private static string?[] RetrieveMarkers(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var wordDoc = WordprocessingDocument.Open(stream, false);
        return wordDoc.MainDocumentPart!.GetXDocument().Descendants(W.p)
            .Select(p => ListItemRetriever.RetrieveListItem(wordDoc, p))
            .ToArray();
    }

    /// <summary>The IR reader, markdown projection and HTML converter call the retriever with no catch around it.</summary>
    private static void AssertConsumersRead(byte[] bytes)
    {
        IrReader.Read(new WmlDocument("list.docx", bytes));
        WmlToMarkdownConverter.Convert(new WmlDocument("list.docx", bytes), new WmlToMarkdownConverterSettings());
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("list.docx", bytes), new WmlToHtmlConverterSettings());
    }

    /// <summary>
    /// A body of list paragraphs, one per (numId, ilvl) item, over a numbering part holding
    /// <paramref name="numbering"/> (or an empty numbering part when null).
    /// </summary>
    private static byte[] BuildList(string? numbering, params (int NumId, int Ilvl)[] items)
    {
        const string Ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        var body = string.Concat(items.Select(item =>
            $"""<w:p><w:pPr><w:numPr><w:ilvl w:val="{item.Ilvl}"/><w:numId w:val="{item.NumId}"/></w:numPr></w:pPr><w:r><w:t>item</w:t></w:r></w:p>"""));

        using var stream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = wordDoc.AddMainDocumentPart();
            main.PutXDocument(XDocument.Parse($"<w:document {Ns}><w:body>{body}</w:body></w:document>"));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse($"<w:styles {Ns}/>"));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings {Ns}/>"));
            var numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
            if (numbering == null)
                numberingPart.FeedData(new MemoryStream());
            else
                numberingPart.PutXDocument(XDocument.Parse($"<w:numbering {Ns}>{numbering}</w:numbering>"));
        }
        return stream.ToArray();
    }

    [Theory]
    // A non-integer numId reads as a missing one: with no style numbering, not a list item.
    [InlineData("abc", "0", null)]
    // A non-integer ilvl reads as level 0, as ListItemRetriever reads it.
    [InlineData("1", "abc", "1.")]
    public void MarkdownAndIr_NonIntegerNumPrValue_ReadsLikeListItemRetriever(string numId, string ilvl, string? expectedMarker)
    {
        // Regression test for #820: WmlToMarkdownConverter cast these values with (int?), so the
        // markdown projection and the IR read (via IsListItemForLayout) threw FormatException.
        var bytes = WithNumPrValues(BuildList(ValidList1, (NumId: 1, Ilvl: 0)), numId, ilvl);
        Assert.Equal(new[] { expectedMarker }, RetrieveMarkers(bytes));

        var paragraph = IrReader.Read(new WmlDocument("list.docx", bytes)).Body.Blocks.OfType<IrParagraph>().Single();
        Assert.Equal(expectedMarker != null, paragraph.IsListItemForLayout);
        Assert.Equal(expectedMarker != null ? 0 : null, paragraph.List?.Ilvl);

        var markdown = WmlToMarkdownConverter.Convert(new WmlDocument("list.docx", bytes), new WmlToMarkdownConverterSettings()).Markdown;
        Assert.Equal(expectedMarker != null, markdown.Contains("1. ", StringComparison.Ordinal));
        Assert.Contains("item", markdown, StringComparison.Ordinal);
    }

    /// <summary>Overwrites the first paragraph's <c>w:numId</c> and <c>w:ilvl</c> values with arbitrary strings.</summary>
    private static byte[] WithNumPrValues(byte[] bytes, string numId, string ilvl)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var wordDoc = WordprocessingDocument.Open(stream, true))
        {
            var numPr = wordDoc.MainDocumentPart!.GetXDocument().Descendants(W.numPr).First();
            numPr.Element(W.numId)!.SetAttributeValue(W.val, numId);
            numPr.Element(W.ilvl)!.SetAttributeValue(W.val, ilvl);
            wordDoc.MainDocumentPart.PutXDocument();
        }
        return stream.ToArray();
    }
}
