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
        // IrReader swallows RetrieveListItem failures, so #814 surfaced only as a first-chance
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
}
