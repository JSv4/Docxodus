// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Wp = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

public class FieldRetrieverTests
{
    [Fact]
    public void ReannotationRefreshesIdsAndInstructionsAfterAFieldSubtreeChanges()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        XElement Field(string instruction) => new XElement(W.p,
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"))),
            new XElement(W.r, new XElement(W.instrText, instruction)),
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate"))),
            new XElement(W.r, new XElement(W.t, "Stored result")),
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))));
        main.PutXDocument(new XDocument(new XElement(W.document, new XElement(W.body, Field("TITLE"), Field("DATE")))));
        main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles)));
        main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
        FieldRetriever.AnnotateWithFieldInfo(main);
        var root = main.GetXDocument().Root!;
        root.Descendants(W.p).First().Remove();
        root.Descendants(W.instrText).Single().ReplaceWith(new XElement(W.instrText, "SEQ Heading"));

        FieldRetriever.AnnotateWithFieldInfo(main);

        Assert.Equal("{SEQ Heading}", FieldRetriever.InstrText(root, 1));
        Assert.True(FieldRetriever.IsFieldResult(root.Descendants(W.t).Single(), "SEQ"));
    }

    [Fact]
    public void IsFieldResultTracksTocAcrossParagraphsWithoutLeakingPastEnd()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(
            stream,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.Document = new Wp.Document(
            new Wp.Body(
                new Wp.Paragraph(
                    new Wp.Run(new Wp.FieldChar { FieldCharType = Wp.FieldCharValues.Begin }),
                    new Wp.Run(new Wp.FieldCode(" TOC \\o \"1-3\" \\h ")),
                    new Wp.Run(new Wp.FieldChar { FieldCharType = Wp.FieldCharValues.Separate }),
                    new Wp.Hyperlink(new Wp.Run(new Wp.Text("First cached entry")))
                    {
                        Anchor = "_Toc1",
                    }),
                new Wp.Paragraph(
                    new Wp.Hyperlink(
                        new Wp.Run(new Wp.Text("Second cached entry")),
                        new Wp.Run(new Wp.FieldChar { FieldCharType = Wp.FieldCharValues.Begin }),
                        new Wp.Run(new Wp.FieldCode(" PAGEREF _Toc2 \\h ")),
                        new Wp.Run(new Wp.FieldChar { FieldCharType = Wp.FieldCharValues.Separate }),
                        new Wp.Run(new Wp.Text("2")),
                        new Wp.Run(new Wp.FieldChar { FieldCharType = Wp.FieldCharValues.End }))
                    {
                        Anchor = "_Toc2",
                    },
                    new Wp.Run(new Wp.FieldChar { FieldCharType = Wp.FieldCharValues.End })),
                new Wp.Paragraph(
                    new Wp.Hyperlink(new Wp.Run(new Wp.Text("Ordinary link")))
                    {
                        Anchor = "_Ordinary",
                    })));
        main.Document.Save();

        FieldRetriever.AnnotateWithFieldInfo(main);
        var root = main.GetXDocument().Root!;
        var firstCachedRun = root.Descendants(W.r)
            .Single(run => run.Descendants(W.t).Any(text => text.Value == "First cached entry"));
        var secondCachedRun = root.Descendants(W.r)
            .Single(run => run.Descendants(W.t).Any(text => text.Value == "Second cached entry"));
        var nestedPageReferenceRun = root.Descendants(W.r)
            .Single(run => run.Descendants(W.t).Any(text => text.Value == "2"));
        var ordinaryRun = root.Descendants(W.r)
            .Single(run => run.Descendants(W.t).Any(text => text.Value == "Ordinary link"));

        var fieldResult = firstCachedRun.Annotation<Stack<FieldRetriever.FieldElementTypeInfo>>()!
            .Single(info => info.FieldElementType == FieldRetriever.FieldElementTypeEnum.Result);
        Assert.Equal("{ TOC \\o \"1-3\" \\h }", FieldRetriever.InstrText(root, fieldResult.Id));
        Assert.True(FieldRetriever.IsFieldResult(firstCachedRun, "TOC"));
        Assert.True(FieldRetriever.IsFieldResult(secondCachedRun, "TOC"));
        Assert.True(FieldRetriever.IsFieldResult(nestedPageReferenceRun, "TOC"));
        Assert.True(FieldRetriever.IsFieldResult(nestedPageReferenceRun, "PAGEREF"));
        Assert.False(FieldRetriever.IsFieldResult(ordinaryRun, "TOC"));
    }
}
