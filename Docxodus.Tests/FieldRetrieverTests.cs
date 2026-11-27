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

    // A simple field's content is its result (issue #1870): a run in a w:fldSimple is in that field's result, as a run
    // in a complex field's result is, and both contexts count when one holds the other.
    [Fact]
    public void IsFieldResultCountsASimpleFieldsContentAsItsResult()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        XElement Run(string text) => new XElement(W.r, new XElement(W.t, text));
        XElement Char(string type) => new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, type)));
        XElement Simple(string instruction, XElement content) => new XElement(W.fldSimple, new XAttribute(W.instr, instruction), content);
        main.PutXDocument(new XDocument(new XElement(W.document, new XElement(W.body,
            new XElement(W.p, Simple(" hyperlink \"https://example.test/\" ", Run("Alone"))),
            new XElement(W.p, Char("begin"), new XElement(W.r, new XElement(W.instrText, " TOC \\o ")), Char("separate"),
                Simple(" HYPERLINK \\l \"_Toc1\" ", Run("Entry")), Run("Plain"), Char("end"))))));
        main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles)));
        main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
        FieldRetriever.AnnotateWithFieldInfo(main);
        XElement Text(string text) => main.GetXDocument().Root!.Descendants(W.t).Single(t => t.Value == text);

        Assert.True(FieldRetriever.IsFieldResult(Text("Alone"), "HYPERLINK"));
        Assert.False(FieldRetriever.IsFieldResult(Text("Alone"), "TOC"));
        Assert.True(FieldRetriever.IsFieldResult(Text("Entry"), "HYPERLINK"));
        Assert.True(FieldRetriever.IsFieldResult(Text("Entry"), "TOC"));
        Assert.False(FieldRetriever.IsFieldResult(Text("Plain"), "HYPERLINK"));
        Assert.True(FieldRetriever.IsFieldResult(Text("Plain"), "TOC"));
        // The field element itself is not in its own result.
        Assert.False(FieldRetriever.IsFieldResult(Text("Alone").Ancestors(W.fldSimple).Single(), "HYPERLINK"));
    }

    // Each field's type is read once per annotation; annotating again reads it again.
    [Fact]
    public void IsFieldResultReadsAFieldsTypeAgainAfterReannotation()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        main.PutXDocument(new XDocument(new XElement(W.document, new XElement(W.body, new XElement(W.p,
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"))),
            new XElement(W.r, new XElement(W.instrText, " TITLE ")),
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate"))),
            new XElement(W.r, new XElement(W.t, "Stored result")),
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))))))));
        main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles)));
        main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
        FieldRetriever.AnnotateWithFieldInfo(main);
        var root = main.GetXDocument().Root!;
        Assert.True(FieldRetriever.IsFieldResult(root.Descendants(W.t).Single(), "TITLE"));

        root.Descendants(W.instrText).Single().Value = " SEQ Figure ";
        FieldRetriever.AnnotateWithFieldInfo(main);

        Assert.True(FieldRetriever.IsFieldResult(root.Descendants(W.t).Single(), "SEQ"));
        Assert.False(FieldRetriever.IsFieldResult(root.Descendants(W.t).Single(), "TITLE"));
    }

    // A field's type is its code's first word, read without building the whole code.
    [Fact]
    public void FieldTypeReadsTheFirstWordOfTheCodeAsInstrTextHasIt()
    {
        using var stream = new MemoryStream();
        using var document = WordprocessingDocument.Create(stream,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        var main = document.AddMainDocumentPart();
        XElement Char(string type) => new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, type)));
        XElement Code(string code) => new XElement(W.r, new XElement(W.instrText, new XAttribute(XNamespace.Xml + "space", "preserve"), code));
        XElement Result(string text) => new XElement(W.r, new XElement(W.t, text));
        XElement[] Field(params XElement[] code) => new[] { Char("begin") }.Concat(code)
            .Concat(new[] { Char("separate"), Result("Shown"), Char("end") }).ToArray();
        main.PutXDocument(new XDocument(new XElement(W.document, new XElement(W.body,
            new XElement(W.p, Field(Code("  HYPER"), Code("LINK \"https://example.test/\" "))),
            new XElement(W.p, Field(new[] { Code(" IF ") }.Concat(Field(Code(" PAGE "))).Concat(new[] { Code(" > 1 \"a\" \"b\" ") }).ToArray())),
            new XElement(W.p, Field(new[] { Code(" ") }.Concat(Field(Code(" REF Name "))).Concat(new[] { Code(" ") }).ToArray())),
            new XElement(W.p, new XElement(W.del, new XAttribute(W.id, "1"), new XAttribute(W.author, "A"),
                Char("begin"), new XElement(W.r, new XElement(W.delInstrText, " TOC \\o ")), Char("separate"),
                new XElement(W.r, new XElement(W.delText, "Old")), Char("end")))))));
        main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles)));
        main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
        FieldRetriever.AnnotateWithFieldInfo(main);
        var root = main.GetXDocument().Root!;

        var types = Enumerable.Range(1, 6).Select(id => FieldRetriever.FieldType(root, id)).ToList();

        // A word split over two runs; a field whose code holds another; that nested field; a code that starts with a
        // nested field, so has no word of its own; the nested field; and a deleted field, whose code is w:delInstrText.
        Assert.Equal(new[] { "HYPERLINK", "IF", "PAGE", "", "REF", "TOC" }, types);
    }
}
