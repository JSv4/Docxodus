// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionFieldDeletionTests
{
    private static readonly XNamespace W = IrTestDocuments.W;
    private const string Field = "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
        "<w:r><w:instrText xml:space=\"preserve\"> REF heading </w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" +
        "<w:r><w:t>Section 4</w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";
    private static byte[] Source(string field = Field) => IrTestDocuments.FromBodyXml(
        "<w:p><w:r><w:t xml:space=\"preserve\">Before </w:t></w:r>" + field +
        "<w:r><w:t xml:space=\"preserve\"> after.</w:t></w:r></w:p>").DocumentByteArray;

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DeletesCompleteFieldAsOneOperation(bool tracked)
    {
        using var session = new DocxSession(Source(), new DocxSessionSettings {
            TrackedChanges = tracked ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept,
            RevisionAuthor = "Reviewer"
        });
        var anchor = session.FindByText("Section 4")!.Anchor.Id;
        var edit = session.DeleteFieldAtSpan(anchor, 7, 9);
        Assert.True(edit.Success, edit.Error?.Message);
        var content = session.Save();
        if (tracked)
        {
            using var saved = WordprocessingDocument.Open(new MemoryStream(content), false);
            Assert.Empty(new OpenXmlValidator().Validate(saved));
            var deletion = Assert.Single(saved.MainDocumentPart!.GetXDocument().Descendants(W + "del"));
            Assert.Single(deletion.Descendants(W + "delInstrText"));
            Assert.Single(deletion.Descendants(W + "delText"));
            var report = DeliverableVerifier.VerifyDeliverable(Source(), content);
            Assert.DoesNotContain(report.Findings, finding => finding.Code == "structure.field_instruction_missing");
        }
        using var accepted = new DocxSession(content);
        if (tracked) Assert.True(accepted.AcceptAllRevisions().Success);
        Assert.Single(accepted.Grep("^Before  after\\.$"));
        using var rejected = new DocxSession(content);
        if (tracked) {
            Assert.True(rejected.RejectAllRevisions().Success);
            Assert.Single(rejected.Grep("^Before Section 4 after\\.$"));
            using var saved = WordprocessingDocument.Open(new MemoryStream(rejected.Save()), false);
            Assert.Single(saved.MainDocumentPart!.GetXDocument().Descendants(W + "instrText"));
        }
    }

    [Theory]
    [InlineData("mixed-marker")]
    [InlineData("nested-field")]
    [InlineData("field-in-wrapper")]
    public void RefusesAmbiguousFieldStructuresWithoutMutation(string shape)
    {
        var field = shape switch
        {
            "mixed-marker" => Field.Replace("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>",
                "<w:r><w:fldChar w:fldCharType=\"begin\"/><w:t>x</w:t></w:r>"),
            "nested-field" => Field.Replace("<w:r><w:t>Section 4</w:t></w:r>",
                "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:t>Section 4</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"),
            _ => "<w:hyperlink w:anchor=\"heading\">" + Field + "</w:hyperlink>",
        };
        using var session = new DocxSession(Source(field));
        var anchor = session.FindByText("Section 4")!.Anchor.Id;
        var before = session.GetAnchorInfo(anchor)!.VisibleText;
        Assert.False(session.DeleteFieldAtSpan(anchor, 7, 9).Success);
        Assert.Equal(before, session.GetAnchorInfo(anchor)!.VisibleText);
        Assert.Equal(0, session.UndoCount);
    }

    [Fact]
    public void RefusesAFieldInsideAnotherFieldThatStartedInThePriorParagraph()
    {
        var doc = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
            "<w:r><w:instrText> TOC </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r></w:p>" +
            "<w:p><w:r><w:t xml:space=\"preserve\">Before </w:t></w:r>" + Field +
            "<w:r><w:t xml:space=\"preserve\"> after.</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>");
        using var session = new DocxSession(doc.DocumentByteArray);
        var anchor = session.FindByText("Section 4")!.Anchor.Id;
        Assert.False(session.DeleteFieldAtSpan(anchor, 7, 9).Success);
    }

    [Fact]
    public void RefusesFieldCarryingInheritedFormattingRevision()
    {
        var revised = Field.Replace("<w:r><w:t>Section 4</w:t></w:r>",
            "<w:r><w:rPr><w:b/><w:rPrChange w:id=\"91\" w:author=\"Counsel\" w:date=\"2026-01-01T00:00:00Z\"><w:rPr/></w:rPrChange></w:rPr><w:t>Section 4</w:t></w:r>");
        using var session = new DocxSession(Source(revised));
        var anchor = session.FindByText("Section 4")!.Anchor.Id;
        Assert.False(session.DeleteFieldAtSpan(anchor, 7, 9).Success);
        Assert.Single(session.Grep("^Before Section 4 after\\.$"));
    }

    [Fact]
    public void RefusesPartialFieldAndLeavesSourceIntact()
    {
        using var session = new DocxSession(Source());
        var anchor = session.FindByText("Section 4")!.Anchor.Id;
        var edit = session.DeleteFieldAtSpan(anchor, 7, 7);
        Assert.False(edit.Success);
        Assert.Single(session.Grep("^Before Section 4 after\\.$"));
    }
}
