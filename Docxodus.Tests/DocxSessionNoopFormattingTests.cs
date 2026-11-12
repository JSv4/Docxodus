// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionNoopFormattingTests
{
    private static readonly XNamespace W = IrTestDocuments.W;
    private const string Prior = "<w:r><w:rPr><w:u w:val=\"single\"/>" +
        "<w:rPrChange w:id=\"9\" w:author=\"Counsel\" w:date=\"2026-01-01T00:00:00Z\"><w:rPr/></w:rPrChange>" +
        "</w:rPr><w:t>Earlier formatting</w:t></w:r>";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FormattingAlreadySatisfiedInsidePriorRevisionDoesNotSplitOrRenumberIt(bool tracked)
    {
        var input = IrTestDocuments.FromBodyXml("<w:p>" + Prior + "</w:p>").DocumentByteArray;
        using var session = new DocxSession(input, new DocxSessionSettings {
            TrackedChanges = tracked ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept,
            RevisionAuthor = "Reviewer"
        });
        var anchor = session.FindByText("Earlier formatting")!.Anchor.Id;
        var result = session.ApplyFormat(anchor, new CharSpan(2, 5), new FormatOp { Underline = true });
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(0, session.UndoCount);
        using var saved = WordprocessingDocument.Open(new MemoryStream(session.Save()), false);
        var revision = Assert.Single(saved.MainDocumentPart!.GetXDocument().Descendants(W + "rPrChange"));
        Assert.Equal("9", (string?)revision.Attribute(W + "id"));
        Assert.Single(saved.MainDocumentPart.GetXDocument().Descendants(W + "r"));
    }
    [Fact]
    public void ANoopBesideARealChangeKeepsTheEarlierRevisionIntact()
    {
        var input = IrTestDocuments.FromBodyXml("<w:p>" + Prior +
            "<w:r><w:t xml:space=\"preserve\"> and later text</w:t></w:r></w:p>").DocumentByteArray;
        using var session = new DocxSession(input, new DocxSessionSettings {
            TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = "Reviewer"
        });
        var anchor = session.FindByText("Earlier formatting")!.Anchor.Id;
        Assert.True(session.ApplyFormat(anchor, new CharSpan(2, 5), new FormatOp { Underline = true }).Success);
        Assert.True(session.ApplyFormat(anchor, new CharSpan(23, 5), new FormatOp { Underline = true }).Success);
        using var saved = WordprocessingDocument.Open(new MemoryStream(session.Save()), false);
        var prior = Assert.Single(saved.MainDocumentPart!.GetXDocument().Descendants(W + "rPrChange"),
            item => (string?)item.Attribute(W + "author") == "Counsel");
        Assert.Equal("9", (string?)prior.Attribute(W + "id"));
        Assert.Equal("Earlier formatting", prior.Parent!.Parent!.Element(W + "t")!.Value);
        Assert.Single(saved.MainDocumentPart.GetXDocument().Descendants(W + "rPrChange"),
            item => (string?)item.Attribute(W + "author") == "Reviewer");
    }
}
