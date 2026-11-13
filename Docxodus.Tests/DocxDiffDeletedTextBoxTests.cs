// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Deleting a run that anchors a text box strikes the run, not the text inside the box (issue #1699). A text
/// box is its own story: the <c>w:del</c> around the anchoring run does not reach into
/// <c>w:txbxContent</c>, so a <c>w:delText</c> there has no deletion of its own and Word refuses the file.
/// Word itself leaves the box's text as <c>w:t</c> when it tracks the deletion of a text box.
/// </summary>
public class DocxDiffDeletedTextBoxTests
{
    private const string BoxText = "Rotate the filters weekly";

    private const string RootAttributes =
        " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"" +
        " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" +
        " xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"" +
        " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"" +
        " xmlns:v=\"urn:schemas-microsoft-com:vml\"";

    /// <summary>The box's story as Word writes it twice: once per <c>mc:AlternateContent</c> branch. It holds
    /// text and a field, so both spellings a deletion changes are present.</summary>
    private const string BoxStory =
        $"<w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">{BoxText} </w:t></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText>DATE</w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p></w:txbxContent>";

    private const string TextBoxRun =
        "<w:r><mc:AlternateContent><mc:Choice Requires=\"wps\"><w:drawing>" +
        "<wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"2\"" +
        " behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\"><wp:simplePos x=\"0\" y=\"0\"/>" +
        "<wp:positionH relativeFrom=\"column\"><wp:posOffset>0</wp:posOffset></wp:positionH>" +
        "<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>228600</wp:posOffset></wp:positionV>" +
        "<wp:extent cx=\"2286000\" cy=\"457200\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>" +
        "<wp:wrapSquare wrapText=\"bothSides\"/><wp:docPr id=\"1\" name=\"Callout\"/><wp:cNvGraphicFramePr/>" +
        "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
        "<wps:wsp><wps:cNvSpPr txBox=\"1\"/><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>" +
        "<a:ext cx=\"2286000\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></wps:spPr>" +
        $"<wps:txbx>{BoxStory}</wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor>" +
        "</w:drawing></mc:Choice><mc:Fallback><w:pict>" +
        "<v:shape id=\"Callout\" type=\"#_x0000_t202\" style=\"position:absolute;width:180pt;height:36pt\">" +
        $"<v:textbox>{BoxStory}</v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>";

    private const string Lead = "<w:r><w:t xml:space=\"preserve\">Maintenance plan. </w:t></w:r>";
    private const string Trailing = "<w:r><w:t>Reviewed yearly.</w:t></w:r>";

    private static WmlDocument Document(params string[] paragraphs) => IrTestDocuments.FromParts(
        string.Concat(paragraphs.Select(runs => $"<w:p>{runs}</w:p>")), rootAttributes: RootAttributes);

    private static readonly WmlDocument WithBox = Document(Lead + TextBoxRun);
    private static readonly WmlDocument WithoutBox = Document(Lead);

    public static TheoryData<string> Engines => new() { "DocxCompare", "DocxDiff" };

    private static WmlDocument Compare(string engine, WmlDocument left, WmlDocument right) =>
        engine == "DocxCompare" ? DocxCompare.Compare(left, right) : DocxDiff.Compare(left, right);

    [Theory]
    [MemberData(nameof(Engines))]
    public void DeletedTextBoxRun_StrikesTheRunAndLeavesTheBoxTextOrdinary(string engine)
    {
        var body = Body(Compare(engine, WithBox, WithoutBox));

        Assert.Empty(DeletedTextOutsideADeletion(body));
        var stories = body.Descendants(W.txbxContent).ToList();
        Assert.Equal(2, stories.Count);
        Assert.All(stories, story =>
        {
            Assert.Contains(story.Ancestors(), ancestor => ancestor.Name == W.del);
            Assert.Equal(BoxText + " ", string.Concat(story.Descendants(W.t).Select(t => t.Value)));
            Assert.Equal("DATE", string.Concat(story.Descendants(W.instrText).Select(t => t.Value)));
        });
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void DeletedTextBoxRun_RejectsToTheOriginalAndAcceptsToTheRevision(string engine)
    {
        var redline = Compare(engine, WithBox, WithoutBox);

        Assert.Equal(Body(WithBox).ToString(), Body(RevisionProcessor.RejectRevisions(redline)).ToString());
        Assert.Equal(Body(WithoutBox).ToString(), Body(RevisionProcessor.AcceptRevisions(redline)).ToString());
    }

    /// <summary>Control: a deletion beside a text box that stays is still spelled as deleted, and the box is
    /// not drawn into it.</summary>
    [Theory]
    [MemberData(nameof(Engines))]
    public void DeletedParagraphBesideATextBox_IsStillSpelledAsDeleted(string engine)
    {
        var body = Body(Compare(engine, Document(Lead + TextBoxRun, Trailing), Document(Lead + TextBoxRun)));

        var deleted = Assert.Single(body.Descendants(W.delText));
        Assert.Equal("Reviewed yearly.", deleted.Value);
        Assert.Contains(deleted.Ancestors(), ancestor => ancestor.Name == W.del);
        Assert.DoesNotContain(body.Descendants(W.txbxContent).Ancestors(), ancestor => ancestor.Name == W.del);
    }

    /// <summary>The shared rename, directly: a deletion spells as deleted only the text it owns. Text in a text
    /// box's story, and text under a revision nested in the deleted element, belong to someone else.</summary>
    [Fact]
    public void SpellTextAsDeleted_RenamesOnlyTheTextTheDeletionOwns()
    {
        var deleted = XElement.Parse(
            $"<w:hyperlink xmlns:w=\"{W.w}\" xmlns:v=\"urn:schemas-microsoft-com:vml\">" +
            "<w:r><w:t>own</w:t><w:instrText>OWN</w:instrText></w:r>" +
            "<w:ins><w:r><w:t>nested</w:t></w:r></w:ins>" +
            $"<w:r><w:pict><v:textbox>{BoxStory}</v:textbox></w:pict></w:r></w:hyperlink>");

        RevisionOps.SpellTextAsDeleted(deleted);

        Assert.Equal("own", Assert.Single(deleted.Descendants(W.delText)).Value);
        Assert.Equal("OWN", Assert.Single(deleted.Descendants(W.delInstrText)).Value);
        Assert.Equal(new[] { "nested", BoxText + " " }, deleted.Descendants(W.t).Select(t => t.Value));
        Assert.Equal("DATE", Assert.Single(deleted.Descendants(W.instrText)).Value);
    }

    /// <summary>Every <c>w:delText</c>/<c>w:delInstrText</c> with no <c>w:del</c> or <c>w:moveFrom</c> around it
    /// inside its own story — the markup Word reports as a corrupt file.</summary>
    private static IEnumerable<XElement> DeletedTextOutsideADeletion(XElement body) =>
        body.Descendants()
            .Where(e => e.Name == W.delText || e.Name == W.delInstrText)
            .Where(text => !text.Ancestors().TakeWhile(a => a.Name != W.txbxContent)
                .Any(a => a.Name == W.del || a.Name == W.moveFrom));

    private static XElement Body(WmlDocument document)
    {
        using var stream = new OpenXmlMemoryStreamDocument(document);
        using var package = stream.GetWordprocessingDocument();
        return package.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }
}
