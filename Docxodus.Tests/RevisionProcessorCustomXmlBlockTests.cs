// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Accepting or rejecting revisions keeps a block-level <c>w:customXml</c> wrapper (issue #913).
/// The deleted-paragraph-mark pass rebuilt each block container from the paragraphs it found
/// inside wrappers and reattached only <c>w:sdt</c> controls, so every block custom-XML wrapper
/// was dropped — with its <c>w:customXmlPr</c> — even in a document with no revisions at all.
/// </summary>
public class RevisionProcessorCustomXmlBlockTests
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Stamp = "w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\"";
    private const string ClauseProps =
        "<w:customXmlPr><w:attr w:name=\"id\" w:val=\"c1\"/></w:customXmlPr>";

    [Fact]
    public void Accept_BlockCustomXmlWithoutRevisions_SurvivesUnchanged()
    {
        var input =
            $"<w:customXml w:uri=\"urn:example\" w:element=\"clause\">{ClauseProps}" +
            "<w:p><w:r><w:t>body cx</w:t></w:r></w:p></w:customXml>" +
            "<w:p><w:r><w:t>after</w:t></w:r></w:p>";

        var body = Accept(input);

        var wrapper = Assert.Single(body.Elements(W.customXml));
        Assert.Equal("clause", (string?)wrapper.Attribute(W.element));
        Assert.Equal("urn:example", (string?)wrapper.Attribute(W.uri));
        Assert.Equal("c1", (string?)wrapper.Element(W.customXmlPr)?.Element(W.attr)?.Attribute(W.val));
        Assert.Equal(W.customXmlPr, wrapper.Elements().First().Name);
        Assert.Equal("body cx", Assert.Single(wrapper.Elements(W.p)).Value);
        Assert.Equal(new[] { "clause", "after" }, body.Elements().Where(e => e.Name != W.sectPr)
            .Select(e => e.Name == W.customXml ? (string)e.Attribute(W.element)! : e.Value));
    }

    [Fact]
    public void Accept_KeepsCustomXmlInCellsAndNestedWithContentControls()
    {
        var input =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr><w:tc>" +
            "<w:customXml w:element=\"cell-clause\"><w:p><w:r><w:t>cell cx</w:t></w:r></w:p></w:customXml>" +
            "</w:tc></w:tr></w:tbl>" +
            "<w:customXml w:element=\"outer\"><w:sdt><w:sdtPr><w:tag w:val=\"inner-control\"/></w:sdtPr>" +
            "<w:sdtContent><w:p><w:r><w:t>control in cx</w:t></w:r></w:p></w:sdtContent></w:sdt></w:customXml>" +
            "<w:sdt><w:sdtPr><w:tag w:val=\"outer-control\"/></w:sdtPr><w:sdtContent>" +
            "<w:customXml w:element=\"inner\"><w:p><w:r><w:t>cx in control</w:t></w:r></w:p></w:customXml>" +
            "</w:sdtContent></w:sdt>" +
            "<w:p><w:r><w:t>after</w:t></w:r></w:p>";

        var body = Accept(input);

        var cellWrapper = Assert.Single(body.Descendants(W.tc).Elements(W.customXml));
        Assert.Equal("cell cx", cellWrapper.Value);

        var outer = Assert.Single(body.Elements(W.customXml));
        Assert.Equal("outer", (string?)outer.Attribute(W.element));
        var innerControl = Assert.Single(outer.Elements(W.sdt));
        Assert.Equal("control in cx", innerControl.Value);

        var outerControl = Assert.Single(body.Elements(W.sdt));
        var inner = Assert.Single(outerControl.Elements(W.sdtContent).Elements(W.customXml));
        Assert.Equal("inner", (string?)inner.Attribute(W.element));
        Assert.Equal("cx in control", inner.Value);
    }

    [Fact]
    public void Accept_DeletedParagraphMarkInsideCustomXml_JoinsParagraphsInsideTheWrapper()
    {
        var input =
            "<w:customXml w:element=\"clause\">" +
            $"<w:p><w:pPr><w:rPr><w:del w:id=\"1\" {Stamp}/></w:rPr></w:pPr><w:r><w:t>first </w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>second</w:t></w:r></w:p>" +
            "</w:customXml>";

        var body = Accept(input);

        var wrapper = Assert.Single(body.Elements(W.customXml));
        Assert.Equal("first second", Assert.Single(wrapper.Elements(W.p)).Value);
    }

    [Fact]
    public void Accept_InsertedContentInsideCustomXml_KeepsTheWrapperAndTheText()
    {
        var input =
            "<w:customXml w:element=\"clause\">" +
            $"<w:p><w:ins w:id=\"1\" {Stamp}><w:r><w:t>added</w:t></w:r></w:ins>" +
            $"<w:del w:id=\"2\" {Stamp}><w:r><w:delText>removed</w:delText></w:r></w:del></w:p>" +
            "</w:customXml>";

        var body = Accept(input);

        var wrapper = Assert.Single(body.Elements(W.customXml));
        Assert.Equal("added", wrapper.Value);
        Assert.Empty(body.Descendants(W.ins));
        Assert.Empty(body.Descendants(W.del));
    }

    [Fact]
    public void Reject_KeepsTheWrapperAroundTheOriginalText()
    {
        var input =
            "<w:customXml w:element=\"clause\">" +
            $"<w:p><w:ins w:id=\"1\" {Stamp}><w:r><w:t>added</w:t></w:r></w:ins>" +
            $"<w:del w:id=\"2\" {Stamp}><w:r><w:delText>original</w:delText></w:r></w:del></w:p>" +
            "</w:customXml>";

        var body = BodyOf(RevisionProcessor.RejectRevisions(Build(input)));

        var wrapper = Assert.Single(body.Elements(W.customXml));
        Assert.Equal("original", wrapper.Value);
    }

    [Fact]
    public void Accept_DeletedCustomXmlEnvelope_IsStillRemoved()
    {
        // The tracked deletion of a whole wrapper (issue #764): the wrapper sits inside a
        // customXmlDelRange and its paragraph mark and runs are deleted. Accept removes it all.
        var input =
            "<w:p><w:r><w:t>before</w:t></w:r></w:p>" +
            $"<w:customXmlDelRangeStart w:id=\"10\" {Stamp}/>" +
            "<w:customXml w:element=\"clause\">" +
            "<w:customXmlDelRangeEnd w:id=\"10\"/>" +
            $"<w:p><w:pPr><w:rPr><w:del w:id=\"11\" {Stamp}/></w:rPr></w:pPr>" +
            $"<w:del w:id=\"12\" {Stamp}><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>" +
            $"<w:customXmlDelRangeStart w:id=\"13\" {Stamp}/>" +
            "</w:customXml>" +
            "<w:customXmlDelRangeEnd w:id=\"13\"/>" +
            "<w:p><w:r><w:t>after</w:t></w:r></w:p>";

        var body = Accept(input);

        Assert.Empty(body.Descendants(W.customXml));
        Assert.DoesNotContain("gone", body.Value);
        Assert.Equal(new[] { "before", "after" }, body.Elements(W.p).Select(p => p.Value));
    }

    [Fact]
    public void Accept_DeletedParagraphMarkBeforeTheWrapper_KeepsTheWrapperAsItsOwnBlock()
    {
        // The wrapper is a boundary, as a block content control is: the paragraph before it keeps
        // its text, and the wrapper is not swallowed into the deleted-mark group.
        var input =
            $"<w:p><w:pPr><w:rPr><w:del w:id=\"1\" {Stamp}/></w:rPr></w:pPr><w:r><w:t>outside</w:t></w:r></w:p>" +
            "<w:customXml w:element=\"clause\"><w:p><w:r><w:t>inside</w:t></w:r></w:p></w:customXml>";

        var body = Accept(input);

        Assert.Equal("inside", Assert.Single(body.Elements(W.customXml)).Value);
        Assert.Equal("outside", Assert.Single(body.Elements(W.p)).Value);
    }

    [Fact]
    public void Accept_DeletedMarkOnTheWrappersLastParagraph_KeepsTheParagraphAfterOutside()
    {
        var input =
            "<w:customXml w:element=\"clause\">" +
            $"<w:p><w:pPr><w:rPr><w:del w:id=\"1\" {Stamp}/></w:rPr></w:pPr><w:r><w:t>inside</w:t></w:r></w:p>" +
            "</w:customXml>" +
            "<w:p><w:r><w:t>outside</w:t></w:r></w:p>";

        var body = Accept(input);

        Assert.Equal("inside", Assert.Single(body.Elements(W.customXml)).Value);
        Assert.Equal("outside", Assert.Single(body.Elements(W.p)).Value);
    }

    private static XElement Accept(string bodyXml) =>
        BodyOf(RevisionProcessor.AcceptRevisions(Build(bodyXml)));

    private static WmlDocument Build(string bodyXml)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults());
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.Document.Save();
            var xDoc = main.GetXDocument();
            xDoc.Root!.Element(W.body)!.ReplaceWith(XElement.Parse($"<w:body xmlns:w=\"{WNs}\">{bodyXml}</w:body>"));
            main.PutXDocument();
        }

        return new WmlDocument("d.docx", stream.ToArray());
    }

    private static XElement BodyOf(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }
}
