// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Tests.Ir;
using Xunit;
using static Docxodus.Tests.DocxBackendReconciliationTests;

namespace Docxodus.Tests;

/// <summary>
/// A paragraph the revised document inserts, carrying an anchored DrawingML text box, keeps its text-box
/// body as <c>w:txbxContent</c> in the main WordprocessingML namespace in the redline (issue #838) — both
/// as a body paragraph and inside a block-level content control, through the two-way and the consolidate
/// renderer. The original's part roots declare only <c>w:</c>, so every drawing namespace the inserted
/// shape uses arrives with content cloned out of the revised document.
/// </summary>
public class DocxDiffInsertedTextBoxTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

    private const string DrawingRootAttributes =
        " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"" +
        " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" +
        " xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"" +
        " xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\"";

    private const string BoxText = "Boxed text";

    private const string Existing = "<w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p>";

    private const string InsertedWithTextBox =
        "<w:p><w:r><w:drawing>" +
        "<wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"1\"" +
        " behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\"><wp:simplePos x=\"0\" y=\"0\"/>" +
        "<wp:positionH relativeFrom=\"column\"><wp:posOffset>0</wp:posOffset></wp:positionH>" +
        "<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>0</wp:posOffset></wp:positionV>" +
        "<wp:extent cx=\"1828800\" cy=\"457200\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>" +
        "<wp:wrapSquare wrapText=\"bothSides\"/><wp:docPr id=\"1\" name=\"Text Box 1\"/><wp:cNvGraphicFramePr/>" +
        "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
        "<wps:wsp><wps:cNvSpPr txBox=\"1\"/><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>" +
        "<a:ext cx=\"1828800\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></wps:spPr>" +
        $"<wps:txbx><w:txbxContent><w:p><w:r><w:t>{BoxText}</w:t></w:r></w:p></w:txbxContent></wps:txbx>" +
        "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>" +
        "<w:r><w:t>With a box.</w:t></w:r></w:p>";

    private const string ContentControl =
        "<w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr><w:sdtContent>{0}</w:sdtContent></w:sdt>";

    private static readonly WmlDocument Original = IrTestDocuments.FromParts(Existing);

    private static WmlDocument Revised(bool inContentControl) => IrTestDocuments.FromParts(
        Existing + (inContentControl ? string.Format(ContentControl, InsertedWithTextBox) : InsertedWithTextBox),
        rootAttributes: DrawingRootAttributes);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Compare_InsertedTextBox_KeepsMainNamespaceTextBoxContent(bool inContentControl) =>
        AssertTextBoxBody(DocxCompare.Compare(Original, Revised(inContentControl)));

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Consolidate_InsertedTextBox_KeepsMainNamespaceTextBoxContent(bool inContentControl) =>
        AssertTextBoxBody(Consolidate(Revised(inContentControl)));

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Compare_InsertedTextBox_AddsNoValidationErrors(bool inContentControl)
    {
        var revised = Revised(inContentControl);
        var output = DocxCompare.Compare(Original, revised).DocumentByteArray;

        NoNewValidationErrors(Original.DocumentByteArray, output);
        NoNewValidationErrors(revised.DocumentByteArray, output);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Consolidate_InsertedTextBox_AddsNoValidationErrors(bool inContentControl)
    {
        var revised = Revised(inContentControl);
        var output = Consolidate(revised).DocumentByteArray;

        NoNewValidationErrors(Original.DocumentByteArray, output);
        NoNewValidationErrors(revised.DocumentByteArray, output);
    }

    private static WmlDocument Consolidate(WmlDocument revised) =>
        DocxDiff.Consolidate(Original, new[] { new DocxDiffReviewer { Author = "Reviewer", Document = revised } });

    /// <summary>
    /// The inserted shape's <c>wps:txbx</c> holds exactly one child, <c>w:txbxContent</c> in the main
    /// namespace, whose text is the source box's text, and the shape itself sits inside a <c>w:ins</c>.
    /// </summary>
    private static void AssertTextBoxBody(WmlDocument redline)
    {
        var body = DocumentRoot(redline);
        var textBox = Assert.Single(body.Descendants(Wps + "txbx"));
        var content = Assert.Single(textBox.Elements());

        Assert.Equal(W + "txbxContent", content.Name);
        Assert.Equal(BoxText, string.Concat(content.Descendants(W + "t").Select(t => t.Value)));
        Assert.Contains(textBox.Ancestors(), ancestor => ancestor.Name == W + "ins");
        Assert.DoesNotContain(body.Descendants(), element => element.Name.LocalName == "txbxContent" && element.Name.Namespace != W);
    }

    /// <summary>The main document part's root as serialized, so its names are the ones a consumer reads.</summary>
    private static XElement DocumentRoot(WmlDocument document)
    {
        using var package = new ZipArchive(new MemoryStream(document.DocumentByteArray));
        using var part = package.GetEntry("word/document.xml")!.Open();
        return XElement.Load(part);
    }
}
