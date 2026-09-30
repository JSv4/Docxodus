// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A comparison that keeps a drawing from each document — the original's deleted, the revised's
/// inserted — keeps the drawing object ids of both (issue #860). Both documents number their drawings
/// from 1, so the output must renumber one <c>wp:docPr/@id</c>, and two copies of one VML shapetype
/// definition must not leave two elements with the same <c>id</c> in one part.
/// </summary>
public class DocxDiffDrawingIdTests
{
    private const string DrawingNamespaces =
        " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"" +
        " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" +
        " xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"";

    private const string VmlNamespaces =
        " xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\"";

    private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";

    /// <summary>A paragraph holding one inline shape drawing, as Word numbers it.</summary>
    private static string Shape(int id, string name, string geometry) =>
        $"<w:p><w:r><w:drawing{DrawingNamespaces}><wp:inline><wp:extent cx=\"914400\" cy=\"914400\"/>" +
        $"<wp:docPr id=\"{id}\" name=\"{name}\"/><a:graphic><a:graphicData " +
        "uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:wsp><wps:cNvSpPr/>" +
        "<wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/></a:xfrm>" +
        $"<a:prstGeom prst=\"{geometry}\"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp>" +
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

    /// <summary>A paragraph holding one VML text box, with the text box shapetype Word writes before it.</summary>
    private static string TextBox(string shapeId, string text) =>
        $"<w:p><w:r><w:pict{VmlNamespaces}><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" " +
        "o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/>" +
        "<v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>" +
        $"<v:shape id=\"{shapeId}\" type=\"#_x0000_t202\" style=\"width:100pt;height:50pt\"><v:textbox>" +
        $"<w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>" +
        "</w:pict></w:r></w:p>";

    private const string Intro = "<w:p><w:r><w:t>Intro</w:t></w:r></w:p>";

    private static readonly WmlDocument Original = IrTestDocuments.FromBodyXml(Intro + Shape(1, "Rectangle 1", "rect"));
    private static readonly WmlDocument Revised = IrTestDocuments.FromBodyXml(Intro + Shape(1, "Oval 1", "ellipse"));

    [Fact]
    public void Compare_DeletedAndInsertedDrawings_GetDistinctDocPrIds()
    {
        var ids = DocPrIds(PartRoot(DocxCompare.Compare(Original, Revised), "word/document.xml"));

        Assert.Equal(2, ids.Length);
        Assert.Equal(ids.Length, ids.Distinct().Count());
    }

    [Fact]
    public void Compare_DeletedAndInsertedDrawingsInAHeader_GetDistinctDocPrIds()
    {
        var original = IrTestDocuments.FromBodyAndHeaderXml(Intro, Shape(1, "Rectangle 1", "rect"));
        var revised = IrTestDocuments.FromBodyAndHeaderXml(Intro, Shape(1, "Oval 1", "ellipse"));

        var ids = DocPrIds(PartRoot(DocxCompare.Compare(original, revised), "word/header1.xml"));

        Assert.Equal(2, ids.Length);
        Assert.Equal(ids.Length, ids.Distinct().Count());
    }

    [Fact]
    public void Consolidate_DrawingsFromBaseAndReviewers_GetDistinctDocPrIds()
    {
        var reviewers = new[]
        {
            new DocxDiffReviewer { Author = "Reviewer A", Document = Revised },
            new DocxDiffReviewer
            {
                Author = "Reviewer B",
                Document = IrTestDocuments.FromBodyXml(Intro + Shape(1, "Triangle 1", "triangle")),
            },
        };

        var ids = DocPrIds(PartRoot(DocxDiff.Consolidate(Original, reviewers), "word/document.xml"));

        Assert.True(ids.Length >= 2, $"expected drawings from more than one source, got {ids.Length}");
        Assert.Equal(ids.Length, ids.Distinct().Count());
    }

    [Fact]
    public void Compare_DeletedAndInsertedTextBoxes_LeaveOneShapetypePerIdAndEveryShapeResolves()
    {
        var output = PartRoot(
            DocxCompare.Compare(
                IrTestDocuments.FromBodyXml(Intro + TextBox("Old box", "Before")),
                IrTestDocuments.FromBodyXml(Intro + TextBox("New box", "After and longer"))),
            "word/document.xml");

        var definitions = output.Descendants(V + "shapetype").Select(d => (string)d.Attribute("id")!).ToArray();
        var shapes = output.Descendants(V + "shape").ToArray();

        Assert.Equal(2, shapes.Length);
        Assert.Equal(definitions.Length, definitions.Distinct().Count());
        Assert.All(shapes, shape => Assert.Contains(((string)shape.Attribute("type")!).TrimStart('#'), definitions));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Compare_DeletedAndInsertedTextBoxes_EveryShapeStillResolvesAfterAcceptingOrRejecting(bool accept)
    {
        var redline = DocxCompare.Compare(
            IrTestDocuments.FromBodyXml(Intro + TextBox("Old box", "Before")),
            IrTestDocuments.FromBodyXml(Intro + TextBox("New box", "After and longer")));

        var output = PartRoot(
            accept ? RevisionProcessor.AcceptRevisions(redline) : RevisionProcessor.RejectRevisions(redline),
            "word/document.xml");

        var definitions = output.Descendants(V + "shapetype").Select(d => (string)d.Attribute("id")!).ToArray();
        var shape = Assert.Single(output.Descendants(V + "shape"));
        Assert.Contains(((string)shape.Attribute("type")!).TrimStart('#'), definitions);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Compare_DeletedAndInsertedDrawings_IntroduceNoValidatorError(bool textBoxes)
    {
        var original = textBoxes ? IrTestDocuments.FromBodyXml(Intro + TextBox("Old box", "Before")) : Original;
        var revised = textBoxes ? IrTestDocuments.FromBodyXml(Intro + TextBox("New box", "After and longer")) : Revised;

        var output = DocxCompare.Compare(original, revised);

        Assert.Empty(ValidatorErrorIds(output).Except(ValidatorErrorIds(original)).Except(ValidatorErrorIds(revised)));
    }

    private static string[] ValidatorErrorIds(WmlDocument document)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(document.DocumentByteArray), false);
        return new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package)
            .Select(error => $"{error.Id}@{error.Node?.LocalName}")
            .ToArray();
    }

    private static string[] DocPrIds(XElement root) =>
        root.Descendants(Wp + "docPr").Select(docPr => (string)docPr.Attribute("id")!).ToArray();

    private static XElement PartRoot(WmlDocument document, string partName)
    {
        using var zip = new ZipArchive(new MemoryStream(document.DocumentByteArray));
        using var stream = zip.GetEntry(partName)!.Open();
        return XDocument.Load(stream).Root!;
    }
}
