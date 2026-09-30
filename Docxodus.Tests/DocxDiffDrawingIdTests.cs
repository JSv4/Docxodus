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
/// from 1, so the output must renumber one <c>wp:docPr/@id</c>. VML text boxes from both documents
/// share a part: each brings the same shape id and its own copy of the shape type definition, and every
/// text box must still find a definition once the revisions are accepted or rejected.
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

    /// <summary>A paragraph holding one VML text box as Word writes it: the text box shape type is
    /// defined once per part, before the first text box, and every text box refers to it.</summary>
    private static string TextBox(string shapeId, string text, bool definesType = true) =>
        $"<w:p><w:r><w:pict{VmlNamespaces}>" +
        (definesType
            ? "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" " +
              "path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/>" +
              "<v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>"
            : string.Empty) +
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

    private const string Middle = "<w:p><w:r><w:t>Middle text stays</w:t></w:r></w:p>";

    /// <summary>Text box comparisons, each shape id numbered by its own document as Word does.</summary>
    public static TheoryData<string, string, string> TextBoxEdits => new()
    {
        {
            "replaced",
            Intro + TextBox("_x0000_s1026", "Before"),
            Intro + TextBox("_x0000_s1026", "After and longer")
        },
        {
            "replaced before an unchanged text box",
            Intro + TextBox("_x0000_s1026", "Alpha") + Middle + TextBox("_x0000_s1027", "Beta stays", definesType: false),
            Intro + TextBox("_x0000_s1026", "Gamma, entirely new") + Middle + TextBox("_x0000_s1027", "Beta stays", definesType: false)
        },
        {
            "replaced before a deleted text box",
            Intro + TextBox("_x0000_s1026", "Alpha") + Middle + TextBox("_x0000_s1027", "Beta goes", definesType: false),
            Intro + TextBox("_x0000_s1026", "Gamma, entirely new") + Middle
        },
        {
            "inserted before an existing text box",
            Intro + TextBox("_x0000_s1026", "Beta stays"),
            Intro + TextBox("_x0000_s1026", "New box") + Middle + TextBox("_x0000_s1027", "Beta stays", definesType: false)
        },
    };

    [Theory]
    [MemberData(nameof(TextBoxEdits))]
    public void Compare_TextBoxes_EveryShapeResolvesAsRedlinedAcceptedAndRejected(string edit, string original, string revised)
    {
        var redline = DocxCompare.Compare(IrTestDocuments.FromBodyXml(original), IrTestDocuments.FromBodyXml(revised));

        foreach (var (view, document) in new[]
                 {
                     ("redline", redline),
                     ("accepted", RevisionProcessor.AcceptRevisions(redline)),
                     ("rejected", RevisionProcessor.RejectRevisions(redline)),
                 })
        {
            var root = PartRoot(document, "word/document.xml");
            var definitions = root.Descendants(V + "shapetype").Select(d => (string)d.Attribute("id")!).ToArray();
            Assert.True(definitions.Length == definitions.Distinct().Count(), $"{edit}, {view}: duplicate shapetype ids");
            Assert.All(root.Descendants(V + "shape"), shape =>
                Assert.True(definitions.Contains(((string)shape.Attribute("type")!).TrimStart('#')),
                    $"{edit}, {view}: shape {(string?)shape.Attribute("id")} has no definition"));
        }
    }

    [Theory]
    [MemberData(nameof(TextBoxEdits))]
    public void Compare_TextBoxes_GetDistinctShapeIds(string edit, string original, string revised)
    {
        var root = PartRoot(
            DocxCompare.Compare(IrTestDocuments.FromBodyXml(original), IrTestDocuments.FromBodyXml(revised)),
            "word/document.xml");

        var ids = root.Descendants(V + "shape").Select(shape => (string)shape.Attribute("id")!).ToArray();
        Assert.True(ids.Length == ids.Distinct().Count(), $"{edit}: {string.Join(", ", ids)}");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Compare_DeletedAndInsertedDrawings_IntroduceNoValidatorError(bool textBoxes)
    {
        var original = textBoxes ? IrTestDocuments.FromBodyXml(Intro + TextBox("_x0000_s1026", "Before")) : Original;
        var revised = textBoxes ? IrTestDocuments.FromBodyXml(Intro + TextBox("_x0000_s1026", "After and longer")) : Revised;

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
