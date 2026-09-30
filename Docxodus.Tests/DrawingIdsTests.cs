// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="DrawingIds"/> on part roots: which drawing and VML ids it renumbers, and how it leaves
/// every shape a <c>v:shapetype</c> definition in every view of the document (issue #860).
/// </summary>
public class DrawingIdsTests
{
    private const string Namespaces =
        " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"" +
        " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"" +
        " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"" +
        " xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\"" +
        " xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"";

    private static XElement Part(string content) => XElement.Parse($"<w:body{Namespaces}>{content}</w:body>");

    private static string DocPr(int id) => $"<wp:docPr id=\"{id}\" name=\"Shape {id}\"/>";

    private static string[] DocPrIds(params XElement[] roots) =>
        roots.SelectMany(root => root.Descendants(WP.docPr)).Select(docPr => (string)docPr.Attribute("id")!).ToArray();

    [Fact]
    public void RenumberDocPrIds_KeepsTheFirstAndMovesLaterDuplicatesAboveTheHighestId()
    {
        var body = Part($"<w:del>{DocPr(1)}</w:del><w:ins>{DocPr(1)}</w:ins>{DocPr(4)}");

        DrawingIds.RenumberDocPrIds(new[] { body });

        Assert.Equal(new[] { "1", "5", "4" }, DocPrIds(body));
    }

    [Fact]
    public void RenumberDocPrIds_LeavesDistinctIdsAndThePartsHoldingThemUnchanged()
    {
        var body = Part(DocPr(1) + DocPr(2));

        Assert.Empty(DrawingIds.RenumberDocPrIds(new[] { body }));
        Assert.Equal(new[] { "1", "2" }, DocPrIds(body));
    }

    [Fact]
    public void RenumberDocPrIds_TreatsIdsAsUniqueAcrossStoryParts()
    {
        var body = Part(DocPr(1));
        var header = Part(DocPr(1));

        var changed = DrawingIds.RenumberDocPrIds(new[] { body, header });

        Assert.Equal(new[] { "1", "2" }, DocPrIds(body, header));
        Assert.Equal(new[] { header }, changed);
    }

    [Fact]
    public void RenumberDocPrIds_GivesTheChoiceAndFallbackCopiesOfOneDrawingOneId()
    {
        var alternative = $"<mc:AlternateContent><mc:Choice Requires=\"wps\">{DocPr(1)}</mc:Choice>" +
            $"<mc:Fallback>{DocPr(1)}</mc:Fallback></mc:AlternateContent>";
        var body = Part($"<w:del>{alternative}</w:del><w:ins>{alternative}</w:ins>");

        DrawingIds.RenumberDocPrIds(new[] { body });

        Assert.Equal(new[] { "1", "1", "2", "2" }, DocPrIds(body));
    }

    [Fact]
    public void RenumberDocPrIds_SeparatesTwoDrawingsInOneBranch()
    {
        var body = Part($"<mc:AlternateContent><mc:Choice Requires=\"wps\">{DocPr(1)}{DocPr(1)}</mc:Choice>" +
            "</mc:AlternateContent>");

        DrawingIds.RenumberDocPrIds(new[] { body });

        Assert.Equal(new[] { "1", "2" }, DocPrIds(body));
    }

    private static string Shapetype(string anchorId, string path = "m,l,21600r21600,l21600,xe") =>
        $"<v:shapetype w14:anchorId=\"{anchorId}\" id=\"_x0000_t202\" coordsize=\"21600,21600\" path=\"{path}\">" +
        "<v:stroke joinstyle=\"miter\"/></v:shapetype>";

    private static string TextBox(string shapeId) => $"<v:shape id=\"{shapeId}\" type=\"#_x0000_t202\"/>";

    /// <summary>A paragraph holding one run of VML, optionally inside a revision mark.</summary>
    private static string Paragraph(string vml, string? revision = null) =>
        revision is null
            ? $"<w:p><w:r><w:pict>{vml}</w:pict></w:r></w:p>"
            : $"<w:p><w:{revision} w:id=\"1\" w:author=\"A\"><w:r><w:pict>{vml}</w:pict></w:r></w:{revision}></w:p>";

    private static string[] ShapetypeIds(XElement root) =>
        root.Descendants(VML.shapetype).Select(d => (string)d.Attribute("id")!).ToArray();

    private static string TypeOf(XElement root, string shapeId) =>
        (string)root.Descendants(VML.shape).Single(s => (string?)s.Attribute("id") == shapeId).Attribute("type")!;

    private static bool IsTracked(XElement element) =>
        element.Ancestors().Any(a => a.Name == W.ins || a.Name == W.del);

    [Fact]
    public void ResolveShapetypes_KeepsAnUntrackedFirstDefinitionAndDropsEquivalentCopies()
    {
        var body = Part(Paragraph(Shapetype("1") + TextBox("A")) + Paragraph(Shapetype("2") + TextBox("B"), "ins"));

        Assert.True(DrawingIds.ResolveShapetypes(body));

        Assert.Equal(new[] { "_x0000_t202" }, ShapetypeIds(body));
        Assert.Equal("#_x0000_t202", TypeOf(body, "B"));
    }

    [Fact]
    public void ResolveShapetypes_MovesEquivalentTrackedCopiesIntoOneUntrackedRun()
    {
        var body = Part(
            Paragraph(Shapetype("1") + TextBox("A"), "del") +
            Paragraph(Shapetype("2") + TextBox("B"), "ins") +
            Paragraph(TextBox("C")));

        DrawingIds.ResolveShapetypes(body);

        var definition = Assert.Single(body.Descendants(VML.shapetype));
        Assert.False(IsTracked(definition));
        Assert.Null(definition.Attribute(W14.w14 + "anchorId"));
        Assert.Same(body.Elements(W.p).First(), definition.Ancestors(W.p).Single());
        Assert.All(new[] { "A", "B", "C" }, id => Assert.Equal("#_x0000_t202", TypeOf(body, id)));
    }

    [Fact]
    public void ResolveShapetypes_MovesALoneInsertedDefinitionThatRejectingWouldTakeFromAnUntrackedShape()
    {
        var body = Part(Paragraph(Shapetype("1") + TextBox("New"), "ins") + Paragraph(TextBox("Old")));

        Assert.True(DrawingIds.ResolveShapetypes(body));

        Assert.False(IsTracked(Assert.Single(body.Descendants(VML.shapetype))));
    }

    [Fact]
    public void ResolveShapetypes_RenamesADifferentDefinitionAndBindsEachShapeToOneThatOutlivesIt()
    {
        var body = Part(
            Paragraph(Shapetype("1") + TextBox("A"), "del") +
            Paragraph(Shapetype("2", path: "m,l,21600,21600xe") + TextBox("B"), "ins") +
            Paragraph(TextBox("C"), "del"));

        DrawingIds.ResolveShapetypes(body);

        Assert.Equal(new[] { "_x0000_t202", "_x0000_t202_1" }, ShapetypeIds(body));
        Assert.Equal("#_x0000_t202", TypeOf(body, "A"));
        Assert.Equal("#_x0000_t202_1", TypeOf(body, "B"));
        Assert.Equal("#_x0000_t202", TypeOf(body, "C")); // the deleted definition, not the nearer inserted one
    }

    [Fact]
    public void ResolveShapetypes_LeavesOneUntrackedDefinitionPerIdUnchanged()
    {
        var body = Part(Paragraph(Shapetype("1") + TextBox("A")) + Paragraph(TextBox("B"), "ins"));

        Assert.False(DrawingIds.ResolveShapetypes(body));
    }

    [Fact]
    public void RenumberVmlIds_MovesALaterDuplicateAboveTheHighestShapeNumberAndItsOleObjectFollows()
    {
        var body = Part(
            Paragraph("<v:shape id=\"_x0000_s1026\"/><o:OLEObject ShapeID=\"_x0000_s1026\"/>", "del") +
            Paragraph("<v:shape id=\"_x0000_s1026\"/><o:OLEObject ShapeID=\"_x0000_s1026\"/>", "ins") +
            Paragraph("<v:shape id=\"_x0000_s1030\"/>"));

        Assert.True(DrawingIds.RenumberVmlIds(body));

        Assert.Equal(new[] { "_x0000_s1026", "_x0000_s1031", "_x0000_s1030" },
            body.Descendants(VML.shape).Select(s => (string)s.Attribute("id")!));
        Assert.Equal(new[] { "_x0000_s1026", "_x0000_s1031" },
            body.Descendants(O.OLEObject).Select(o => (string)o.Attribute("ShapeID")!));
    }

    [Fact]
    public void RenumberVmlIds_SuffixesANamedDuplicateAndLeavesAlternativeCopiesAlone()
    {
        var alternative = "<mc:AlternateContent><mc:Choice Requires=\"wps\"><v:rect id=\"Box\"/></mc:Choice>" +
            "<mc:Fallback><v:rect id=\"Box\"/></mc:Fallback></mc:AlternateContent>";
        var body = Part(Paragraph(alternative, "del") + Paragraph(alternative, "ins"));

        DrawingIds.RenumberVmlIds(body);

        Assert.Equal(new[] { "Box", "Box", "Box_1", "Box_1" },
            body.Descendants(VML.vml + "rect").Select(s => (string)s.Attribute("id")!));
    }
}
