// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="DrawingIds"/> on part roots: which drawing ids it renumbers, and how it resolves a
/// <c>v:shapetype</c> id defined more than once in one part (issue #860).
/// </summary>
public class DrawingIdsTests
{
    private const string Namespaces =
        " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"" +
        " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"" +
        " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"" +
        " xmlns:v=\"urn:schemas-microsoft-com:vml\"" +
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

    private static string[] ShapetypeIds(XElement root) =>
        root.Descendants(VML.shapetype).Select(d => (string)d.Attribute("id")!).ToArray();

    private static string TypeOf(XElement root, string shapeId) =>
        (string)root.Descendants(VML.shape).Single(s => (string?)s.Attribute("id") == shapeId).Attribute("type")!;

    [Fact]
    public void ResolveShapetypes_RemovesAnEquivalentCopyTheKeptDefinitionOutlives()
    {
        var body = Part(Shapetype("1") + TextBox("A") + $"<w:ins>{Shapetype("2")}{TextBox("B")}</w:ins>");

        Assert.True(DrawingIds.ResolveShapetypes(body));

        Assert.Equal(new[] { "_x0000_t202" }, ShapetypeIds(body));
        Assert.Equal("#_x0000_t202", TypeOf(body, "B"));
    }

    [Fact]
    public void ResolveShapetypes_RenamesAnEquivalentCopyThatRejectingWouldOrphan()
    {
        var body = Part($"<w:ins>{Shapetype("1")}{TextBox("A")}</w:ins><w:del>{Shapetype("2")}{TextBox("B")}</w:del>");

        DrawingIds.ResolveShapetypes(body);

        Assert.Equal(new[] { "_x0000_t202", "_x0000_t202_1" }, ShapetypeIds(body));
        Assert.Equal("#_x0000_t202", TypeOf(body, "A"));
        Assert.Equal("#_x0000_t202_1", TypeOf(body, "B"));
    }

    [Fact]
    public void ResolveShapetypes_RenamesADifferentDefinitionAndTheShapesThatFollowIt()
    {
        var body = Part(Shapetype("1") + TextBox("A") + Shapetype("2", path: "m,l,21600,21600xe") + TextBox("B"));

        DrawingIds.ResolveShapetypes(body);

        Assert.Equal(new[] { "_x0000_t202", "_x0000_t202_1" }, ShapetypeIds(body));
        Assert.Equal("#_x0000_t202", TypeOf(body, "A"));
        Assert.Equal("#_x0000_t202_1", TypeOf(body, "B"));
    }

    [Fact]
    public void ResolveShapetypes_LeavesAPartWithOneDefinitionPerIdUnchanged()
    {
        var body = Part(Shapetype("1") + TextBox("A") + TextBox("B"));

        Assert.False(DrawingIds.ResolveShapetypes(body));
    }
}
