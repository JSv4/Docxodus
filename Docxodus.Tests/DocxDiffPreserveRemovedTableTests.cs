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
/// Under <see cref="DocxDiffSettings.PreserveInputRevisions"/>, a table in the revised input that accepting
/// removes (every row tracked-deleted, or its content moved away) is carried into the redline with its rows
/// still marked deleted, in its place, instead of disappearing (issue #866).
/// </summary>
public class DocxDiffPreserveRemovedTableTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    private const string Before = "<w:p><w:r><w:t>Before</w:t></w:r></w:p>";
    private const string Tail = "<w:p><w:r><w:t>After the table</w:t></w:r></w:p>";

    private static string P(string text) => $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static string Table(params string[] rows) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>" +
        string.Concat(rows) + "</w:tbl>";

    private static string DeletedRow(string text, int id) =>
        $"<w:tr><w:trPr><w:del w:id=\"{id}\" w:author=\"Alice\" w:date=\"2020-01-01T00:00:00Z\"/></w:trPr>" +
        $"<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc></w:tr>";

    private static readonly string WhollyDeletedTable = Table(DeletedRow("gone one", 1), DeletedRow("gone two", 2));

    /// <summary>Every row deleted and its text moved out, as Word writes a cut-and-pasted table.</summary>
    private static readonly string MovedAwayTable =
        Table("<w:tr><w:trPr><w:del w:id=\"3\" w:author=\"Alice\"/></w:trPr><w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>" +
              "<w:p><w:pPr><w:rPr><w:del w:id=\"4\" w:author=\"Alice\"/></w:rPr></w:pPr>" +
              "<w:moveFrom w:id=\"5\" w:author=\"Alice\"><w:del w:id=\"6\" w:author=\"Alice\"><w:r><w:delText>Moved</w:delText></w:r>" +
              "</w:del></w:moveFrom></w:p></w:tc></w:tr>");

    private static WmlDocument Preserve(string leftBody, string rightBody) =>
        DocxDiff.Compare(
            IrTestDocuments.FromBodyXml(leftBody),
            IrTestDocuments.FromBodyXml(rightBody),
            new DocxDiffSettings { PreserveInputRevisions = true });

    private static XElement Body(WmlDocument document)
    {
        using var package = new ZipArchive(new MemoryStream(document.DocumentByteArray));
        using var part = package.GetEntry("word/document.xml")!.Open();
        return XElement.Load(part).Element(W + "body")!;
    }

    private static XElement CarriedTable(WmlDocument output)
    {
        var table = Assert.Single(Body(output).Descendants(W + "tbl"));
        Assert.All(table.Elements(W + "tr"), row => Assert.NotNull(row.Element(W + "trPr")?.Element(W + "del")));
        return table;
    }

    private static string[] BlockTexts(XElement body) =>
        body.Elements()
            .Where(e => e.Name == W + "p" || e.Name == W + "tbl")
            .Select(e => string.Concat(e.Descendants().Where(d => d.Name == W + "t" || d.Name == W + "delText").Select(d => d.Value)))
            .ToArray();

    private static string AcceptedText(WmlDocument doc) =>
        string.Concat(Body(RevisionProcessor.AcceptRevisions(doc)).Descendants(W + "t").Select(t => t.Value));

    [Fact]
    public void WhollyDeletedTable_BetweenEqualParagraphs_IsCarriedInPlaceWithItsDeletedRows()
    {
        var right = Before + WhollyDeletedTable + P("Kept") + Tail;

        var output = Preserve(Before + P("Kept") + Tail, right);

        var table = CarriedTable(output);
        Assert.Equal(new[] { "Alice", "Alice" },
            table.Elements(W + "tr").Select(row => (string?)row.Element(W + "trPr")!.Element(W + "del")!.Attribute(W + "author")));
        Assert.Equal(new[] { "Before", "gone onegone two", "Kept", "After the table" }, BlockTexts(Body(output)));
    }

    [Fact]
    public void WhollyDeletedTable_AcceptRemovesItAndRejectRestoresIt()
    {
        var right = Before + WhollyDeletedTable + P("Kept") + Tail;

        var output = Preserve(Before + P("Kept") + Tail, right);

        Assert.Equal(AcceptedText(IrTestDocuments.FromBodyXml(right)), AcceptedText(output));
        Assert.Empty(Body(RevisionProcessor.AcceptRevisions(output)).Descendants(W + "tbl"));
        var rejected = Body(RevisionProcessor.RejectRevisions(output));
        Assert.Equal("gone onegone two", string.Concat(Assert.Single(rejected.Descendants(W + "tbl")).Descendants(W + "t").Select(t => t.Value)));
    }

    [Fact]
    public void MovedAwayTable_IsCarriedInPlaceWithItsDeletedRows()
    {
        var right = Before + MovedAwayTable + P("Kept") + Tail;

        var output = Preserve(Before + P("Kept") + Tail, right);

        CarriedTable(output);
        Assert.Equal(AcceptedText(IrTestDocuments.FromBodyXml(right)), AcceptedText(output));
    }

    [Fact]
    public void WhollyDeletedTable_BeforeAnInsertedParagraph_StaysDeletedNotInserted()
    {
        var output = Preserve(Before + Tail, Before + WhollyDeletedTable + P("Brand new") + Tail);

        var table = CarriedTable(output);
        Assert.Empty(table.Descendants(W + "ins"));
        Assert.Equal(new[] { "Before", "gone onegone two", "Brand new", "After the table" }, BlockTexts(Body(output)));
    }

    [Fact]
    public void WhollyDeletedTable_BeforeAModifiedCleanParagraph_KeepsThatParagraphsFineRedline()
    {
        var output = Preserve(
            Before + P("Kept original words") + Tail,
            Before + WhollyDeletedTable + P("Kept changed words") + Tail);

        CarriedTable(output);
        // The clean paragraph is still diffed word by word: its common words are neither inserted nor deleted.
        var paragraph = Body(output).Elements(W + "p").Single(p => p.Descendants(W + "t").Any(t => t.Value.Contains("Kept")));
        Assert.Contains(paragraph.Descendants(W + "t"),
            t => t.Value.Contains("Kept") && !t.Ancestors(W + "ins").Any());
        Assert.Contains(paragraph.Descendants(W + "delText"), t => t.Value.Contains("original"));
    }

    [Fact]
    public void WhollyDeletedTable_AddsNoValidationErrors()
    {
        var right = IrTestDocuments.FromBodyXml(Before + WhollyDeletedTable + P("Kept") + Tail);

        var output = DocxDiff.Compare(
            IrTestDocuments.FromBodyXml(Before + P("Kept") + Tail), right,
            new DocxDiffSettings { PreserveInputRevisions = true });

        NoNewValidationErrors(right.DocumentByteArray, output.DocumentByteArray);
    }
}
