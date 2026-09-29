// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Tests.Ir;
using Xunit;
using static Docxodus.Tests.DocxBackendReconciliationTests;

namespace Docxodus.Tests;

/// <summary>
/// Package integrity of comparison output (issue #840): bookmark ids stay unique and paired when both
/// documents' bookmarks survive, every note reference resolves to a note, and a part type the main
/// document may relate to only once is never related twice. Each case goes through the public entry
/// points — <see cref="DocxCompare.Compare"/> and <see cref="DocxDiff.Consolidate"/>.
/// </summary>
public class DocxDiffPackageIntegrityTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    private const string NoteBoilerplate =
        "<w:{0} w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:{0}>" +
        "<w:{0} w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:{0}>";

    private static string Notes(string kind, params string[] texts) =>
        string.Format(NoteBoilerplate, kind) + string.Concat(texts.Select((text, i) =>
            $"<w:{kind} w:id=\"{i + 1}\"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:{kind}>"));

    private static string Paragraph(string text, string extraRuns = "") =>
        $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>{extraRuns}</w:p>";

    private static string Cell(string innerXml) => $"<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>{innerXml}</w:tc>";

    /// <summary>A one-row table whose row carries bookmark <paramref name="name"/> (id 0) around its cells —
    /// the row-level placement Word uses for a bookmark that spans a whole row.</summary>
    private static string TableWithRowBookmark(string name, params string[] cellTexts) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid>" +
        string.Concat(cellTexts.Select(_ => "<w:gridCol w:w=\"2000\"/>")) + "</w:tblGrid>" +
        $"<w:tr><w:bookmarkStart w:id=\"0\" w:name=\"{name}\"/>" +
        string.Concat(cellTexts.Select(text => Cell(Paragraph(text)))) + "<w:bookmarkEnd w:id=\"0\"/></w:tr></w:tbl>";

    // ---------------------------------------------------------------- 1. duplicate bookmark ids

    public static IEnumerable<object[]> BookmarkPairs()
    {
        // A deleted paragraph's bookmark and an inserted table's row-level bookmark, both numbered 0: both
        // survive in the redline.
        yield return new object[]
        {
            "table",
            "<w:p><w:bookmarkStart w:id=\"0\" w:name=\"LeftText\"/><w:r><w:t>Quarterly revenue by region</w:t></w:r>" +
            "<w:bookmarkEnd w:id=\"0\"/></w:p>",
            TableWithRowBookmark("RightRow", "Headcount", "Plan", "Fiscal year"),
        };

        // A deleted table and an inserted table, each with a row-level bookmark numbered 0 and no other
        // bookmark in the body.
        yield return new object[]
        {
            "tables",
            TableWithRowBookmark("LeftRow", "Quarterly revenue by region") + Paragraph("Shared closing line."),
            Paragraph("Shared closing line.") + TableWithRowBookmark("RightRow", "Headcount", "Plan", "Fiscal year"),
        };

        // A bookmark inside deleted math and a run-level bookmark on inserted text, both numbered 0.
        yield return new object[]
        {
            "math",
            "<w:p><m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath>" +
            "<m:r><m:t>x</m:t></m:r><w:bookmarkStart w:id=\"0\" w:name=\"LeftMath\"/><m:r><m:t>=1</m:t></m:r>" +
            "<w:bookmarkEnd w:id=\"0\"/></m:oMath></m:oMathPara></w:p>",
            "<w:p><w:bookmarkStart w:id=\"0\" w:name=\"RightText\"/><w:r><w:t>Entirely different prose here</w:t></w:r>" +
            "<w:bookmarkEnd w:id=\"0\"/></w:p>",
        };
    }

    [Theory]
    [MemberData(nameof(BookmarkPairs))]
    public void Compare_BothSidesBookmarks_GetUniquePairedIds(string scenario, string leftBody, string rightBody)
    {
        _ = scenario;
        var (left, right) = (IrTestDocuments.FromParts(leftBody), IrTestDocuments.FromParts(rightBody));

        AssertUniquePairedBookmarks(DocxCompare.Compare(left, right));
    }

    [Theory]
    [MemberData(nameof(BookmarkPairs))]
    public void Consolidate_BothSidesBookmarks_GetUniquePairedIds(string scenario, string leftBody, string rightBody)
    {
        _ = scenario;
        var (left, right) = (IrTestDocuments.FromParts(leftBody), IrTestDocuments.FromParts(rightBody));

        AssertUniquePairedBookmarks(Consolidate(left, right));
    }

    [Theory]
    [MemberData(nameof(BookmarkPairs))]
    public void Compare_BothSidesBookmarks_AcceptKeepsRightsAndRejectKeepsLefts(
        string scenario, string leftBody, string rightBody)
    {
        _ = scenario;
        var (left, right) = (IrTestDocuments.FromParts(leftBody), IrTestDocuments.FromParts(rightBody));
        var redline = DocxCompare.Compare(left, right);

        Assert.Equal(BookmarkSpans(right), BookmarkSpans(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal(BookmarkSpans(left), BookmarkSpans(RevisionProcessor.RejectRevisions(redline)));
    }

    [Theory]
    [MemberData(nameof(BookmarkPairs))]
    public void Compare_BothSidesBookmarks_AddsNoValidationErrors(string scenario, string leftBody, string rightBody)
    {
        _ = scenario;
        var (left, right) = (IrTestDocuments.FromParts(leftBody), IrTestDocuments.FromParts(rightBody));

        NoNewValidationErrors(left.DocumentByteArray, DocxCompare.Compare(left, right).DocumentByteArray);
    }

    // ---------------------------------------------------------------- 2. dangling note references

    /// <summary>The original's two footnoted paragraphs are rewritten into one with a single footnote, so
    /// the first original note pairs with the revised note while its own reference is deleted.</summary>
    private static readonly WmlDocument FootnotedOriginal = IrTestDocuments.FromParts(
        Paragraph("Alpha clause about payment terms.", "<w:r><w:footnoteReference w:id=\"1\"/></w:r>") +
        Paragraph("Beta clause about delivery dates.", "<w:r><w:footnoteReference w:id=\"2\"/></w:r>"),
        footnotesInnerXml: Notes("footnote", "Net thirty days from invoice.", "Delivery within ten business days."));

    private static readonly WmlDocument FootnotedRevised = IrTestDocuments.FromParts(
        Paragraph("Gamma governing law provision replaces everything.", "<w:r><w:footnoteReference w:id=\"1\"/></w:r>"),
        footnotesInnerXml: Notes("footnote", "Net forty-five days from invoice."));

    /// <summary>The original has no footnotes part at all.</summary>
    private static readonly WmlDocument UnfootnotedOriginal = IrTestDocuments.FromParts(
        Paragraph("Alpha clause about payment terms."));

    /// <summary>The revised document references a footnote whose definition is empty
    /// (<c>&lt;w:footnote w:id="1"/&gt;</c>), as the WC064 fixture's is.</summary>
    private static readonly WmlDocument EmptyFootnoteRevised = IrTestDocuments.FromParts(
        Paragraph("Alpha clause about payment terms.", "<w:r><w:footnoteReference w:id=\"1\"/></w:r>"),
        footnotesInnerXml: string.Format(NoteBoilerplate, "footnote") + "<w:footnote w:id=\"1\"/>");

    public static IEnumerable<object[]> NotePairs()
    {
        yield return new object[] { "rewritten", FootnotedOriginal, FootnotedRevised };
        yield return new object[] { "empty-note", UnfootnotedOriginal, EmptyFootnoteRevised };
    }

    [Theory]
    [MemberData(nameof(NotePairs))]
    public void Compare_EveryNoteReferenceResolves(string scenario, WmlDocument left, WmlDocument right)
    {
        _ = scenario;
        var redline = DocxCompare.Compare(left, right);

        AssertNoteReferencesResolve(redline);
        AssertNoteReferencesResolve(RevisionProcessor.AcceptRevisions(redline));
        AssertNoteReferencesResolve(RevisionProcessor.RejectRevisions(redline));
    }

    [Theory]
    [MemberData(nameof(NotePairs))]
    public void Consolidate_EveryNoteReferenceResolves(string scenario, WmlDocument left, WmlDocument right)
    {
        _ = scenario;
        AssertNoteReferencesResolve(Consolidate(left, right));
    }

    // ---------------------------------------------------------------- 3. duplicated singleton relationships

    private const string DrawingNamespaces =
        " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"" +
        " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" +
        " xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"" +
        " xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"";

    private static string Picture(string relationshipId) =>
        "<w:r><w:drawing><wp:inline><wp:extent cx=\"9525\" cy=\"9525\"/><wp:docPr id=\"1\" name=\"Picture\"/>" +
        "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic>" +
        "<pic:nvPicPr><pic:cNvPr id=\"1\" name=\"Picture\"/><pic:cNvPicPr/></pic:nvPicPr>" +
        $"<pic:blipFill><a:blip r:embed=\"{relationshipId}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
        "<pic:spPr/></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

    private static string PriceTable(string price, string extraRuns) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
        "<w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>" +
        $"<w:tr>{Cell(Paragraph("Item"))}{Cell(Paragraph("Price"))}</w:tr>" +
        $"<w:tr>{Cell(Paragraph("Widget"))}{Cell(Paragraph(price, extraRuns))}</w:tr></w:tbl>";

    /// <summary>A document with endnotes whose main-part relationship ids are fixed, so the original's picture
    /// and the revised document's endnotes part can share one id — as they routinely do in real documents.</summary>
    private static WmlDocument WithEndnotes(string bodyInnerXml, string endnotesInnerXml, string endnotesId,
        string? imageId = null)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>("rId1").Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>("rId2").Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            Write(main.AddNewPart<EndnotesPart>(endnotesId), $"<w:endnotes xmlns:w=\"{W}\">{endnotesInnerXml}</w:endnotes>");
            if (imageId is not null)
            {
                using var image = main.AddNewPart<ImagePart>("image/png", imageId).GetStream(FileMode.Create);
                image.Write(IrTestDocuments.TinyPng);
            }
            Write(main, $"<w:document xmlns:w=\"{W}\"{DrawingNamespaces}><w:body>{bodyInnerXml}</w:body></w:document>");
        }
        return new WmlDocument("endnotes.docx", stream.ToArray());
    }

    private static void Write(OpenXmlPart part, string xml)
    {
        using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
        writer.Write(xml);
    }

    /// <summary>Both documents carry endnotes; the revised one edits the first, adds a second, and drops the
    /// picture from a table cell whose text also changed. The original's picture is related as rId9 — the id
    /// the revised document gives its endnotes part.</summary>
    private static readonly WmlDocument EndnotedOriginal = WithEndnotes(
        PriceTable("Ten dollars", Picture("rId9")) +
        Paragraph("Terms apply.", "<w:r><w:endnoteReference w:id=\"1\"/></w:r>"),
        Notes("endnote", "Prices exclude tax."), endnotesId: "rId3", imageId: "rId9");

    private static readonly WmlDocument EndnotedRevised = WithEndnotes(
        PriceTable("Twelve dollars", "") +
        Paragraph("Terms apply.", "<w:r><w:endnoteReference w:id=\"1\"/></w:r>") +
        Paragraph("Shipping is extra.", "<w:r><w:endnoteReference w:id=\"2\"/></w:r>"),
        Notes("endnote", "Prices include tax.", "Shipping is quoted per order."), endnotesId: "rId9");

    [Fact]
    public void Compare_NeverRelatesASingletonPartTwice()
    {
        var redline = DocxCompare.Compare(EndnotedOriginal, EndnotedRevised);

        AssertSingletonRelationshipsUnique(redline);
        using var opened = WordprocessingDocument.Open(new MemoryStream(redline.DocumentByteArray), false);
        Assert.NotNull(opened.MainDocumentPart!.EndnotesPart);
    }

    [Fact]
    public void Compare_DeletedPictureKeepsTheOriginalsImage()
    {
        var redline = DocxCompare.Compare(EndnotedOriginal, EndnotedRevised);

        using var opened = WordprocessingDocument.Open(new MemoryStream(redline.DocumentByteArray), false);
        var main = opened.MainDocumentPart!;
        var embed = main.GetXDocument().Descendants(XName.Get("blip", "http://schemas.openxmlformats.org/drawingml/2006/main"))
            .Select(b => (string)b.Attribute(XName.Get("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"))!)
            .Single();
        Assert.IsAssignableFrom<ImagePart>(main.GetPartById(embed));
    }

    [Fact]
    public void Consolidate_NeverRelatesASingletonPartTwice() =>
        AssertSingletonRelationshipsUnique(Consolidate(EndnotedOriginal, EndnotedRevised));

    [Fact]
    public void Compare_EndnotesAddedAndRemoved_AddsNoValidationErrors() =>
        NoNewValidationErrors(EndnotedOriginal.DocumentByteArray,
            DocxCompare.Compare(EndnotedOriginal, EndnotedRevised).DocumentByteArray);

    // ---------------------------------------------------------------- helpers

    private static WmlDocument Consolidate(WmlDocument original, WmlDocument revised) =>
        DocxDiff.Consolidate(original, new[] { new DocxDiffReviewer { Author = "Reviewer", Document = revised } });

    private static XDocument Part(WmlDocument document, string name)
    {
        using var package = new ZipArchive(new MemoryStream(document.DocumentByteArray));
        var entry = package.GetEntry(name);
        if (entry is null)
            return new XDocument();
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static void AssertUniquePairedBookmarks(WmlDocument document)
    {
        var body = Part(document, "word/document.xml");
        var starts = body.Descendants(W + "bookmarkStart").Select(s => (string)s.Attribute(W + "id")!).ToList();
        var ends = body.Descendants(W + "bookmarkEnd").Select(e => (string)e.Attribute(W + "id")!).ToList();

        Assert.Equal(starts.Distinct().Count(), starts.Count);
        Assert.Equal(starts.OrderBy(id => id), ends.OrderBy(id => id));
    }

    /// <summary>Each bookmark's name with the text its start/end pair encloses, in document order.</summary>
    private static List<string> BookmarkSpans(WmlDocument document)
    {
        var body = Part(document, "word/document.xml").Root!;
        return body.Descendants(W + "bookmarkStart").Select(start =>
        {
            var id = (string)start.Attribute(W + "id")!;
            var end = body.Descendants(W + "bookmarkEnd").Single(e => (string)e.Attribute(W + "id")! == id);
            var text = string.Concat(start.ElementsAfterSelf().Concat(start.Parent!.ElementsAfterSelf())
                .SelectMany(e => e.DescendantsAndSelf())
                .TakeWhile(e => e != end)
                .Where(e => e.Name == W + "t" || e.Name.LocalName == "t")
                .Select(e => e.Value));
            return $"{(string)start.Attribute(W + "name")!}:{text}";
        }).ToList();
    }

    private static void AssertNoteReferencesResolve(WmlDocument document)
    {
        foreach (var kind in new[] { "footnote", "endnote" })
        {
            var defined = Part(document, $"word/{kind}s.xml").Descendants(W + kind)
                .Select(n => (string)n.Attribute(W + "id")!).ToHashSet();
            var referenced = Part(document, "word/document.xml").Descendants(W + (kind + "Reference"))
                .Select(r => (string)r.Attribute(W + "id")!);
            Assert.All(referenced, id => Assert.Contains(id, defined));
        }
    }

    private static void AssertSingletonRelationshipsUnique(WmlDocument document)
    {
        XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        var types = Part(document, "word/_rels/document.xml.rels").Descendants(packageRelationships + "Relationship")
            .Select(r => ((string)r.Attribute("Type")!).Split('/').Last())
            .Where(type => type is "styles" or "settings" or "numbering" or "fontTable" or "theme" or "webSettings"
                or "footnotes" or "endnotes" or "comments")
            .ToList();

        Assert.Equal(types.Distinct().OrderBy(t => t), types.OrderBy(t => t));
    }
}
