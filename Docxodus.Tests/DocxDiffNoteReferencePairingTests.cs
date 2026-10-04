// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Tests.Ir;
using Xunit;
using static Docxodus.Tests.DocxBackendReconciliationTests;

namespace Docxodus.Tests;

/// <summary>
/// A note reference the body diff keeps unchanged must name the original's note after reject and the revised
/// document's note after accept, even when note contents alone would pair the notes differently (issue #865).
/// </summary>
public class DocxDiffNoteReferencePairingTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    private static string Boilerplate(string kind) =>
        $"<w:{kind} w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:{kind}>" +
        $"<w:{kind} w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:{kind}>";

    private static string Notes(string kind, params string[] texts) =>
        Boilerplate(kind) + string.Concat(texts.Select((text, i) =>
            $"<w:{kind} w:id=\"{i + 1}\"><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:{kind}>"));

    private static string Reference(string kind, int id) => $"<w:r><w:{kind}Reference w:id=\"{id}\"/></w:r>";

    private static string Paragraph(string text, string extraRuns = "") =>
        $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>{extraRuns}</w:p>";

    private static string CellTable(params string[] paragraphs) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid>" +
        string.Concat(paragraphs.Select(p => $"<w:tr><w:tc><w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/></w:tcPr>{p}</w:tc></w:tr>")) +
        "</w:tbl>";

    private static WmlDocument Doc(string kind, string body, params string[] notes) =>
        kind == "footnote"
            ? IrTestDocuments.FromParts(body, footnotesInnerXml: Notes(kind, notes))
            : IrTestDocuments.FromParts(body, endnotesInnerXml: Notes(kind, notes));

    private static XDocument Part(WmlDocument document, string name)
    {
        using var package = new ZipArchive(new MemoryStream(document.DocumentByteArray));
        var entry = package.GetEntry(name);
        if (entry is null)
            return new XDocument();
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    /// <summary>The text of the note each reference names, in reference order.</summary>
    private static List<string> NotesInReferenceOrder(WmlDocument document, string kind)
    {
        var notes = Part(document, $"word/{kind}s.xml").Descendants(W + kind)
            .GroupBy(n => (string)n.Attribute(W + "id")!)
            .ToDictionary(g => g.Key, g => string.Concat(g.First().Descendants(W + "t").Select(t => t.Value)));
        return Part(document, "word/document.xml").Descendants(W + kind + "Reference")
            .Select(r => notes.GetValueOrDefault((string)r.Attribute(W + "id")!, "<missing>"))
            .ToList();
    }

    private static void AssertReferencesReadTheirNotes(string kind, WmlDocument left, WmlDocument right)
    {
        var redline = DocxCompare.Compare(left, right);

        Assert.Equal(NotesInReferenceOrder(right, kind), NotesInReferenceOrder(RevisionProcessor.AcceptRevisions(redline), kind));
        Assert.Equal(NotesInReferenceOrder(left, kind), NotesInReferenceOrder(RevisionProcessor.RejectRevisions(redline), kind));
    }

    /// <summary>Both paragraphs are rewritten, so both references stay in place, while by content alone the
    /// original's second note ("Banana note") would pair with the revised first.</summary>
    [Theory]
    [InlineData("footnote")]
    [InlineData("endnote")]
    public void RewrittenParagraphs_KeptReferences_ReadTheirOwnNotesBothWays(string kind)
    {
        var left = Doc(kind,
            Paragraph("Apple one words", Reference(kind, 1)) + Paragraph("Banana two words", Reference(kind, 2)),
            "Apple note", "Banana note");
        var right = Doc(kind,
            Paragraph("Cherry fresh words entirely", Reference(kind, 1)) + Paragraph("Date more words entirely", Reference(kind, 2)),
            "Banana note", "Cherry note");

        AssertReferencesReadTheirNotes(kind, left, right);
    }

    [Fact]
    public void RewrittenTableCellParagraphs_KeptReferences_ReadTheirOwnNotesBothWays()
    {
        var left = Doc("footnote",
            CellTable(Paragraph("Apple one words", Reference("footnote", 1)), Paragraph("Banana two words", Reference("footnote", 2))),
            "Apple note", "Banana note");
        var right = Doc("footnote",
            CellTable(Paragraph("Cherry fresh words entirely", Reference("footnote", 1)),
                Paragraph("Date more words entirely", Reference("footnote", 2))),
            "Banana note", "Cherry note");

        AssertReferencesReadTheirNotes("footnote", left, right);
    }

    [Fact]
    public void UnchangedParagraphs_KeptReferences_ReadTheirOwnNotesBothWays()
    {
        // The paragraphs are equal, so the references are too, but the notes were edited so that by content the
        // original's second note looks like the revised first.
        var left = Doc("footnote",
            Paragraph("Apple one words", Reference("footnote", 1)) + Paragraph("Banana two words", Reference("footnote", 2)),
            "Apple note", "Banana note");
        var right = Doc("footnote",
            Paragraph("Apple one words", Reference("footnote", 1)) + Paragraph("Banana two words", Reference("footnote", 2)),
            "Banana note", "Cherry note");

        AssertReferencesReadTheirNotes("footnote", left, right);
    }

    [Fact]
    public void InsertedReferenceBeforeKeptOnes_StillPairsEachKeptReferenceWithItsNote()
    {
        var left = Doc("footnote",
            Paragraph("Apple one words", Reference("footnote", 1)) + Paragraph("Banana two words", Reference("footnote", 2)),
            "Apple note", "Banana note");
        var right = Doc("footnote",
            Paragraph("Zulu brand new paragraph", Reference("footnote", 1)) +
            Paragraph("Apple one words", Reference("footnote", 2)) + Paragraph("Banana two words", Reference("footnote", 3)),
            "Zulu note", "Apple note", "Banana note");

        AssertReferencesReadTheirNotes("footnote", left, right);
    }

    [Fact]
    public void RewrittenParagraphs_KeptReferences_AddNoValidationErrors()
    {
        var left = Doc("footnote",
            Paragraph("Apple one words", Reference("footnote", 1)) + Paragraph("Banana two words", Reference("footnote", 2)),
            "Apple note", "Banana note");
        var right = Doc("footnote",
            Paragraph("Cherry fresh words entirely", Reference("footnote", 1)) + Paragraph("Date more words entirely", Reference("footnote", 2)),
            "Banana note", "Cherry note");

        NoNewValidationErrors(left.DocumentByteArray, DocxCompare.Compare(left, right).DocumentByteArray);
    }
}
