// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1156: a complex field that opens in one body paragraph and closes in a later one (a Word table of
/// contents without its content control) must survive a comparison. Accepting the redline leaves exactly the
/// revised document's field characters, rejecting it the original's, each in the same paragraph and the same
/// place among the paragraph's text and tabs.
/// </summary>
public class BodyFieldSpanDiffTests
{
    public enum FieldEnd
    {
        /// <summary>Word's own shape: the field's end character sits alone in a trailing paragraph.</summary>
        OwnParagraph,

        /// <summary>The end character closes the last entry paragraph.</summary>
        LastEntry,
    }

    private const string Tail = "Closing paragraph after the contents.";

    private static string P(string text) => $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    /// <summary>One entry as Word writes it: a hyperlink around the heading text, a tab and a PAGEREF field.
    /// The first entry opens the TOC field itself before its hyperlink. The anchor comes from the entry's first
    /// word, so editing the rest of the text leaves the anchor alone, as Word does.</summary>
    private static string Entry(string text, bool opensToc, bool closesToc)
    {
        var anchor = "_Toc" + text.Split(' ')[0];
        var sb = new StringBuilder("<w:p>");
        if (opensToc)
            sb.Append("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>")
                .Append("<w:r><w:instrText xml:space=\"preserve\"> TOC \\o \"1-3\" \\h \\z \\u </w:instrText></w:r>")
                .Append("<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>");
        sb.Append($"<w:hyperlink w:anchor=\"{anchor}\" w:history=\"1\">")
            .Append($"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r><w:r><w:tab/></w:r>")
            .Append("<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>")
            .Append($"<w:r><w:instrText xml:space=\"preserve\"> PAGEREF {anchor} \\h </w:instrText></w:r>")
            .Append("<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>1</w:t></w:r>")
            .Append("<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:hyperlink>");
        if (closesToc)
            sb.Append("<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>");
        return sb.Append("</w:p>").ToString();
    }

    /// <summary>The paragraphs of a table of contents at body level, with no content control around them.</summary>
    private static string Contents(FieldEnd end, string[] entries)
    {
        var sb = new StringBuilder(P("Contents"));
        for (var i = 0; i < entries.Length; i++)
            sb.Append(Entry(entries[i], opensToc: i == 0, closesToc: end == FieldEnd.LastEntry && i == entries.Length - 1));
        if (end == FieldEnd.OwnParagraph)
            sb.Append("<w:p><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>");
        return sb.ToString();
    }

    private static WmlDocument Doc(FieldEnd end, string[] entries) =>
        IrTestDocuments.FromBodyXml(Contents(end, entries) + P(Tail));

    public static IEnumerable<object[]> ContentsEdits()
    {
        string[] three = { "Scope", "Term", "Fees" };
        foreach (var end in new[] { FieldEnd.OwnParagraph, FieldEnd.LastEntry })
        {
            yield return new object[] { end, "entry inserted first", new[] { "Term", "Fees" }, three };
            yield return new object[] { end, "entry inserted last", new[] { "Scope", "Term" }, three };
            yield return new object[] { end, "first entry deleted", three, new[] { "Term", "Fees" } };
            yield return new object[] { end, "middle entry deleted", three, new[] { "Scope", "Fees" } };
            yield return new object[] { end, "last entry deleted", three, new[] { "Scope", "Term" } };
            yield return new object[] { end, "entry text edited", three, new[] { "Scope", "Term and renewal", "Fees" } };
            yield return new object[] { end, "entry text shortened", new[] { "Scope", "Term and renewal", "Fees" }, three };
            yield return new object[] { end, "entry text replaced", three, new[] { "Scope", "Term", "Costs" } };
        }
    }

    [Theory]
    [MemberData(nameof(ContentsEdits))]
    public void Compare_BodyLevelContentsEntriesChanged_AcceptAndRejectKeepEachSidesFieldCharacters(
        FieldEnd end, string name, string[] leftEntries, string[] rightEntries)
    {
        var original = Doc(end, leftEntries);
        var revised = Doc(end, rightEntries);

        AssertFieldsSurviveRoundTrip(original, revised, DocxCompare.Compare(original, revised), name);
    }

    [Fact]
    public void Compare_EntryTextEdited_LeavesTheTabBeforeThePageField()
    {
        string[] three = { "Scope", "Term", "Fees" };
        var original = Doc(FieldEnd.OwnParagraph, three);
        var revised = Doc(FieldEnd.OwnParagraph, new[] { "Scope", "Term and renewal", "Fees" });

        var redline = DocxCompare.Compare(original, revised).DocumentByteArray;
        var accepted = Paragraphs(DocxDiffOps.AcceptRevisions(redline));
        var rejected = Paragraphs(DocxDiffOps.RejectRevisions(redline));
        Assert.Equal("\"Term and renewal\" TAB BEGIN [PAGEREF _TocTerm \\h] SEPARATE \"1\" END",
            Assert.Single(accepted, p => p.StartsWith("\"Term", StringComparison.Ordinal)));
        Assert.Equal("\"Term\" TAB BEGIN [PAGEREF _TocTerm \\h] SEPARATE \"1\" END",
            Assert.Single(rejected, p => p.StartsWith("\"Term", StringComparison.Ordinal)));
    }

    /// <summary>The same contents inside one table cell: the cell's identity carries its paragraphs' field
    /// structure, so moving the field's end between them is still a change.</summary>
    [Theory]
    [InlineData("last entry deleted")]
    [InlineData("entry inserted last")]
    public void Compare_ContentsInATableCell_AcceptAndRejectKeepEachSidesFieldCharacters(string name)
    {
        string[] three = { "Scope", "Term", "Fees" }, two = { "Scope", "Term" };
        static WmlDocument InCell(string[] entries) => IrTestDocuments.FromBodyXml(
            "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"9000\"/></w:tblGrid>" +
            "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"9000\" w:type=\"dxa\"/></w:tcPr>" + Contents(FieldEnd.LastEntry, entries) +
            "</w:tc></w:tr></w:tbl>" + P(Tail));
        var (original, revised) = name == "last entry deleted" ? (InCell(three), InCell(two)) : (InCell(two), InCell(three));

        AssertFieldsSurviveRoundTrip(original, revised, DocxCompare.Compare(original, revised), name);
    }

    /// <summary>Consolidating a reviewer who deleted the last entry (moving the field's end into the entry before
    /// it) with one who edited an earlier entry keeps the field balanced in both views.</summary>
    [Fact]
    public void Consolidate_ReviewerDeletesTheEntryHoldingTheFieldEnd_KeepsEveryFieldBalanced()
    {
        string[] three = { "Scope", "Term", "Fees" };
        var baseline = Doc(FieldEnd.LastEntry, three);
        var first = Doc(FieldEnd.LastEntry, new[] { "Scope", "Term" });
        var second = Doc(FieldEnd.LastEntry, new[] { "Scope and purpose", "Term", "Fees" });
        var merged = Doc(FieldEnd.LastEntry, new[] { "Scope and purpose", "Term" });

        var redline = DocxCompare.Consolidate(baseline,
            [new DocxDiffReviewer { Document = first, Author = "One" }, new DocxDiffReviewer { Document = second, Author = "Two" }]);

        AssertFieldsSurviveRoundTrip(baseline, merged, redline, "consolidate");
    }

    // ---- a table of contents Word wrote -------------------------------------------------------

    private static readonly DirectoryInfo FixtureDir = new("../../../../TestFiles/");

    private static WmlDocument WordContents() =>
        new(Path.Combine(FixtureDir.FullName, "HC022-Table-Of-Contents.docx"));

    /// <summary>The fixture's table of contents with its content control removed, so the field spans body
    /// paragraphs directly. Word writes the field's end character alone in a trailing paragraph; for
    /// <see cref="FieldEnd.LastEntry"/> it is moved to the end of the last entry instead.</summary>
    private static WmlDocument WordContentsAtBodyLevel(FieldEnd end, Action<List<XElement>>? editEntries = null) =>
        EditBody(WordContents(), body =>
        {
            var toc = body.Descendants(W.sdt).First(s => s.Descendants(W.docPartGallery).Any());
            var content = toc.Element(W.sdtContent)!.Elements().ToList();
            toc.ReplaceWith(content);
            var entries = body.Elements(W.p).Where(p => p.Elements(W.hyperlink).Any()).ToList();
            if (end == FieldEnd.LastEntry)
            {
                var closing = entries[^1].ElementsAfterSelf(W.p).First();
                entries[^1].Add(closing.Elements(W.r).Where(r => r.Element(W.fldChar) is not null).Select(r => new XElement(r)));
                closing.Remove();
            }
            editEntries?.Invoke(entries);
        });

    public static IEnumerable<object[]> WordContentsEdits() =>
        from end in new[] { FieldEnd.OwnParagraph, FieldEnd.LastEntry }
        from name in new[] { "last entry removed", "last entry appended", "entry text edited" }
        select new object[] { end, name };

    [Theory]
    [MemberData(nameof(WordContentsEdits))]
    public void Compare_WordAuthoredContentsAtBodyLevel_AcceptAndRejectKeepEachSidesFieldCharacters(FieldEnd end, string name)
    {
        var baseline = WordContentsAtBodyLevel(end);
        WmlDocument original = baseline, revised;
        switch (name)
        {
            case "last entry removed":
                revised = WordContentsAtBodyLevel(end, entries => RemoveEntryKeepingFieldEnd(entries[^1], entries[^2]));
                break;
            case "last entry appended":
                original = WordContentsAtBodyLevel(end, entries => RemoveEntryKeepingFieldEnd(entries[^1], entries[^2]));
                revised = baseline;
                break;
            default:
                revised = WordContentsAtBodyLevel(end, entries =>
                {
                    var heading = entries[1].Descendants(W.t).First(t => t.Value.Trim().Length > 0);
                    heading.Value += " and more";
                });
                break;
        }

        AssertFieldsSurviveRoundTrip(original, revised, DocxCompare.Compare(original, revised), name);
    }

    /// <summary>Removes an entry paragraph; a field character that closed the table of contents inside it moves to
    /// the end of the entry before it, which is what Word does when the last entry is deleted.</summary>
    private static void RemoveEntryKeepingFieldEnd(XElement entry, XElement previous)
    {
        var trailingEnds = entry.Elements(W.r).Where(r => r.Element(W.fldChar) is { } f &&
            (string?)f.Attribute(W.fldCharType) == "end").ToList();
        foreach (var run in trailingEnds)
            previous.Add(new XElement(run));
        entry.Remove();
    }

    // ---- assertions ---------------------------------------------------------------------------

    /// <summary>The redline adds no validation error, and accepting / rejecting it leaves each paragraph with the
    /// field characters, instructions, text and tabs of the revised / original document, in the same order.</summary>
    private static void AssertFieldsSurviveRoundTrip(WmlDocument original, WmlDocument revised, WmlDocument redline, string name)
    {
        var failures = FieldShapeMismatch(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray), revised.DocumentByteArray, "accept, " + name)
            + FieldShapeMismatch(DocxDiffOps.RejectRevisions(redline.DocumentByteArray), original.DocumentByteArray, "reject, " + name);
        Assert.True(failures.Length == 0, failures);
        AssertNoNewValidationErrors(redline, original, revised);
    }

    private static string FieldShapeMismatch(byte[] actual, byte[] expected, string because)
    {
        var want = string.Join("\n", Paragraphs(expected));
        var got = string.Join("\n", Paragraphs(actual));
        return want == got ? "" : $"field shape ({because}):\nexpected:\n{want}\n---\ngot:\n{got}\n";
    }

    /// <summary>Each body paragraph as the sequence of its field characters, instructions, tabs and text runs
    /// (adjacent text joined), in document order.</summary>
    private static List<string> Paragraphs(byte[] bytes)
    {
        var result = new List<string>();
        foreach (var p in Body(bytes).Descendants(W.p))
        {
            var parts = new List<string>();
            var text = new StringBuilder();
            var instr = new StringBuilder();
            void Flush()
            {
                if (text.Length > 0) { parts.Add("\"" + text + "\""); text.Clear(); }
                if (instr.Length > 0) { parts.Add("[" + instr.ToString().Trim() + "]"); instr.Clear(); }
            }
            foreach (var e in p.Descendants())
            {
                if (e.Name == W.t)
                {
                    if (instr.Length > 0) Flush();
                    text.Append(e.Value);
                }
                else if (e.Name == W.instrText)
                {
                    if (text.Length > 0) Flush();
                    instr.Append(e.Value);
                }
                else if (e.Name == W.tab && e.Parent?.Name == W.r)
                {
                    Flush();
                    parts.Add("TAB");
                }
                else if (e.Name == W.fldChar)
                {
                    Flush();
                    parts.Add(((string?)e.Attribute(W.fldCharType) ?? "?").ToUpperInvariant());
                }
            }
            Flush();
            result.Add(string.Join(" ", parts));
        }
        return result;
    }

    private static XElement Body(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    private static WmlDocument EditBody(WmlDocument document, Action<XElement> edit)
    {
        using var stream = new MemoryStream();
        stream.Write(document.DocumentByteArray, 0, document.DocumentByteArray.Length);
        using (var package = WordprocessingDocument.Open(stream, true))
        {
            var main = package.MainDocumentPart!;
            edit(main.GetXDocument().Root!.Element(W.body)!);
            main.PutXDocument();
        }
        return new WmlDocument("revised.docx", stream.ToArray());
    }

    /// <summary>Every validation error of <paramref name="redline"/> also occurs in one of the
    /// <paramref name="inputs"/>: the comparison introduced none.</summary>
    private static void AssertNoNewValidationErrors(WmlDocument redline, params WmlDocument[] inputs)
    {
        static HashSet<string> Errors(WmlDocument document)
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            return new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package)
                .Select(e => e.Description).ToHashSet(StringComparer.Ordinal);
        }
        var known = inputs.SelectMany(Errors).ToHashSet(StringComparer.Ordinal);
        Assert.All(Errors(redline), e => Assert.Contains(e, known));
    }
}
