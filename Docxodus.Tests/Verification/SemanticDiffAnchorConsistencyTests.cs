// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

/// <summary>
/// One change set must speak one anchor vocabulary: package-detector records use the same
/// hdr{N}/ftr{N} scope names as the IR, and a fact whose containing block the IR aligned in
/// place (text edit, duplicate-shift) is not a Move.
/// </summary>
public class SemanticDiffAnchorConsistencyTests
{
    private const string W = IrTestDocuments.W;
    private const string RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    [Fact]
    public void Header_scopes_agree_between_package_and_ir_records()
    {
        var left = WithTwoHeadersRelsReversed(
            "<w:p><w:r><w:t xml:space=\"preserve\">Alpha one</w:t></w:r>" +
            "<w:bookmarkStart w:id=\"1\" w:name=\"target\"/><w:bookmarkEnd w:id=\"1\"/></w:p>",
            "Beta header");
        var right = WithTwoHeadersRelsReversed(
            "<w:p><w:r><w:t xml:space=\"preserve\">Alpha two</w:t></w:r>" +
            "<w:bookmarkStart w:id=\"1\" w:name=\"renamed\"/><w:bookmarkEnd w:id=\"1\"/></w:p>",
            "Beta header");

        var result = SemanticDiff.Compare(left, right);

        var headerScopes = result.Changes
            .Where(change => change.PartUri == "/word/header1.xml")
            .SelectMany(change => new[] { change.LeftScope, change.RightScope })
            .Where(scope => scope is not null)
            .Distinct()
            .ToArray();
        Assert.True(headerScopes.Length == 1,
            "every change in /word/header1.xml must carry the same scope name, got: "
            + string.Join(", ", headerScopes));
    }

    [Fact]
    public void Editing_text_around_a_bookmark_is_not_a_bookmark_move()
    {
        var (left, right) = TextEditedAroundABookmark();

        var result = SemanticDiff.Compare(left, right);

        Assert.Contains(result.Changes, change => change.Family == SemanticChangeFamily.Text);
        Assert.DoesNotContain(result.Changes, change =>
            change.Family == SemanticChangeFamily.Bookmark);
    }

    [Fact]
    public void Inserting_a_duplicate_paragraph_does_not_move_a_later_bookmark()
    {
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Unique</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Repeat</w:t></w:r>" +
            "<w:bookmarkStart w:id=\"2\" w:name=\"bm2\"/><w:bookmarkEnd w:id=\"2\"/></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Unique</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Repeat</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Repeat</w:t></w:r>" +
            "<w:bookmarkStart w:id=\"2\" w:name=\"bm2\"/><w:bookmarkEnd w:id=\"2\"/></w:p>");

        var result = SemanticDiff.Compare(left, right);

        Assert.Contains(result.Changes, change =>
            change.Family == SemanticChangeFamily.BlockStructure
            && change.Operation == SemanticChangeOperation.Insert);
        Assert.DoesNotContain(result.Changes, change =>
            change.Family == SemanticChangeFamily.Bookmark);
    }

    [Fact]
    public void A_bookmark_relocated_to_another_paragraph_is_still_a_move()
    {
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>First</w:t></w:r>" +
            "<w:bookmarkStart w:id=\"1\" w:name=\"bm1\"/><w:bookmarkEnd w:id=\"1\"/></w:p>" +
            "<w:p><w:r><w:t>Second</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>First</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>Second</w:t></w:r>" +
            "<w:bookmarkStart w:id=\"1\" w:name=\"bm1\"/><w:bookmarkEnd w:id=\"1\"/></w:p>");

        var result = SemanticDiff.Compare(left, right);

        var move = Assert.Single(result.Changes, change =>
            change.Family == SemanticChangeFamily.Bookmark);
        Assert.Equal(SemanticChangeOperation.Move, move.Operation);
        Assert.False(string.IsNullOrEmpty(move.MoveId));
    }

    /// <summary>
    /// The bounded evidence budget governs findings the caller is actually handed, so a package
    /// draft dropped as an aligned relocation must not consume it. The two facts meet on one
    /// statement in SemanticDiffEngine: the alignment filter has to run BEFORE the budget check.
    /// Reversing them makes the cap fire on evidence that is never published — a caller whose
    /// document merely had text edited around a bookmark would get a limit failure instead of a
    /// report.
    ///
    /// Calibration note: the assertion has teeth only while this fixture still produces an
    /// aligned relocation for the filter to drop. Its sibling
    /// <see cref="Editing_text_around_a_bookmark_is_not_a_bookmark_move"/> pins that the draft is
    /// suppressed, and <see cref="A_bookmark_relocated_to_another_paragraph_is_still_a_move"/>
    /// pins that the detector does emit Move drafts for this construct.
    /// </summary>
    [Fact]
    public void An_aligned_relocation_does_not_consume_the_bounded_change_budget()
    {
        var (left, right) = TextEditedAroundABookmark();
        var options = new SemanticDiffOptions();

        // The exact number of findings the caller is handed — every one of them published.
        var published = SemanticDiff.Compare(left, right, options).Changes.Count;
        Assert.True(published > 0, "the fixture must produce findings for the budget to bound");

        // A budget of exactly that many is sufficient: nothing published is dropped, and the
        // filtered aligned relocation is not evidence anyone has to retain. Counting the
        // unfiltered draft here would overshoot the budget by one and throw.
        var bounded = SemanticDiff.CompareBounded(left, right, options, published);

        Assert.Equal(published, bounded.Changes.Count);
    }

    /// <summary>
    /// The companion to <see cref="An_aligned_relocation_does_not_consume_the_bounded_change_budget"/>:
    /// the budget is still live at the boundary. Without this, that test would also pass against a
    /// cap that never fires at all.
    /// </summary>
    [Fact]
    public void The_bounded_change_budget_still_fails_closed_one_finding_short()
    {
        var (left, right) = TextEditedAroundTwoBookmarks();
        var options = new SemanticDiffOptions();

        var published = SemanticDiff.Compare(left, right, options).Changes.Count;
        Assert.True(published > 1, "a budget one short of the published count must be reachable");

        Assert.Throws<SemanticChangeLimitExceededException>(() =>
            SemanticDiff.CompareBounded(left, right, options, published - 1));
    }

    /// <summary>
    /// Text edited on both sides of an untouched bookmark: the IR aligns the containing paragraph
    /// in place, so the package detector's relocation draft describes a block that did not move.
    /// </summary>
    /// <summary>
    /// The same construct in two paragraphs, so the published finding count is high enough for a
    /// budget one short of it to be a legal (positive) budget.
    /// </summary>
    private static (WmlDocument Left, WmlDocument Right) TextEditedAroundTwoBookmarks() => (
        IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t xml:space=\"preserve\">hello </w:t></w:r>" +
            "<w:bookmarkStart w:id=\"1\" w:name=\"bm1\"/><w:bookmarkEnd w:id=\"1\"/>" +
            "<w:r><w:t>world</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t xml:space=\"preserve\">second </w:t></w:r>" +
            "<w:bookmarkStart w:id=\"2\" w:name=\"bm2\"/><w:bookmarkEnd w:id=\"2\"/>" +
            "<w:r><w:t>line</w:t></w:r></w:p>"),
        IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t xml:space=\"preserve\">hello brave </w:t></w:r>" +
            "<w:bookmarkStart w:id=\"1\" w:name=\"bm1\"/><w:bookmarkEnd w:id=\"1\"/>" +
            "<w:r><w:t>world</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t xml:space=\"preserve\">second edited </w:t></w:r>" +
            "<w:bookmarkStart w:id=\"2\" w:name=\"bm2\"/><w:bookmarkEnd w:id=\"2\"/>" +
            "<w:r><w:t>line</w:t></w:r></w:p>"));

    private static (WmlDocument Left, WmlDocument Right) TextEditedAroundABookmark() => (
        IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t xml:space=\"preserve\">hello </w:t></w:r>" +
            "<w:bookmarkStart w:id=\"1\" w:name=\"bm1\"/><w:bookmarkEnd w:id=\"1\"/>" +
            "<w:r><w:t>world</w:t></w:r></w:p>"),
        IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t xml:space=\"preserve\">hello brave </w:t></w:r>" +
            "<w:bookmarkStart w:id=\"1\" w:name=\"bm1\"/><w:bookmarkEnd w:id=\"1\"/>" +
            "<w:r><w:t>world</w:t></w:r></w:p>"));

    /// <summary>
    /// Two headers whose main-part relationship order is REVERSED relative to their part-name
    /// digits, reproducing ordinary Word output where hdr digits and part-collection order
    /// disagree (fixture DB002's shape).
    /// </summary>
    private static WmlDocument WithTwoHeadersRelsReversed(
        string firstHeaderInnerXml, string secondHeaderText)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var headerOne = main.AddNewPart<HeaderPart>();
            var headerOneId = main.GetIdOfPart(headerOne);
            WritePartXml(headerOne, $"<w:hdr xmlns:w=\"{W}\">{firstHeaderInnerXml}</w:hdr>");

            var headerTwo = main.AddNewPart<HeaderPart>();
            var headerTwoId = main.GetIdOfPart(headerTwo);
            WritePartXml(headerTwo,
                $"<w:hdr xmlns:w=\"{W}\"><w:p><w:r><w:t xml:space=\"preserve\">{secondHeaderText}</w:t></w:r></w:p></w:hdr>");

            WritePartXml(main,
                $"<w:document xmlns:w=\"{W}\" xmlns:r=\"{RNs}\"><w:body>" +
                "<w:p><w:r><w:t>Body paragraph</w:t></w:r></w:p>" +
                "<w:sectPr><w:titlePg/>" +
                $"<w:headerReference w:type=\"default\" r:id=\"{headerOneId}\"/>" +
                $"<w:headerReference w:type=\"first\" r:id=\"{headerTwoId}\"/>" +
                "</w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("two-headers.docx", ReverseMainRels(ms.ToArray()));
    }

    private static byte[] ReverseMainRels(byte[] package)
    {
        using var stream = new MemoryStream();
        stream.Write(package);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("word/_rels/document.xml.rels")!;
            XDocument rels;
            using (var input = entry.Open())
                rels = XDocument.Load(input);
            var reversed = rels.Root!.Elements().Reverse().ToArray();
            rels.Root.RemoveNodes();
            rels.Root.Add(reversed);
            using var output = entry.Open();
            output.SetLength(0);
            var bytes = Encoding.UTF8.GetBytes(rels.ToString(SaveOptions.DisableFormatting));
            output.Write(bytes);
        }
        return stream.ToArray();
    }

    private static void WritePartXml(OpenXmlPart part, string xml)
    {
        using var partStream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(partStream, Encoding.UTF8, 1024, leaveOpen: false);
        writer.Write(xml);
    }
}
