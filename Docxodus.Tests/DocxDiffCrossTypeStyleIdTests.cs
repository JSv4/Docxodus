// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
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
/// A style id names one style across all types, so a comparison whose two documents use one id for styles of
/// different types (a paragraph style in one, a character style in the other) must not write both under that id
/// (issue #1184). The original document's style keeps the id. The revised document's style is left out when
/// nothing in the redline uses it, as Word's compare does; when inserted content uses it, it is imported under
/// a fresh id and those references follow it.
/// </summary>
public class DocxDiffCrossTypeStyleIdTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    private const string Normal =
        "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>";

    private const string ParagraphDocId =
        "<w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"DocID\"><w:name w:val=\"DocID\"/>" +
        "<w:basedOn w:val=\"Normal\"/><w:rPr><w:sz w:val=\"16\"/></w:rPr></w:style>";

    private const string CharacterDocId =
        "<w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"DocID\"><w:name w:val=\"DocID\"/>" +
        "<w:rPr><w:b/><w:sz w:val=\"16\"/></w:rPr></w:style>";

    public static TheoryData<string> Engines => new() { "DocxCompare", "DocxDiff" };

    private static WmlDocument Compare(string engine, WmlDocument left, WmlDocument right) =>
        engine == "DocxCompare" ? DocxCompare.Compare(left, right) : DocxDiff.Compare(left, right);

    private static string P(string text, string pPr = "") =>
        $"<w:p>{pPr}<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static string StyledRun(string styleId, string text) =>
        $"<w:r><w:rPr><w:rStyle w:val=\"{styleId}\"/></w:rPr><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private static WmlDocument Doc(string body, string styles, string? footnotes = null) =>
        IrTestDocuments.FromParts(body + "<w:sectPr/>", Normal + styles, footnotesInnerXml: footnotes);

    // ------------------------------------------------------------------------- unreferenced revised style

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_RevisedStyleOfAnotherTypeUnderTheOriginalsId_IsLeftOutWhenUnused(string engine)
    {
        var left = Doc(P("The original paragraph."), ParagraphDocId);
        var right = Doc(P("The revised paragraph."), CharacterDocId);

        var redline = Compare(engine, left, right);

        var docId = Assert.Single(Styles(redline), s => StyleId(s) == "DocID");
        Assert.Equal("paragraph", (string?)docId.Attribute(W + "type"));
        AssertNoCharacterStyleNamedDocId(redline);
        Assert.Equal("16", (string?)docId.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val"));
        AssertUniqueStyleIds(redline);
        AssertNoNewValidationErrors(redline, left, right);
        Assert.Equal("The revised paragraph.", BodyText(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal("The original paragraph.", BodyText(RevisionProcessor.RejectRevisions(redline)));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_OriginalStillUsingItsStyle_KeepsItAndLeavesTheUnusedRevisedStyleOut(string engine)
    {
        var docIdParagraph = P("Document 123", "<w:pPr><w:pStyle w:val=\"DocID\"/></w:pPr>");
        var left = Doc(P("The original paragraph.") + docIdParagraph, ParagraphDocId);
        var right = Doc(P("The revised paragraph.") + docIdParagraph.Replace("<w:pStyle w:val=\"DocID\"/>", ""),
            CharacterDocId);

        var redline = Compare(engine, left, right);

        Assert.Equal("paragraph", (string?)Assert.Single(Styles(redline), s => StyleId(s) == "DocID").Attribute(W + "type"));
        AssertNoCharacterStyleNamedDocId(redline);
        AssertUniqueStyleIds(redline);
        AssertNoNewValidationErrors(redline, left, right);
        Assert.Contains("DocID", ParagraphStyles(RevisionProcessor.RejectRevisions(redline)));
    }

    // --------------------------------------------------------------------------- referenced revised style

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_InsertedRunUsingTheRevisedCharacterStyle_ImportsItUnderAFreshId(string engine)
    {
        var left = Doc(P("The original paragraph."), ParagraphDocId);
        var right = Doc(P("The original paragraph.") + $"<w:p>{StyledRun("DocID", "Document 123")}</w:p>", CharacterDocId);

        var redline = Compare(engine, left, right);

        AssertUniqueStyleIds(redline);
        AssertNoNewValidationErrors(redline, left, right);
        Assert.Equal("paragraph", (string?)Assert.Single(Styles(redline), s => StyleId(s) == "DocID").Attribute(W + "type"));
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var runStyle = Assert.Single(RunStyles(accepted));
        Assert.NotEqual("DocID", runStyle);
        var imported = Assert.Single(Styles(accepted), s => StyleId(s) == runStyle);
        Assert.Equal("character", (string?)imported.Attribute(W + "type"));
        Assert.NotNull(imported.Element(W + "rPr")?.Element(W + "b"));
        Assert.Equal("The original paragraph.Document 123", BodyText(accepted));
        Assert.Equal("The original paragraph.", BodyText(RevisionProcessor.RejectRevisions(redline)));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_InsertedParagraphUsingTheRevisedParagraphStyle_ImportsItUnderAFreshId(string engine)
    {
        // The original names a character style DocID; the revised document holds its character style under
        // RunID and adds a paragraph style DocID that its inserted paragraph uses.
        var left = Doc($"<w:p>{StyledRun("DocID", "Kept run")}</w:p>", CharacterDocId);
        var right = Doc($"<w:p>{StyledRun("RunID", "Kept run")}</w:p>" +
            P("Inserted paragraph", "<w:pPr><w:pStyle w:val=\"DocID\"/></w:pPr>"),
            CharacterDocId.Replace("DocID", "RunID") + ParagraphDocId);

        var redline = Compare(engine, left, right);

        AssertUniqueStyleIds(redline);
        AssertNoNewValidationErrors(redline, left, right);
        Assert.Equal("character", (string?)Assert.Single(Styles(redline), s => StyleId(s) == "DocID").Attribute(W + "type"));
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var paragraphStyle = Assert.Single(ParagraphStyles(accepted));
        Assert.NotEqual("DocID", paragraphStyle);
        Assert.Equal("paragraph", (string?)Assert.Single(Styles(accepted), s => StyleId(s) == paragraphStyle).Attribute(W + "type"));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_RevisedStyleBasedOnTheCollidingStyle_FollowsItToTheFreshId(string engine)
    {
        const string derived =
            "<w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"DocIDItalic\"><w:name w:val=\"DocID Italic\"/>" +
            "<w:basedOn w:val=\"DocID\"/><w:rPr><w:i/></w:rPr></w:style>";
        var left = Doc(P("The original paragraph."), ParagraphDocId);
        var right = Doc(P("The original paragraph.") + $"<w:p>{StyledRun("DocIDItalic", "Document 123")}</w:p>",
            CharacterDocId + derived);

        var redline = Compare(engine, left, right);

        AssertUniqueStyleIds(redline);
        AssertNoNewValidationErrors(redline, left, right);
        var styles = Styles(redline).ToList();
        var italic = Assert.Single(styles, s => StyleId(s) == "DocIDItalic");
        var basedOn = (string?)italic.Element(W + "basedOn")?.Attribute(W + "val");
        Assert.NotEqual("DocID", basedOn);
        Assert.Equal("character", (string?)Assert.Single(styles, s => StyleId(s) == basedOn).Attribute(W + "type"));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_RevisedStyleOnlyLinkedFromAnotherCopiedStyle_IsCopiedUnderAFreshId(string engine)
    {
        // The revised paragraph style Quote is copied in (the original has no Quote); its w:link names the
        // character style DocID, so that style is used even though no run names it.
        const string quote =
            "<w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"Quote\"><w:name w:val=\"Quote\"/>" +
            "<w:basedOn w:val=\"Normal\"/><w:link w:val=\"DocID\"/><w:rPr><w:i/></w:rPr></w:style>";
        var left = Doc(P("The original paragraph."), ParagraphDocId);
        var right = Doc(P("The revised paragraph."), CharacterDocId + quote);

        var redline = Compare(engine, left, right);

        AssertUniqueStyleIds(redline);
        AssertNoNewValidationErrors(redline, left, right);
        var styles = Styles(redline).ToList();
        var link = (string?)Assert.Single(styles, s => StyleId(s) == "Quote").Element(W + "link")?.Attribute(W + "val");
        Assert.NotEqual("DocID", link);
        Assert.Equal("character", (string?)Assert.Single(styles, s => StyleId(s) == link).Attribute(W + "type"));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_RevisedNumberingLevelNamingTheCollidingParagraphStyle_FollowsItToTheFreshId(string engine)
    {
        const string characterLead =
            "<w:style w:type=\"character\" w:customStyle=\"1\" w:styleId=\"ListLead\"><w:name w:val=\"ListLead\"/>" +
            "<w:rPr><w:b/></w:rPr></w:style>";
        const string paragraphLead =
            "<w:style w:type=\"paragraph\" w:customStyle=\"1\" w:styleId=\"ListLead\"><w:name w:val=\"ListLead\"/>" +
            "<w:basedOn w:val=\"Normal\"/><w:pPr><w:numPr><w:numId w:val=\"1\"/></w:numPr></w:pPr></w:style>";
        const string numbering =
            "<w:abstractNum w:abstractNumId=\"0\"><w:multiLevelType w:val=\"singleLevel\"/>" +
            "<w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:pStyle w:val=\"ListLead\"/>" +
            "<w:lvlText w:val=\"%1.\"/><w:lvlJc w:val=\"left\"/></w:lvl></w:abstractNum>" +
            "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>";
        var left = Doc(P("The original paragraph."), characterLead);
        var right = IrTestDocuments.FromParts(
            P("The original paragraph.") + P("First item", "<w:pPr><w:pStyle w:val=\"ListLead\"/></w:pPr>") + "<w:sectPr/>",
            Normal + paragraphLead, numberingInnerXml: numbering);

        var redline = Compare(engine, left, right);

        AssertUniqueStyleIds(redline);
        AssertNoNewValidationErrors(redline, left, right);
        var styles = Styles(redline).ToList();
        Assert.Equal("character", (string?)Assert.Single(styles, s => StyleId(s) == "ListLead").Attribute(W + "type"));
        var paragraphStyle = Assert.Single(ParagraphStyles(RevisionProcessor.AcceptRevisions(redline)));
        Assert.Equal("paragraph", (string?)Assert.Single(styles, s => StyleId(s) == paragraphStyle).Attribute(W + "type"));
        var levelStyles = Part(redline, "word/numbering.xml").Descendants(W + "lvl").Elements(W + "pStyle")
            .Select(e => (string?)e.Attribute(W + "val")).ToList();
        Assert.Equal(new[] { paragraphStyle }, levelStyles);
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_OriginalAndInsertedContentEachUsingTheirOwnStyle_KeepBoth(string engine)
    {
        var docIdParagraph = P("Document 123", "<w:pPr><w:pStyle w:val=\"DocID\"/></w:pPr>");
        var left = Doc(docIdParagraph, ParagraphDocId);
        var right = Doc(P("Document 123") + $"<w:p>{StyledRun("DocID", "Reference 456")}</w:p>", CharacterDocId);

        var redline = Compare(engine, left, right);

        AssertUniqueStyleIds(redline);
        AssertNoNewValidationErrors(redline, left, right);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        Assert.Equal(new[] { "DocID" }, ParagraphStyles(rejected));
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var runStyle = Assert.Single(RunStyles(accepted));
        Assert.Equal("character", (string?)Assert.Single(Styles(accepted), s => StyleId(s) == runStyle).Attribute(W + "type"));
    }

    [Theory]
    [MemberData(nameof(Engines))]
    public void Compare_InsertedFootnoteRunUsingTheRevisedStyle_FollowsItToTheFreshId(string engine)
    {
        const string separators =
            "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
            "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>";
        const string body = "<w:p><w:r><w:t>Body</w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>";
        var left = Doc(body, ParagraphDocId, separators + "<w:footnote w:id=\"1\">" + P("Note") + "</w:footnote>");
        var right = Doc(body, CharacterDocId,
            separators + $"<w:footnote w:id=\"1\"><w:p><w:r><w:t xml:space=\"preserve\">Note </w:t></w:r>{StyledRun("DocID", "cited")}</w:p></w:footnote>");

        var redline = Compare(engine, left, right);

        AssertUniqueStyleIds(redline);
        AssertNoNewValidationErrors(redline, left, right);
        var noteRunStyles = Part(redline, "word/footnotes.xml").Descendants(W + "rStyle")
            .Select(s => (string?)s.Attribute(W + "val")).ToList();
        var runStyle = Assert.Single(noteRunStyles);
        Assert.Equal("character", (string?)Assert.Single(Styles(redline), s => StyleId(s) == runStyle).Attribute(W + "type"));
    }

    [Fact]
    public void Consolidate_ReviewerStyleOfAnotherTypeUnderTheBasesId_KeepsStyleIdsUnique()
    {
        var baseDoc = Doc(P("The original paragraph."), ParagraphDocId);
        // Reviewer A's DocID (italic) is unused and left out; reviewer B's (bold) is used and copied in.
        var unused = Doc(P("The revised paragraph."), CharacterDocId.Replace("<w:b/>", "<w:i/>"));
        var used = Doc(P("The original paragraph.") + $"<w:p>{StyledRun("DocID", "Document 123")}</w:p>", CharacterDocId);

        var consolidated = DocxCompare.Consolidate(baseDoc, new[]
        {
            new DocxDiffReviewer { Author = "Reviewer A", Document = unused },
            new DocxDiffReviewer { Author = "Reviewer B", Document = used },
        });

        AssertUniqueStyleIds(consolidated);
        AssertNoNewValidationErrors(consolidated, baseDoc, unused, used);
        Assert.Equal("paragraph", (string?)Assert.Single(Styles(consolidated), s => StyleId(s) == "DocID").Attribute(W + "type"));
        var runStyle = Assert.Single(RunStyles(consolidated));
        var imported = Assert.Single(Styles(consolidated), s => StyleId(s) == runStyle);
        Assert.Equal("character", (string?)imported.Attribute(W + "type"));
        Assert.NotNull(imported.Element(W + "rPr")?.Element(W + "b"));
        Assert.Null(imported.Element(W + "rPr")?.Element(W + "i"));
    }

    [Fact]
    public void Consolidate_TwoReviewersUsingTheirOwnStyleUnderTheBasesId_TheFirstReviewersCopyServesBoth()
    {
        // Same-type ids already resolve to the first reviewer's definition; a cross-type id does the same
        // rather than copying one fresh style per reviewer.
        var baseDoc = Doc(P("The original paragraph."), ParagraphDocId);
        var first = Doc(P("The original paragraph.") + $"<w:p>{StyledRun("DocID", "First addition")}</w:p>", CharacterDocId);
        var second = Doc($"<w:p>{StyledRun("DocID", "Second addition")}</w:p>" + P("The original paragraph."),
            CharacterDocId.Replace("<w:b/>", "<w:i/>"));

        var consolidated = DocxCompare.Consolidate(baseDoc, new[]
        {
            new DocxDiffReviewer { Author = "Reviewer A", Document = first },
            new DocxDiffReviewer { Author = "Reviewer B", Document = second },
        });

        AssertUniqueStyleIds(consolidated);
        AssertNoNewValidationErrors(consolidated, baseDoc, first, second);
        var copied = Assert.Single(Styles(consolidated), s => (string?)s.Attribute(W + "type") == "character" &&
            (string?)s.Element(W + "name")?.Attribute(W + "val") == "DocID");
        Assert.NotNull(copied.Element(W + "rPr")?.Element(W + "b"));
        Assert.Equal(new[] { StyleId(copied), StyleId(copied) }, RunStyles(consolidated));
    }

    // ---------------------------------------------------------------------------------------- helpers

    private static XDocument Part(WmlDocument document, string path)
    {
        using var package = new ZipArchive(new MemoryStream(document.DocumentByteArray));
        using var stream = package.GetEntry(path)!.Open();
        return XDocument.Load(stream);
    }

    private static IEnumerable<XElement> Styles(WmlDocument document) =>
        Part(document, "word/styles.xml").Root!.Elements(W + "style");

    private static string? StyleId(XElement style) => (string?)style.Attribute(W + "styleId");

    private static XElement Body(WmlDocument document) => Part(document, "word/document.xml").Root!.Element(W + "body")!;

    private static string BodyText(WmlDocument document) =>
        string.Concat(Body(document).Descendants(W + "t").Select(t => t.Value));

    private static List<string?> RunStyles(WmlDocument document) =>
        Body(document).Descendants(W + "rStyle").Select(s => (string?)s.Attribute(W + "val")).ToList();

    private static List<string?> ParagraphStyles(WmlDocument document) =>
        Body(document).Descendants(W + "pStyle").Select(s => (string?)s.Attribute(W + "val")).ToList();

    /// <summary>The revised document's character style DocID was left out, under any id.</summary>
    private static void AssertNoCharacterStyleNamedDocId(WmlDocument document) =>
        Assert.DoesNotContain(Styles(document), s => (string?)s.Attribute(W + "type") == "character" &&
            (string?)s.Element(W + "name")?.Attribute(W + "val") == "DocID");

    private static void AssertUniqueStyleIds(WmlDocument document)
    {
        var duplicates = Styles(document).GroupBy(StyleId).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
        Assert.Empty(duplicates);
    }

    /// <summary>The redline adds no validation error its inputs lack (compared by description).</summary>
    private static void AssertNoNewValidationErrors(WmlDocument redline, params WmlDocument[] inputs)
    {
        static HashSet<string> Errors(WmlDocument document)
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = WordprocessingDocument.Open(stream, false);
            return new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package)
                .Select(e => $"{e.Id}: {e.Description}").ToHashSet(StringComparer.Ordinal);
        }
        var known = inputs.SelectMany(Errors).ToHashSet(StringComparer.Ordinal);
        Assert.All(Errors(redline), e => Assert.Contains(e, known));
    }
}
