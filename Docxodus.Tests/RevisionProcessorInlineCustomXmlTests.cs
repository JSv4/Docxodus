// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Rejecting revisions reverses a <c>w:ins</c> or <c>w:del</c> held by an inline (run-level)
/// <c>w:customXml</c> element (issue #1176). Reject works by turning each insertion into a deletion
/// and each deletion into an insertion, then accepting; the turn-around skipped revisions whose
/// container was an inline custom-XML element, so reject kept the inserted text and dropped the
/// original text there.
/// </summary>
public class RevisionProcessorInlineCustomXmlTests
{
    private const string Stamp = "w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\"";

    /// <summary>One insertion and one deletion, the content an inline wrapper holds in each case.</summary>
    private const string Revisions =
        $"<w:ins w:id=\"1\" {Stamp}><w:r><w:t>added</w:t></w:r></w:ins>" +
        $"<w:del w:id=\"2\" {Stamp}><w:r><w:delText>original</w:delText></w:r></w:del>";

    private static string InlineCustomXml(string content) =>
        $"<w:customXml w:uri=\"urn:test\" w:element=\"term\"><w:customXmlPr><w:attr w:name=\"id\" w:val=\"t1\"/></w:customXmlPr>{content}</w:customXml>";

    /// <summary>Paragraph shapes that put the revisions inside an inline custom-XML element, alone or
    /// nested with the other run containers (each case's text is the paragraph's whole content).</summary>
    private static readonly Dictionary<string, string> ShapeBodies = new()
    {
        ["in the paragraph"] = "<w:p>" + InlineCustomXml(Revisions) + "</w:p>",
        ["in another inline custom-XML element"] = "<w:p>" + InlineCustomXml(InlineCustomXml(Revisions)) + "</w:p>",
        ["in a hyperlink"] = "<w:p><w:hyperlink w:anchor=\"target\">" + InlineCustomXml(Revisions) + "</w:hyperlink></w:p>",
        ["in a simple field"] = "<w:p><w:fldSimple w:instr=\" AUTHOR \">" + InlineCustomXml(Revisions) + "</w:fldSimple></w:p>",
        ["in a smart tag"] = "<w:p><w:smartTag w:element=\"place\">" + InlineCustomXml(Revisions) + "</w:smartTag></w:p>",
        ["in an inline content control"] = "<w:p><w:sdt><w:sdtPr/><w:sdtContent>" + InlineCustomXml(Revisions) + "</w:sdtContent></w:sdt></w:p>",
        ["around an inline content control"] = "<w:p>" + InlineCustomXml("<w:sdt><w:sdtPr/><w:sdtContent>" + Revisions + "</w:sdtContent></w:sdt>") + "</w:p>",
        ["around a smart tag"] = "<w:p>" + InlineCustomXml("<w:smartTag w:element=\"place\">" + Revisions + "</w:smartTag>") + "</w:p>",
        ["in a paragraph inside a block custom-XML element"] = "<w:customXml w:element=\"clause\"><w:p>" + InlineCustomXml(Revisions) + "</w:p></w:customXml>",
        ["in a table cell"] = "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr><w:tc><w:p>" + InlineCustomXml(Revisions) + "</w:p></w:tc></w:tr></w:tbl>",
    };

    public static TheoryData<string> Shapes() => new(ShapeBodies.Keys);

    [Theory]
    [MemberData(nameof(Shapes))]
    public void Reject_RevisionsInsideAnInlineCustomXml_RestoresTheOriginalText(string shape)
    {
        var input = Build(ShapeBodies[shape]);

        var body = BodyOf(RevisionProcessor.RejectRevisions(input));

        Assert.Equal("original", TextOf(body));
        AssertNoRevisionMarkup(body);
        AssertWrappersKept(input, body);
    }

    [Theory]
    [MemberData(nameof(Shapes))]
    public void Accept_RevisionsInsideAnInlineCustomXml_KeepsTheInsertedText(string shape)
    {
        var input = Build(ShapeBodies[shape]);

        var body = BodyOf(RevisionProcessor.AcceptRevisions(input));

        Assert.Equal("added", TextOf(body));
        AssertNoRevisionMarkup(body);
        AssertWrappersKept(input, body);
    }

    [Fact]
    public void Reject_InlineCustomXmlBetweenUnchangedRuns_RestoresTheOriginalSentence()
    {
        var bodyXml =
            "<w:p><w:r><w:t xml:space=\"preserve\">The </w:t></w:r>" +
            InlineCustomXml(
                $"<w:del w:id=\"1\" {Stamp}><w:r><w:delText>Seller</w:delText></w:r></w:del>" +
                $"<w:ins w:id=\"2\" {Stamp}><w:r><w:t>Vendor</w:t></w:r></w:ins>") +
            "<w:r><w:t xml:space=\"preserve\"> shall deliver.</w:t></w:r></w:p>";

        var rejected = BodyOf(RevisionProcessor.RejectRevisions(Build(bodyXml)));
        var accepted = BodyOf(RevisionProcessor.AcceptRevisions(Build(bodyXml)));

        Assert.Equal("The Seller shall deliver.", TextOf(rejected));
        Assert.Equal("The Vendor shall deliver.", TextOf(accepted));
        Assert.Equal("Seller", Assert.Single(rejected.Descendants(W.customXml)).Value);
    }

    [Fact]
    public void Accept_HyperlinkWhoseRunsSitInAnInlineCustomXml_KeepsTheLink()
    {
        var bodyXml = "<w:p><w:hyperlink w:anchor=\"target\">" +
            InlineCustomXml("<w:r><w:t>linked</w:t></w:r>") + "</w:hyperlink></w:p>";

        var body = BodyOf(RevisionProcessor.AcceptRevisions(Build(bodyXml)));

        Assert.Equal("linked", Assert.Single(body.Descendants(W.hyperlink)).Value);
        Assert.Single(body.Descendants(W.customXml));
    }

    [Fact]
    public void Reject_InsertedParagraphHoldingAnInlineCustomXml_RemovesTheParagraph()
    {
        var bodyXml =
            "<w:p><w:r><w:t>kept</w:t></w:r></w:p>" +
            $"<w:p><w:pPr><w:rPr><w:ins w:id=\"1\" {Stamp}/></w:rPr></w:pPr>" +
            InlineCustomXml($"<w:ins w:id=\"2\" {Stamp}><w:r><w:t>new paragraph</w:t></w:r></w:ins>") + "</w:p>" +
            "<w:p><w:r><w:t>after</w:t></w:r></w:p>";

        var body = BodyOf(RevisionProcessor.RejectRevisions(Build(bodyXml)));

        Assert.Equal(new[] { "kept", "after" }, body.Elements(W.p).Select(p => p.Value));
    }

    [Fact]
    public void Accept_DeletedParagraphMarkAfterAnInlineCustomXml_JoinsTheParagraphs()
    {
        var bodyXml =
            $"<w:p><w:pPr><w:rPr><w:del w:id=\"1\" {Stamp}/></w:rPr></w:pPr>" +
            InlineCustomXml("<w:r><w:t xml:space=\"preserve\">first </w:t></w:r>") + "</w:p>" +
            "<w:p><w:r><w:t>second</w:t></w:r></w:p>";

        var body = BodyOf(RevisionProcessor.AcceptRevisions(Build(bodyXml)));

        Assert.Equal("first second", Assert.Single(body.Elements(W.p)).Value);
    }

    [Fact]
    public void Reject_InlineCustomXmlInsertedWithItsContent_IsRemoved()
    {
        // Word marks an inserted custom-XML element with an insert range around each of its two tags.
        var bodyXml = "<w:p><w:r><w:t xml:space=\"preserve\">A </w:t></w:r>" +
            $"<w:customXmlInsRangeStart w:id=\"10\" {Stamp}/>" +
            "<w:customXml w:uri=\"urn:test\" w:element=\"term\"><w:customXmlInsRangeEnd w:id=\"10\"/>" +
            $"<w:ins w:id=\"1\" {Stamp}><w:r><w:t>added</w:t></w:r></w:ins>" +
            $"<w:customXmlInsRangeStart w:id=\"11\" {Stamp}/></w:customXml><w:customXmlInsRangeEnd w:id=\"11\"/>" +
            "<w:r><w:t xml:space=\"preserve\"> Z</w:t></w:r></w:p>";

        var rejected = BodyOf(RevisionProcessor.RejectRevisions(Build(bodyXml)));
        var accepted = BodyOf(RevisionProcessor.AcceptRevisions(Build(bodyXml)));

        Assert.Equal("A  Z", TextOf(rejected));
        Assert.Empty(rejected.Descendants(W.customXml));
        Assert.Equal("A added Z", TextOf(accepted));
        Assert.Single(accepted.Descendants(W.customXml));
    }

    private static string TextOf(XElement body) =>
        string.Concat(body.Descendants().Where(e => e.Name == W.t || e.Name == W.delText).Select(e => e.Value));

    private static void AssertNoRevisionMarkup(XElement body)
    {
        Assert.Empty(body.Descendants(W.ins));
        Assert.Empty(body.Descendants(W.del));
        Assert.Empty(body.Descendants(W.delText));
    }

    /// <summary>Accept and reject resolve the revisions, never the wrappers around them.</summary>
    private static void AssertWrappersKept(WmlDocument input, XElement body)
    {
        var before = BodyOf(input);
        foreach (var name in new[] { W.customXml, W.hyperlink, W.fldSimple, W.smartTag, W.sdt })
            Assert.Equal(before.Descendants(name).Count(), body.Descendants(name).Count());
    }

    private static WmlDocument Build(string bodyXml) => IrTestDocuments.FromBodyXml(bodyXml + "<w:sectPr/>");

    private static XElement BodyOf(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }
}
