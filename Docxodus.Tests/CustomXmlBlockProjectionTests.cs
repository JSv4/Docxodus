// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Docxodus.Ir;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1114: paragraphs inside a block-level <c>w:customXml</c> wrapper are shown by the markdown
/// projection, and an edit inside one comes out of a comparison as a tracked, reversible change.
/// </summary>
public class CustomXmlBlockProjectionTests
{
    private const string Clause = """
        <w:customXml w:uri="urn:example" w:element="clause">
          <w:customXmlPr><w:attr w:name="id" w:val="c1"/></w:customXmlPr>
          <w:p><w:r><w:t>Clause one: backups are verified monthly.</w:t></w:r></w:p>
        </w:customXml>
        """;

    private const string Intro = "<w:p><w:r><w:t>Intro paragraph.</w:t></w:r></w:p>";
    private const string Closing = "<w:p><w:r><w:t>Clause two: restoration is tested once a year.</w:t></w:r></w:p>";

    // ---- projection ---------------------------------------------------------------------------

    [Fact]
    public void Project_ShowsAParagraphInsideABlockCustomXmlWrapper_InDocumentOrder()
    {
        using var session = new DocxSession(IrTestDocuments.FromBodyXml(Intro + Clause + Closing).DocumentByteArray);
        var projection = session.Project();
        var clauseId = AnchorWithText(session, projection, "Clause one: backups are verified monthly.");

        var markdown = projection.Markdown;
        var intro = markdown.IndexOf("Intro paragraph.", StringComparison.Ordinal);
        var clause = markdown.IndexOf("{#" + clauseId + "} Clause one: backups are verified monthly.", StringComparison.Ordinal);
        var closing = markdown.IndexOf("Clause two:", StringComparison.Ordinal);
        Assert.True(intro >= 0 && clause > intro && closing > clause, markdown);
    }

    [Fact]
    public void Project_IrEmitterMatchesTheConverter_ForNestedWrappersListsTablesAndControls()
    {
        var body = """
            <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Outer item.</w:t></w:r></w:p>
            <w:customXml w:uri="urn:example" w:element="section">
              <w:customXmlPr/>
              <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Wrapped item.</w:t></w:r></w:p>
              <w:customXml w:uri="urn:example" w:element="clause">
                <w:p><w:r><w:t>Nested wrapper text.</w:t></w:r></w:p>
                <w:tbl><w:tr><w:tc><w:p><w:r><w:t>Wrapped cell.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
              </w:customXml>
              <w:sdt><w:sdtPr><w:tag w:val="t"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Control in wrapper.</w:t></w:r></w:p></w:sdtContent></w:sdt>
            </w:customXml>
            <w:sdt><w:sdtPr><w:tag w:val="u"/></w:sdtPr><w:sdtContent>
              <w:customXml w:uri="urn:example" w:element="inner"><w:p><w:r><w:t>Wrapper in control.</w:t></w:r></w:p></w:customXml>
            </w:sdtContent></w:sdt>
            <w:p><w:r><w:t>Tail.</w:t></w:r></w:p>
            """;
        var bytes = IrTestDocuments.FromBodyXml(body).DocumentByteArray;
        var settings = new WmlToMarkdownConverterSettings();

        var oracle = WmlToMarkdownConverter.Convert(new WmlDocument("a.docx", bytes), settings);
        var ir = IrMarkdownEmitter.Emit(
            IrReader.Read(new WmlDocument("a.docx", bytes), new IrReaderOptions { RetainSources = false }),
            settings).ToProjection();

        Assert.Equal(oracle.Markdown, ir.Markdown);
        Assert.Equal(oracle.AnchorIndex.Keys, ir.AnchorIndex.Keys);
        Assert.Contains("Wrapped item.", oracle.Markdown);
        Assert.Contains("Nested wrapper text.", oracle.Markdown);
        Assert.Contains("| Wrapped cell. |", oracle.Markdown);
        // A content control's payload stays out of the markdown, wherever the wrappers nest.
        Assert.DoesNotContain("Control in wrapper.", oracle.Markdown);
        Assert.DoesNotContain("Wrapper in control.", oracle.Markdown);
    }

    [Fact]
    public void Project_ShowsWrappedParagraphsInFootnotesAndComments_AndTheIrEmitterAgrees()
    {
        var bytes = WrapNoteAndCommentParagraphs(IrTestDocuments.WithMixedReviewParts(
            "Header text", "Footer text", "Wrapped footnote text.", "Endnote text", "Wrapped comment text.").DocumentByteArray);
        var settings = new WmlToMarkdownConverterSettings();

        var oracle = WmlToMarkdownConverter.Convert(new WmlDocument("a.docx", bytes), settings);
        var ir = IrMarkdownEmitter.Emit(
            IrReader.Read(new WmlDocument("a.docx", bytes), new IrReaderOptions { RetainSources = false }),
            settings).ToProjection();

        Assert.Contains("Wrapped footnote text.", oracle.Markdown);
        Assert.Contains("Wrapped comment text.", oracle.Markdown);
        Assert.Equal(oracle.Markdown, ir.Markdown);
    }

    // ---- comparison ---------------------------------------------------------------------------

    [Fact]
    public void Compare_EditInsideABlockCustomXmlWrapper_IsTrackedAndReversible()
    {
        var left = IrTestDocuments.FromBodyXml(Intro + Clause + Closing);
        var right = IrTestDocuments.FromBodyXml(Intro + Clause.Replace("monthly", "weekly") + Closing);

        foreach (var redline in new[] { DocxCompare.Compare(left, right), DocxDiff.Compare(left, right, new DocxDiffSettings()) })
        {
            Assert.NotEmpty(Body(redline.DocumentByteArray).Descendants(W.del));
            Assert.NotEmpty(Body(redline.DocumentByteArray).Descendants(W.ins));
            AssertRoundTrip(left, right, redline);
        }
    }

    [Fact]
    public void Compare_EditInsideACustomXmlWrapperInATableCell_IsTrackedAndReversible()
    {
        static WmlDocument InCell(string clause) => IrTestDocuments.FromBodyXml(
            Intro + "<w:tbl><w:tr><w:tc>" + clause + "</w:tc></w:tr></w:tbl>" + Closing);
        var left = InCell(Clause);
        var right = InCell(Clause.Replace("monthly", "weekly"));

        var redline = DocxCompare.Compare(left, right);

        Assert.NotEmpty(Body(redline.DocumentByteArray).Descendants(W.ins));
        AssertRoundTrip(left, right, redline);
    }

    [Fact]
    public void Compare_EditInsideACustomXmlWrapperInAHeader_IsTrackedAndReversible()
    {
        var left = IrTestDocuments.FromBodyAndHeaderXml(Intro, Clause);
        var right = IrTestDocuments.FromBodyAndHeaderXml(Intro, Clause.Replace("monthly", "weekly"));

        var redline = DocxCompare.Compare(left, right);

        Assert.NotEmpty(Header(redline.DocumentByteArray).Descendants(W.ins));
        Assert.Equal(Shape(Header(right.DocumentByteArray)),
            Shape(Header(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray))));
        Assert.Equal(Shape(Header(left.DocumentByteArray)),
            Shape(Header(DocxDiffOps.RejectRevisions(redline.DocumentByteArray))));
    }

    [Fact]
    public void Compare_InsertedAndDeletedCustomXmlWrappers_CarryPairedRangeMarkers()
    {
        var without = IrTestDocuments.FromBodyXml(Intro + Closing);
        var with = IrTestDocuments.FromBodyXml(Intro + Clause + Closing);

        var inserted = DocxCompare.Compare(without, with);
        Assert.Equal(2, Body(inserted.DocumentByteArray).Descendants(W.customXmlInsRangeStart).Count());
        AssertRoundTrip(without, with, inserted);

        var deleted = DocxCompare.Compare(with, without);
        Assert.Equal(2, Body(deleted.DocumentByteArray).Descendants(W.customXmlDelRangeStart).Count());
        AssertRoundTrip(with, without, deleted);
    }

    [Fact]
    public void Compare_EditOutsideACustomXmlWrapper_LeavesTheWrapperUntouched()
    {
        var left = IrTestDocuments.FromBodyXml(Intro + Clause + Closing);
        var right = IrTestDocuments.FromBodyXml(Intro + Clause + Closing.Replace("once a year", "twice a year"));

        var redline = DocxCompare.Compare(left, right);

        var body = Body(redline.DocumentByteArray);
        var wrapper = Assert.Single(body.Descendants(W.customXml));
        Assert.DoesNotContain(wrapper.DescendantsAndSelf().Concat(body.Elements()), IsRevisionMarkup);
        Assert.NotEmpty(body.Descendants(W.ins));
        AssertRoundTrip(left, right, redline);
    }

    [Fact]
    public void Session_TrackedEditInsideAWrapper_AcceptedThenCompared_ReportsTheEdit()
    {
        var original = IrTestDocuments.FromBodyXml(Intro + Clause + Closing);
        using var session = new DocxSession(original.DocumentByteArray,
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var clauseId = AnchorWithText(session, session.Project(), "Clause one: backups are verified monthly.");
        var text = session.GetAnchorInfo(clauseId)!.TextPreview!;
        Assert.True(session.ReplaceTextAtSpan(clauseId, text.IndexOf("monthly", StringComparison.Ordinal), "monthly".Length, "weekly").Success);
        Assert.True(session.AcceptAllRevisions().Success);
        var accepted = new WmlDocument("accepted.docx", session.Save());

        var redline = DocxCompare.Compare(original, accepted);

        Assert.Contains(Body(redline.DocumentByteArray).Descendants(W.ins).Descendants(W.t), t => t.Value.Contains("weekly"));
        AssertRoundTrip(original, accepted, redline);
    }

    [Fact]
    public void Consolidate_ReviewerEditInsideACustomXmlWrapper_IsTrackedAndReversible()
    {
        var baseline = IrTestDocuments.FromBodyXml(Intro + Clause + Closing);
        var reviewed = IrTestDocuments.FromBodyXml(Intro + Clause.Replace("monthly", "weekly") + Closing);

        var redline = DocxCompare.Consolidate(baseline,
            [new DocxDiffReviewer { Document = reviewed, Author = "Reviewer" }]);

        Assert.NotEmpty(Body(redline.DocumentByteArray).Descendants(W.ins));
        AssertRoundTrip(baseline, reviewed, redline);
    }

    // ---- semantic diff ------------------------------------------------------------------------

    [Fact]
    public void SemanticDiff_ReportsTheTextEditInsideAWrapper_AndLabelsWrapperChangesAsCustomXml()
    {
        var without = IrTestDocuments.FromBodyXml(Intro + Closing);
        var with = IrTestDocuments.FromBodyXml(Intro + Clause + Closing);
        var edited = IrTestDocuments.FromBodyXml(Intro + Clause.Replace("monthly", "weekly") + Closing);

        var edit = SemanticDiff.Compare(with, edited).Changes;
        Assert.Contains(edit, change => change.Family == SemanticChangeFamily.Text);
        // A text edit leaves the wrapper itself as it was.
        Assert.DoesNotContain(edit, change => change.Family is SemanticChangeFamily.OpaquePackagePart
            or SemanticChangeFamily.ContentControl or SemanticChangeFamily.BlockStructure);

        var retagged = IrTestDocuments.FromBodyXml(Intro + Clause.Replace("w:val=\"c1\"", "w:val=\"c2\"") + Closing);
        Assert.Contains(SemanticDiff.Compare(with, retagged).Changes, change =>
            change.Family == SemanticChangeFamily.BlockStructure && change.Path == "custom_xml.envelope");

        var insert = SemanticDiff.Compare(without, with).Changes;
        Assert.Contains(insert, change => change.Operation == SemanticChangeOperation.Insert
            && change.Family == SemanticChangeFamily.BlockStructure && change.Path == "custom_xml");
        Assert.DoesNotContain(insert, change => change.Family == SemanticChangeFamily.ContentControl);
    }

    // ---- helpers ------------------------------------------------------------------------------

    private static string AnchorWithText(DocxSession session, MarkdownProjection projection, string text) =>
        projection.AnchorIndex.Keys.Single(id => session.GetAnchorInfo(id)?.TextPreview == text);

    private static bool IsRevisionMarkup(XElement e) =>
        e.Name == W.ins || e.Name == W.del || e.Name == W.delText
        || e.Name.LocalName.StartsWith("customXmlIns", StringComparison.Ordinal)
        || e.Name.LocalName.StartsWith("customXmlDel", StringComparison.Ordinal);

    /// <summary>accept(redline) reads as <paramref name="revised"/> and reject(redline) as
    /// <paramref name="original"/> — paragraph text and the custom-XML wrappers around it alike — with
    /// no revision markup left in either view.</summary>
    private static void AssertRoundTrip(WmlDocument original, WmlDocument revised, WmlDocument redline)
    {
        Assert.Equal(Shape(Body(revised.DocumentByteArray)), Shape(Body(DocxDiffOps.AcceptRevisions(redline.DocumentByteArray))));
        Assert.Equal(Shape(Body(original.DocumentByteArray)), Shape(Body(DocxDiffOps.RejectRevisions(redline.DocumentByteArray))));
    }

    /// <summary>Each paragraph's text, prefixed by the custom-XML wrappers around it (element name and
    /// properties), followed by every revision element left in the story.</summary>
    private static string Shape(XElement body)
    {
        var paragraphs = body.Descendants(W.p).Select(p =>
            string.Concat(p.Ancestors(W.customXml).Reverse().Select(Wrapper))
            + string.Concat(p.Descendants(W.t).Select(t => t.Value)));
        var revisions = body.Descendants().Where(IsRevisionMarkup).Select(e => e.Name.LocalName);
        return string.Join("|", paragraphs) + " // " + string.Join(",", revisions);

        static string Wrapper(XElement c) =>
            "[" + (string?)c.Attribute(W.element)
            + string.Concat(c.Elements(W.customXmlPr).Elements(W.attr)
                .Select(a => $" {(string?)a.Attribute(W.name)}={(string?)a.Attribute(W.val)}"))
            + "]";
    }

    /// <summary>Wrap the paragraphs of every footnote and comment in a <c>w:customXml</c> element.</summary>
    private static byte[] WrapNoteAndCommentParagraphs(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            foreach (var part in new OpenXmlPart[] { main.FootnotesPart!, main.WordprocessingCommentsPart! })
            {
                var xml = part.GetXDocument();
                foreach (var container in xml.Root!.Elements().Where(e => e.Name == W.footnote || e.Name == W.comment)
                             .Where(e => e.Attribute(W.type) is null).ToList())
                {
                    var paragraphs = container.Elements(W.p).ToList();
                    paragraphs.Remove();
                    container.Add(new XElement(W.customXml, new XAttribute(W.element, "note"), paragraphs));
                }
                part.PutXDocument();
            }
        }
        return stream.ToArray();
    }

    private static XElement Header(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return new XElement(document.MainDocumentPart!.HeaderParts.Single().GetXDocument().Root!);
    }

    private static XElement Body(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return new XElement(document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!);
    }
}
