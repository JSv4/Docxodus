#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// The editor has to be able to RENDER what it can edit. Footnote/endnote bodies live in their own
/// parts, so two things must hold for the browser editor to show and edit them: the editor's render
/// profile must emit the notes section and the citation markers, and the note paragraphs must carry
/// <c>data-anchor</c> stamps so they are addressable as ordinary editable blocks.
/// Test IDs use the DS34x range.
/// </summary>
public class DocxSessionNoteRenderTests
{
    private static HtmlConversionOptions EditorProfile() => new()
    {
        CssClassPrefix = "docx-",
        FabricateCssClasses = true,
        StampAnchors = true,
        RenderFootnotesAndEndnotes = true,
    };

    /// <summary>Body paragraph citing a footnote, plus a second note that is never cited.</summary>
    private static byte[] BuildCitedFootnoteDoc()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body(
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Body text citing a note.")),
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 1 }))));
            var fnPart = main.AddNewPart<FootnotesPart>();
            using var s = fnPart.GetStream(FileMode.Create);
            using var w = new StreamWriter(s);
            w.Write("""
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
                  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
                  <w:footnote w:id="1"><w:p><w:r><w:t>THE NOTE BODY.</w:t></w:r></w:p></w:footnote>
                </w:footnotes>
                """);
        }
        return ms.ToArray();
    }

    [Fact]
    public void DS340_EditorRender_EmitsTheCitationMarkerAndTheNotesSection()
    {
        using var session = new DocxSession(BuildCitedFootnoteDoc());
        var html = HtmlConversionOps.ConvertToHtml(session, EditorProfile());

        // The body citation renders as a numbered superscript marker…
        Assert.Contains("class=\"footnote-ref\"", html);
        Assert.Contains("<sup>1</sup>", html);
        // …and the note body renders in the footnotes section.
        Assert.Contains("THE NOTE BODY.", html);
    }

    [Fact]
    public void DS341_EditorRender_StampsAnchorsOnNoteParagraphs()
    {
        using var session = new DocxSession(BuildCitedFootnoteDoc());
        var html = HtmlConversionOps.ConvertToHtml(session, EditorProfile());

        // The note paragraph must be addressable, or the editor can show it and not edit it.
        var noteParaAnchor = session.Project().AnchorIndex.Values
            .Single(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "fn");
        Assert.Contains($"data-anchor=\"{noteParaAnchor.Unid}\"", html);
    }

    /// <summary>
    /// The anchors the render stamps must be the SAME ones the session resolves — the editor looks
    /// up each `data-anchor` unid in the projection to find the op target. A different Unid scheme
    /// (or a scope the projection doesn't index) leaves the block rendered but unwired.
    /// </summary>
    [Fact]
    public void DS342_StampedNoteAnchors_ResolveThroughTheSession()
    {
        using var session = new DocxSession(BuildCitedFootnoteDoc());
        var html = HtmlConversionOps.ConvertToHtml(session, EditorProfile());

        var index = session.Project().AnchorIndex;
        var stamped = Regex.Matches(html, "data-anchor=\"([^\"]+)\"")
            .Select(m => m.Groups[1].Value).ToList();
        Assert.NotEmpty(stamped);
        foreach (var unid in stamped)
            Assert.Contains(index.Values, t => t.Unid == unid);

        // And specifically: the note paragraph's unid is among them.
        var notePara = index.Values.Single(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "fn");
        Assert.Contains(notePara.Unid, stamped);
    }

    [Fact]
    public void DS343_EditingAStampedNoteParagraph_UpdatesTheNoteAndRerendersIt()
    {
        using var session = new DocxSession(BuildCitedFootnoteDoc());
        var notePara = session.Project().AnchorIndex.Values
            .Single(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "fn").Anchor.Id;

        // The full editing round trip the GUI performs: edit the block, re-render just that block.
        var edited = session.ReplaceText(notePara, "REWRITTEN NOTE BODY.");
        Assert.True(edited.Success, edited.Error?.Message);

        var newAnchor = edited.Modified.Single().Id;
        var blockHtml = HtmlConversionOps.RenderBlockHtml(session, newAnchor, EditorProfile());
        Assert.Contains("REWRITTEN NOTE BODY.", blockHtml);

        // …and the citation still resolves to the edited note in the full render.
        var full = HtmlConversionOps.ConvertToHtml(session, EditorProfile());
        Assert.Contains("REWRITTEN NOTE BODY.", full);
        Assert.Contains("class=\"footnote-ref\"", full);
        Assert.DoesNotContain("THE NOTE BODY.", full);
    }

    /// <summary>
    /// The stateless <see cref="HtmlConversionOps.RenderBlockHtml(byte[], string, HtmlConversionOptions)"/>
    /// overload has to resolve a note anchor too — it assigns Unids itself, and before the note parts
    /// were included a footnote-paragraph anchor could never be found on that path.
    /// </summary>
    [Fact]
    public void DS344_StatelessRenderBlockHtml_ResolvesANoteParagraphAnchor()
    {
        var bytes = BuildCitedFootnoteDoc();
        string anchor;
        using (var session = new DocxSession(bytes))
            anchor = session.Project().AnchorIndex.Values
                .Single(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "fn").Anchor.Id;

        var html = HtmlConversionOps.RenderBlockHtml(bytes, anchor, EditorProfile());
        Assert.Contains("THE NOTE BODY.", html);
    }

    /// <summary>
    /// Word-reserved notes must never surface as editable blocks. <c>separator</c> and
    /// <c>continuationSeparator</c> are filtered by the projector; <c>continuationNotice</c> — which
    /// real documents carry (the NVCA model certificate has one at id 1) — was NOT, so it projected
    /// as a user note and would render as a stray empty footnote with no citation.
    /// </summary>
    [Fact]
    public void DS345_ContinuationNotice_IsNotProjectedAsAUserNote()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body(
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Body.")),
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 2 }))));
            var fnPart = main.AddNewPart<FootnotesPart>();
            using var s = fnPart.GetStream(FileMode.Create);
            using var w = new StreamWriter(s);
            w.Write("""
                <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
                  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
                  <w:footnote w:type="continuationNotice" w:id="1"><w:p/></w:footnote>
                  <w:footnote w:id="2"><w:p><w:r><w:t>REAL NOTE.</w:t></w:r></w:p></w:footnote>
                </w:footnotes>
                """);
        }

        using var session = new DocxSession(ms.ToArray());
        var noteDefs = session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "fn").ToList();
        var only = Assert.Single(noteDefs);
        Assert.Equal("2", (string?)only.Resolve(session.LiveDocument)!
            .Attribute(System.Xml.Linq.XName.Get("id",
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main")));
    }
}
