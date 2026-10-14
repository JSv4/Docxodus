#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Cross-document relationship-id collisions in <see cref="DocxDiff.Compare"/>'s markup renderer.
/// A right-sourced clone's <c>r:id</c> may collide with a LEFT relationship of a different KIND
/// (part relationship, e.g. comments.xml or an image — not just another hyperlink): the import must
/// remap to a fresh id instead of recreating the right relationship under the taken id, which
/// System.IO.Packaging rejects with "'rIdN' ID conflicts with the ID of an existing relationship".
/// Regression coverage for the corpus family docx_lots_of_comments_* that crashed exactly there.
/// </summary>
public class DocxDiffRelationshipRemapTests
{
    private const string RightTarget = "https://example.com/right-target";

    private static WmlDocument LeftWithCommentsPartAt(string relId, params string[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                paragraphs.Select(text => new Paragraph(new Run(new Text(text))))));
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()));
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            // Occupy the colliding id with a NON-hyperlink (part) relationship.
            mainPart.AddNewPart<WordprocessingCommentsPart>(relId).Comments = new Comments();
            doc.Save();
        }
        return new WmlDocument("left.docx", stream.ToArray());
    }

    private static WmlDocument RightWithHyperlinkAt(string relId, string linkedText, params string[] leadingParagraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body(leadingParagraphs.Select(text => new Paragraph(new Run(new Text(text)))));
            body.Append(new Paragraph(
                new Hyperlink(new Run(new Text(linkedText))) { Id = relId }));
            mainPart.Document = new Document(body);
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
                new ParagraphPropertiesDefault()));
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            mainPart.AddHyperlinkRelationship(new Uri(RightTarget), true, relId);
            doc.Save();
        }
        return new WmlDocument("right.docx", stream.ToArray());
    }

    private static List<string> BodyTexts(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var body = wdoc.MainDocumentPart?.Document.Body;
        return body is null
            ? new List<string>()
            : body.Descendants<Paragraph>().Select(p => p.InnerText).ToList();
    }

    [Fact]
    public void Compare_RightHyperlinkIdCollidesWithLeftPartRelationship_RemapsAndRoundTrips()
    {
        var left = LeftWithCommentsPartAt("rId12", "Common paragraph.");
        var right = RightWithHyperlinkAt("rId12", "the link", "Common paragraph.");

        var result = DocxDiff.Compare(left, right);

        var accepted = RevisionProcessor.AcceptRevisions(result);
        var rejected = RevisionProcessor.RejectRevisions(result);
        Assert.Equal(new List<string> { "Common paragraph.", "the link" }, BodyTexts(accepted));
        Assert.Equal(new List<string> { "Common paragraph." }, BodyTexts(rejected));

        // The inserted hyperlink must RESOLVE to the right-side target through whatever id the
        // renderer assigned (the literal id is free to change; the resolution contract is not).
        using var stream = new MemoryStream(accepted.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var main = wdoc.MainDocumentPart!;
        var hyperlink = main.Document.Body!.Descendants<Hyperlink>().Single();
        Assert.False(string.IsNullOrEmpty(hyperlink.Id));
        var rel = main.HyperlinkRelationships.Single(r => r.Id == hyperlink.Id!.Value);
        Assert.Equal(RightTarget, rel.Uri.ToString());
    }

    [Fact]
    public void Compare_RightExternalIdCollidesWithLeftPartRelationship_DoesNotThrow()
    {
        // Same collision class, external (non-hyperlink) relationship: an r:id on right-sourced
        // content whose id names a left PART. The compare must not throw and must round-trip.
        var left = LeftWithCommentsPartAt("rId9", "Shared.");
        var right = RightWithHyperlinkAt("rId9", "external link", "Shared.");

        var result = DocxDiff.Compare(left, right);

        Assert.Equal(new List<string> { "Shared.", "external link" },
            BodyTexts(RevisionProcessor.AcceptRevisions(result)));
        Assert.Equal(new List<string> { "Shared." },
            BodyTexts(RevisionProcessor.RejectRevisions(result)));
    }
}
