#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Cross-references to whole paragraphs (hidden <c>_Ref</c> bookmarks), PAGEREF and the
/// \n / \w number switches, and the editor's reference-field render marker.</summary>
public class DocxSessionCrossReferenceTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string[] BodyParagraphs(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body")
            .Select(t => t.Anchor.Id).ToArray();

    private static XElement Field(DocxSession session, string anchorId) =>
        Assert.Single(XElement.Parse(session.Raw.GetXml(anchorId)).Descendants(W + "fldSimple"));

    [Fact]
    public void XR001_AddReferenceBookmark_SpansTheParagraphAndIsReused()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(session);
        var text = session.GetAnchorInfo(paragraphs[0])!.VisibleText;

        var made = session.AddReferenceBookmark(paragraphs[0]);

        Assert.True(made.Success, made.Error?.Message);
        Assert.Matches(new Regex(@"^_Ref\d{9}$"), made.BookmarkName);
        var bookmark = Assert.Single(session.ListBookmarks(), b => b.Name == made.BookmarkName);
        Assert.Equal(text, bookmark.Text);
        Assert.Equal(new DocumentRange(paragraphs[0], 0, paragraphs[0], text.Length), bookmark.Range);

        var again = session.AddReferenceBookmark(paragraphs[0]);
        Assert.True(again.Success, again.Error?.Message);
        Assert.Equal(made.BookmarkName, again.BookmarkName);
        Assert.Single(session.ListBookmarks(), b => b.Name.StartsWith("_Ref", StringComparison.Ordinal));

        // The reference itself targets the reserved name, and its cached result is the paragraph's text.
        var reference = session.InsertCrossReference(paragraphs[1], 0, made.BookmarkName!,
            new CrossReferenceOptions { Hyperlink = true });
        Assert.True(reference.Success, reference.Error?.Message);
        Assert.Equal(text, Field(session, paragraphs[1]).Value);

        Assert.True(session.Undo());
        Assert.True(session.Undo());
        Assert.DoesNotContain(session.ListBookmarks(), b => b.Name == made.BookmarkName);
    }

    [Fact]
    public void XR002_AddReferenceBookmark_RefusesAMissingAnchor()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var result = session.AddReferenceBookmark("p:nope");
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.AnchorNotFound, result.Error!.Code);
    }

    [Fact]
    public void XR003_PageNumber_WritesPageRefWithTheCallersPageOrADirtyOne()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(session);
        Assert.True(session.AddBookmark("target", DocumentRange.In(paragraphs[0], new CharSpan(0, 4))).Success);

        var known = session.InsertCrossReference(paragraphs[1], 0, "target",
            new CrossReferenceOptions { PageNumber = true, Hyperlink = true, ReferenceNumber = true, CachedPage = "7" });
        Assert.True(known.Success, known.Error?.Message);
        var field = Field(session, paragraphs[1]);
        var instr = (string)field.Attribute(W + "instr")!;
        Assert.StartsWith(" PAGEREF target ", instr);
        Assert.Contains("\\h", instr);
        Assert.DoesNotContain("\\r", instr);
        Assert.Equal("7", field.Value);
        Assert.Null(field.Attribute(W + "dirty"));
        Assert.True(session.Undo());

        var unknown = session.InsertCrossReference(paragraphs[1], 0, "target",
            new CrossReferenceOptions { PageNumber = true, IncludePosition = true });
        Assert.True(unknown.Success, unknown.Error?.Message);
        field = Field(session, paragraphs[1]);
        Assert.Contains("\\p", (string)field.Attribute(W + "instr")!);
        Assert.Equal("on page 1", field.Value);
        Assert.Equal("true", (string?)field.Attribute(W + "dirty"));
    }

    [Fact]
    public void XR004_NumberSwitches_OneAtMostWithFullContextFirst()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(session);
        Assert.True(session.AddBookmark("target", DocumentRange.In(paragraphs[0], new CharSpan(0, 4))).Success);

        var full = session.InsertCrossReference(paragraphs[1], 0, "target",
            new CrossReferenceOptions { NumberFullContext = true, NumberNoContext = true, ReferenceNumber = true });
        Assert.True(full.Success, full.Error?.Message);
        var instr = (string)Field(session, paragraphs[1]).Attribute(W + "instr")!;
        Assert.Contains("\\w", instr);
        Assert.DoesNotContain("\\n", instr);
        Assert.DoesNotContain("\\r", instr);
        // An unnumbered target caches 0, as for \r.
        Assert.Equal("0", Field(session, paragraphs[1]).Value);
        Assert.True(session.Undo());

        var none = session.InsertCrossReference(paragraphs[1], 0, "target",
            new CrossReferenceOptions { NumberNoContext = true });
        Assert.True(none.Success, none.Error?.Message);
        Assert.Contains("\\n", (string)Field(session, paragraphs[1]).Attribute(W + "instr")!);
    }

    [Fact]
    public void XR005_StampReferenceFields_MarksSimpleAndComplexReferences()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(session);
        Assert.True(session.AddBookmark("target", DocumentRange.In(paragraphs[0], new CharSpan(0, 4))).Success);
        Assert.True(session.InsertCrossReference(paragraphs[1], 0, "target",
            new CrossReferenceOptions { Hyperlink = true }).Success);
        var bytes = session.Save();

        var plain = HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions());
        Assert.DoesNotContain("data-field=\"REF\"", plain);

        var simple = Reference(HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { StampReferenceFields = true }));
        Assert.Equal("target", (string?)simple.Attribute("data-field-target"));
        Assert.Equal("true", (string?)simple.Attribute("data-field-hyperlink"));

        // The same field as complex-field runs (how Word itself writes most of them).
        var complexBytes = Mutate(bytes, document =>
        {
            var part = document.MainDocumentPart!;
            var xml = part.GetXDocument();
            var field = xml.Descendants(W + "fldSimple").Single();
            var result = field.Elements(W + "r").Single();
            field.ReplaceWith(
                new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"))),
                new XElement(W + "r", new XElement(W + "instrText", " PAGEREF target ")),
                new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))),
                result,
                new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))));
            part.PutXDocument();
        });
        var complex = Reference(HtmlConversionOps.ConvertToHtml(complexBytes,
            new HtmlConversionOptions { StampReferenceFields = true }), "PAGEREF");
        Assert.Equal("target", (string?)complex.Attribute("data-field-target"));
        Assert.Null(complex.Attribute("data-field-hyperlink"));
        Assert.Equal(session.GetAnchorInfo(paragraphs[0])!.VisibleText[..4], complex.Value);
    }

    [Fact]
    public void XR006_ReferenceFieldInstr_ParsesTheThreeKinds()
    {
        Assert.Equal(("REF", "_Ref123456789", true), ReferenceFieldInstr.TryParse(" REF _Ref123456789 \\r \\h "));
        Assert.Equal(("PAGEREF", "x", false), ReferenceFieldInstr.TryParse("pageref x \\p"));
        Assert.Equal(("NOTEREF", "n", true), ReferenceFieldInstr.TryParse(" NOTEREF n \\H"));
        Assert.Null(ReferenceFieldInstr.TryParse(" PAGE "));
        Assert.Null(ReferenceFieldInstr.TryParse(" REF \\h "));
        Assert.Null(ReferenceFieldInstr.TryParse(null));
    }

    private static XElement Reference(string html, string kind = "REF") =>
        Assert.Single(XDocument.Parse(html).Descendants()
            .Where(e => (string?)e.Attribute("data-field") == kind));

    private static byte[] Mutate(byte[] bytes, Action<WordprocessingDocument> mutate)
    {
        var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true)) mutate(document);
        return stream.ToArray();
    }
}
