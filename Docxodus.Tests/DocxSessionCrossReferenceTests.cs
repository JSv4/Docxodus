using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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

    private const string Ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    /// <summary>List paragraphs (text, level), numbered by <paramref name="lvlTexts"/>, then "outside".</summary>
    private static DocxSession NumberedSession((string Format, string Text)[] lvlTexts, params (string Text, int Level)[] items)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            void Write(OpenXmlPart part, string xml)
            {
                using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write));
                writer.Write(xml);
            }
            Write(main.AddNewPart<StyleDefinitionsPart>(),
                $"<w:styles {Ns}><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style></w:styles>");
            Write(main.AddNewPart<NumberingDefinitionsPart>(), $"<w:numbering {Ns}><w:abstractNum w:abstractNumId=\"0\">"
                + string.Concat(lvlTexts.Select((l, i) => $"<w:lvl w:ilvl=\"{i}\"><w:start w:val=\"1\"/><w:numFmt w:val=\"{l.Format}\"/><w:lvlText w:val=\"{l.Text}\"/></w:lvl>"))
                + "</w:abstractNum><w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num></w:numbering>");
            Write(main, $"<w:document {Ns}><w:body>" + string.Concat(items.Select(i =>
                $"<w:p><w:pPr><w:numPr><w:ilvl w:val=\"{i.Level}\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>{i.Text}</w:t></w:r></w:p>"))
                + "<w:p><w:r><w:t>outside</w:t></w:r></w:p></w:body></w:document>");
        }
        return new DocxSession(stream.ToArray());
    }

    private static string AnchorOf(DocxSession session, string text) =>
        session.Project().AnchorIndex.Values.Single(t => t.Anchor.Scope == "body" && t.TextPreview == text).Anchor.Id;

    /// <summary>Reference <paramref name="target"/> from <paramref name="from"/>: the cached number, and "?" when dirty.</summary>
    private static string Number(DocxSession session, string target, string from, CrossReferenceOptions options)
    {
        var bookmark = session.AddReferenceBookmark(AnchorOf(session, target));
        Assert.True(bookmark.Success, bookmark.Error?.Message);
        var paragraph = AnchorOf(session, from);
        var result = session.InsertCrossReference(paragraph, 0, bookmark.BookmarkName!, options);
        Assert.True(result.Success, result.Error?.Message);
        var field = Field(session, paragraph);
        Assert.True(session.Undo());
        return field.Value + (field.Attribute(W + "dirty") is null ? "" : "?");
    }

    [Fact]
    public void XR007_NumberSwitches_ShowTheNumberWordShows()
    {
        using var session = NumberedSession(
            new[] { ("decimal", "%1."), ("lowerLetter", "(%2)"), ("lowerRoman", "(%3)") },
            ("one", 0), ("a", 1), ("i", 2), ("ii", 2), ("b", 1), ("two", 0));
        var none = new CrossReferenceOptions { NumberNoContext = true };
        var full = new CrossReferenceOptions { NumberFullContext = true };
        var relative = new CrossReferenceOptions { ReferenceNumber = true };

        Assert.Equal("(ii)", Number(session, "ii", "outside", none));
        Assert.Equal("1(a)(ii)", Number(session, "ii", "outside", full));
        Assert.Equal("1", Number(session, "one", "outside", full));
        Assert.Equal("2", Number(session, "two", "outside", relative));
        // relative to where the field is: the levels both share are left out
        Assert.Equal("(ii)", Number(session, "ii", "i", relative));
        Assert.Equal("(a)(ii)", Number(session, "ii", "b", relative));
        Assert.Equal("1(a)(ii)", Number(session, "ii", "two", relative));
        // from outside the list Word's text is not known: cached without context, and dirty
        Assert.Equal("(ii)?", Number(session, "ii", "outside", relative));
    }

    /// <summary>Word's own example for \\w: paragraph "ii." under "a." under "1." reads 1.a.ii.</summary>
    [Fact]
    public void XR008_FullContext_JoinsNumbersWithPeriods_AndStopsAtOneThatShowsItsUpperLevels()
    {
        var full = new CrossReferenceOptions { NumberFullContext = true };
        using (var periods = NumberedSession(
            new[] { ("decimal", "%1."), ("lowerLetter", "%2."), ("lowerRoman", "%3.") }, ("one", 0), ("a", 1), ("i", 2), ("ii", 2)))
            Assert.Equal("1.a.ii", Number(periods, "ii", "outside", full));
        using (var legal = NumberedSession(
            new[] { ("upperLetter", "%1."), ("decimal", "%2."), ("decimal", "%2.%3"), ("lowerLetter", "(%4)") },
            ("A", 0), ("1", 1), ("1.1", 2), ("a", 3)))
            Assert.Equal("1.1(a)", Number(legal, "a", "outside", full));
        // words in a joined number: Word's text is not known
        using (var words = NumberedSession(
            new[] { ("ordinalText", "%1"), ("lowerLetter", "(%2)") }, ("first", 0), ("a", 1)))
            Assert.Equal("(a)?", Number(words, "a", "outside", full));
    }

    /// <summary>The NVCA model charter holds 100+ REF numbers as Word cached them: every one the
    /// engine can work out must be Word's.</summary>
    [Fact]
    public void XR009_ReferenceNumber_AgreesWithTheResultsWordCached()
    {
        using var document = WordprocessingDocument.Open(Path.Combine(
            "../../../../TestFiles", "NVCA-Model-COI.docx"), false);
        var body = document.MainDocumentPart!.GetXDocument().Root!.Element(W + "body")!;
        int known = 0, unknown = 0;
        foreach (var (instr, result, paragraph) in CachedReferences(body))
        {
            var tokens = instr.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var numberSwitch = new[] { "\\n", "\\w", "\\r" }.FirstOrDefault(tokens.Contains);
            var target = body.Descendants(W + "bookmarkStart").FirstOrDefault(b => (string?)b.Attribute(W + "name") == tokens[1])
                ?.Ancestors(W + "p").FirstOrDefault();
            if (numberSwitch is null || target is null) continue;
            var number = ReferenceNumber.Resolve(document, target, paragraph, numberSwitch, tokens.Contains("\\t"));
            if (number is null) { unknown++; continue; }
            Assert.True(number == result, $"{instr}: Word {result}, engine {number}");
            known++;
        }
        Assert.True(known > 100 && unknown < 10, $"{known} known, {unknown} unknown");
    }

    private static IEnumerable<(string Instr, string Result, XElement Paragraph)> CachedReferences(XElement body)
    {
        var open = new Stack<(StringBuilder Instr, StringBuilder Result)>();
        bool inResult = false;
        foreach (var e in body.Descendants())
        {
            if (e.Name == W + "instrText" && open.Count > 0 && !inResult) open.Peek().Instr.Append(e.Value);
            else if (e.Name == W + "t" && open.Count > 0 && inResult) open.Peek().Result.Append(e.Value);
            else if (e.Name == W + "fldChar")
            {
                var type = (string?)e.Attribute(W + "fldCharType");
                if (type == "begin") { open.Push((new StringBuilder(), new StringBuilder())); inResult = false; }
                else if (type == "separate") inResult = true;
                else if (type == "end" && open.Count > 0)
                {
                    var (instr, result) = open.Pop();
                    inResult = open.Count > 0;
                    if (instr.ToString().TrimStart().StartsWith("REF ", StringComparison.Ordinal))
                        yield return (instr.ToString().Trim(), result.ToString(), e.Ancestors(W + "p").First());
                }
            }
        }
    }

    [Fact]
    public void XR010_AddReferenceBookmark_SpansAParagraphThatHoldsAField()
    {
        using var session = new DocxSession(Mutate(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), document =>
        {
            var part = document.MainDocumentPart!;
            part.GetXDocument().Descendants(W + "p").First().Add(
                new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "begin"))),
                new XElement(W + "r", new XElement(W + "instrText", " SEQ Figure ")),
                new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "separate"))),
                new XElement(W + "r", new XElement(W + "t", "1")),
                new XElement(W + "r", new XElement(W + "fldChar", new XAttribute(W + "fldCharType", "end"))));
            part.PutXDocument();
        }));
        var paragraph = BodyParagraphs(session)[0];

        var made = session.AddReferenceBookmark(paragraph);

        Assert.True(made.Success, made.Error?.Message);
        Assert.Equal(session.GetAnchorInfo(paragraph)!.VisibleText, session.ListBookmarks().Single(b => b.Name == made.BookmarkName).Text);
        Assert.Equal(made.BookmarkName, session.AddReferenceBookmark(paragraph).BookmarkName);
    }

    [Fact]
    public void XR011_EditorRenderOptions_ReferenceFields_StampsThem()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(session);
        Assert.True(session.AddBookmark("target", DocumentRange.In(paragraphs[0], new CharSpan(0, 4))).Success);
        Assert.True(session.InsertCrossReference(paragraphs[1], 0, "target").Success);
        int handle = SessionRegistry.OpenSession(session.Save(), new DocxSessionSettings());
        try
        {
            Assert.DoesNotContain("data-field-target", DocxSessionOps.RenderEditorHtml(handle, "{}"));
            Assert.Contains("data-field-target=\"target\"", DocxSessionOps.RenderEditorHtml(handle, "{\"referenceFields\":true}"));
        }
        finally { SessionRegistry.CloseSession(handle); }
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
