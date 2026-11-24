using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary><see cref="DocxSession.UpdateFields"/>: a table of contents built from the headings,
/// REF / PAGEREF / SEQ results made current, one undo, a range, the update-on-open request.</summary>
public class DocxSessionUpdateFieldsTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string[] BodyParagraphs(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind is "p" or "h" && t.Anchor.Scope == "body")
            .Select(t => t.Anchor.Id).ToArray();

    private static XElement Body(DocxSession session)
    {
        using var stream = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(stream, false);
        return XDocument.Parse(new StreamReader(doc.MainDocumentPart!.GetStream()).ReadToEnd()).Root!.Element(W + "body")!;
    }

    private static bool AsksForUpdate(DocxSession session)
    {
        using var stream = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(stream, false);
        var part = doc.MainDocumentPart!.DocumentSettingsPart;
        if (part is null) return false;
        return XDocument.Parse(new StreamReader(part.GetStream()).ReadToEnd()).Descendants(W + "updateFields").Any();
    }

    private static string Text(XElement paragraph) =>
        string.Concat(paragraph.Descendants().Where(e => e.Name == W + "t" || (e.Name == W + "tab" && e.Parent?.Name == W + "r"))
            .Select(e => e.Name == W + "tab" ? "\t" : e.Value));

    /// <summary>Two headings with a table of contents above them.</summary>
    private static (DocxSession Session, string[] Headings) WithContents()
    {
        // (outline levels on the paragraphs themselves: the bare test document has no heading styles)
        var stream = new MemoryStream();
        stream.Write(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var part = document.MainDocumentPart!;
            var xml = part.GetXDocument();
            var ps = xml.Descendants(W + "p").ToList();
            for (int i = 0; i < 2; i++)
            {
                var pPr = ps[i].Element(W + "pPr");
                if (pPr is null) { pPr = new XElement(W + "pPr"); ps[i].AddFirst(pPr); }
                pPr.Add(new XElement(W + "outlineLvl", new XAttribute(W + "val", i.ToString())));
            }
            part.PutXDocument();
        }
        var session = new DocxSession(stream.ToArray());
        var headings = BodyParagraphs(session);
        Assert.True(session.InsertTableOfContents(headings[0], Position.Before, new TableOfContentsOptions()).Success);
        return (session, session.ListOutline().Select(h => h.AnchorId).ToArray());
    }

    [Fact]
    public void UF001_TableOfContents_IsBuiltFromTheHeadings_WithPagesLinksAndBookmarks()
    {
        var (session, headings) = WithContents();
        using var _ = session;
        Assert.True(AsksForUpdate(session));
        var texts = session.ListOutline().Select(h => h.Text).ToArray();

        var result = session.UpdateFields(new UpdateFieldsOptions
        {
            Pages = new Dictionary<string, string> { [headings[0]] = "2", [headings[1]] = "5" },
        });

        Assert.True(result.Success, result.Error?.Message);
        var body = Body(session);
        var entries = body.Descendants(W + "p")
            .Where(p => (string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") is "TOC1" or "TOC2").ToList();
        Assert.Equal(2, entries.Count);
        Assert.EndsWith(texts[0] + "\t2", Text(entries[0]));
        Assert.EndsWith(texts[1] + "\t5", Text(entries[1]));
        Assert.Equal("TOC2", (string?)entries[1].Element(W + "pPr")!.Element(W + "pStyle")!.Attribute(W + "val"));
        // each entry links to a _Toc bookmark on its heading, and its page is a PAGEREF to it
        var link = entries[0].Descendants(W + "hyperlink").Single();
        var bookmark = (string)link.Attribute(W + "anchor")!;
        Assert.Matches(new Regex(@"^_Toc\d{9}$"), bookmark);
        Assert.Contains(session.ListBookmarks(), b => b.Name == bookmark && b.Text == texts[0]);
        Assert.Contains(entries[0].Descendants(W + "instrText"), t => t.Value.Contains("PAGEREF " + bookmark));
        // one TOC field still spans the entries, and the document no longer asks Word to update it
        Assert.Single(body.Descendants(W + "instrText"), t => t.Value.TrimStart().StartsWith("TOC"));
        Assert.False(AsksForUpdate(session));
        Assert.True(session.ListOutline().Count == 2);

        // again: the same bookmarks, the new pages
        Assert.True(session.UpdateFields(new UpdateFieldsOptions
        {
            Pages = new Dictionary<string, string> { [headings[0]] = "3", [headings[1]] = "4" },
        }).Success);
        Assert.Equal(2, session.ListBookmarks().Count(b => b.Name.StartsWith("_Toc", StringComparison.Ordinal)));
        Assert.Contains(Body(session).Descendants(W + "p"), p => Text(p).EndsWith(texts[1] + "\t4", StringComparison.Ordinal));

        // one undo takes the last update back
        Assert.True(session.Undo());
        Assert.Contains(Body(session).Descendants(W + "p"), p => Text(p).EndsWith(texts[1] + "\t5", StringComparison.Ordinal));
    }

    [Fact]
    public void UF002_TableOfContents_WithoutHeadings_SaysSo()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.True(session.InsertTableOfContents(BodyParagraphs(session)[0], Position.Before, new TableOfContentsOptions()).Success);
        Assert.True(session.UpdateFields().Success);
        Assert.Contains(Body(session).Descendants(W + "p"), p => Text(p) == "No table of contents entries found.");
    }

    [Fact]
    public void UF003_Ref_TakesTheBookmarksTextAsItIsNow()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(session);
        Assert.True(session.AddBookmark("target", DocumentRange.In(paragraphs[0], new CharSpan(0, 4))).Success);
        Assert.True(session.InsertCrossReference(paragraphs[1], 0, "target").Success);
        var before = Body(session).Descendants(W + "fldSimple").Single().Value;
        Assert.True(session.ReplaceTextAtSpan(paragraphs[0], 0, 1, "Z").Success);

        Assert.True(session.UpdateFields().Success);

        var after = Body(session).Descendants(W + "fldSimple").Single().Value;
        Assert.Equal("Z" + before[1..], after);
    }

    [Fact]
    public void UF004_PageRef_TakesThePageFromTheMap_AndAMissingBookmarkSaysSo()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paragraphs = BodyParagraphs(session);
        Assert.True(session.AddBookmark("target", DocumentRange.In(paragraphs[0], new CharSpan(0, 4))).Success);
        Assert.True(session.InsertCrossReference(paragraphs[1], 0, "target", new CrossReferenceOptions { PageNumber = true }).Success);
        var unid = paragraphs[0][(paragraphs[0].LastIndexOf(':') + 1)..];
        Assert.True(session.UpdateFields(new UpdateFieldsOptions { Pages = new Dictionary<string, string> { [unid] = "9" } }).Success);
        var field = Body(session).Descendants(W + "fldSimple").Single();
        Assert.Equal("9", field.Value);
        Assert.Null(field.Attribute(W + "dirty"));
    }

    [Fact]
    public void UF005_Seq_IsNumberedAgainInDocumentOrder_AndARangeUpdatesOnlyItsFields()
    {
        var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var part = document.MainDocumentPart!;
            var xml = part.GetXDocument();
            var ps = xml.Descendants(W + "p").ToList();
            XElement Seq(string instr, string shown) => new XElement(W + "fldSimple", new XAttribute(W + "instr", instr),
                new XElement(W + "r", new XElement(W + "t", shown)));
            ps[0].Add(Seq(" SEQ Figure \\* ARABIC ", "7"), Seq(" SEQ Figure ", "7"));
            ps[1].Add(Seq(" SEQ Figure \\* ROMAN ", "7"), Seq(" SEQ Table ", "7"));
            part.PutXDocument();
        }
        using var session = new DocxSession(stream.ToArray());
        var paragraphs = BodyParagraphs(session);

        // only the second paragraph: its fields numbered as the whole document counts them
        Assert.True(session.UpdateFields(new UpdateFieldsOptions
        {
            Range = DocumentRange.In(paragraphs[1], new CharSpan(0, 0)),
        }).Success);
        Assert.Equal(new[] { "7", "7", "III", "1" }, Body(session).Descendants(W + "fldSimple").Select(f => f.Value).ToArray());

        Assert.True(session.UpdateFields().Success);
        Assert.Equal(new[] { "1", "2", "III", "1" }, Body(session).Descendants(W + "fldSimple").Select(f => f.Value).ToArray());
    }

    [Fact]
    public void UF007_TablesOfContentsOnly_LeavesOtherFields()
    {
        var (session, _) = WithContents();
        using var _s = session;
        var paragraphs = BodyParagraphs(session);
        var last = paragraphs[^1];
        Assert.True(session.AddBookmark("target", DocumentRange.In(last, new CharSpan(0, 2))).Success);
        Assert.True(session.InsertCrossReference(last, 0, "target", new CrossReferenceOptions { PageNumber = true, CachedPage = "8" }).Success);
        Assert.True(session.UpdateFields(new UpdateFieldsOptions
        {
            TablesOfContentsOnly = true,
            Pages = new Dictionary<string, string> { [last] = "3" },
        }).Success);
        var body = Body(session);
        Assert.Equal("8", body.Descendants(W + "fldSimple").Single().Value);
        Assert.Contains(body.Descendants(W + "p"), p => (string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") == "TOC1");
        // (part of an update: the document still asks Word for the rest)
        Assert.True(AsksForUpdate(session));
    }

    [Fact]
    public void UF008_ARangeInAContentsEntry_RebuildsTheTable()
    {
        var (session, headings) = WithContents();
        using var _s = session;
        Assert.True(session.UpdateFields().Success);
        var entry = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p" && session.Raw.GetXml(t.Anchor.Id).Contains("TOC1")).Anchor.Id;
        Assert.True(session.ReplaceTextAtSpan(headings[0], 0, 1, "Z").Success);
        var result = session.UpdateFields(new UpdateFieldsOptions { Range = DocumentRange.In(entry, new CharSpan(1, 0)) });
        Assert.True(result.Success, result.Error?.Message);
        Assert.Contains(Body(session).Descendants(W + "p"),
            p => (string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") == "TOC1" && Text(p).StartsWith("Z", StringComparison.Ordinal));
    }

    [Fact]
    public void UF009_TheTitleIsNotNumbered_AndAnUnnumberedEntryHasOnlyThePageTab()
    {
        var (session, _) = WithContents();
        using var _s = session;
        Assert.True(session.UpdateFields().Success);
        using var stream = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(stream, false);
        var styles = XDocument.Parse(new StreamReader(doc.MainDocumentPart!.StyleDefinitionsPart!.GetStream()).ReadToEnd());
        var title = styles.Descendants(W + "style").Single(s => (string?)s.Attribute(W + "styleId") == "TOCHeading");
        Assert.Equal("0", (string?)title.Descendants(W + "numId").Single().Attribute(W + "val"));
        var entry = Body(session).Descendants(W + "p")
            .First(p => (string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val") == "TOC1");
        var tab = Assert.Single(entry.Element(W + "pPr")!.Element(W + "tabs")!.Elements(W + "tab"));
        Assert.Equal("right", (string?)tab.Attribute(W + "val"));
    }

    // ─── What it leaves alone, and what it must not lose ────────────────────

    private const string Ns = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
        + "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:v=\"urn:schemas-microsoft-com:vml\" "
        + "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"";

    /// <summary>A document with heading styles that asks Word to update its fields.</summary>
    private static DocxSession Open(string body, string styles = "", string sectPr = "<w:sectPr/>", string numbering = "")
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
            Write(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles {Ns}>"
                + "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>"
                + "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:pPr><w:outlineLvl w:val=\"0\"/></w:pPr></w:style>"
                + "<w:style w:type=\"paragraph\" w:styleId=\"Heading2\"><w:name w:val=\"heading 2\"/><w:pPr><w:outlineLvl w:val=\"1\"/></w:pPr></w:style>"
                + "<w:style w:type=\"paragraph\" w:styleId=\"Clause\"><w:name w:val=\"Clause\"/><w:pPr><w:outlineLvl w:val=\"0\"/></w:pPr></w:style>"
                + styles + "</w:styles>");
            Write(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings {Ns}><w:updateFields w:val=\"true\"/></w:settings>");
            if (numbering.Length > 0) Write(main.AddNewPart<NumberingDefinitionsPart>(), $"<w:numbering {Ns}>{numbering}</w:numbering>");
            Write(main, $"<w:document {Ns}><w:body>{body}{sectPr}</w:body></w:document>");
        }
        return new DocxSession(stream.ToArray());
    }

    private static string P(string content, string? style = null, string pPr = "") =>
        $"<w:p><w:pPr>{(style is null ? "" : $"<w:pStyle w:val=\"{style}\"/>")}{pPr}</w:pPr>{content}</w:p>";
    private static string R(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";
    private static string Char(string type, string attributes = "") => $"<w:r><w:fldChar w:fldCharType=\"{type}\"{attributes}/></w:r>";
    private static string Start(string instr, string attributes = "") =>
        Char("begin", attributes) + $"<w:r><w:instrText xml:space=\"preserve\">{instr}</w:instrText></w:r>" + Char("separate");
    private static string Field(string instr, string result, string attributes = "") => Start(instr, attributes) + R(result) + Char("end");
    private static string Marked(string name, int id, string content) =>
        $"<w:bookmarkStart w:id=\"{id}\" w:name=\"{name}\"/>{content}<w:bookmarkEnd w:id=\"{id}\"/>";
    private const string Tab = "<w:r><w:tab/></w:r>";

    /// <summary>A contents table as Word writes one: its entries, and the paragraph that closes the field.</summary>
    private static string Contents(string instr, string closing, params string[] entries) =>
        string.Concat(entries.Select((e, i) => P((i == 0 ? Start(instr) : "") + e, "TOC1"))) + closing;
    private static string Closing(string after = "", string pPr = "") => P(Char("end") + after, null, pPr);

    private static string[] Lines(DocxSession session) => Body(session).Elements(W + "p").Select(Text).ToArray();

    [Theory]
    [InlineData(" TOC \\h \\z \\c \"Figure\" ")]
    [InlineData(" TOC \\t \"Clause,1\" \\h ")]
    [InlineData(" TOC \\o \"1-3\" \\b part ")]
    [InlineData(" TOC \\f A ")]
    [InlineData(" TOC \\o \"1-3\" \\p \"-\" ")]
    public void UF010_ATableThatIsNotOnlyOfHeadings_IsLeftAlone(string instr)
    {
        using var session = Open(Contents(instr, Closing(), R("Figure 1 A cat") + Tab + R("3")) + P(R("Intro"), "Heading1"));
        var before = Body(session).ToString();

        Assert.True(session.UpdateFields().Success);

        Assert.Equal(before, Body(session).ToString());
        Assert.True(AsksForUpdate(session));
        Assert.False(session.Undo());
    }

    [Fact]
    public void UF011_ATableOfFiguresTheSessionInserted_StaysForWordToFill()
    {
        using var session = Open(P(R("Intro"), "Heading1") + P(R("Figure ") + Field(" SEQ Figure \\* ARABIC ", "9"), "Caption"));
        Assert.True(session.InsertTableOfFigures(session.ListOutline()[0].AnchorId, Position.Before).Success);

        Assert.True(session.UpdateFields().Success);

        Assert.Equal(new[] { "", "Intro", "Figure 1" }, Lines(session));
        Assert.True(AsksForUpdate(session));
    }

    [Fact]
    public void UF012_WithoutPages_EntriesKeepTheirPage_AndWordIsStillAskedToUpdate()
    {
        using var session = Open(Contents(" TOC \\o \"1-3\" \\h ", Closing(), R("Old name") + Tab + Field(" PAGEREF _Toc1 \\h ", "7"))
            + P(Marked("_Toc1", 1, R("Intro")), "Heading1") + P(R("New"), "Heading1"));

        Assert.True(session.UpdateFields().Success);

        Assert.Equal(new[] { "Intro\t7", "New\t1", "", "Intro", "New" }, Lines(session));
        var dirty = Body(session).Descendants(W + "fldChar").Where(f => f.Attribute(W + "dirty") is not null).ToList();
        Assert.Equal("New\t1", Text(Assert.Single(dirty).Ancestors(W + "p").First()));
        Assert.True(AsksForUpdate(session));
        Assert.Contains(session.ListBookmarks(), b => b.Name == "_Toc1");
    }

    [Theory]
    [InlineData("")]
    [InlineData("This agreement is made today.")]
    public void UF013_TheParagraphThatClosesTheTable_KeepsItsSectionBreakAndText(string text)
    {
        using var session = Open(Contents(" TOC \\o \"1-3\" ",
                Closing(text.Length > 0 ? R(text) : "", "<w:sectPr><w:pgNumType w:fmt=\"lowerRoman\"/></w:sectPr>"), R("Intro") + Tab + R("7"))
            + P(R("Intro"), "Heading1") + P(R("Scope"), "Heading1"));

        Assert.True(session.UpdateFields().Success);

        Assert.Equal(new[] { "Intro\t1", "Scope\t1", text, "Intro", "Scope" }, Lines(session));
        var section = Body(session).Elements(W + "p").ElementAt(2).Element(W + "pPr")!.Element(W + "sectPr");
        Assert.Equal("lowerRoman", (string?)section!.Element(W + "pgNumType")!.Attribute(W + "fmt"));
        PackageValidation.AssertValid(session.Save(), DocumentFormat.OpenXml.FileFormatVersions.Office2007);
    }

    [Fact]
    public void UF014_ASectionBreakInTheLastEntry_AndTextBesideTheField_Stay()
    {
        using var session = Open(
            P(R("Contents: ") + Start(" TOC \\o \"1-3\" ") + R("Intro") + Tab + R("7") + Char("end") + R(" end"), "TOC1", "<w:sectPr/>")
            + P(R("Intro"), "Heading1") + P(R("Scope"), "Heading1"));

        Assert.True(session.UpdateFields().Success);

        Assert.Equal(new[] { "Contents: Intro\t1", "Scope\t1 end", "Intro", "Scope" }, Lines(session));
        Assert.NotNull(Body(session).Elements(W + "p").ElementAt(1).Element(W + "pPr")!.Element(W + "sectPr"));
    }

    [Fact]
    public void UF015_HeadingsInATable_OrHoldingAField_AreListed()
    {
        using var session = Open(Contents(" TOC \\o \"1-3\" ", Closing(), R("x"))
            + P(R("Schedule ") + Field(" SEQ Schedule ", "1") + R(" Fees"), "Heading1")
            + "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"900\"/></w:tblGrid><w:tr><w:tc>" + P(R("In a cell"), "Heading2") + "</w:tc></w:tr></w:tbl>"
            + P(R("Custom, only an outline level"), "Clause"));

        Assert.True(session.UpdateFields().Success);

        Assert.Equal(new[] { "Schedule 1 Fees\t1", "In a cell\t1" }, Lines(session).Take(2));
        Assert.Equal("", Lines(session)[2]);
        Assert.Contains(session.ListBookmarks(), b => b.Name.StartsWith("_Toc", StringComparison.Ordinal) && b.Text == "Schedule 1 Fees");
    }

    [Fact]
    public void UF016_OutlineLevels_AndPageNumberSwitches()
    {
        using var session = Open(Contents(" TOC \\o \"1-2\" \\n 2-2 \\u ", Closing(), R("x"))
            + P(R("One"), "Heading1") + P(R("Two"), "Heading2") + P(R("Custom"), "Clause"));
        Assert.True(session.UpdateFields().Success);
        Assert.Equal(new[] { "One\t1", "Two", "Custom\t1" }, Lines(session).Take(3));
    }

    [Fact]
    public void UF017_TheDocumentsOwnContentsStyle_IsUsed_WhateverItsId()
    {
        using var session = Open(Contents(" TOC \\o \"1-3\" ", Closing(), R("x")) + P(R("Intro"), "Heading1") + P(R("Sub"), "Heading2"),
            "<w:style w:type=\"paragraph\" w:styleId=\"TDC1\"><w:name w:val=\"toc 1\"/><w:pPr><w:tabs><w:tab w:val=\"right\" w:leader=\"dot\" w:pos=\"8000\"/></w:tabs></w:pPr></w:style>",
            "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/></w:sectPr>");

        Assert.True(session.UpdateFields().Success);

        var pPrs = Body(session).Elements(W + "p").Take(2).Select(p => p.Element(W + "pPr")!).ToArray();
        Assert.Equal("TDC1", (string?)pPrs[0].Element(W + "pStyle")!.Attribute(W + "val"));
        Assert.Null(pPrs[0].Element(W + "tabs")); // the style has the tab
        Assert.Equal("TOC2", (string?)pPrs[1].Element(W + "pStyle")!.Attribute(W + "val"));
        Assert.Equal("9016", (string?)pPrs[1].Element(W + "tabs")!.Element(W + "tab")!.Attribute(W + "pos")); // A4
    }

    [Theory]
    [InlineData("<w:bookmarkStart w:id=\"8\" w:name=\"_Cut\"/>")]
    [InlineData("<w:ins w:id=\"8\" w:author=\"Bob\"><w:r><w:t>x</w:t></w:r></w:ins>")]
    public void UF018_AResultThatHoldsMoreThanText_IsLeftAlone(string extra)
    {
        using var session = Open(Contents(" TOC \\o \"1-3\" ", Closing(), R("Old") + extra)
            + P(Marked("b", 1, R("Intro")), "Heading1") + P(Start(" REF b ") + R("Old") + extra.Replace("\"8\"", "\"9\"").Replace("_Cut", "_Cut2") + Char("end")));
        var before = Body(session).ToString();
        Assert.True(session.UpdateFields().Success);
        Assert.Equal(before, Body(session).ToString());
    }

    [Fact]
    public void UF019_LockedFields_KeepTheirResult()
    {
        using var session = Open(P(Marked("b", 1, R("New"))) + P(Field(" REF b ", "Old", " w:fldLock=\"1\""))
            + P("<w:fldSimple w:instr=\" SEQ Figure \" w:fldLock=\"true\">" + R("8") + "</w:fldSimple>") + P(Field(" SEQ Figure ", "8")));
        Assert.True(session.UpdateFields().Success);
        Assert.Equal(new[] { "New", "Old", "8", "2" }, Lines(session));
        // Word would leave the locked ones too, so nothing is left for it to update
        Assert.False(AsksForUpdate(session));
    }

    [Fact]
    public void UF020_Sequences_RestartAtHeadings_AndKeepTheirFormat()
    {
        const string figure = " SEQ Figure \\* ARABIC \\s 1 ";
        using var session = Open(
            P(R("One"), "Heading1") + P(Field(figure, "9")) + P(Field(figure, "9")) + P(R("Sub"), "Heading2") + P(Field(figure, "9"))
            + P(R("Two"), "Heading1") + P(Field(figure, "9"))
            + P(Field(" SEQ Table \\* ROMAN \\* MERGEFORMAT ", "9") + R(" ") + Field(" SEQ Table \\* alphabetic ", "9") + R(" ")
                + Field(" SEQ Table \\* Ordinal ", "3rd") + R(" ") + Field(" SEQ Table b ", "9") + R(" ") + Field(" SEQ Table \\r5 ", "9") + R(" ") + Field(" SEQ Table \\c ", "9")));
        Assert.True(session.UpdateFields().Success);
        Assert.Equal(new[] { "One", "1", "2", "Sub", "3", "Two", "1", "I b 3rd 9 5 5" }, Lines(session));
        Assert.True(AsksForUpdate(session));
    }

    [Fact]
    public void UF021_ACaptionInADrawing_IsCountedOnce_AndItsFallbackCopyFollows()
    {
        static string Box(string shown)
        {
            var text = "<w:txbxContent>" + P(R("Figure ") + Field(" SEQ Figure ", shown)) + "</w:txbxContent>";
            return $"<w:p><w:r><mc:AlternateContent><mc:Choice Requires=\"wps\"><w:drawing><wps:txbx>{text}</wps:txbx></w:drawing></mc:Choice>"
                + $"<mc:Fallback><w:pict><v:textbox>{text}</v:textbox></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>";
        }
        using var session = Open(Box("8") + Box("8") + P(R("Figure ") + Field(" SEQ Figure ", "8")));
        Assert.True(session.UpdateFields().Success);
        Assert.Equal(new[] { "Figure 1", "Figure 1", "Figure 2", "Figure 2", "Figure 3" },
            Body(session).Descendants(W + "p").Where(p => !p.Descendants(W + "p").Any()).Select(Text).ToArray());
    }

    /// <summary>The NVCA model charter as Word saved it: its 100+ REF results are already right.</summary>
    [Fact]
    public void UF022_ResultsWordCached_AreNotChanged()
    {
        using var session = new DocxSession(File.ReadAllBytes(Path.Combine("../../../../TestFiles", "NVCA-Model-COI.docx")));
        var before = Lines(session);
        Assert.True(session.UpdateFields().Success);
        Assert.Equal(before, Lines(session));
    }

    [Fact]
    public void UF023_ReferenceSwitches()
    {
        const string numbering = "<w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\"><w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"Section %1.\"/></w:lvl></w:abstractNum>"
            + "<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>";
        using var session = Open(P(Marked("n", 1, R("Term")), null, "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>")
            + P("<w:bookmarkStart w:id=\"2\" w:name=\"two\"/>" + R("first")) + P(R("second") + "<w:bookmarkEnd w:id=\"2\"/>")
            + P(Field(" REF n \\r \\h ", "old")) + P(Field(" REF n \\w \\t \\h ", "old")) + P(Field(" REF n \\p ", "old"))
            + P(Field(" REF n \\* Upper ", "TERM")) + P(Field(" REF two ", "first second")) + P(Field(" REF N ", "Term"))
            + P(Field(" REF gone ", "old")) + P(Field(" PAGEREF gone ", "4")), numbering: numbering);

        Assert.True(session.UpdateFields().Success);

        Assert.Equal(new[]
        {
            "Section 1", "1", "above",      // the number, its numerals only, where the target is
            "TERM", "first second", "Term", // a text format, two paragraphs, another case: kept
            "Error! Reference source not found.", "Error! Bookmark not defined.",
        }, Lines(session).Skip(3));
        Assert.True(AsksForUpdate(session));
    }

    [Fact]
    public void UF024_AFieldInAHeader_KeepsTheRequestToUpdate()
    {
        using var session = Open(P(Marked("b", 1, R("New"))) + P(Field(" REF b ", "Old")));
        Assert.True(session.SetHeaderText(BodyParagraphs(session)[0], HeaderFooterKind.Default, "Title").Success);
        Assert.True(session.UpdateFields().Success);
        Assert.False(AsksForUpdate(session));

        using var stream = new MemoryStream();
        stream.Write(session.Save());
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var header = document.MainDocumentPart!.HeaderParts.Single();
            header.GetXDocument().Descendants(W + "p").First().Add(XElement.Parse($"<w:fldSimple {Ns} w:instr=\" DOCPROPERTY Client \"/>"));
            header.PutXDocument();
            var settings = document.MainDocumentPart.DocumentSettingsPart!;
            settings.GetXDocument().Root!.Add(new XElement(W + "updateFields"));
            settings.PutXDocument();
        }
        using var withHeaderField = new DocxSession(stream.ToArray());
        Assert.True(withHeaderField.UpdateFields().Success);
        Assert.True(AsksForUpdate(withHeaderField));
    }

    [Fact]
    public void UF026_AHeadingWithSeveralContentsBookmarks_KeepsTheOneTheTableUses()
    {
        using var session = Open(Contents(" TOC \\o \"1-3\" \\h ", Closing(), R("Intro") + Tab + Field(" PAGEREF _Toc22 \\h ", "7"))
            + P(Marked("_Toc11", 1, Marked("_Toc22", 2, R("Intro"))), "Heading1"));
        Assert.True(session.UpdateFields().Success);
        Assert.Equal("Intro\t7", Lines(session)[0]);
        Assert.Contains(Body(session).Descendants(W + "instrText"), i => i.Value.Contains("_Toc22"));
        Assert.Equal(2, session.ListBookmarks().Count);
    }

    [Fact]
    public void UF025_TheResultSaysWhatChanged_AndNoChangeTakesNoUndoStep()
    {
        using var session = Open(Contents(" TOC \\o \"1-3\" ", Closing(), R("x"), R("y")) + P(R("Intro"), "Heading1"));
        var old = BodyParagraphs(session).Take(2).ToArray();

        var result = session.UpdateFields();

        Assert.Equal(old, result.Removed.Select(a => a.Id));
        Assert.Equal("Intro\t1", session.Project().AnchorIndex[Assert.Single(result.Created).Id].TextPreview);
        Assert.Contains(session.ListOutline()[0].AnchorId, result.Modified.Select(a => a.Id));
        Assert.True(session.Undo());

        using var current = Open(P(Marked("b", 1, R("Same"))) + P(Field(" REF b ", "Same") + Field(" DATE ", "today")));
        long version = current.Version;
        Assert.True(current.UpdateFields().Success);
        Assert.Equal(version, current.Version);
        Assert.False(current.Undo());
    }

    [Fact]
    public void UF006_RefusedWhileChangesAreTracked()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        Assert.False(session.UpdateFields().Success);
    }
}
