// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Docxodus;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="DocxSession.ListOutline"/> (the headings Word's Navigation Pane lists) and
/// <see cref="DocxSession.MoveBlockRange"/> (a heading moved with everything under it).
/// </summary>
public class DocxSessionOutlineTests
{
    private const string Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // A localized heading style (Spanish Word: id "Ttulo1", name "heading 1"), a style that is a
    // heading only through basedOn, a paragraph with its own outline level, and body text.
    private const string Styles =
        "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>"
        + "<w:style w:type=\"paragraph\" w:styleId=\"Ttulo1\"><w:name w:val=\"heading 1\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:outlineLvl w:val=\"0\"/></w:pPr></w:style>"
        + "<w:style w:type=\"paragraph\" w:styleId=\"Ttulo2\"><w:name w:val=\"heading 2\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:outlineLvl w:val=\"1\"/></w:pPr></w:style>"
        + "<w:style w:type=\"paragraph\" w:styleId=\"Chapter\"><w:name w:val=\"Chapter\"/><w:basedOn w:val=\"Ttulo1\"/></w:style>"
        + "<w:style w:type=\"paragraph\" w:styleId=\"Cuatro\"><w:name w:val=\"heading 4\"/><w:basedOn w:val=\"Normal\"/></w:style>"
        + "<w:style w:type=\"paragraph\" w:styleId=\"Plain\"><w:name w:val=\"Plain\"/><w:basedOn w:val=\"Ttulo1\"/><w:pPr><w:outlineLvl w:val=\"9\"/></w:pPr></w:style>";

    private static string P(string text, string? style = null, string? pPrExtra = null) =>
        "<w:p>" + (style is null && pPrExtra is null ? "" : "<w:pPr>" + (style is null ? "" : $"<w:pStyle w:val=\"{style}\"/>") + pPrExtra + "</w:pPr>")
        + $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static byte[] Build(string body)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            Write(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles xmlns:w=\"{Wns}\">{Styles}</w:styles>");
            Write(main, $"<w:document xmlns:w=\"{Wns}\"><w:body>{body}<w:sectPr/></w:body></w:document>");
        }
        return ms.ToArray();
    }

    private static void Write(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(xml);
    }

    private static string[] Texts(DocxSession session) =>
        session.Project().AnchorIndex.Values.Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h" or "li")
            .Select(t => t.FullText).ToArray();

    private static string Outline(DocxSession session) =>
        string.Join(" | ", session.ListOutline().Select(e => $"{e.Level}:{e.Text}"));

    private static string AnchorOf(DocxSession session, string text) =>
        session.Project().AnchorIndex.Values.First(t => t.Anchor.Scope == "body" && t.FullText == text).Anchor.Id;

    [Fact]
    public void OL001_Outline_ResolvesLevels_AsWord_ByStyleChain_AndOwnLevel()
    {
        using var session = new DocxSession(Build(
            P("Intro", "Ttulo1") + P("body") + P("Sub", "Ttulo2") + P("Chap", "Chapter")
            + P("Own", null, "<w:outlineLvl w:val=\"2\"/>") + P("NotAHeading", "Plain")
            + "<w:tbl><w:tr><w:tc>" + P("InTable", "Ttulo1") + "</w:tc></w:tr></w:tbl>"));

        Assert.Equal("1:Intro | 2:Sub | 1:Chap | 3:Own", Outline(session));
        var first = session.ListOutline()[0];
        Assert.Equal("Ttulo1", first.StyleId);
        Assert.Equal(AnchorOf(session, "Intro"), first.AnchorId);
        // a heading's part runs to the next heading at its level or above: Intro takes body and
        // Sub; Sub ends before Chap; Own (the last) runs to the end of the body, the table included
        var outline = session.ListOutline();
        Assert.Equal(AnchorOf(session, "Sub"), outline[0].LastAnchorId);
        Assert.Equal(AnchorOf(session, "Sub"), outline[1].LastAnchorId);
        Assert.Equal(outline[3].LastAnchorId, outline[2].LastAnchorId);
        Assert.StartsWith("tbl:", outline[3].LastAnchorId);
    }

    [Fact]
    public void OL002_MoveBlockRange_MovesAHeadingWithItsText_AsOneUndo()
    {
        using var session = new DocxSession(Build(
            P("A", "Ttulo1") + P("a1") + P("a2") + P("B", "Ttulo1") + P("b1")));
        var res = session.MoveBlockRange(AnchorOf(session, "B"), AnchorOf(session, "b1"), AnchorOf(session, "A"), Position.Before);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(new[] { "B", "b1", "A", "a1", "a2" }, Texts(session));
        Assert.Equal(2, res.Modified.Count);
        Assert.True(session.Undo());
        Assert.Equal(new[] { "A", "a1", "a2", "B", "b1" }, Texts(session));

        // and down, after the last block
        var down = session.MoveBlockRange(AnchorOf(session, "A"), AnchorOf(session, "a2"), AnchorOf(session, "b1"), Position.After);
        Assert.True(down.Success, down.Error?.Message);
        Assert.Equal(new[] { "B", "b1", "A", "a1", "a2" }, Texts(session));
        using var reopened = new DocxSession(session.Save());
        Assert.Equal(Texts(session), Texts(reopened));
    }

    [Fact]
    public void OL003_MoveBlockRange_Refuses_TargetInside_SectionBreak_AndCutRange()
    {
        using var session = new DocxSession(Build(
            P("A", "Ttulo1") + P("a1") + P("brk", null, "<w:sectPr/>") + P("B", "Ttulo1")
            + "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"Cross\"/><w:r><w:t>C</w:t></w:r></w:p>"
            + "<w:p><w:r><w:t>c1</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>" + P("D")));

        var inside = session.MoveBlockRange(AnchorOf(session, "A"), AnchorOf(session, "a1"), AnchorOf(session, "a1"), Position.After);
        var acrossBreak = session.MoveBlockRange(AnchorOf(session, "B"), AnchorOf(session, "B"), AnchorOf(session, "A"), Position.Before);
        var cut = session.MoveBlockRange(AnchorOf(session, "c1"), AnchorOf(session, "D"), AnchorOf(session, "B"), Position.Before);
        var whole = session.MoveBlockRange(AnchorOf(session, "C"), AnchorOf(session, "c1"), AnchorOf(session, "B"), Position.Before);

        Assert.False(inside.Success);
        Assert.False(acrossBreak.Success);
        Assert.False(cut.Success);
        Assert.True(whole.Success, whole.Error?.Message);
        Assert.Equal(new[] { "A", "a1", "brk", "C", "c1", "B", "D" }, Texts(session));
    }

    private static string Sdt(string content) =>
        "<w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr><w:sdtContent>" + content + "</w:sdtContent></w:sdt>";

    [Fact]
    public void OL005_MoveBlockRange_MovesAPartThatEndsInAContentControl()
    {
        using var session = new DocxSession(Build(
            P("A", "Ttulo1") + P("a1") + Sdt(P("c1") + P("c2")) + P("B", "Ttulo1") + P("b1")));
        var outline = session.ListOutline();

        var res = session.MoveBlockRange(outline[0].AnchorId, outline[0].LastAnchorId, outline[1].LastAnchorId, Position.After);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(new[] { "B", "b1", "A", "a1", "c1", "c2" }, Texts(session));
        Assert.Equal(4, res.Modified.Count);
        Assert.Contains("c2", session.Raw.GetXml(AnchorOf(session, "c2")));
        // and back, to after the part that now ends in the control: beside the control, not inside it
        outline = session.ListOutline();
        var back = session.MoveBlockRange(outline[0].AnchorId, outline[0].LastAnchorId, outline[1].LastAnchorId, Position.After);
        Assert.True(back.Success, back.Error?.Message);
        Assert.Equal(new[] { "A", "a1", "c1", "c2", "B", "b1" }, Texts(session));
        Assert.Equal(new[] { "c1", "c2" }, ControlTexts(session));
    }

    private static string[] ControlTexts(DocxSession session)
    {
        using var ms = new MemoryStream(session.Save());
        using var document = WordprocessingDocument.Open(ms, false);
        return document.MainDocumentPart!.GetXDocument().Descendants(W.sdtContent).Single().Elements(W.p).Select(p => p.Value).ToArray();
    }

    [Fact]
    public void OL006_MoveBlockRange_PastAContentControl_IsNotANoOp()
    {
        using var session = new DocxSession(Build(
            P("A", "Ttulo1") + P("a1") + Sdt(P("c1")) + P("B", "Ttulo1")));

        var res = session.MoveBlockRange(AnchorOf(session, "A"), AnchorOf(session, "a1"), AnchorOf(session, "B"), Position.Before);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(new[] { "c1", "A", "a1", "B" }, Texts(session));
    }

    [Fact]
    public void OL007_MoveBlockRange_RefusesToSplitAContentControl_OrToLandInsideOne()
    {
        using var session = new DocxSession(Build(
            P("A", "Ttulo1") + P("a1") + Sdt(P("c1") + P("c2")) + P("B", "Ttulo1")));
        var before = Texts(session);

        Assert.False(session.MoveBlockRange(AnchorOf(session, "A"), AnchorOf(session, "c1"), AnchorOf(session, "B"), Position.After).Success);
        Assert.False(session.MoveBlockRange(AnchorOf(session, "B"), AnchorOf(session, "B"), AnchorOf(session, "c1"), Position.After).Success);
        Assert.False(session.MoveBlockRange(AnchorOf(session, "c2"), AnchorOf(session, "B"), AnchorOf(session, "A"), Position.Before).Success);
        Assert.Equal(before, Texts(session));
        // within the control is its own run of blocks
        Assert.True(session.MoveBlockRange(AnchorOf(session, "c2"), AnchorOf(session, "c2"), AnchorOf(session, "c1"), Position.Before).Success);
        Assert.Equal(new[] { "A", "a1", "c2", "c1", "B" }, Texts(session));
    }

    /// <summary>Word ends a bookmark over whole paragraphs directly in the body, after the last one.</summary>
    [Fact]
    public void OL008_MoveBlockRange_TakesTheBookmarkThatCoversTheRun()
    {
        using var session = new DocxSession(Build(
            "<w:p><w:pPr><w:pStyle w:val=\"Ttulo1\"/></w:pPr><w:bookmarkStart w:id=\"3\" w:name=\"Span\"/><w:r><w:t>A</w:t></w:r></w:p>"
            + P("a1") + "<w:bookmarkEnd w:id=\"3\"/>" + P("B", "Ttulo1") + P("b1")));
        var text = session.ListBookmarks().Single().Text;

        var res = session.MoveBlockRange(AnchorOf(session, "A"), AnchorOf(session, "a1"), AnchorOf(session, "b1"), Position.After);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(new[] { "B", "b1", "A", "a1" }, Texts(session));
        var bookmark = session.ListBookmarks().Single();
        Assert.True(bookmark.IsPaired && bookmark.IsValid);
        Assert.Equal(text, bookmark.Text);
        // and the next run lands outside it, not between its last paragraph and its end
        Assert.True(session.MoveBlockRange(AnchorOf(session, "B"), AnchorOf(session, "b1"), AnchorOf(session, "a1"), Position.After).Success);
        Assert.Equal(text, session.ListBookmarks().Single().Text);
    }

    [Fact]
    public void OL009_MoveBlockRange_RefusesToCutABookmarkWithABodyLevelMarker()
    {
        using var session = new DocxSession(Build(
            P("A", "Ttulo1") + "<w:bookmarkStart w:id=\"3\" w:name=\"Span\"/>" + P("a1") + P("B", "Ttulo1") + P("b1")
            + "<w:bookmarkEnd w:id=\"3\"/>" + P("C", "Ttulo1") + P("c1")));
        var before = Texts(session);

        Assert.False(session.MoveBlockRange(AnchorOf(session, "A"), AnchorOf(session, "a1"), AnchorOf(session, "c1"), Position.After).Success);
        Assert.Equal(before, Texts(session));
        Assert.True(session.ListBookmarks().Single().IsValid);
    }

    [Fact]
    public void OL010_MoveBlockRange_RefusesToLandInsideAFieldsResult()
    {
        const string begin = "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> TOC \\o </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>";
        const string end = "<w:p><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        using var session = new DocxSession(Build(
            $"<w:p>{begin}<w:r><w:t>entry one</w:t></w:r></w:p>" + P("entry two") + end + P("A", "Ttulo1") + P("B", "Ttulo1")));

        var res = session.MoveBlockRange(AnchorOf(session, "B"), AnchorOf(session, "B"), AnchorOf(session, "entry two"), Position.Before);

        Assert.Equal(EditErrorCode.InvalidPosition, res.Error!.Code);
        Assert.True(session.MoveBlockRange(AnchorOf(session, "B"), AnchorOf(session, "B"), AnchorOf(session, "A"), Position.Before).Success);
    }

    private static string TextBox(string content) =>
        "<w:r><w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:textbox><w:txbxContent>" + content + "</w:txbxContent></v:textbox></v:shape></w:pict></w:r>";

    [Fact]
    public void OL012_ATextBoxInAHeading_IsNotItsText_AndItsSectionIsNotTheDocuments()
    {
        using var session = new DocxSession(Build(
            $"<w:p><w:pPr><w:pStyle w:val=\"Ttulo1\"/></w:pPr><w:r><w:t>Chapter title</w:t></w:r>{TextBox(P("Sidebar", null, "<w:sectPr/>"))}</w:p>"
            + P("a1") + P("B", "Ttulo1") + P("b1")));
        var outline = session.ListOutline();
        Assert.Equal(new[] { "Chapter title", "B" }, outline.Select(e => e.Text));

        var res = session.MoveBlockRange(outline[0].AnchorId, outline[0].LastAnchorId, outline[1].LastAnchorId, Position.After);

        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal(new[] { "B", "Chapter title" }, session.ListOutline().Select(e => e.Text));
    }

    [Fact]
    public void OL013_ASectionBreakInATableCell_IsStillOne()
    {
        using var session = new DocxSession(Build(P("A", "Ttulo1") + P("B", "Ttulo1")
            + "<w:tbl><w:tr><w:tc>" + P("cell", null, "<w:sectPr/>") + "</w:tc></w:tr></w:tbl>" + P("C", "Ttulo1")));
        Assert.False(session.MoveBlockRange(AnchorOf(session, "A"), AnchorOf(session, "A"), AnchorOf(session, "C"), Position.After).Success);
    }

    [Fact]
    public void OL014_MoveBlockRange_ReportsItsPatch()
    {
        using var session = new DocxSession(Build(P("A", "Ttulo1") + P("a1") + P("B", "Ttulo1")));
        var before = session.Project().Markdown;

        var res = session.MoveBlockRange(AnchorOf(session, "A"), AnchorOf(session, "a1"), AnchorOf(session, "B"), Position.After);

        Assert.NotNull(res.Patch);
        Assert.NotEqual(before, session.Project().Markdown);
        Assert.Contains("a1", res.Patch!.Markdown);
        using var quiet = new DocxSession(Build(P("A", "Ttulo1") + P("B", "Ttulo1")), new DocxSessionSettings { EmitMarkdownPatch = false });
        Assert.Null(quiet.MoveBlockRange(AnchorOf(quiet, "A"), AnchorOf(quiet, "A"), AnchorOf(quiet, "B"), Position.After).Patch);
    }

    private static string Tbl(string text) =>
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"900\"/></w:tblGrid><w:tr><w:tc>" + P(text) + "</w:tc></w:tr></w:tbl>";

    private static string Blocks(DocxSession session)
    {
        using var ms = new MemoryStream(session.Save());
        using var document = WordprocessingDocument.Open(ms, false);
        return string.Join(" ", document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!.Elements()
            .Select(e => e.Name == W.tbl ? $"[{e.Value}]" : e.Name == W.p ? (e.Value.Length > 0 ? e.Value : "-") : e.Name.LocalName));
    }

    [Fact]
    public void OL015_MoveBlockRange_KeepsAParagraphBetweenTables_AndAfterTheLast()
    {
        using var between = new DocxSession(Build(Tbl("T1") + P("sep") + Tbl("T2") + P("tail")));
        var res = between.MoveBlockRange(AnchorOf(between, "sep"), AnchorOf(between, "sep"), AnchorOf(between, "tail"), Position.After);
        Assert.True(res.Success, res.Error?.Message);
        Assert.Equal("[T1] - [T2] tail sep sectPr", Blocks(between));
        Assert.Single(res.Created);
        Assert.True(between.Undo());
        Assert.Equal("[T1] sep [T2] tail sectPr", Blocks(between));

        // a part that ends in a table, to the end of the body
        using var toEnd = new DocxSession(Build(P("A", "Ttulo1") + Tbl("T") + P("B", "Ttulo1") + P("b1")));
        var outline = toEnd.ListOutline();
        Assert.True(toEnd.MoveBlockRange(outline[0].AnchorId, outline[0].LastAnchorId, outline[1].LastAnchorId, Position.After).Success);
        Assert.Equal("B b1 A [T] - sectPr", Blocks(toEnd));

        // the part after a table, away from the end
        using var fromEnd = new DocxSession(Build(P("A", "Ttulo1") + Tbl("T") + P("B", "Ttulo1") + P("b1")));
        outline = fromEnd.ListOutline();
        Assert.True(fromEnd.MoveBlockRange(outline[1].AnchorId, outline[1].LastAnchorId, outline[0].AnchorId, Position.Before).Success);
        Assert.Equal("B b1 A [T] - sectPr", Blocks(fromEnd));
    }

    [Fact]
    public void OL016_ADefaultStyleSpelledOn_AndAFallbackThatIsShown_Count()
    {
        using var session = new DocxSession(Build(
            "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:zz=\"urn:unknown\">"
            + "<mc:Choice Requires=\"zz\">" + P("hidden", "Ttulo1") + "</mc:Choice><mc:Fallback>" + P("shown", "Ttulo1") + "</mc:Fallback></mc:AlternateContent>" + P("b1")));
        Assert.Equal("1:shown", Outline(session));

        using var ms = new MemoryStream();
        ms.Write(Build(P("plain")));
        ms.Position = 0;
        using (var document = WordprocessingDocument.Open(ms, true))
        {
            var part = document.MainDocumentPart!.StyleDefinitionsPart!;
            var normal = part.GetXDocument().Root!.Elements(W.style).First();
            normal.SetAttributeValue(W._default, "on");
            normal.Add(new XElement(W.pPr, new XElement(W.outlineLvl, new XAttribute(W.val, "2"))));
            part.PutXDocument();
        }
        using var byDefault = new DocxSession(ms.ToArray());
        Assert.Equal("3:plain", Outline(byDefault));
    }

    [Fact]
    public void OL011_Outline_ABuiltInHeadingStyleWithoutALevel_HasItsOwn()
    {
        using var session = new DocxSession(Build(P("Four", "Cuatro") + P("body")));
        Assert.Equal("4:Four", Outline(session));
    }

    [Fact]
    public void OL004_MoveBlockRange_AlreadyInPlace_TakesNoUndoStep()
    {
        using var session = new DocxSession(Build(P("A", "Ttulo1") + P("a1") + P("B", "Ttulo1")));
        var res = session.MoveBlockRange(AnchorOf(session, "A"), AnchorOf(session, "a1"), AnchorOf(session, "B"), Position.Before);
        Assert.True(res.Success);
        Assert.False(session.Undo());
    }
}
