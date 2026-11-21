// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Text;
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

    [Fact]
    public void OL004_MoveBlockRange_AlreadyInPlace_TakesNoUndoStep()
    {
        using var session = new DocxSession(Build(P("A", "Ttulo1") + P("a1") + P("B", "Ttulo1")));
        var res = session.MoveBlockRange(AnchorOf(session, "A"), AnchorOf(session, "a1"), AnchorOf(session, "B"), Position.Before);
        Assert.True(res.Success);
        Assert.False(session.Undo());
    }
}
