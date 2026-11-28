using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>What a field nested in another field's code holds does not show in HTML (issue #1875), but the session counts
/// its text as the paragraph's. The browser editor diffs a block's text against the session's and sends its offsets to the
/// session, so the block must count what the session counts: the hidden text renders as an empty field marker holding it
/// in <c>data-field-cached</c>, which the editor counts as it counts any field's cached text.</summary>
public class WmlToHtmlConverterFieldCodeEditorTests
{
    private static string Run(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private static string Code(string code) => $"<w:r><w:instrText xml:space=\"preserve\">{code}</w:instrText></w:r>";

    private static string Field(string instruction, string result) =>
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" + instruction + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
        + result + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    // Each shape: what the paragraph shows, then its content. The DOCPROPERTY one sits in a part the converter already
    // evaluated fields in before issue #1875.
    public static TheoryData<string, string, string> Shapes() => new()
    {
        { "If", "If: First tail", Field(Code(" IF ") + Field(Code(" PAGE "), Run("1")) + Code(" &gt; 1 \"Later\" \"First\" "), Run("First")) },
        { "Merge", "Merge: Visit tail", Field(Code(" HYPERLINK \"") + Field(Code(" MERGEFIELD Url "), Run("https://example.test/merged"))
            + Code("\" "), Run("Visit")) },
        { "Ref", "Ref: Yes tail", Field(Code(" IF ") + Field(Code(" REF Mark1 \\h "), Run("target")) + Code(" = \"target\" \"Yes\" \"No\" "),
            Run("Yes")) },
        { "Property", "Property: Yes tail", Field(Code(" IF ") + Field(Code(" DOCPROPERTY Company "), Run("Acme"))
            + Code(" = \"Acme\" \"Yes\" \"No\" "), Run("Yes")) },
        { "Simple", "Simple: First tail", Field(Code(" IF ") + $"<w:fldSimple w:instr=\" PAGE \">{Run("1")}</w:fldSimple>"
            + Code(" &gt; 1 \"Later\" \"First\" "), Run("First")) },
    };

    private const string Notes = "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
        + "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>";

    /// <summary>The stories the paragraph is in, by the prefix its text starts with there.</summary>
    private static readonly string[] Stories = { "", "Header ", "Note " };

    private static void Put(OpenXmlPart part, string xml) => part.PutXDocument(XDocument.Parse(xml));

    private static byte[] Docx(string label, string field)
    {
        string Paragraph(string story) => $"<w:p>{Run($"{story}{label}: ")}{field}{Run(" tail")}</w:p>";
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var ns = $"xmlns:w=\"{W.w}\" xmlns:r=\"{R.r}\"";
            Put(main.AddNewPart<HeaderPart>("rIdHeader"), $"<w:hdr {ns}>{Paragraph("Header ")}</w:hdr>");
            Put(main.AddNewPart<FootnotesPart>(), $"<w:footnotes {ns}>{Notes}<w:footnote w:id=\"1\">"
                + $"<w:p><w:r><w:footnoteRef/></w:r>{Run("Note")}</w:p>{Paragraph("Note ")}<w:p>{Run("End")}</w:p></w:footnote></w:footnotes>");
            Put(main, $"""
                <w:document {ns}><w:body>
                  {Paragraph("")}
                  <w:p><w:bookmarkStart w:id="9" w:name="Mark1"/>{Run("target")}<w:bookmarkEnd w:id="9"/><w:r><w:footnoteReference w:id="1"/></w:r></w:p>
                  <w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="12240" w:h="15840"/>
                    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>
                </w:body></w:document>
                """);
            Put(main.AddNewPart<StyleDefinitionsPart>(), $"<w:styles {ns}/>");
            Put(main.AddNewPart<DocumentSettingsPart>(), $"<w:settings {ns}/>");
        }
        return stream.ToArray();
    }

    /// <summary>The paragraph in the story <paramref name="story"/> names, by the text the session holds for it.</summary>
    private static string AnchorOf(DocxSession session, string story, string label) =>
        session.Project().AnchorIndex.Values.Select(t => t.Anchor.Id)
            .Single(id => id.StartsWith("p:", StringComparison.Ordinal)
                && session.GetAnchorInfo(id)?.VisibleText.StartsWith($"{story}{label}: ", StringComparison.Ordinal) == true);

    /// <summary>A block's text as the editor counts it (<c>blockContentText</c> in npm/src/editor.ts): a field marker
    /// counts its <c>data-field-cached</c> text, else what it shows; a tab and a break one character each; a list marker
    /// nothing; bidi marks nothing.</summary>
    private static string EditorText(XElement block)
    {
        var text = new StringBuilder();
        void Walk(XNode node)
        {
            if (node is XText t) text.Append(t.Value.Where(c => c is not ('‎' or '‏' or '؜')
                && c is not (>= '‪' and <= '‮') && c is not (>= '⁦' and <= '⁩')).ToArray());
            if (node is not XElement e || e.Attribute("data-list-marker") != null) return;
            if (e.Attribute("data-field") != null) text.Append((string?)e.Attribute("data-field-cached") ?? e.Value);
            else if (e.Attribute("data-docx-tab") != null) text.Append('\t');
            else if (e.Name.LocalName == "br") text.Append('\n');
            else foreach (var child in e.Nodes()) Walk(child);
        }
        Walk(block);
        return text.ToString();
    }

    private static string Unid(string anchorId) => anchorId.Substring(anchorId.LastIndexOf(':') + 1);

    private static XElement Block(string html, string anchorId) =>
        XElement.Parse(html, LoadOptions.PreserveWhitespace).DescendantsAndSelf()
            .First(e => (string?)e.Attribute("data-anchor") == Unid(anchorId));

    private static string Profile(bool paginated, bool referenceFields, bool stampAnchors) =>
        $"{{\"paginated\":{(paginated ? "true" : "false")},\"referenceFields\":{(referenceFields ? "true" : "false")},"
        + $"\"stampAnchors\":{(stampAnchors ? "true" : "false")}}}";

    private static void ForEachProfile(Action<string> act)
    {
        foreach (var paginated in new[] { false, true })
            foreach (var referenceFields in new[] { false, true })
                foreach (var stampAnchors in new[] { false, true })
                    act(Profile(paginated, referenceFields, stampAnchors));
    }

    [Theory]
    [MemberData(nameof(Shapes))]
    public void TheEditorsRendersCountWhatTheSessionCounts(string label, string shown, string field)
    {
        int handle = DocxSessionOps.OpenSession(Docx(label, field), null);
        try
        {
            var session = SessionRegistry.Get(handle);
            foreach (var story in Stories)
            {
                var anchor = AnchorOf(session, story, label);
                var counted = session.GetAnchorInfo(anchor)!.VisibleText;
                ForEachProfile(profile =>
                {
                    var block = Block(DocxSessionOps.RenderEditorBlockHtml(handle, anchor, profile), anchor);
                    Assert.Equal(story + shown, block.Value);
                    Assert.Equal(counted, EditorText(block));
                    // The full render stamps the body's and the notes' paragraphs; a running story is edited through
                    // its block render.
                    if (story != "Header ")
                        Assert.Equal(counted, EditorText(Block(DocxSessionOps.RenderEditorHtml(handle, profile), anchor)));
                });
            }
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    // What the editor sends when the user types, measured in the block's own text, is where the session puts the
    // text: at the end of the block, before the field, or at the start of the field's result (the hidden code text
    // before it is behind the caret there, and the session puts text typed at a boundary inside a field after the
    // field). The typed text shows, and the field keeps its code and result.
    [Theory]
    [MemberData(nameof(Shapes))]
    public void TypingAtTheEditorsOffsetLosesNothing(string label, string shown, string field)
    {
        var result = shown.Substring(label.Length + 2, shown.Length - label.Length - 2 - " tail".Length);
        foreach (var story in Stories)
            foreach (var referenceFields in new[] { false, true })
                foreach (var place in new[] { "end", "field", "result" })
                {
                    int handle = DocxSessionOps.OpenSession(Docx(label, field), null);
                    try
                    {
                        var session = SessionRegistry.Get(handle);
                        var anchor = AnchorOf(session, story, label);
                        var counted = session.GetAnchorInfo(anchor)!.VisibleText;
                        var block = Block(DocxSessionOps.RenderEditorBlockHtml(handle, anchor, Profile(story == "Header ", referenceFields, false)), anchor);
                        var text = EditorText(block);
                        var at = place switch
                        {
                            "end" => text.Length,
                            "field" => story.Length + label.Length + 2,
                            _ => text.LastIndexOf(result + " tail", StringComparison.Ordinal),
                        };
                        Assert.True(session.ReplaceTextAtSpan(anchor, at, 0, "X").Success);
                        Assert.Equal(counted, session.GetAnchorInfo(anchor)!.VisibleText.Replace("X", ""));
                        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("edited.docx", session.Save()), new WmlToHtmlConverterSettings
                        {
                            FabricateCssClasses = false,
                            RenderHeadersAndFooters = true,
                            RenderFootnotesAndEndnotes = true,
                        });
                        var edited = Assert.Single(html.Descendants(Xhtml.p), p => p.Value.StartsWith($"{story}{label}: ", StringComparison.Ordinal)).Value;
                        Assert.Equal(story + shown, edited.Replace("X", ""));
                        Assert.Single(edited, c => c == 'X');
                        if (place == "end") Assert.EndsWith("X", edited);
                        if (place == "field") Assert.StartsWith($"{story}{label}: X", edited);
                    }
                    finally
                    {
                        DocxSessionOps.CloseSession(handle);
                    }
                }
    }
}
