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

    private static byte[] Docx(string label, string field)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var ns = $"xmlns:w=\"{W.w}\" xmlns:r=\"{R.r}\"";
            main.PutXDocument(XDocument.Parse($"""
                <w:document {ns}><w:body>
                  <w:p>{Run($"{label}: ")}{field}{Run(" tail")}</w:p>
                  <w:p><w:bookmarkStart w:id="9" w:name="Mark1"/>{Run("target")}<w:bookmarkEnd w:id="9"/></w:p>
                  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
                </w:body></w:document>
                """));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse($"<w:styles {ns}/>"));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings {ns}/>"));
        }
        return stream.ToArray();
    }

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
            var anchor = session.ListBlocks().Body.First(u => u.Kind is "p").Id;
            var counted = session.GetAnchorInfo(anchor)!.VisibleText;
            ForEachProfile(profile =>
            {
                var block = Block(DocxSessionOps.RenderEditorBlockHtml(handle, anchor, profile), anchor);
                Assert.Equal(shown, block.Value);
                Assert.Equal(counted, EditorText(block));
                var full = Block(DocxSessionOps.RenderEditorHtml(handle, profile), anchor);
                Assert.Equal(counted, EditorText(full));
            });
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    // What the editor sends when the user types "!" at the end of the block, measured in the block's own text, lands
    // after "tail", and the field keeps its code and result.
    [Theory]
    [MemberData(nameof(Shapes))]
    public void TypingAtTheEditorsOffsetEditsWhereTheUserTyped(string label, string shown, string field)
    {
        foreach (var referenceFields in new[] { false, true })
        {
            int handle = DocxSessionOps.OpenSession(Docx(label, field), null);
            try
            {
                var session = SessionRegistry.Get(handle);
                var anchor = session.ListBlocks().Body.First(u => u.Kind is "p").Id;
                var before = session.GetAnchorInfo(anchor)!.VisibleText;
                var block = Block(DocxSessionOps.RenderEditorBlockHtml(handle, anchor, Profile(false, referenceFields, false)), anchor);
                Assert.True(session.ReplaceTextAtSpan(anchor, EditorText(block).Length, 0, "!").Success);
                Assert.Equal(before + "!", session.GetAnchorInfo(anchor)!.VisibleText);
                var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("edited.docx", session.Save()),
                    new WmlToHtmlConverterSettings { FabricateCssClasses = false });
                Assert.Equal(shown + "!", html.Descendants(Xhtml.p).First().Value);
            }
            finally
            {
                DocxSessionOps.CloseSession(handle);
            }
        }
    }
}
