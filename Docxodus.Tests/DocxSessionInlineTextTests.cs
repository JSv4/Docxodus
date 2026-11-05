using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Ir;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionInlineTextTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    [Fact]
    public void A_complex_field_can_cross_an_inline_custom_xml_boundary()
    {
        var doc = IrTestDocuments.FromBodyXml("<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/>" +
            "<w:instrText> REF top </w:instrText></w:r><w:customXml w:element=\"clause\">" +
            "<w:r><w:fldChar w:fldCharType=\"separate\"/><w:t>Result</w:t></w:r></w:customXml>" +
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>");
        var oracle = WmlToMarkdownConverter.Convert(new WmlDocument(doc), new());
        Assert.Contains("} Result", oracle.Markdown);
        Assert.Equal(oracle.Markdown, IrMarkdownEmitter.Emit(IrReader.Read(new WmlDocument(doc)), new()).Markdown);
    }

    [Theory]
    [InlineData("<w:tab/>")]
    [InlineData("<w:br/>")]
    [InlineData("<w:cr/>")]
    [InlineData("<w:ptab w:alignment=\"right\" w:relativeTo=\"margin\" w:leader=\"none\"/>")]
    [InlineData("<w:noBreakHyphen/>")]
    [InlineData("<w:softHyphen/>")]
    [InlineData("<w:sym w:font=\"Symbol\" w:char=\"03A9\"/>")]
    public void Suppress_and_sequential_anchors_keep_atom_only_paragraphs_in_document_order(string atom)
    {
        var doc = IrTestDocuments.FromBodyXml("<w:p><w:r>" + atom + "</w:r></w:p>" +
            "<w:p><w:r><w:t>Control</w:t></w:r></w:p>");
        var settings = new WmlToMarkdownConverterSettings
        {
            EmptyParagraphs = EmptyParagraphMode.Suppress, AnchorIdRendering = AnchorIdRendering.Sequential,
        };
        var oracle = WmlToMarkdownConverter.Convert(new WmlDocument(doc), settings);
        var ir = IrMarkdownEmitter.Emit(IrReader.Read(new WmlDocument(doc)), settings);
        Assert.Equal(oracle.Markdown, ir.Markdown);
        Assert.Equal(oracle.AnchorIndex.Keys, ir.AnchorIndex.Keys);
        Assert.Equal(oracle.AnchorIndex.Values.Select(a => a.TextPreview), ir.AnchorIndex.Values.Select(a => a.TextPreview));
    }

    [Theory]
    [InlineData("ptab", "alignment")]
    [InlineData("customXml", "element")]
    public void Display_facts_do_not_hide_carrier_metadata_changes_from_comparison(string element, string attribute)
    {
        WmlDocument Document(string value) => IrTestDocuments.FromBodyXml("<w:p>" +
            (element == "ptab" ? "<w:r><w:t>A</w:t><w:ptab w:alignment=\"" + value +
                "\" w:relativeTo=\"margin\" w:leader=\"none\"/><w:t>B</w:t></w:r>" :
                "<w:customXml w:element=\"" + value + "\"><w:r><w:t>AB</w:t></w:r></w:customXml>") + "</w:p>");
        var compared = DocxDiff.Compare(Document("left"), Document("right"));
        foreach (var (bytes, expected) in new[]
        {
            (RevisionProcessor.AcceptRevisions(compared).DocumentByteArray, "right"),
            (RevisionProcessor.RejectRevisions(compared).DocumentByteArray, "left"),
        })
        {
            using var saved = WordprocessingDocument.Open(new MemoryStream(bytes), false);
            var carrier = Assert.Single(saved.MainDocumentPart!.GetXDocument().Descendants(W + element));
            Assert.Equal(expected, (string?)carrier.Attribute(W + attribute));
        }
    }

    [Theory]
    [InlineData("br")]
    [InlineData("cr")]
    public void Formatted_breaks_have_the_same_projection_in_both_readers(string kind)
    {
        var doc = IrTestDocuments.FromBodyXml("<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t><w:" +
            kind + "/><w:t>B</w:t></w:r></w:p>");
        Assert.Equal(WmlToMarkdownConverter.Convert(new WmlDocument(doc), new()).Markdown,
            IrMarkdownEmitter.Emit(IrReader.Read(new WmlDocument(doc)), new()).Markdown);
    }

    [Fact]
    public void A_formatted_insertion_after_a_break_preserves_the_break()
    {
        using var session = new DocxSession(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:br/><w:t>document</w:t></w:r></w:p>").DocumentByteArray,
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchor = session.FindByText("document")!.Anchor.Id;
        var edit = session.ReplaceTextAtSpanWithFormat(anchor, 5, 0, " inserted ", new FormatOp { Bold = true });
        Assert.True(edit.Success, edit.Error?.Message);
        Assert.True(session.AcceptAllRevisions().Success);
        Assert.Single(session.Grep("^\ndocu inserted ment$"));
        using var saved = WordprocessingDocument.Open(new MemoryStream(session.Save()), false);
        Assert.Single(saved.MainDocumentPart!.GetXDocument().Descendants(W + "br"));
        Assert.NotNull(Assert.Single(saved.MainDocumentPart.GetXDocument().Descendants(W + "t"),
            t => t.Value == " inserted ").Parent!.Element(W + "rPr")?.Element(W + "b"));
    }

    [Theory]
    [InlineData("<w:ptab w:alignment=\"right\" w:relativeTo=\"margin\" w:leader=\"none\"/>", "\t")]
    [InlineData("<w:softHyphen/>", "\u00ad")]
    [InlineData("<w:noBreakHyphen/>", "\u2011")]
    public void Html_keeps_the_same_visible_atom_spelling(string atom, string text)
    {
        var doc = IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>A</w:t>" + atom + "<w:t>B</w:t></w:r></w:p>");
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument(doc), new());
        var paragraph = Assert.Single(html.Descendants().Where(e => e.Name.LocalName == "p"));
        Assert.Equal("A" + text + "B", paragraph.Value.Replace("\u200e", ""));
    }

    [Theory]
    [InlineData("<w:tab/>", "\t", "\t")]
    [InlineData("<w:br/>", "\n", "  \n")]
    [InlineData("<w:cr/>", "\n", "  \n")]
    [InlineData("<w:noBreakHyphen/>", "\u2011", "\u2011")]
    [InlineData("<w:softHyphen/>", "\u00ad", "\u00ad")]
    [InlineData("<w:sym w:font=\"Symbol\" w:char=\"03A9\"/>", "\u03a9", "\u03a9")]
    [InlineData("<w:ptab w:alignment=\"right\" w:relativeTo=\"margin\" w:leader=\"none\"/>", "\t", "\t")]
    public void Visible_run_atoms_are_searchable_and_consumed_by_span_edits(string atom, string text, string markdown)
    {
        using var session = new DocxSession(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>A</w:t>" + atom + "<w:t>B</w:t></w:r></w:p>").DocumentByteArray);
        var projection = session.Project();
        var anchor = Assert.Single(projection.AnchorIndex.Values, a => a.Anchor.Kind == "p").Anchor.Id;

        Assert.Single(session.Grep("^" + Regex.Escape("A" + text + "B") + "$"));
        Assert.Contains("} A" + markdown + "B", projection.Markdown);
        var edit = session.ReplaceTextAtSpan(anchor, 1, 1, "X");

        Assert.True(edit.Success, edit.Error?.Message);
        Assert.Single(session.Grep("^AXB$"));
        Assert.Contains("} AXB", session.Project().Markdown);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Suppress_keeps_tab_only_paragraphs_and_header_scopes(bool header)
    {
        const string tab = "<w:p><w:r><w:tab/></w:r></w:p>";
        var doc = header ? IrTestDocuments.FromBodyAndHeaderXml("<w:p/>", tab) : IrTestDocuments.FromBodyXml(tab);
        var settings = new WmlToMarkdownConverterSettings { EmptyParagraphs = EmptyParagraphMode.Suppress };
        var oracle = WmlToMarkdownConverter.Convert(new WmlDocument(doc), settings);
        var ir = IrMarkdownEmitter.Emit(IrReader.Read(new WmlDocument(doc)), settings);

        Assert.Contains('\t', oracle.Markdown);
        Assert.Single(oracle.AnchorIndex.Values, a => a.Anchor.Kind == "p");
        Assert.Equal(oracle.Markdown, ir.Markdown);
    }

    [Theory]
    [InlineData("{0}")]
    [InlineData("<w:sdt><w:sdtContent>{0}</w:sdtContent></w:sdt>")]
    [InlineData("<w:hyperlink w:anchor=\"top\">{0}</w:hyperlink>")]
    [InlineData("<w:ins w:id=\"1\" w:author=\"Other\">{0}</w:ins>")]
    [InlineData("<w:fldSimple w:instr=\" REF top \">{0}</w:fldSimple>")]
    public void An_anchored_textbox_does_not_join_its_host_paragraph_text(string envelope)
    {
        const string run = "<w:r><w:t>AB</w:t><w:pict><v:shape><v:textbox><w:txbxContent>" +
            "<w:p><w:r><w:t>Boxed note</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>";
        var doc = IrTestDocuments.FromBodyXmlWithDrawingNamespaces("<w:p>" + string.Format(envelope, run) + "</w:p>");
        using var session = new DocxSession(doc.DocumentByteArray);

        var host = Assert.Single(session.Grep("^AB$"));
        var box = Assert.Single(session.Grep("^Boxed note$"));
        Assert.Empty(session.Grep("ABBoxed"));
        var projection = session.Project();
        Assert.DoesNotContain("} ABBoxed", projection.Markdown);
        // Keep source identity across the IR's revision normalization; removing an insertion
        // wrapper before assigning IDs otherwise changes the nested textbox paragraph's anchor.
        var anchored = new WmlDocument("anchored.docx", session.Save(persistAnchorIds: true));
        Assert.Equal(projection.Markdown, IrMarkdownEmitter.Emit(IrReader.Read(anchored), new()).Markdown);
        Assert.True(session.ReplaceMatch(box, "Changed note").Success);
        Assert.Single(session.Grep("^AB$"));
        Assert.True(session.ReplaceMatch(host, "CD").Success);
        Assert.Single(session.Grep("^Changed note$"));
        Assert.Single(session.Grep("^CD$"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Inline_custom_xml_projects_and_addresses_its_text(bool control)
    {
        var inline =
            "<w:customXml w:uri=\"urn:test\" w:element=\"clause\"><w:r><w:t>Wrapped value</w:t></w:r>" +
            "</w:customXml>";
        var doc = IrTestDocuments.FromBodyXml("<w:p>" +
            (control ? "<w:sdt><w:sdtContent>" + inline + "</w:sdtContent></w:sdt>" : inline) + "</w:p>");
        using var session = new DocxSession(doc.DocumentByteArray);

        Assert.Single(session.Grep("^Wrapped value$"));
        Assert.Contains("} Wrapped value", session.Project().Markdown);
        Assert.Equal(session.Project().Markdown,
            IrMarkdownEmitter.Emit(IrReader.Read(new WmlDocument(doc)), new()).Markdown);
        Assert.True(session.ReplaceMatch(Assert.Single(session.Grep("value")), "text").Success);
        Assert.Single(session.Grep("^Wrapped text$"));
    }

    [Theory]
    [InlineData("<w:ptab w:alignment=\"right\" w:relativeTo=\"margin\" w:leader=\"dot\"/>")]
    [InlineData("<w:sym w:font=\"Symbol\" w:char=\"03A9\"/>")]
    [InlineData("<w:br w:type=\"page\" w:clear=\"all\"/>")]
    [InlineData("<w:cr/>")]
    [InlineData("<w:noBreakHyphen/>")]
    [InlineData("<w:softHyphen/>")]
    public void Rejecting_a_tracked_atom_replacement_restores_the_original_carrier(string atom)
    {
        using var session = new DocxSession(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t>" + atom + "<w:t>B</w:t></w:r></w:p>")
            .DocumentByteArray, new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchor = session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p").Anchor.Id;
        Assert.True(session.ReplaceTextAtSpan(anchor, 1, 1, "X").Success);
        var saved = session.Save();
        using var accepted = new DocxSession(saved);
        Assert.True(accepted.AcceptAllRevisions().Success);
        Assert.Single(accepted.Grep("^AXB$"));
        Assert.True(session.RejectAllRevisions().Success);
        using var restored = WordprocessingDocument.Open(new MemoryStream(session.Save()), false);
        var expected = XElement.Parse("<root xmlns:w=\"" + W + "\">" + atom + "</root>").Elements().Single();
        var actual = Assert.Single(restored.MainDocumentPart!.GetXDocument().Descendants(expected.Name));
        Assert.Equal(expected.Attributes().Select(a => (a.Name, a.Value)), actual.Attributes().Select(a => (a.Name, a.Value)));
        Assert.NotNull(actual.Parent!.Element(W + "rPr")?.Element(W + "b"));
    }

    [Theory]
    [InlineData("zzzz")]
    [InlineData("0000")]
    [InlineData("D800")]
    [InlineData("FFFF")]
    public void Invalid_symbol_codes_stay_opaque_and_do_not_break_projection(string code)
    {
        var doc = IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>A</w:t><w:sym w:font=\"Symbol\" w:char=\"" +
            code + "\"/><w:t>B</w:t></w:r></w:p>");
        using var session = new DocxSession(doc.DocumentByteArray);
        Assert.Single(session.Grep("^AB$"));
        Assert.Equal(session.Project().Markdown, IrMarkdownEmitter.Emit(IrReader.Read(new WmlDocument(doc)), new()).Markdown);
    }

    [Theory]
    [InlineData(TableMergeContent.Append)]
    [InlineData(TableMergeContent.Reject)]
    public void A_tab_only_cell_is_not_empty_when_merging(TableMergeContent mode)
    {
        using var session = new DocxSession(IrTestDocuments.FromBodyXml(
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2400\"/><w:gridCol w:w=\"2400\"/></w:tblGrid><w:tr>" +
            "<w:tc><w:tcPr/><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
            "<w:tc><w:tcPr/><w:p><w:r><w:tab/></w:r></w:p></w:tc></w:tr></w:tbl>").DocumentByteArray);
        var first = session.Project().AnchorIndex.Values.First(a => a.Anchor.Kind == "tc").Anchor.Id;

        var edit = session.MergeCells(first, 1, 2, new TableMergeOptions { Content = mode });

        Assert.Equal(mode == TableMergeContent.Append, edit.Success);
        using var saved = WordprocessingDocument.Open(new MemoryStream(session.Save()), false);
        Assert.Single(saved.MainDocumentPart!.GetXDocument().Descendants(W + "tab"));
    }

    [Theory]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(true, false)]
    public void Filling_a_control_emits_the_changed_markdown_unless_patches_are_disabled(bool inline, bool patches)
    {
        const string value = "<w:r><w:t>Original value</w:t></w:r>";
        var control = "<w:sdt><w:sdtPr><w:id w:val=\"5\"/><w:text/></w:sdtPr><w:sdtContent>" +
            (inline ? value : "<w:p>" + value + "</w:p>") + "</w:sdtContent></w:sdt>";
        using var session = new DocxSession(IrTestDocuments.FromBodyXml(inline ? "<w:p>" + control + "</w:p>" : control)
            .DocumentByteArray, new DocxSessionSettings { EmitMarkdownPatch = patches });
        session.Project();

        var edit = session.FillContentControlText(Assert.Single(session.ListContentControls()).AnchorId, "Updated value");

        Assert.True(edit.Success, edit.Error?.Message);
        if (patches) Assert.Contains("Updated value", Assert.IsType<MarkdownPatch>(edit.Patch).Markdown);
        else Assert.Null(edit.Patch);
    }
}
