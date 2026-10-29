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
    [InlineData(false)]
    [InlineData(true)]
    public void An_anchored_textbox_does_not_join_its_host_paragraph_text(bool control)
    {
        const string run = "<w:r><w:t>AB</w:t><w:pict><v:shape><v:textbox><w:txbxContent>" +
            "<w:p><w:r><w:t>Boxed note</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>";
        var inline = control ? "<w:sdt><w:sdtContent>" + run + "</w:sdtContent></w:sdt>" : run;
        using var session = new DocxSession(IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
            "<w:p>" + inline + "</w:p>").DocumentByteArray);

        Assert.Single(session.Grep("^AB$"));
        Assert.Single(session.Grep("^Boxed note$"));
        Assert.Empty(session.Grep("ABBoxed"));
    }

    [Fact]
    public void Inline_custom_xml_inside_a_control_projects_its_text()
    {
        var doc = IrTestDocuments.FromBodyXml("<w:p><w:sdt><w:sdtContent>" +
            "<w:customXml w:uri=\"urn:test\" w:element=\"clause\"><w:r><w:t>Wrapped value</w:t></w:r>" +
            "</w:customXml></w:sdtContent></w:sdt></w:p>");
        using var session = new DocxSession(doc.DocumentByteArray);

        Assert.Single(session.Grep("^Wrapped value$"));
        Assert.Contains("} Wrapped value", session.Project().Markdown);
        Assert.Equal(session.Project().Markdown,
            IrMarkdownEmitter.Emit(IrReader.Read(new WmlDocument(doc)), new()).Markdown);
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
