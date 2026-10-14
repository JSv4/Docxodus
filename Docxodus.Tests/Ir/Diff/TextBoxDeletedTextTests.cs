#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Word offers to recover a file when <c>w:delText</c> or <c>w:delInstrText</c> sits inside
/// <c>w:txbxContent</c>. One such element is enough. The same elements in the body, or in a
/// header or footer that has no text box, stay deleted text and Word opens them.
/// Proved on docxodus redlines 1499, 1528, 1614, 1679, 1719, 1734, 1807, 1810, 1815, 1821, and 1823.
/// Several of those files carry the defect in more than one part. Each part is checked on its own.
/// </summary>
public class TextBoxDeletedTextTests
{
    private const string VmlBox = """
        <w:p><w:r><w:pict><v:shape style="width:100pt;height:40pt"><v:textbox><w:txbxContent>
        <w:p><w:r><w:t>one-boxed</w:t></w:r></w:p>
        </w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>
        """;

    private const string ShapeBox = """
        <w:p><w:r><w:drawing><wps:txbx><w:txbxContent>
        <w:p><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r></w:p>
        </w:txbxContent></wps:txbx></w:drawing></w:r></w:p>
        """;

    private const string BothCopies = """
        <w:p><w:r><mc:AlternateContent>
        <mc:Choice Requires="wps"><w:drawing><wps:txbx><w:txbxContent>
        <w:p><w:r><w:t>choice-copy</w:t></w:r></w:p>
        </w:txbxContent></wps:txbx></w:drawing></mc:Choice>
        <mc:Fallback><w:pict><v:shape style="width:100pt;height:40pt"><v:textbox><w:txbxContent>
        <w:p><w:r><w:t>fallback-copy</w:t></w:r></w:p>
        </w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>
        </mc:AlternateContent></w:r></w:p>
        """;

    [Fact]
    public void Render_keeps_deleted_text_ordinary_inside_a_text_box_and_deleted_in_the_body()
    {
        var left = IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
            VmlBox + ShapeBox + BothCopies +
            "<w:p><w:r><w:t>body-stays</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
            "<w:p><w:r><w:t>kept</w:t></w:r></w:p>");

        var rendered = Render(left, right);
        AssertNoDeletedTextInsideTextBoxes(rendered);
        Assert.Contains(TextBoxInstructions(rendered), text => text.Contains("PAGE"));

        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var body = doc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        var bodyDeleted = body.Descendants(W.delText)
            .Where(text => !text.Ancestors(W.txbxContent).Any())
            .Select(text => text.Value);
        Assert.Contains("body-stays", bodyDeleted);

        var accepted = RevisionProcessor.AcceptRevisions(rendered);
        var rejected = RevisionProcessor.RejectRevisions(rendered);
        Assert.DoesNotContain("one-boxed", VisibleText(accepted));
        Assert.DoesNotContain("body-stays", VisibleText(accepted));
        Assert.Contains("one-boxed", VisibleText(rejected));
        Assert.Contains("choice-copy", VisibleText(rejected));
        Assert.Contains("fallback-copy", VisibleText(rejected));
        Assert.Contains("body-stays", VisibleText(rejected));
        AssertNoDeletedTextInsideTextBoxes(rejected);
        Assert.Contains(TextBoxInstructions(rejected), text => text.Contains("PAGE"));
    }

    [Fact]
    public void Render_checks_each_header_and_footer_on_its_own()
    {
        var left = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(
                    new[] { "section-one" },
                    Headers: new[] { ("default", "rIdH1") },
                    Footers: new[] { ("default", "rIdF1") }),
                new HeaderFooterFixtures.Section(
                    new[] { "section-two" },
                    Headers: new[] { ("default", "rIdH2") },
                    Footers: new[] { ("default", "rIdF2") }),
            },
            headerParts: new Dictionary<string, string[]>
            {
                ["rIdH1"] = new[] { VmlParagraph("header-boxed"), "header-plain" },
                ["rIdH2"] = new[] { "header-no-box" },
            },
            footerParts: new Dictionary<string, string[]>
            {
                ["rIdF1"] = new[] { VmlParagraph("footer-one") },
                ["rIdF2"] = new[] { "<w:p><w:r><w:instrText xml:space=\"preserve\"> NUMPAGES </w:instrText></w:r></w:p>" },
            });
        var right = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(
                    new[] { "section-one" },
                    Headers: new[] { ("default", "rIdH1") },
                    Footers: new[] { ("default", "rIdF1") }),
                new HeaderFooterFixtures.Section(
                    new[] { "section-two" },
                    Headers: new[] { ("default", "rIdH2") },
                    Footers: new[] { ("default", "rIdF2") }),
            },
            headerParts: new Dictionary<string, string[]>
            {
                ["rIdH1"] = new[] { "kept-header" },
                ["rIdH2"] = new[] { "zzzz" },
            },
            footerParts: new Dictionary<string, string[]>
            {
                ["rIdF1"] = new[] { "kept-footer" },
                ["rIdF2"] = new[] { "kept-footer-two" },
            });

        var rendered = Render(left, right);
        AssertNoDeletedTextInsideTextBoxes(rendered);

        using var ms = new MemoryStream(rendered.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        var headers = main.HeaderParts.Select(part => part.GetXDocument().Root!).ToList();
        var footers = main.FooterParts.Select(part => part.GetXDocument().Root!).ToList();

        Assert.Contains(headers, root =>
            root.Descendants(W.delText).Any(text =>
                text.Value.Contains("header-plain") && !text.Ancestors(W.txbxContent).Any()));
        Assert.Contains(headers, root =>
            root.Descendants(W.delText).Any(text =>
                text.Value.Contains("header-no-box") && !text.Ancestors(W.txbxContent).Any()));
        Assert.Equal(2, footers.Count);
        Assert.All(footers, root =>
        {
            Assert.Empty(root.Descendants(W.txbxContent).SelectMany(box => box.Descendants(W.delText)));
            Assert.Empty(root.Descendants(W.txbxContent).SelectMany(box => box.Descendants(W.delInstrText)));
        });
        Assert.Contains(footers, root => root.Descendants(W.txbxContent).Any(box => box.Value.Contains("footer-one")));
        Assert.Contains(footers, root => root.Descendants(W.delInstrText).Any(text =>
            text.Value.Contains("NUMPAGES") && !text.Ancestors(W.txbxContent).Any()));
    }

    [Fact]
    public void Session_tracked_fill_of_a_content_control_inside_a_text_box_keeps_its_text_ordinary()
    {
        var original = IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
            "<w:p><w:r><w:pict><v:shape style=\"width:100pt;height:40pt\"><v:textbox><w:txbxContent>" +
            "<w:p><w:sdt><w:sdtPr><w:id w:val=\"7\"/></w:sdtPr><w:sdtContent>" +
            "<w:r><w:t>boxed-control</w:t></w:r>" +
            "</w:sdtContent></w:sdt></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>");
        using var session = new DocxSession(original.DocumentByteArray, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Reviewer",
        });
        var control = session.ListContentControls().Single();
        var filled = session.FillContentControlText(control.AnchorId, "replacement");
        Assert.True(filled.Success, filled.Error?.Message);

        var tracked = new WmlDocument("tracked.docx", session.Save());
        AssertNoDeletedTextInsideTextBoxes(tracked);
        using var ms = new MemoryStream(tracked.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var boxes = doc.MainDocumentPart!.GetXDocument().Descendants(W.txbxContent).ToList();
        Assert.Contains(boxes, box => box.Descendants(W.t).Any(text => text.Value.Contains("boxed-control")));
        Assert.Contains(boxes, box => box.Descendants(W.t).Any(text => text.Value.Contains("replacement")));
    }

    [Fact]
    public void Session_delete_of_a_text_box_paragraph_keeps_its_text_ordinary()
    {
        var original = IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
            "<w:p><w:r><w:t>before</w:t></w:r></w:p>" + VmlBox + ShapeBox);
        using var session = new DocxSession(original.DocumentByteArray, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Reviewer",
        });
        // The box's own paragraph and the paragraph that contains the drawing both preview
        // "one-boxed". Document order lists the containing paragraph first. Deleting that one
        // is the shape Word recovered from: the drawing run is wrapped, and the text inside
        // the box must stay ordinary.
        var textBoxes = session.Project().AnchorIndex.Values
            .Where(anchor => anchor.Anchor.Kind == "p" && anchor.TextPreview.Contains("one-boxed"))
            .ToList();
        Assert.Equal(2, textBoxes.Count);
        var deleted = session.DeleteBlock(textBoxes[0].Anchor.Id);
        Assert.True(deleted.Success, deleted.Error?.Message);
        var tracked = new WmlDocument("tracked.docx", session.Save());
        AssertNoDeletedTextInsideTextBoxes(tracked);

        using var ms = new MemoryStream(tracked.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var body = doc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        Assert.Contains(body.Descendants(W.txbxContent), box => box.Value.Contains("one-boxed"));
        Assert.DoesNotContain(body.Descendants(W.txbxContent), box =>
            box.Descendants(W.delText).Any() || box.Descendants(W.delInstrText).Any());
    }

    [Fact]
    public void Session_tracked_span_replace_inside_a_text_box_keeps_its_text_ordinary()
    {
        var original = IrTestDocuments.FromBodyXmlWithDrawingNamespaces(VmlBox);
        using var session = new DocxSession(original.DocumentByteArray, new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Reviewer",
        });
        var textBoxes = session.Project().AnchorIndex.Values
            .Where(anchor => anchor.Anchor.Kind == "p" && anchor.TextPreview.Contains("one-boxed"))
            .ToList();
        Assert.Equal(2, textBoxes.Count);
        // A span replace builds its deleted slice directly (CloneRunWithRevisionText),
        // bypassing ConvertTextToDeletedText — the deleted run must keep w:t inside
        // the box all the same.
        var replaced = session.ReplaceTextAtSpan(
            textBoxes[1].Anchor.Id, 0, "one-boxed".Length, "two-boxed");
        Assert.True(replaced.Success, replaced.Error?.Message);

        var tracked = new WmlDocument("tracked.docx", session.Save());
        AssertNoDeletedTextInsideTextBoxes(tracked);
        using var ms = new MemoryStream(tracked.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var boxes = doc.MainDocumentPart!.GetXDocument().Descendants(W.txbxContent).ToList();
        Assert.Contains(boxes, box => box.Descendants(W.t).Any(text => text.Value.Contains("one-boxed")));
        Assert.Contains(boxes, box => box.Descendants(W.t).Any(text => text.Value.Contains("two-boxed")));
    }

    private static WmlDocument Render(WmlDocument left, WmlDocument right)
    {
        var settings = new IrDiffSettings();
        var irLeft = IrReader.Read(left, new IrReaderOptions { RetainSources = false, RevisionView = RevisionView.Accept });
        var irRight = IrReader.Read(right, new IrReaderOptions { RetainSources = false, RevisionView = RevisionView.Accept });
        var script = IrEditScriptBuilder.Build(irLeft, irRight, settings);
        return IrMarkupRenderer.Render(script, left, right, settings);
    }

    private static void AssertNoDeletedTextInsideTextBoxes(WmlDocument document)
    {
        using var ms = new MemoryStream(document.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        var roots = new List<XElement> { main.GetXDocument().Root! };
        roots.AddRange(main.HeaderParts.Select(part => part.GetXDocument().Root!));
        roots.AddRange(main.FooterParts.Select(part => part.GetXDocument().Root!));
        foreach (var root in roots)
        {
            foreach (var box in root.Descendants(W.txbxContent))
            {
                Assert.Empty(box.Descendants(W.delText));
                Assert.Empty(box.Descendants(W.delInstrText));
            }
        }
    }

    private static IEnumerable<string> TextBoxInstructions(WmlDocument document)
    {
        using var ms = new MemoryStream(document.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        var roots = new List<XElement> { main.GetXDocument().Root! };
        roots.AddRange(main.HeaderParts.Select(part => part.GetXDocument().Root!));
        roots.AddRange(main.FooterParts.Select(part => part.GetXDocument().Root!));
        return roots.SelectMany(root => root.Descendants(W.txbxContent))
            .SelectMany(box => box.Descendants(W.instrText))
            .Select(text => text.Value)
            .ToList();
    }

    private static string VisibleText(WmlDocument document)
    {
        using var ms = new MemoryStream(document.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        var roots = new List<XElement> { main.GetXDocument().Root! };
        roots.AddRange(main.HeaderParts.Select(part => part.GetXDocument().Root!));
        roots.AddRange(main.FooterParts.Select(part => part.GetXDocument().Root!));
        return string.Concat(roots.SelectMany(root => root.Descendants(W.t).Select(text => text.Value)));
    }

    private static string VmlParagraph(string text) =>
        "<w:p><w:r><w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
        "style=\"width:100pt;height:40pt\"><v:textbox><w:txbxContent><w:p><w:r><w:t>" +
        text + "</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>";
}
