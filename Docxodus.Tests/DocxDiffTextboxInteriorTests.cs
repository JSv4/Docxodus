#nullable enable

using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace OxPt;

/// <summary>
/// A change confined to a textbox must be tracked INSIDE the box — Word's shape — rather than deleting and
/// re-inserting the whole drawing, which renders as two stacked floating boxes and reads as "not detected".
/// </summary>
public class DocxDiffTextboxInteriorTests
{
    private const string Namespaces =
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
        "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
        "xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
        "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" " +
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
        "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"";

    private static WmlDocument Doc(string bodyInner, string? headerInner = null)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var sectPr = "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>";
            if (headerInner is not null)
            {
                var header = main.AddNewPart<HeaderPart>("rIdHdr");
                Write(header, $"<w:hdr {Namespaces}>{headerInner}</w:hdr>");
                sectPr = "<w:sectPr><w:headerReference w:type=\"default\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                    "r:id=\"rIdHdr\"/><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>";
            }

            Write(main, $"<w:document {Namespaces}><w:body>{bodyInner}{sectPr}</w:body></w:document>");
        }
        return new WmlDocument("textbox.docx", ms.ToArray());
    }

    private static void Write(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream, new UTF8Encoding(false));
        writer.Write(xml);
    }

    /// <summary>A DrawingML textbox wrapped in mc:AlternateContent with a VML fallback — what Word writes.</summary>
    private static string DrawingMlBox(string boxText, string leadText = "") =>
        "<w:p>" +
        (leadText.Length == 0 ? "" : $"<w:r><w:t xml:space=\"preserve\">{leadText}</w:t></w:r>") +
        "<w:r><mc:AlternateContent><mc:Choice Requires=\"wps\"><w:drawing>" +
        "<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"2292350\" cy=\"414655\"/>" +
        "<wp:docPr id=\"1\" name=\"TextBox 1\"/><a:graphic>" +
        "<a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
        "<wps:wsp><wps:cNvSpPr txBox=\"1\"/><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>" +
        "<a:ext cx=\"2292350\" cy=\"414655\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></wps:spPr>" +
        $"<wps:txbx><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">{boxText}</w:t></w:r></w:p></w:txbxContent></wps:txbx>" +
        "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>" +
        "<mc:Fallback><w:pict><v:shape id=\"fb1\" style=\"width:180pt;height:32pt\"/></w:pict></mc:Fallback>" +
        "</mc:AlternateContent></w:r></w:p>";

    /// <summary>A plain VML textbox (no AlternateContent).</summary>
    private static string VmlBox(string boxText, string shapeId = "s1") =>
        "<w:p><w:r><w:pict>" +
        $"<v:shape id=\"{shapeId}\" style=\"width:200pt;height:50pt\">" +
        $"<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">{boxText}</w:t></w:r></w:p>" +
        "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>";

    private static string Para(string text) =>
        $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static void AssertValid(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        Assert.Empty(new OpenXmlValidator().Validate(wDoc).Select(e => e.Description));
    }

    private static string PartXml(byte[] bytes, bool header = false)
    {
        using var ms = new MemoryStream(bytes.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return header
            ? wDoc.MainDocumentPart!.HeaderParts.First().Header.OuterXml
            : wDoc.MainDocumentPart!.Document.Body!.OuterXml;
    }

    /// <summary>The visible text of every textbox interior, boxes separated by " || ".</summary>
    private static string BoxText(byte[] bytes, bool header = false)
    {
        var matches = Regex.Matches(PartXml(bytes, header), "(?s)<w:txbxContent>.*?</w:txbxContent>");
        return string.Join(" || ", matches.Select(m =>
            Regex.Replace(Regex.Replace(m.Value, "<[^>]+>", ""), @"\s+", " ").Trim()));
    }

    private static int Count(string xml, string tag) => Regex.Matches(xml, tag).Count;

    // ------------------------------------------------------------------ the reported bug

    [Fact]
    public void A_changed_textbox_is_tracked_inside_the_box_not_duplicated()
    {
        var compared = DocxDiff.Compare(Doc(DrawingMlBox("This is textbox 1")), Doc(DrawingMlBox("This is textbox 2")));
        AssertValid(compared);
        var xml = PartXml(compared.DocumentByteArray);

        Assert.Equal(1, Count(xml, "<w:drawing"));        // ONE drawing, not a deleted + an inserted copy
        Assert.Equal(1, Count(xml, "<w:txbxContent>"));
        // …with the revision markup INSIDE the box.
        var interior = Regex.Match(xml, "(?s)<w:txbxContent>.*?</w:txbxContent>").Value;
        Assert.Contains("<w:ins ", interior);
        Assert.Contains("<w:del ", interior);
    }

    [Fact]
    public void A_changed_textbox_round_trips()
    {
        var compared = DocxDiff.Compare(Doc(DrawingMlBox("This is textbox 1")), Doc(DrawingMlBox("This is textbox 2")));

        Assert.Equal("This is textbox 2", BoxText(DocxDiffOps.AcceptRevisions(compared.DocumentByteArray)));
        Assert.Equal("This is textbox 1", BoxText(DocxDiffOps.RejectRevisions(compared.DocumentByteArray)));
    }

    [Fact]
    public void A_changed_VML_textbox_is_tracked_inside_the_box_too()
    {
        var compared = DocxDiff.Compare(Doc(VmlBox("box one")), Doc(VmlBox("box two")));
        AssertValid(compared);

        Assert.Equal(1, Count(PartXml(compared.DocumentByteArray), "<w:txbxContent>"));
        Assert.Equal("box two", BoxText(DocxDiffOps.AcceptRevisions(compared.DocumentByteArray)));
        Assert.Equal("box one", BoxText(DocxDiffOps.RejectRevisions(compared.DocumentByteArray)));
    }

    [Fact]
    public void A_textbox_in_a_header_is_tracked_inside_the_box()
    {
        var compared = DocxDiff.Compare(
            Doc(Para("Body."), VmlBox("header box one")),
            Doc(Para("Body."), VmlBox("header box two")));
        AssertValid(compared);

        Assert.Equal(1, Count(PartXml(compared.DocumentByteArray, header: true), "<w:txbxContent>"));
        Assert.Equal("header box two", BoxText(DocxDiffOps.AcceptRevisions(compared.DocumentByteArray), header: true));
        Assert.Equal("header box one", BoxText(DocxDiffOps.RejectRevisions(compared.DocumentByteArray), header: true));
    }

    // ------------------------------------------------------------------ interactions

    [Fact]
    public void A_paragraph_that_changes_around_its_textbox_tracks_both()
    {
        var compared = DocxDiff.Compare(
            Doc(DrawingMlBox("box one", leadText: "Before alpha ")),
            Doc(DrawingMlBox("box two", leadText: "Before beta ")));
        AssertValid(compared);

        Assert.Equal("box two", BoxText(DocxDiffOps.AcceptRevisions(compared.DocumentByteArray)));
        Assert.Equal("box one", BoxText(DocxDiffOps.RejectRevisions(compared.DocumentByteArray)));

        var accepted = PartXml(DocxDiffOps.AcceptRevisions(compared.DocumentByteArray));
        var rejected = PartXml(DocxDiffOps.RejectRevisions(compared.DocumentByteArray));
        Assert.Contains("beta", accepted);
        Assert.DoesNotContain("alpha", accepted);
        Assert.Contains("alpha", rejected);
        Assert.DoesNotContain("beta", rejected);
    }

    [Fact]
    public void Only_the_changed_box_of_two_is_marked()
    {
        var compared = DocxDiff.Compare(
            Doc(VmlBox("first box", "a1") + VmlBox("second one", "a2")),
            Doc(VmlBox("first box", "a1") + VmlBox("second two", "a2")));
        AssertValid(compared);

        // Each box is emitted once, and the revision markup sits in the SECOND one only — the round-trip
        // assertions below hold under the wholesale path too, so they alone would not prove this.
        var interiors = Regex.Matches(PartXml(compared.DocumentByteArray), "(?s)<w:txbxContent>.*?</w:txbxContent>")
            .Select(m => m.Value).ToList();
        Assert.Equal(2, interiors.Count);
        Assert.DoesNotContain("<w:ins ", interiors[0]);
        Assert.DoesNotContain("<w:del ", interiors[0]);
        Assert.Contains("<w:ins ", interiors[1]);
        Assert.Contains("<w:del ", interiors[1]);

        Assert.Equal("first box || second two", BoxText(DocxDiffOps.AcceptRevisions(compared.DocumentByteArray)));
        Assert.Equal("first box || second one", BoxText(DocxDiffOps.RejectRevisions(compared.DocumentByteArray)));
    }

    [Fact]
    public void An_unchanged_textbox_produces_no_revisions()
    {
        var left = Doc(DrawingMlBox("same box"));
        Assert.Empty(DocxDiff.GetRevisions(left, Doc(DrawingMlBox("same box"))));
    }

    [Fact]
    public void CompareTextboxes_false_keeps_the_wholesale_replacement()
    {
        // With the inner diff switched off there are no nested ops to render, so the box is del/ins'd
        // wholesale by the host paragraph — the documented granularity trade-off, still round-tripping.
        var compared = DocxDiff.Compare(
            Doc(VmlBox("box one")), Doc(VmlBox("box two")),
            new DocxDiffSettings { CompareTextboxes = false });
        AssertValid(compared);

        Assert.Equal(2, Count(PartXml(compared.DocumentByteArray), "<w:txbxContent>"));
        Assert.Equal("box two", BoxText(DocxDiffOps.AcceptRevisions(compared.DocumentByteArray)));
        Assert.Equal("box one", BoxText(DocxDiffOps.RejectRevisions(compared.DocumentByteArray)));
    }

    [Fact]
    public void A_document_that_gains_a_textbox_paragraph_still_round_trips()
    {
        // A whole added paragraph is an InsertBlock — the interior guard is never consulted here.
        var compared = DocxDiff.Compare(
            Doc(VmlBox("only box", "b1")),
            Doc(VmlBox("only box", "b1") + VmlBox("added box", "b2")));
        AssertValid(compared);

        Assert.Contains("added box", BoxText(DocxDiffOps.AcceptRevisions(compared.DocumentByteArray)));
        Assert.DoesNotContain("added box", BoxText(DocxDiffOps.RejectRevisions(compared.DocumentByteArray)));
    }

    [Fact]
    public void A_paragraph_that_gains_a_second_textbox_still_round_trips()
    {
        // A surplus box WITHIN one paragraph: the emitted interiors and the nested diffs must still pair, or
        // TryRenderTextboxInteriors commits nothing and the conservative whole-block path runs instead.
        static string TwoBoxRun(string first, string? second) =>
            "<w:p><w:r><w:pict>" +
            "<v:shape id=\"c1\" style=\"width:200pt;height:50pt\">" +
            $"<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">{first}</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r>" +
            (second is null ? "" :
                "<w:r><w:pict><v:shape id=\"c2\" style=\"width:200pt;height:50pt\">" +
                $"<v:textbox><w:txbxContent><w:p><w:r><w:t xml:space=\"preserve\">{second}</w:t></w:r></w:p>" +
                "</w:txbxContent></v:textbox></v:shape></w:pict></w:r>") +
            "</w:p>";

        var compared = DocxDiff.Compare(Doc(TwoBoxRun("kept box", null)), Doc(TwoBoxRun("kept box", "extra box")));
        AssertValid(compared);

        Assert.Contains("extra box", BoxText(DocxDiffOps.AcceptRevisions(compared.DocumentByteArray)));
        Assert.DoesNotContain("extra box", BoxText(DocxDiffOps.RejectRevisions(compared.DocumentByteArray)));
        Assert.Contains("kept box", BoxText(DocxDiffOps.RejectRevisions(compared.DocumentByteArray)));
    }
}
