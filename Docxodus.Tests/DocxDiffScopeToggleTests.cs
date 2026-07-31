#nullable enable

using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace OxPt;

/// <summary>
/// Word Compare's "Textboxes" comparison option (<see cref="DocxDiffSettings.CompareTextboxes"/>) — a
/// granularity switch: off, a changed box is deleted-and-reinserted wholesale instead of inner-diffed.
/// </summary>
public class DocxDiffScopeToggleTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// One body paragraph whose run hosts a VML textbox containing a paragraph. Callers pass DISTINCT
    /// shape ids when the box content differs: the output then keeps both the deleted and the inserted
    /// shape, and the renderer does not renumber VML shape ids, so shared ids make the output
    /// schema-invalid regardless of this option (reproduced with it on as well — a pre-existing renderer
    /// gap, tracked separately).
    /// </summary>
    private static WmlDocument TextboxDoc(string bodyText, string boxText, string shapeId)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(new DocDefaults(
                new RunPropertiesDefault(new RunPropertiesBaseStyle(
                    new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" }))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            using var stream = main.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(
                $"<w:document xmlns:w=\"{W}\" xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                "xmlns:w10=\"urn:schemas-microsoft-com:office:word\"><w:body>" +
                $"<w:p><w:r><w:t xml:space=\"preserve\">{bodyText}</w:t></w:r>" +
                $"<w:r><w:pict><v:shape id=\"{shapeId}\" type=\"#_x0000_t202\" style=\"width:200pt;height:50pt\">" +
                "<v:textbox><w:txbxContent>" +
                $"<w:p><w:r><w:t xml:space=\"preserve\">{boxText}</w:t></w:r></w:p>" +
                "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>");
        }
        return new WmlDocument("textbox.docx", ms.ToArray());
    }

    private static void AssertValid(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        Assert.Empty(new OpenXmlValidator().Validate(wDoc).Select(e => e.Description));
    }

    private static string BodyXml(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray.ToArray());
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.Document.Body!.OuterXml;
    }

    private static WmlDocument ChangedBoxLeft => TextboxDoc("Body.", "in box one", "s1");
    private static WmlDocument ChangedBoxRight => TextboxDoc("Body.", "in box two", "s2");

    [Fact]
    public void CompareTextboxes_defaults_to_true()
    {
        Assert.True(new DocxDiffSettings().CompareTextboxes);
    }

    [Fact]
    public void Default_reports_a_textbox_change_as_a_per_box_inner_diff()
    {
        var json = DocxDiff.GetEditScriptJson(ChangedBoxLeft, ChangedBoxRight);
        Assert.Contains("textboxDiffs", json);
    }

    [Fact]
    public void CompareTextboxes_false_drops_the_per_box_inner_diff()
    {
        var json = DocxDiff.GetEditScriptJson(ChangedBoxLeft, ChangedBoxRight,
            new DocxDiffSettings { CompareTextboxes = false });
        Assert.DoesNotContain("textboxDiffs", json);
    }

    [Fact]
    public void CompareTextboxes_false_still_tracks_the_change_as_a_wholesale_box_replacement()
    {
        // The point of the option: granularity, not suppression. The box is del/ins'd by the host
        // paragraph's run diff, so both texts survive in the markup under revision wrappers.
        var compared = DocxDiff.Compare(ChangedBoxLeft, ChangedBoxRight,
            new DocxDiffSettings { CompareTextboxes = false });

        AssertValid(compared);
        var body = BodyXml(compared);
        Assert.Contains("in box one", body);
        Assert.Contains("in box two", body);
        Assert.Contains("<w:ins ", body);
        Assert.Contains("<w:del ", body);
    }

    [Fact]
    public void CompareTextboxes_false_round_trips()
    {
        var compared = DocxDiff.Compare(ChangedBoxLeft, ChangedBoxRight,
            new DocxDiffSettings { CompareTextboxes = false });

        var accepted = BodyXml(new WmlDocument("a.docx",
            Docxodus.Internal.DocxDiffOps.AcceptRevisions(compared.DocumentByteArray)));
        var rejected = BodyXml(new WmlDocument("r.docx",
            Docxodus.Internal.DocxDiffOps.RejectRevisions(compared.DocumentByteArray)));

        Assert.Contains("in box two", accepted);
        Assert.DoesNotContain("in box one", accepted);
        Assert.Contains("in box one", rejected);
        Assert.DoesNotContain("in box two", rejected);
    }

    [Fact]
    public void CompareTextboxes_false_still_reports_a_body_change()
    {
        // Same shape id and same box content on both sides, so the only difference is body text and
        // the assertion cannot pass for a textbox-related reason.
        var revisions = DocxDiff.GetRevisions(
            TextboxDoc("Body one.", "same box", "s1"),
            TextboxDoc("Body two.", "same box", "s1"),
            new DocxDiffSettings { CompareTextboxes = false });

        Assert.NotEmpty(revisions);
    }

    [Fact]
    public void CompareTextboxes_false_is_a_no_op_when_the_box_is_unchanged()
    {
        var left = TextboxDoc("Body.", "same box", "s1");
        var right = TextboxDoc("Body.", "same box", "s1");

        Assert.Empty(DocxDiff.GetRevisions(left, right, new DocxDiffSettings { CompareTextboxes = false }));
        Assert.Empty(DocxDiff.GetRevisions(left, right));
    }

    [Fact]
    public void CompareTextboxes_parses_from_the_settings_wire()
    {
        var withDetail = Docxodus.Internal.DocxDiffOps.GetEditScriptJson(
            ChangedBoxLeft.DocumentByteArray, ChangedBoxRight.DocumentByteArray, null);
        var withoutDetail = Docxodus.Internal.DocxDiffOps.GetEditScriptJson(
            ChangedBoxLeft.DocumentByteArray, ChangedBoxRight.DocumentByteArray,
            "{\"compareTextboxes\":false}");

        Assert.Contains("textboxDiffs", withDetail);
        Assert.DoesNotContain("textboxDiffs", withoutDetail);
    }
}
