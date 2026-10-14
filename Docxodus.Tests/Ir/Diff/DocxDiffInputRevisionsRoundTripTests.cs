#nullable enable

using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;
using WordType = DocumentFormat.OpenXml.WordprocessingDocumentType;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Content-preservation round trip when BOTH input documents already carry UN-accepted tracked changes
/// (inline <c>w:ins</c>/<c>w:del</c> and a paragraph <c>w:moveFrom</c>/<c>w:moveTo</c>) — the "redline of a
/// redline" case. Default <see cref="DocxDiffSettings"/> compares the ACCEPTED VIEW of each side, so the
/// contract is: <c>accept(Compare(l,r))</c> reproduces the accepted view of the RIGHT input, and
/// <c>reject(...)</c> the accepted view of the LEFT — with no content ever lost. The pre-existing revision
/// scenario is otherwise only exercised by the pre-accept/preserve tests and the fuzzer (which reads under
/// <c>RevisionView.Accept</c> and never seeds carried-over revisions), so this pins the default-settings
/// content guarantee directly. Documents are synthesized here; no external fixtures.
/// </summary>
public class DocxDiffInputRevisionsRoundTripTests
{
    private const string Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string A = "Prior Author";
    private const string D = "2020-01-01T00:00:00Z";

    private static string Ins(string text) =>
        $"<w:ins w:id=\"801\" w:author=\"{A}\" w:date=\"{D}\"><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:ins>";

    private static string Del(string text) =>
        $"<w:del w:id=\"802\" w:author=\"{A}\" w:date=\"{D}\"><w:r><w:delText xml:space=\"preserve\">{text}</w:delText></w:r></w:del>";

    /// <summary>A paragraph tracked-MOVED away (moveFrom) and its destination paragraph (moveTo). In the
    /// accepted view the moveFrom paragraph vanishes and only the moveTo copy survives.</summary>
    private static string MovePair(string movedText) =>
        "<w:p><w:moveFromRangeStart w:id=\"810\" w:name=\"mv1\" w:author=\"" + A + "\" w:date=\"" + D + "\"/>" +
        $"<w:moveFrom w:id=\"811\" w:author=\"{A}\" w:date=\"{D}\"><w:r><w:t xml:space=\"preserve\">{movedText}</w:t></w:r></w:moveFrom>" +
        "<w:moveFromRangeEnd w:id=\"810\"/></w:p>" +
        "<w:p><w:moveToRangeStart w:id=\"812\" w:name=\"mv1\" w:author=\"" + A + "\" w:date=\"" + D + "\"/>" +
        $"<w:moveTo w:id=\"813\" w:author=\"{A}\" w:date=\"{D}\"><w:r><w:t xml:space=\"preserve\">{movedText}</w:t></w:r></w:moveTo>" +
        "<w:moveToRangeEnd w:id=\"812\"/></w:p>";

    private static WmlDocument Doc(string name, string leadText)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();

            var body =
                $"<w:document xmlns:w=\"{Wns}\"><w:body>" +
                // plain lead paragraph (varies per side so the diff has real work)
                $"<w:p><w:r><w:t xml:space=\"preserve\">{leadText} </w:t></w:r>" +
                // a pre-existing tracked insertion + deletion inline
                Ins("inserted-word") + Del("deleted-word") + "</w:p>" +
                // a stable paragraph both sides share verbatim
                "<w:p><w:r><w:t xml:space=\"preserve\">Stable shared paragraph.</w:t></w:r></w:p>" +
                // a pre-existing tracked MOVE (moveFrom + moveTo)
                MovePair("Relocated sentence.") +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>" +
                "</w:body></w:document>";
            using var w = new StreamWriter(main.GetStream(FileMode.Create));
            w.Write(body);
        }
        return new WmlDocument(name, ms.ToArray());
    }

    /// <summary>Whole-document visible text under the ACCEPTED view (all tracked changes accepted: del/moveFrom
    /// gone, ins/moveTo kept), whitespace-normalized — the content that must be preserved.</summary>
    private static string AcceptedText(WmlDocument d)
    {
        var accepted = RevisionProcessor.AcceptRevisions(d);
        using var ms = new MemoryStream(accepted.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        var xml = wDoc.MainDocumentPart!.Document;
        var text = string.Concat(xml.Descendants()
            .Where(e => e.LocalName == "t" || e.LocalName == "tab")
            .Select(e => e.LocalName == "tab" ? " " : e.InnerText));
        return string.Join(" ", text.Split((char[]?)null, System.StringSplitOptions.RemoveEmptyEntries));
    }

    [Fact]
    public void RoundTrip_PreservesContent_WhenInputsCarryTrackedChangesAndMoves()
    {
        var left = Doc("left.docx", "Left lead");
        var right = Doc("right.docx", "Right lead edited");

        var redline = DocxDiff.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);

        var leftAccepted = AcceptedText(left);
        var rightAccepted = AcceptedText(right);

        // Sanity: the accepted views genuinely differ (the diff had real work), and the pre-existing move
        // collapsed to a single copy of the relocated sentence on each side (no double-count).
        Assert.NotEqual(leftAccepted, rightAccepted);
        Assert.Contains("Relocated sentence.", rightAccepted);
        Assert.Equal(1, CountOccurrences(rightAccepted, "Relocated sentence."));

        // The contract: accept ≡ accepted view of RIGHT, reject ≡ accepted view of LEFT — zero content loss,
        // even though both inputs carried un-accepted tracked insertions, deletions, and a move.
        Assert.Equal(rightAccepted, AcceptedText(accepted));
        Assert.Equal(leftAccepted, AcceptedText(rejected));
    }

    private static int CountOccurrences(string haystack, string needle)
    {
        int n = 0, i = 0;
        while ((i = haystack.IndexOf(needle, i, System.StringComparison.Ordinal)) >= 0) { n++; i += needle.Length; }
        return n;
    }

    // ---------------------------------------------------------------- table scope

    private static WmlDocument TableDoc(string name, string cellLead)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();

            string Cell(string inner) =>
                "<w:tc><w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/></w:tcPr><w:p>" + inner + "</w:p></w:tc>";
            var body =
                $"<w:document xmlns:w=\"{Wns}\"><w:body>" +
                "<w:tbl><w:tblPr><w:tblW w:w=\"8000\" w:type=\"dxa\"/></w:tblPr>" +
                "<w:tblGrid><w:gridCol w:w=\"4000\"/><w:gridCol w:w=\"4000\"/></w:tblGrid>" +
                // row 1: first cell has a pre-existing tracked ins+del; text varies per side
                "<w:tr>" +
                Cell($"<w:r><w:t xml:space=\"preserve\">{cellLead} </w:t></w:r>" + Ins("cell-inserted") + Del("cell-deleted")) +
                Cell("<w:r><w:t xml:space=\"preserve\">Fixed cell.</w:t></w:r>") +
                "</w:tr>" +
                // row 2: stable content both sides
                "<w:tr>" + Cell("<w:r><w:t xml:space=\"preserve\">R2C1 stable.</w:t></w:r>") +
                Cell("<w:r><w:t xml:space=\"preserve\">R2C2 stable.</w:t></w:r>") + "</w:tr>" +
                "</w:tbl>" +
                "<w:p><w:r><w:t xml:space=\"preserve\">Trailing paragraph.</w:t></w:r></w:p>" +
                "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>" +
                "</w:body></w:document>";
            using var w = new StreamWriter(main.GetStream(FileMode.Create));
            w.Write(body);
        }
        return new WmlDocument(name, ms.ToArray());
    }

    [Fact]
    public void RoundTrip_PreservesTableContent_WhenCellsCarryTrackedChanges()
    {
        var left = TableDoc("tleft.docx", "Left cell");
        var right = TableDoc("tright.docx", "Right cell edited");

        var redline = DocxDiff.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);

        var leftAccepted = AcceptedText(left);
        var rightAccepted = AcceptedText(right);

        Assert.NotEqual(leftAccepted, rightAccepted);
        Assert.Contains("cell-inserted", rightAccepted);   // ins survived accept
        Assert.DoesNotContain("cell-deleted", rightAccepted); // del removed on accept
        Assert.Contains("R2C2 stable.", rightAccepted);    // untouched cell content kept

        // Content contract across the table scope: no cell text lost on accept or reject.
        Assert.Equal(rightAccepted, AcceptedText(accepted));
        Assert.Equal(leftAccepted, AcceptedText(rejected));
    }
}
