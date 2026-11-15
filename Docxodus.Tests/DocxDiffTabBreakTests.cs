using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A tab or line break in a paragraph does not change how <see cref="DocxCompare"/> diffs the text around it
/// (issue #1759). The reference compare treats either like a space, so each tabbed pair below draws the same
/// revisions as its plain control.
/// </summary>
public class DocxDiffTabBreakTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Space = "<w:r><w:t xml:space=\"preserve\"> </w:t></w:r>";
    private const string Tab = "<w:r><w:tab/></w:r>";
    private const string LineBreak = "<w:r><w:br/></w:r>";

    private static string R(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private static WmlDocument Doc(params string[] paragraphRuns) =>
        IrTestDocuments.FromBodyXml(string.Concat(paragraphRuns.Select(runs => $"<w:p>{runs}</w:p>")));

    [Theory]
    [InlineData("", Space, "")] // control: no tab or break
    [InlineData(Tab, Space, "")]
    [InlineData("", Space, Tab)]
    [InlineData(LineBreak, Space, "")]
    [InlineData("", Tab, "")] // a tab or break between the replaced words
    [InlineData("", LineBreak, "")]
    public void Replaced_phrase_is_one_insertion_and_one_deletion(string before, string join, string after)
    {
        // No punctuation, so no lone-punctuation rule decides whether a tab between the phrases' words may anchor.
        var left = Doc(before + R("Every harbor pilot logs tide") + join + R("charts before noon") + after);
        var right = Doc(before + R("Every harbor pilot checks wind gauges") + join + R("and buoy lamps before noon") + after);
        var result = DocxCompare.Compare(left, right);

        Assert.Equal(new[] { Plain(before) + "Every harbor pilot " +
            $"[+checks wind gauges{Plain(join)}and buoy lamps][-logs tide{Plain(join)}charts] before noon" + Plain(after) },
            Paragraphs(result));
        AssertRoundTrip(result, left, right);
    }

    [Theory]
    [InlineData("")] // control: no tab
    [InlineData(Tab)]
    [InlineData(LineBreak)]
    public void Comma_between_changed_words_goes_with_them(string lead)
    {
        // A lone shared comma is not kept on its own: the reference replaces "Ropes, nets" whole.
        var left = Doc(lead + R("Ropes, nets and hooks are stowed."));
        var right = Doc(lead + R("Sails, oars and cleats are stowed."));
        var result = DocxCompare.Compare(left, right);

        Assert.Equal(new[] { Plain(lead) + "[+Sails, oars][-Ropes, nets] and [+cleats][-hooks] are stowed." }, Paragraphs(result));
        AssertRoundTrip(result, left, right);
    }

    [Theory]
    [InlineData(false)] // control: "(3) text" in one run
    [InlineData(true)]  // tab, "(3)", tab, text
    public void Rewritten_clause_is_replaced_whole(bool tabbed)
    {
        const string oldText = "For courier by with for warehouse on supplier voucher for by ledger supplier manifest courier with.";
        const string newText = "Lease garden tenant of for utilities and utilities on fixtures notice repair fixtures by notice tenant.";
        string Clause(string text) => tabbed ? Tab + R("(3)") + Tab + R(text) : R("(3) " + text);
        string Label(string text) => (tabbed ? "\t(3)\t" : "(3) ") + text;
        var left = Doc(R("Schedule of works"), Clause(oldText), R("Signed and dated below."));
        var right = Doc(R("Schedule of works"), Clause(newText), R("Signed and dated below."));
        var result = DocxCompare.Compare(left, right);

        Assert.Equal(new[] { "Schedule of works", $"[+{Label(newText)}]", $"[-{Label(oldText)}]", "Signed and dated below." },
            Paragraphs(result));
        AssertRoundTrip(result, left, right);
    }

    /// <summary>The text of <paramref name="runs"/>, as <see cref="Marked"/> writes it.</summary>
    private static string Plain(string runs) => Marked(XElement.Parse($"<w:p xmlns:w=\"{W}\">{runs}</w:p>"));

    private static string[] Paragraphs(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var package = WordprocessingDocument.Open(stream, false);
        return XDocument.Load(package.MainDocumentPart!.GetStream()).Root!.Element(W + "body")!
            .Elements(W + "p").Select(Marked).ToArray();
    }

    /// <summary>A paragraph's text, a tab as <c>\t</c> and a break as <c>\n</c>, with each inserted span written
    /// <c>[+…]</c> and each deleted span <c>[-…]</c>.</summary>
    private static string Marked(XElement paragraph)
    {
        var text = new StringBuilder();
        char open = '=';
        foreach (var e in paragraph.Descendants().Where(e => e.Parent!.Name == W + "r"))
        {
            string piece = e.Name == W + "t" || e.Name == W + "delText" ? e.Value
                : e.Name == W + "tab" ? "\t" : e.Name == W + "br" ? "\n" : "";
            if (piece.Length == 0)
                continue;
            char kind = e.Ancestors(W + "ins").Any() ? '+' : e.Ancestors(W + "del").Any() ? '-' : '=';
            if (kind != open)
            {
                text.Append(open == '=' ? "" : "]").Append(kind == '=' ? "" : "[" + kind);
                open = kind;
            }
            text.Append(piece);
        }
        return text.Append(open == '=' ? "" : "]").ToString();
    }

    private static void AssertRoundTrip(WmlDocument result, WmlDocument left, WmlDocument right)
    {
        Assert.Equal(Paragraphs(right), Paragraphs(RevisionProcessor.AcceptRevisions(result)));
        Assert.Equal(Paragraphs(left), Paragraphs(RevisionProcessor.RejectRevisions(result)));
    }
}
