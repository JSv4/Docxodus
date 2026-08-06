#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A table row aligned against a text box used to hang <see cref="WmlComparer.Compare"/> forever:
/// every branch of the structural walk in <c>DoLcsAlgorithm</c> needed one side to be a word, so two
/// different non-word kinds advanced neither counter.
/// </summary>
public class WmlComparerRowVersusTextboxTests
{
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";

    private static WmlDocument Doc(params XElement[] bodyChildren)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.PutXDocument(new XDocument(
                new XElement(W.document,
                    new XAttribute(XNamespace.Xmlns + "w", W.w),
                    new XAttribute(XNamespace.Xmlns + "v", V),
                    new XElement(W.body, bodyChildren))));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(
                new XDocument(new XElement(W.styles, new XAttribute(XNamespace.Xmlns + "w", W.w))));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(
                new XDocument(new XElement(W.settings, new XAttribute(XNamespace.Xmlns + "w", W.w))));
        }
        return new WmlDocument("test.docx", ms.ToArray());
    }

    private static XElement Para(string text) =>
        new(W.p, new XElement(W.r, new XElement(W.t, text)));

    private static XElement Table(params string[] cells) =>
        new(W.tbl,
            new XElement(W.tblPr,
                new XElement(W.tblW, new XAttribute(W._w, "0"), new XAttribute(W.type, "auto"))),
            new XElement(W.tblGrid,
                cells.Select(_ => new XElement(W.gridCol, new XAttribute(W._w, "3000")))),
            new XElement(W.tr, cells.Select(c => new XElement(W.tc, Para(c)))));

    private static XElement TextboxPara(string text) =>
        new(W.p,
            new XElement(W.r,
                new XElement(W.pict,
                    new XElement(V + "shape",
                        new XAttribute("id", "shape1"),
                        new XAttribute("style", "width:200pt;height:50pt"),
                        new XElement(V + "textbox",
                            new XElement(W.txbxContent, Para(text)))))));

    /// <summary>
    /// A regression here does not fail, it hangs, so the assertion is a deadline. The comparison
    /// takes well under a second once the walk advances.
    /// </summary>
    private static WmlDocument CompareWithin(WmlDocument left, WmlDocument right, int timeoutMs = 10_000)
    {
        WmlDocument? result = null;
        Exception? failure = null;

        var thread = new Thread(() =>
        {
            try
            {
                result = WmlComparer.Compare(left, right,
                    new WmlComparerSettings { DateTimeForRevisions = "2000-01-01T00:00:00Z" });
            }
            catch (Exception ex)
            {
                failure = ex;
            }
        })
        { IsBackground = true };

        thread.Start();
        Assert.True(thread.Join(timeoutMs),
            $"WmlComparer.Compare did not finish within {timeoutMs} ms - the structural walk is not advancing.");

        if (failure != null)
            throw new Xunit.Sdk.XunitException("Compare threw: " + failure);

        return result!;
    }

    private static XElement Body(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        return wDoc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    /// <summary>All the text the document carries, deleted text included.</summary>
    private static string AllText(XElement body) =>
        string.Concat(body.Descendants()
            .Where(d => d.Name == W.t || d.Name == W.delText)
            .Select(d => d.Value));

    /// <summary>RT001 — a table row on the left facing a text box on the right terminates.</summary>
    [Fact]
    public void RT001_RowAgainstTextbox_Terminates()
    {
        var body = Body(CompareWithin(
            Doc(Para("Alpha"), Table("cell"), Para("Omega")),
            Doc(Para("Alpha"), TextboxPara("boxed"), Para("Omega"))));

        // Terminating is not enough - neither side's content may be dropped on the floor.
        var text = AllText(body);
        Assert.Contains("cell", text);
        Assert.Contains("boxed", text);
        Assert.Contains("Alpha", text);
        Assert.Contains("Omega", text);
    }

    /// <summary>RT002 — and the mirror image, text box on the left facing a row on the right.</summary>
    [Fact]
    public void RT002_TextboxAgainstRow_Terminates()
    {
        var body = Body(CompareWithin(
            Doc(Para("Alpha"), TextboxPara("boxed"), Para("Omega")),
            Doc(Para("Alpha"), Table("cell"), Para("Omega"))));

        var text = AllText(body);
        Assert.Contains("cell", text);
        Assert.Contains("boxed", text);
    }

    /// <summary>
    /// RT003 — the mismatch reached with differing prose on both sides, so the walk meets it with
    /// words in play rather than only at the structural boundary.
    /// </summary>
    [Fact]
    public void RT003_RowAgainstTextboxWithDifferingProse_Terminates()
    {
        var body = Body(CompareWithin(
            Doc(Para("Alpha one"), Table("cell", "second"), Para("Omega tail")),
            Doc(Para("Alpha two"), TextboxPara("boxed text"), Para("Omega end"))));

        var text = AllText(body);
        Assert.Contains("cell", text);
        Assert.Contains("boxed text", text);
    }

}
