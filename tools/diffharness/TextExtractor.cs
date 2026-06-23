#nullable enable
using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace DiffHarness;

/// <summary>
/// Content-focused text extraction for round-trip verification, robust to the part-name churn and
/// header/footer part duplication that the diff markup renderer introduces. The body and the note
/// stores are compared exactly; headers/footers are compared as a DEDUPLICATED set of per-part texts
/// (the renderer may emit duplicate header/footer parts with identical content, which must not register
/// as a content difference) while the raw part count is reported separately as a bloat metric.
/// </summary>
internal static class TextExtractor
{
    public static DocText Extract(byte[] docxBytes)
    {
        using var ms = new MemoryStream();
        ms.Write(docxBytes, 0, docxBytes.Length);
        ms.Position = 0;
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart
            ?? throw new InvalidOperationException("document has no MainDocumentPart");

        var body = PartText(main.Document?.Body);
        var headerFooter = main.HeaderParts.Select(h => PartText(h.Header))
            .Concat(main.FooterParts.Select(f => PartText(f.Footer)))
            .ToList();
        // Note stores are compared as a sorted MULTISET of per-note texts: Compare renumbers footnotes/
        // endnotes (the WmlComparer oracle does too), so part-order/id-keyed comparison gives false
        // positives — content preservation (no note text gained or lost) is the real invariant.
        var footnotes = NoteTexts<DocumentFormat.OpenXml.Wordprocessing.Footnote>(main.FootnotesPart?.Footnotes);
        var endnotes = NoteTexts<DocumentFormat.OpenXml.Wordprocessing.Endnote>(main.EndnotesPart?.Endnotes);
        return new DocText(body, headerFooter, footnotes, endnotes);
    }

    /// <summary>The sorted list of each note's concatenated text (empty separator notes drop out).</summary>
    private static List<string> NoteTexts<T>(DocumentFormat.OpenXml.OpenXmlElement? root)
        where T : DocumentFormat.OpenXml.OpenXmlElement
    {
        if (root is null) return new List<string>();
        return root.Descendants<T>()
            .Select(PartText)
            .Where(t => t.Length > 0)
            .OrderBy(t => t, StringComparer.Ordinal)
            .ToList();
    }

    private static string PartText(DocumentFormat.OpenXml.OpenXmlElement? root)
    {
        if (root is null) return string.Empty;
        var sb = new StringBuilder();
        foreach (var t in root.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
            sb.Append(t.Text);
        return sb.ToString();
    }
}

/// <summary>Partitioned document text. Body compared exactly; notes &amp; header/footer as multisets
/// (Compare renumbers notes and may emit duplicate header/footer parts — both must be order-robust).</summary>
internal sealed record DocText(
    string Body, List<string> HeaderFooterParts, List<string> Footnotes, List<string> Endnotes)
{
    public bool BodyEquals(DocText other) => Body == other.Body;

    public bool NotesEqual(DocText other) =>
        Footnotes.SequenceEqual(other.Footnotes) && Endnotes.SequenceEqual(other.Endnotes);

    /// <summary>Header/footer content equality ignoring duplicate parts (dedup by text).</summary>
    public bool HeaderFooterSetEquals(DocText other) =>
        new HashSet<string>(HeaderFooterParts).SetEquals(other.HeaderFooterParts);

    public int HeaderFooterPartCount => HeaderFooterParts.Count;
}
