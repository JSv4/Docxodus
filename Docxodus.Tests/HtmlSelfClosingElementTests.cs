#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// The converter's output is an XML tree that every consumer serializes with
/// <c>XElement.ToString</c>, and XML self-closes an element with no content. HTML does not read it
/// that way: outside foreign content a parser ignores the trailing slash, so <c>&lt;span /&gt;</c>
/// opens a span that stays open and adopts every following sibling until some later closing tag is
/// spent on it. The browser then builds a different tree from the one we emitted.
/// </summary>
/// <remarks>
/// Found through a layout symptom with no visible cause (issue #688): a footnote beginning with a
/// reference mark and a tab emits an empty wrapper span for the mark, and in the browser the note's
/// whole text ended up inside the tab's fixed-width box, two characters to a line. The invariant is
/// asserted over the whole rendered charter rather than over that one span, because the next empty
/// element the converter learns to emit will not be this one. Test IDs use the SC0xx range.
/// </remarks>
public class HtmlSelfClosingElementTests
{
    /// <summary>ECMA/WHATWG void elements — the only tags allowed to stand alone in HTML.</summary>
    private const string VoidElements =
        "area|base|br|col|embed|hr|img|input|link|meta|param|source|track|wbr";

    private static string PaginatedHtml(string fixtureName)
    {
        var bytes = File.ReadAllBytes(Path.Combine("../../../../TestFiles/", fixtureName));
        using var ms = new MemoryStream();
        ms.Write(bytes);
        using var doc = WordprocessingDocument.Open(ms, true);
        return WmlToHtmlConverter.ConvertToHtml(doc, new WmlToHtmlConverterSettings
        {
            RenderPagination = PaginationMode.Paginated,
            RenderHeadersAndFooters = true,
            RenderFootnotesAndEndnotes = true,
        }).ToString();
    }

    /// <summary>
    /// A real charter exercises the shapes that produce empty elements — note reference marks,
    /// bookmark anchors, tab wrappers, empty runs — across body, running and note stories.
    /// </summary>
    [Theory]
    [InlineData("NVCA-Model-COI.docx")]
    [InlineData("HC031-Complicated-Document.docx")]
    public void SC001_NoNonVoidElementSerializesSelfClosing(string fixtureName)
    {
        var offenders = Regex
            .Matches(PaginatedHtml(fixtureName), @"<(?<tag>[A-Za-z][\w:-]*)\b[^>]*?/>")
            .Where(match => !Regex.IsMatch(
                match.Groups["tag"].Value, $"^(?:{VoidElements})$", RegexOptions.IgnoreCase))
            .Select(match => match.Value.Length > 120
                ? match.Value.Substring(0, 120) + "…"
                : match.Value)
            .Distinct(StringComparer.Ordinal)
            .Take(10)
            .ToList();

        Assert.True(
            offenders.Count == 0,
            "self-closing non-void elements reparent their siblings in an HTML parser:\n"
            + string.Join("\n", offenders));
    }
}
