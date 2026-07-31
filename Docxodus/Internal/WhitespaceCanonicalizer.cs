#nullable enable

using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Canonicalizes a document's whitespace so whitespace-only differences cannot register as revisions —
/// the shared implementation behind <c>WmlComparerSettings.CompareWhitespace</c> and
/// <see cref="DocxDiffSettings.CompareWhitespace"/> (Word Compare's "White space" option).
/// </summary>
/// <remarks>
/// Both engines canonicalize their INPUTS rather than comparing whitespace-blind, and for the same
/// reason in each: their diff cores pair the two sides position-for-position.
/// <c>WmlComparer.FlattenToComparisonUnitAtomList</c> zips an Equal sequence's two atom streams, and
/// <c>IrTokenDiffer</c>'s edit stream is 1:1 per token — so in both, one side carrying more whitespace
/// than the other is not representable as "equal". Applying the same transform to both sides keeps the
/// streams the same length wherever the canonical text matches. The cost is that the produced document
/// carries canonical whitespace rather than either input's verbatim spacing.
/// </remarks>
internal static class WhitespaceCanonicalizer
{
    private static readonly Regex WhitespaceRun = new Regex(@"\s+", RegexOptions.Compiled);

    /// <summary>
    /// Returns <paramref name="doc"/> with each paragraph's whitespace canonicalized in the main
    /// document part and the footnotes/endnotes parts, or the same instance when nothing changed.
    /// </summary>
    public static WmlDocument Canonicalize(WmlDocument doc)
    {
        using var streamDoc = new OpenXmlMemoryStreamDocument(doc);
        var anyChanged = false;
        using (var wDoc = streamDoc.GetWordprocessingDocument())
        {
            var mainPart = wDoc.MainDocumentPart;
            var parts = new List<OpenXmlPart> { mainPart };
            if (mainPart.FootnotesPart != null)
                parts.Add(mainPart.FootnotesPart);
            if (mainPart.EndnotesPart != null)
                parts.Add(mainPart.EndnotesPart);

            foreach (var part in parts)
            {
                var xDoc = part.GetXDocument();
                var changed = false;
                foreach (var para in xDoc.Descendants(W.p))
                    changed |= CanonicalizeParagraph(para);
                if (changed)
                    part.PutXDocument();
                anyChanged |= changed;
            }
        }
        return anyChanged ? streamDoc.GetModifiedWmlDocument() : doc;
    }

    private static bool CanonicalizeParagraph(XElement para)
    {
        // Only content that survives an AcceptRevisions pass may join a whitespace run — w:delText, or a
        // w:tab/w:br under w:del/w:moveFrom, would swallow a space the accepted text still needs.
        var inlines = para
            .Descendants()
            .Where(d => d.Parent != null && d.Parent.Name == W.r)
            .Where(d => d.Name == W.t || d.Name == W.tab || d.Name == W.br)
            .Where(d => d.Ancestors(W.p).First() == para)
            .Where(SurvivesAccept)
            .ToList();

        var changed = false;
        var previousWasSpace = true;    // also trims the paragraph's leading whitespace
        XElement? lastText = null;
        foreach (var inline in inlines)
        {
            if (inline.Name != W.t)
            {
                if (lastText != null)
                    changed |= SetTextValue(lastText, lastText.Value.TrimEnd(' '));
                previousWasSpace = true;
                continue;
            }

            var collapsed = WhitespaceRun.Replace(inline.Value, " ");
            if (previousWasSpace && collapsed.Length != 0 && collapsed[0] == ' ')
                collapsed = collapsed.Substring(1);
            if (collapsed.Length != 0)
                previousWasSpace = collapsed[collapsed.Length - 1] == ' ';
            changed |= SetTextValue(inline, collapsed);
            if (collapsed.Length != 0)
                lastText = inline;
        }

        if (lastText != null)
            changed |= SetTextValue(lastText, lastText.Value.TrimEnd(' '));

        return changed;
    }

    private static bool SurvivesAccept(XElement element) =>
        !element.Ancestors().Any(a => a.Name == W.del || a.Name == W.moveFrom);

    private static bool SetTextValue(XElement text, string value)
    {
        if (text.Value == value)
            return false;
        text.SetValue(value);
        var needsPreserve = value.Length != 0 && (value[0] == ' ' || value[value.Length - 1] == ' ');
        text.SetAttributeValue(XNamespace.Xml + "space", needsPreserve ? "preserve" : null);
        return true;
    }
}
