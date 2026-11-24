using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// The paragraph number a <c>REF</c> field shows for its target, by its number switch. The rules are
/// read off the results Word cached in real agreements (<c>DocxSessionCrossReferenceTests</c> compares
/// with them): a trailing period is dropped; a number that already shows its upper levels ("2.3.1") is
/// complete; one that does not ("(i)") takes the numbers above it, up to the first that does.
/// </summary>
internal static class ReferenceNumber
{
    internal const string NoContext = "\\n", RelativeContext = "\\r", FullContext = "\\w";

    /// <summary>The number Word shows for a paragraph that has none.</summary>
    private const string Unnumbered = "0";

    /// <summary>
    /// The number of <paramref name="target"/> as <paramref name="numberSwitch"/> shows it in
    /// <paramref name="fieldParagraph"/>, or null when Word's own text for it is not known: a format in
    /// words or with literal text in a joined number, or a relative reference from outside the list.
    /// <paramref name="numeralsOnly"/> is the field's <c>\t</c> switch.
    /// </summary>
    internal static string? Resolve(WordprocessingDocument doc, XElement target, XElement fieldParagraph,
        string numberSwitch, bool numeralsOnly = false)
    {
        if (ListNumberResolver.Resolve(target, doc) is null) return Unnumbered;
        var settings = new ListItemRetrieverSettings();
        if (ListItemRetriever.RetrieveListItemLevels(doc, target, numeralsOnly, settings) is not { } number) return Unnumbered;
        int own = number.Levels.Count - 1;
        if (numeralsOnly && !number.Levels[own].IsPlain) return null;
        if (number.Levels[own].Text is not { Length: > 0 } ownNumber) return Unnumbered;
        string ownText = WithoutPeriod(ownNumber);
        if (numberSwitch == NoContext || own == 0 || number.Levels[own].ShowsOtherLevels) return ownText;

        int top = own;
        while (top > 0 && !number.Levels[top].ShowsOtherLevels) top--;
        if (numberSwitch == RelativeContext)
        {
            if (ListItemRetriever.RetrieveListItemLevels(doc, fieldParagraph, numeralsOnly, settings) is not { } here
                || here.AbstractNumId != number.AbstractNumId)
                return null;
            int shared = here.Counters.Zip(number.Counters).TakeWhile(pair => pair.First == pair.Second).Count();
            top = System.Math.Min(System.Math.Max(top, shared), own);
        }

        string text = ownText;
        for (int level = own - 1; level >= top; level--)
        {
            var upper = number.Levels[level];
            if (upper.Text is null || !upper.IsPlain || !number.Levels[level + 1].IsPlain || !JoinablePattern.IsMatch(upper.Text)) return null;
            text = (char.IsLetterOrDigit(text[0]) ? upper.Text : WithoutPeriod(upper.Text)) + text;
        }
        return text;
    }

    /// <summary>Upper-level numbers whose joined form is known: "1", "1.", "2.3.1" and "(a)".</summary>
    private static readonly Regex JoinablePattern = new(@"^(\([A-Za-z0-9]+\)|[A-Za-z0-9]+(\.[A-Za-z0-9]+)*\.?)$", RegexOptions.CultureInvariant);

    private static string WithoutPeriod(string number) => number.Length > 1 ? number.TrimEnd('.') : number;
}
