#nullable enable

using System;

namespace Docxodus.Internal;

/// <summary>
/// Recognizes the two page-number field instructions in a complex field's <c>instrText</c> —
/// <c>PAGE</c> and <c>NUMPAGES</c> — together with the format argument of a <c>\*</c>
/// general-formatting switch, if the field carries one.
/// </summary>
/// <remarks>
/// <para>
/// Deliberately narrow, and deliberately NOT <c>FieldRetriever.ParseField</c>: that parser has a
/// field-type allowlist that excludes <c>NUMPAGES</c>, and it splits switches from their arguments
/// (<c>\*</c> lands in <c>Switches</c> while <c>roman</c> lands in <c>Arguments</c>), so pairing
/// them back up would be guesswork. Page numbering needs exactly the pair below, so it owns the
/// three lines that read it rather than widening a parser five other call sites depend on.
/// </para>
/// <para>
/// The consumer is the paginated renderer: a header/footer is cloned onto every page, so the
/// field's cached result would otherwise show the same number on all of them. Marking the result in
/// the HTML lets the paginator substitute each page's real number. <c>npm/src/page-number-format.ts</c>
/// is the browser-side counterpart that renders the value.
/// </para>
/// </remarks>
internal static class PageNumberFieldInstr
{
    /// <summary>The recognized page-number field kinds, spelled as they appear in the HTML marker.</summary>
    internal const string Page = "PAGE";

    /// <summary>See <see cref="Page"/>.</summary>
    internal const string NumPages = "NUMPAGES";

    /// <summary>
    /// Parse <paramref name="instrText"/>. Returns the field kind (<see cref="Page"/> /
    /// <see cref="NumPages"/>) and the <c>\*</c> switch argument verbatim — case matters, since
    /// <c>roman</c> and <c>ROMAN</c> are different formats — or <c>null</c> when the instruction is
    /// some other field.
    /// </summary>
    internal static (string Kind, string? FormatSwitch)? TryParse(string? instrText)
    {
        if (string.IsNullOrWhiteSpace(instrText)) return null;
        var tokens = instrText.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length == 0) return null;

        var kind = tokens[0].ToUpperInvariant() switch
        {
            Page => Page,
            NumPages => NumPages,
            _ => null,
        };
        if (kind is null) return null;

        string? formatSwitch = null;
        for (int i = 1; i < tokens.Length - 1; i++)
        {
            if (tokens[i] == "\\*")
            {
                formatSwitch = tokens[i + 1];
                break;
            }
        }

        // \* MERGEFORMAT is a formatting-preservation switch, not a number format — Word writes it
        // alongside real formats and it must not be mistaken for one.
        if (string.Equals(formatSwitch, "MERGEFORMAT", StringComparison.OrdinalIgnoreCase))
            formatSwitch = null;

        return (kind, formatSwitch);
    }
}
