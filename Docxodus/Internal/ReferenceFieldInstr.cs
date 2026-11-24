using System;
using System.Linq;

namespace Docxodus.Internal;

/// <summary>
/// Recognizes the cross-reference field instructions — <c>REF</c>, <c>PAGEREF</c> and
/// <c>NOTEREF</c> — with the bookmark each one targets and whether its <c>\h</c> switch makes it a
/// jump to that bookmark.
/// </summary>
/// <remarks>
/// The consumer is <see cref="WmlToHtmlConverterSettings.StampReferenceFields"/>, which an editor's
/// render options turn on: the field's result is wrapped in a <c>data-field</c> marker so an editor can keep the cached
/// result whole (typing cannot land inside it) and follow the reference. Like
/// <see cref="PageNumberFieldInstr"/>, deliberately not <c>FieldRetriever.ParseField</c>, whose
/// field-type allowlist leaves out <c>PAGEREF</c> and <c>NOTEREF</c>.
/// </remarks>
internal static class ReferenceFieldInstr
{
    /// <summary>
    /// Parse <paramref name="instrText"/>: the field kind spelled upper-case (<c>REF</c>,
    /// <c>PAGEREF</c>, <c>NOTEREF</c>), its bookmark name, and whether it carries <c>\h</c>; or
    /// <c>null</c> for any other field, or one without a bookmark argument.
    /// </summary>
    internal static (string Kind, string Target, bool Hyperlink)? TryParse(string? instrText)
    {
        if (string.IsNullOrWhiteSpace(instrText)) return null;
        var tokens = instrText.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length < 2) return null;
        var kind = tokens[0].ToUpperInvariant();
        if (kind is not ("REF" or "PAGEREF" or "NOTEREF")) return null;
        var target = tokens[1].Trim('"');
        if (target.Length == 0 || target[0] == '\\') return null;
        bool hyperlink = tokens.Skip(2).Any(t => string.Equals(t, "\\h", StringComparison.OrdinalIgnoreCase));
        return (kind, target, hyperlink);
    }
}
