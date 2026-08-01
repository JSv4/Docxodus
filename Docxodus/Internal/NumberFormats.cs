#nullable enable

namespace Docxodus.Internal;

/// <summary>
/// The single owner of the <see cref="NumberFormat"/> ↔ token mappings.
/// </summary>
/// <remarks>
/// <para>
/// <see cref="NumberFormat"/> is this library's name for the OOXML <c>ST_NumberFormat</c> simple
/// type, which appears in two places the <c>DocxSession</c> surface writes: <c>w:numFmt</c> (list
/// numbering) and <c>w:pgNumType/@w:fmt</c> (section page numbering). Both use the SAME token
/// spellings, so they share one table here rather than each carrying a private copy — the JSON wire
/// format uses those same tokens, so <c>DocxSessionJson</c> and <c>BlockMetadataOps</c> both
/// delegate here.
/// </para>
/// <para>
/// The field <c>\*</c> general-formatting switch (ECMA-376 §17.16.4.1) is a SEPARATE vocabulary
/// with different spellings, and case is significant — <c>roman</c> renders <c>i, ii, iii</c> while
/// <c>ROMAN</c> renders <c>I, II, III</c>. <see cref="ToFieldSwitch"/> owns that translation.
/// </para>
/// <para>
/// The enum is the INTERSECTION of what both vocabularies express. Formats that exist on only one
/// side (<c>\* DollarText</c> has no <c>ST_NumberFormat</c>; <c>chicago</c>/<c>decimalZero</c> have
/// no switch) are deliberately absent: including one would force a mapper to lie.
/// </para>
/// </remarks>
internal static class NumberFormats
{
    /// <summary>The <c>ST_NumberFormat</c> token — valid for <c>w:numFmt/@w:val</c> and
    /// <c>w:pgNumType/@w:fmt</c> alike, and the JSON wire spelling.</summary>
    internal static string ToOoxml(NumberFormat format) => format switch
    {
        NumberFormat.UpperLetter => "upperLetter",
        NumberFormat.LowerLetter => "lowerLetter",
        NumberFormat.UpperRoman => "upperRoman",
        NumberFormat.LowerRoman => "lowerRoman",
        NumberFormat.Bullet => "bullet",
        _ => "decimal",
    };

    /// <summary>Inverse of <see cref="ToOoxml"/>. Lenient by design: an unrecognized or absent token
    /// reads back as <see cref="NumberFormat.Decimal"/>, matching how Word treats a format it does
    /// not implement. Callers that must distinguish "absent" from "decimal" check for the attribute
    /// themselves before calling.</summary>
    internal static NumberFormat ParseOoxml(string? token) => token switch
    {
        "bullet" => NumberFormat.Bullet,
        "upperLetter" => NumberFormat.UpperLetter,
        "lowerLetter" => NumberFormat.LowerLetter,
        "upperRoman" => NumberFormat.UpperRoman,
        "lowerRoman" => NumberFormat.LowerRoman,
        _ => NumberFormat.Decimal,
    };

    /// <summary>
    /// The argument of a field's <c>\*</c> general-formatting switch, or <c>null</c> for a format
    /// that has no switch equivalent (<see cref="NumberFormat.Bullet"/> — see
    /// <see cref="IsPageNumberFormat"/>). Case is load-bearing.
    /// </summary>
    internal static string? ToFieldSwitch(NumberFormat format) => format switch
    {
        NumberFormat.UpperLetter => "ALPHABETIC",
        NumberFormat.LowerLetter => "alphabetic",
        NumberFormat.UpperRoman => "ROMAN",
        NumberFormat.LowerRoman => "roman",
        NumberFormat.Decimal => "Arabic",
        _ => null,
    };

    /// <summary>
    /// Whether <paramref name="format"/> can number PAGES. Everything but
    /// <see cref="NumberFormat.Bullet"/> can: a bullet is a valid list format but neither
    /// <c>w:pgNumType/@w:fmt</c> nor the <c>\*</c> switch has any such notion, so accepting it
    /// would mean silently writing something else.
    /// </summary>
    internal static bool IsPageNumberFormat(NumberFormat format) => format != NumberFormat.Bullet;

    /// <summary>
    /// Render <paramref name="value"/> in <paramref name="format"/> — "1" → <c>i</c>, <c>A</c>,
    /// <c>1</c>. Used to seed a page-number field's cached result so a renderer that does not
    /// recompute fields shows a number consistent with the format, rather than always "1".
    /// Delegates to the list-label formatter, which already owns every ST_NumberFormat algorithm.
    /// </summary>
    internal static string Render(int value, NumberFormat format) =>
        ListItemTextGetter_Default.GetListItemText("en-US", value, ToOoxml(format));
}
