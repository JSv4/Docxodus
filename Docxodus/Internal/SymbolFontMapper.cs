using System.Collections.Generic;
using System.Text;

namespace Docxodus.Internal;

/// <summary>
/// Maps symbol-font glyphs (Symbol / Wingdings, which Word stores in the U+F000 private-use range)
/// to their Unicode equivalents. List bullets in particular render as a blank box in a browser that
/// lacks the proprietary font; emitting the real Unicode glyph (e.g. U+F0B7 → U+2022 "•") renders
/// everywhere. Covers the common list-bullet glyphs Word and <see cref="NumberingFactory"/> use;
/// unmapped characters are returned unchanged so nothing is lost.
/// </summary>
internal static class SymbolFontMapper
{
    // Keyed by (lowercased primary font family, char code). Both the private-use form (0xF0xx, how
    // Word stores them) and the bare 0x00xx form are included.
    private static readonly Dictionary<(string Font, int Code), string> Map = new()
    {
        // Symbol
        { ("symbol", 0xF0B7), "•" }, { ("symbol", 0x00B7), "•" }, // • bullet
        // Wingdings
        { ("wingdings", 0xF0A7), "▪" }, { ("wingdings", 0x00A7), "▪" }, // ▪ small black square
        { ("wingdings", 0xF06E), "■" }, // ■ black square
        { ("wingdings", 0xF075), "◆" }, // ◆ black diamond
        { ("wingdings", 0xF0FC), "✓" }, // ✓ check mark
        { ("wingdings", 0xF0FE), "☑" }, // ☑ checked box
        { ("wingdings", 0xF0D8), "➢" }, // ➢ three-d arrowhead
    };

    // Each symbol font's ascent in ems: its hhea ascender, equal to its OS/2 winAscent, over 2048 units
    // per em (Windows 11 copies; none has a line gap). Word lifts a list marker's line to its font's
    // ascent and ignores its descent, so a mapped glyph keeps the ascent of the font it replaces
    // (#1567): a 12pt Symbol bullet over 12pt Arial makes a 14.61pt line, not Arial's 13.80pt.
    private static readonly Dictionary<string, decimal> Ascents = new()
    {
        { "symbol", 2059m / 2048m }, { "wingdings", 1841m / 2048m }, { "wingdings 2", 1727m / 2048m },
        { "wingdings 3", 1900m / 2048m }, { "webdings", 1638m / 2048m },
    };

    /// <summary>The lowercased primary family of a CSS font-family value (drops quotes + fallbacks).</summary>
    private static string PrimaryFamily(string fontFamily) =>
        fontFamily.Split(',')[0].Trim().Trim('\'', '"').ToLowerInvariant();

    /// <summary>
    /// The families after the primary one in a CSS font-family value, or null when there are none:
    /// what a mapped glyph renders in once the symbol face is dropped.
    /// </summary>
    public static string? FallbackFamilies(string fontFamily)
    {
        var comma = fontFamily.IndexOf(',');
        var rest = comma < 0 ? "" : fontFamily[(comma + 1)..].Trim();
        return rest.Length == 0 ? null : rest;
    }

    /// <summary>True if <paramref name="fontFamily"/>'s primary family is a known symbol font.</summary>
    public static bool IsSymbolFont(string? fontFamily) =>
        fontFamily != null && Ascents.ContainsKey(PrimaryFamily(fontFamily));

    /// <summary>The ascent, in ems, of the symbol font <paramref name="fontFamily"/> names first.</summary>
    public static decimal Ascent(string fontFamily) => Ascents[PrimaryFamily(fontFamily)];

    /// <summary>
    /// Map each character of <paramref name="text"/> from <paramref name="fontFamily"/> to Unicode.
    /// Unmapped characters are left unchanged.
    /// </summary>
    public static string MapText(string text, string fontFamily)
    {
        var font = PrimaryFamily(fontFamily);
        var sb = new StringBuilder(text.Length);
        foreach (var ch in text)
            sb.Append(Map.TryGetValue((font, ch), out var mapped) ? mapped : ch.ToString());
        return sb.ToString();
    }
}
