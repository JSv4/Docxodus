#nullable enable

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

    private static readonly HashSet<string> SymbolFonts = new()
    {
        "symbol", "wingdings", "wingdings 2", "wingdings 3", "webdings",
    };

    /// <summary>The lowercased primary family of a CSS font-family value (drops quotes + fallbacks).</summary>
    private static string PrimaryFamily(string fontFamily) =>
        fontFamily.Split(',')[0].Trim().Trim('\'', '"').ToLowerInvariant();

    /// <summary>True if <paramref name="fontFamily"/>'s primary family is a known symbol font.</summary>
    public static bool IsSymbolFont(string? fontFamily) =>
        fontFamily != null && SymbolFonts.Contains(PrimaryFamily(fontFamily));

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
