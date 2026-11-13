using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

// Issue #1707: text in a symbol font maps to Unicode wherever Unicode has a faithful equivalent.
public class SymbolFontMapperTests
{
    [Theory]
    // The form-control glyphs, each in its private-use and bare form.
    [InlineData("Wingdings", '\uF0A8', "☐")]
    [InlineData("Wingdings", '\u00A8', "☐")]
    [InlineData("Wingdings", '\uF06F', "☐")]
    [InlineData("Wingdings", '\uF0FD', "☒")]
    [InlineData("Wingdings", '\uF0FE', "☑")]
    [InlineData("Wingdings", '\uF0FC', "✓")]
    [InlineData("Wingdings", '\uF0FB', "✗")]
    [InlineData("Wingdings 2", '\uF050', "✓")]
    [InlineData("Wingdings 2", '\uF052', "☑")]
    [InlineData("Wingdings 2", '\uF054', "☒")]
    // The list bullets the mapper handled before it had full tables keep their glyphs.
    [InlineData("Symbol", '\uF0B7', "•")]
    [InlineData("Wingdings", '\uF0A7', "▪")]
    [InlineData("Wingdings", '\uF06E', "■")]
    [InlineData("Wingdings", '\uF0D8', "➢")]
    // Adobe Symbol: Greek, mathematical operators, and the serif (R) the table leaves in private use.
    [InlineData("Symbol", 'a', "α")]
    [InlineData("Symbol", '\uF057', "Ω")]
    [InlineData("Symbol", '\uF0A5', "∞")]
    [InlineData("Symbol", '\uF0B9', "≠")]
    [InlineData("Symbol", '\uF0D2', "®")]
    [InlineData("Symbol", '1', "1")]
    // Unicode 7.0's Wingdings, Wingdings 3 and Webdings additions, beyond the BMP.
    [InlineData("Wingdings", '\uF04A', "☺")]
    [InlineData("Wingdings", '\uF021', "\U0001F589")]
    [InlineData("Wingdings 3", '\uF0B6', "\U0001F83A")]
    [InlineData("Webdings", '\uF061', "✔")]
    // A code point browsers draw as a colour emoji asks for its text form.
    [InlineData("Wingdings", '\uF043', "\U0001F44D\uFE0E")]
    // Wingdings 0x36 is the plain hourglass, not the one with flowing sand (U+23F3).
    [InlineData("Wingdings", '\uF036', "\u231B\uFE0E")]
    // A bare no-break space or soft hyphen in a symbol font is the font's glyph at that byte, as Word draws it.
    [InlineData("Wingdings", '\u00A0', "\u2B1D")]
    [InlineData("Wingdings", '\u00AD', "\U0001F7CF")]
    [InlineData("Symbol", '\u00AD', "↑")]
    // A CSS stack and any case name the same font.
    [InlineData("'WINGDINGS', serif", '\uF0FE', "☑")]
    public void MapsFaithfulGlyphs(string font, char stored, string expected) =>
        Assert.Equal(expected, SymbolFontMapper.Map(stored, font));

    [Theory]
    // No faithful equivalent: the Windows logo, Symbol's bracket pieces, empty slots.
    [InlineData("Wingdings", '\uF0FF')]
    [InlineData("Symbol", '\uF0E6')]
    [InlineData("Symbol", '\uF060')]
    // Adobe's table has the euro at Symbol 0xA0; the Symbol font Word draws with has no glyph there.
    [InlineData("Symbol", '\uF0A0')]
    [InlineData("Symbol", '\u00A0')]
    [InlineData("Wingdings 3", '\uF0F1')]
    [InlineData("Wingdings", '\uF07F')]
    // Outside the font's byte range, or in a font that is not a symbol font.
    [InlineData("Wingdings", '\uF100')]
    [InlineData("Wingdings", 'Ā')]
    [InlineData("Wingdings", '\u001F')]
    [InlineData("Calibri", '\uF0A8')]
    public void LeavesUnmappableCharactersAlone(string font, char stored)
    {
        Assert.Null(SymbolFontMapper.Map(stored, font));
        Assert.Equal(stored.ToString(), SymbolFontMapper.MapText(stored.ToString(), font));
    }

    [Fact]
    public void MapTextKeepsUnmappedCharactersInPlace() =>
        Assert.Equal("☐\uF0FF☑", SymbolFontMapper.MapText("\uF0A8\uF0FF\uF0FE", "Wingdings"));
}
