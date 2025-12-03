#nullable enable
// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;

namespace Docxodus
{
    /// <summary>
    /// Platform-independent color struct replacing SKColor.
    /// Stores ARGB color values without SkiaSharp dependency.
    /// </summary>
    public readonly struct DocxColor : IEquatable<DocxColor>
    {
        public byte Alpha { get; }
        public byte Red { get; }
        public byte Green { get; }
        public byte Blue { get; }

        public DocxColor(byte red, byte green, byte blue, byte alpha = 255)
        {
            Red = red;
            Green = green;
            Blue = blue;
            Alpha = alpha;
        }

        public static DocxColor FromArgb(int alpha, int red, int green, int blue)
            => new((byte)red, (byte)green, (byte)blue, (byte)alpha);

        public static DocxColor FromArgb(int red, int green, int blue)
            => new((byte)red, (byte)green, (byte)blue);

        public static DocxColor FromArgb(byte red, byte green, byte blue)
            => new(red, green, blue);

        public static DocxColor FromArgb(byte alpha, byte red, byte green, byte blue)
            => new(red, green, blue, alpha);

        public static DocxColor FromArgb(int argb)
            => new((byte)(argb >> 16), (byte)(argb >> 8), (byte)argb, (byte)(argb >> 24));

        public int ToArgb()
            => (Alpha << 24) | (Red << 16) | (Green << 8) | Blue;

        public string ToHex() => $"#{Red:X2}{Green:X2}{Blue:X2}";
        public string ToHexWithAlpha() => $"#{Alpha:X2}{Red:X2}{Green:X2}{Blue:X2}";

        // Static colors that don't conflict with property names
        public static DocxColor Empty => new(0, 0, 0, 0);
        public static DocxColor Transparent => new(0, 0, 0, 0);
        public static DocxColor Black => new(0, 0, 0);
        public static DocxColor White => new(255, 255, 255);

        public bool Equals(DocxColor other)
            => Alpha == other.Alpha && Red == other.Red && Green == other.Green && Blue == other.Blue;

        public override bool Equals(object? obj) => obj is DocxColor c && Equals(c);
        public override int GetHashCode() => ToArgb();
        public static bool operator ==(DocxColor left, DocxColor right) => left.Equals(right);
        public static bool operator !=(DocxColor left, DocxColor right) => !left.Equals(right);

        public override string ToString() => $"DocxColor [A={Alpha}, R={Red}, G={Green}, B={Blue}]";
    }

    /// <summary>
    /// Color name lookup and static color constants.
    /// Use DocxColors.Red, DocxColors.Blue etc. for named colors.
    /// </summary>
    public static class DocxColors
    {
        // Basic colors
        public static readonly DocxColor Black = new(0, 0, 0);
        public static readonly DocxColor White = new(255, 255, 255);
        public static readonly DocxColor Red = new(255, 0, 0);
        public static readonly DocxColor Green = new(0, 128, 0);
        public static readonly DocxColor Blue = new(0, 0, 255);
        public static readonly DocxColor Yellow = new(255, 255, 0);
        public static readonly DocxColor Cyan = new(0, 255, 255);
        public static readonly DocxColor Magenta = new(255, 0, 255);

        // Grays
        public static readonly DocxColor Gray = new(128, 128, 128);
        public static readonly DocxColor DarkGray = new(169, 169, 169);
        public static readonly DocxColor LightGray = new(211, 211, 211);
        public static readonly DocxColor Silver = new(192, 192, 192);
        public static readonly DocxColor DimGray = new(105, 105, 105);
        public static readonly DocxColor Gainsboro = new(220, 220, 220);

        // Standard HTML/CSS colors
        public static readonly DocxColor Aqua = new(0, 255, 255);
        public static readonly DocxColor Fuchsia = new(255, 0, 255);
        public static readonly DocxColor Lime = new(0, 255, 0);
        public static readonly DocxColor Maroon = new(128, 0, 0);
        public static readonly DocxColor Navy = new(0, 0, 128);
        public static readonly DocxColor Olive = new(128, 128, 0);
        public static readonly DocxColor Purple = new(128, 0, 128);
        public static readonly DocxColor Teal = new(0, 128, 128);

        // Extended colors
        public static readonly DocxColor AliceBlue = new(240, 248, 255);
        public static readonly DocxColor AntiqueWhite = new(250, 235, 215);
        public static readonly DocxColor Aquamarine = new(127, 255, 212);
        public static readonly DocxColor Azure = new(240, 255, 255);
        public static readonly DocxColor Beige = new(245, 245, 220);
        public static readonly DocxColor Bisque = new(255, 228, 196);
        public static readonly DocxColor BlanchedAlmond = new(255, 235, 205);
        public static readonly DocxColor BlueViolet = new(138, 43, 226);
        public static readonly DocxColor Brown = new(165, 42, 42);
        public static readonly DocxColor BurlyWood = new(222, 184, 135);
        public static readonly DocxColor CadetBlue = new(95, 158, 160);
        public static readonly DocxColor Chartreuse = new(127, 255, 0);
        public static readonly DocxColor Chocolate = new(210, 105, 30);
        public static readonly DocxColor Coral = new(255, 127, 80);
        public static readonly DocxColor CornflowerBlue = new(100, 149, 237);
        public static readonly DocxColor Cornsilk = new(255, 248, 220);
        public static readonly DocxColor Crimson = new(220, 20, 60);
        public static readonly DocxColor DarkBlue = new(0, 0, 139);
        public static readonly DocxColor DarkCyan = new(0, 139, 139);
        public static readonly DocxColor DarkGoldenrod = new(184, 134, 11);
        public static readonly DocxColor DarkGreen = new(0, 100, 0);
        public static readonly DocxColor DarkKhaki = new(189, 183, 107);
        public static readonly DocxColor DarkMagenta = new(139, 0, 139);
        public static readonly DocxColor DarkOliveGreen = new(85, 107, 47);
        public static readonly DocxColor DarkOrange = new(255, 140, 0);
        public static readonly DocxColor DarkOrchid = new(153, 50, 204);
        public static readonly DocxColor DarkRed = new(139, 0, 0);
        public static readonly DocxColor DarkSalmon = new(233, 150, 122);
        public static readonly DocxColor DarkSeaGreen = new(143, 188, 139);
        public static readonly DocxColor DarkSlateBlue = new(72, 61, 139);
        public static readonly DocxColor DarkSlateGray = new(47, 79, 79);
        public static readonly DocxColor DarkTurquoise = new(0, 206, 209);
        public static readonly DocxColor DarkViolet = new(148, 0, 211);
        public static readonly DocxColor DeepPink = new(255, 20, 147);
        public static readonly DocxColor DeepSkyBlue = new(0, 191, 255);
        public static readonly DocxColor DodgerBlue = new(30, 144, 255);
        public static readonly DocxColor Firebrick = new(178, 34, 34);
        public static readonly DocxColor FloralWhite = new(255, 250, 240);
        public static readonly DocxColor ForestGreen = new(34, 139, 34);
        public static readonly DocxColor GhostWhite = new(248, 248, 255);
        public static readonly DocxColor Gold = new(255, 215, 0);
        public static readonly DocxColor Goldenrod = new(218, 165, 32);
        public static readonly DocxColor GreenYellow = new(173, 255, 47);
        public static readonly DocxColor Honeydew = new(240, 255, 240);
        public static readonly DocxColor HotPink = new(255, 105, 180);
        public static readonly DocxColor IndianRed = new(205, 92, 92);
        public static readonly DocxColor Indigo = new(75, 0, 130);
        public static readonly DocxColor Ivory = new(255, 255, 240);
        public static readonly DocxColor Khaki = new(240, 230, 140);
        public static readonly DocxColor Lavender = new(230, 230, 250);
        public static readonly DocxColor LavenderBlush = new(255, 240, 245);
        public static readonly DocxColor LawnGreen = new(124, 252, 0);
        public static readonly DocxColor LemonChiffon = new(255, 250, 205);
        public static readonly DocxColor LightBlue = new(173, 216, 230);
        public static readonly DocxColor LightCoral = new(240, 128, 128);
        public static readonly DocxColor LightCyan = new(224, 255, 255);
        public static readonly DocxColor LightGoldenrodYellow = new(250, 250, 210);
        public static readonly DocxColor LightGreen = new(144, 238, 144);
        public static readonly DocxColor LightPink = new(255, 182, 193);
        public static readonly DocxColor LightSalmon = new(255, 160, 122);
        public static readonly DocxColor LightSeaGreen = new(32, 178, 170);
        public static readonly DocxColor LightSkyBlue = new(135, 206, 250);
        public static readonly DocxColor LightSlateGray = new(119, 136, 153);
        public static readonly DocxColor LightSteelBlue = new(176, 196, 222);
        public static readonly DocxColor LightYellow = new(255, 255, 224);
        public static readonly DocxColor LimeGreen = new(50, 205, 50);
        public static readonly DocxColor Linen = new(250, 240, 230);
        public static readonly DocxColor MediumAquamarine = new(102, 205, 170);
        public static readonly DocxColor MediumBlue = new(0, 0, 205);
        public static readonly DocxColor MediumOrchid = new(186, 85, 211);
        public static readonly DocxColor MediumPurple = new(147, 112, 219);
        public static readonly DocxColor MediumSeaGreen = new(60, 179, 113);
        public static readonly DocxColor MediumSlateBlue = new(123, 104, 238);
        public static readonly DocxColor MediumSpringGreen = new(0, 250, 154);
        public static readonly DocxColor MediumTurquoise = new(72, 209, 204);
        public static readonly DocxColor MediumVioletRed = new(199, 21, 133);
        public static readonly DocxColor MidnightBlue = new(25, 25, 112);
        public static readonly DocxColor MintCream = new(245, 255, 250);
        public static readonly DocxColor MistyRose = new(255, 228, 225);
        public static readonly DocxColor Moccasin = new(255, 228, 181);
        public static readonly DocxColor NavajoWhite = new(255, 222, 173);
        public static readonly DocxColor OldLace = new(253, 245, 230);
        public static readonly DocxColor OliveDrab = new(107, 142, 35);
        public static readonly DocxColor Orange = new(255, 165, 0);
        public static readonly DocxColor OrangeRed = new(255, 69, 0);
        public static readonly DocxColor Orchid = new(218, 112, 214);
        public static readonly DocxColor PaleGoldenrod = new(238, 232, 170);
        public static readonly DocxColor PaleGreen = new(152, 251, 152);
        public static readonly DocxColor PaleTurquoise = new(175, 238, 238);
        public static readonly DocxColor PaleVioletRed = new(219, 112, 147);
        public static readonly DocxColor PapayaWhip = new(255, 239, 213);
        public static readonly DocxColor PeachPuff = new(255, 218, 185);
        public static readonly DocxColor Peru = new(205, 133, 63);
        public static readonly DocxColor Pink = new(255, 192, 203);
        public static readonly DocxColor Plum = new(221, 160, 221);
        public static readonly DocxColor PowderBlue = new(176, 224, 230);
        public static readonly DocxColor RosyBrown = new(188, 143, 143);
        public static readonly DocxColor RoyalBlue = new(65, 105, 225);
        public static readonly DocxColor SaddleBrown = new(139, 69, 19);
        public static readonly DocxColor Salmon = new(250, 128, 114);
        public static readonly DocxColor SandyBrown = new(244, 164, 96);
        public static readonly DocxColor SeaGreen = new(46, 139, 87);
        public static readonly DocxColor SeaShell = new(255, 245, 238);
        public static readonly DocxColor Sienna = new(160, 82, 45);
        public static readonly DocxColor SkyBlue = new(135, 206, 235);
        public static readonly DocxColor SlateBlue = new(106, 90, 205);
        public static readonly DocxColor SlateGray = new(112, 128, 144);
        public static readonly DocxColor Snow = new(255, 250, 250);
        public static readonly DocxColor SpringGreen = new(0, 255, 127);
        public static readonly DocxColor SteelBlue = new(70, 130, 180);
        public static readonly DocxColor Tan = new(210, 180, 140);
        public static readonly DocxColor Thistle = new(216, 191, 216);
        public static readonly DocxColor Tomato = new(255, 99, 71);
        public static readonly DocxColor Transparent = new(0, 0, 0, 0);
        public static readonly DocxColor Turquoise = new(64, 224, 208);
        public static readonly DocxColor Violet = new(238, 130, 238);
        public static readonly DocxColor Wheat = new(245, 222, 179);
        public static readonly DocxColor WhiteSmoke = new(245, 245, 245);
        public static readonly DocxColor YellowGreen = new(154, 205, 50);

        private static readonly Dictionary<string, DocxColor> NamedColors = new(StringComparer.OrdinalIgnoreCase)
        {
            // Basic colors
            { "Black", Black },
            { "White", White },
            { "Red", Red },
            { "Green", Green },
            { "Blue", Blue },
            { "Yellow", Yellow },
            { "Cyan", Cyan },
            { "Magenta", Magenta },

            // Grays
            { "Gray", Gray },
            { "Grey", Gray },
            { "DarkGray", DarkGray },
            { "DarkGrey", DarkGray },
            { "LightGray", LightGray },
            { "LightGrey", LightGray },
            { "Silver", Silver },
            { "DimGray", DimGray },
            { "DimGrey", DimGray },
            { "Gainsboro", Gainsboro },

            // Standard HTML/CSS colors
            { "Aqua", Aqua },
            { "Fuchsia", Fuchsia },
            { "Lime", Lime },
            { "Maroon", Maroon },
            { "Navy", Navy },
            { "Olive", Olive },
            { "Purple", Purple },
            { "Teal", Teal },

            // Extended colors
            { "AliceBlue", AliceBlue },
            { "AntiqueWhite", AntiqueWhite },
            { "Aquamarine", Aquamarine },
            { "Azure", Azure },
            { "Beige", Beige },
            { "Bisque", Bisque },
            { "BlanchedAlmond", BlanchedAlmond },
            { "BlueViolet", BlueViolet },
            { "Brown", Brown },
            { "BurlyWood", BurlyWood },
            { "CadetBlue", CadetBlue },
            { "Chartreuse", Chartreuse },
            { "Chocolate", Chocolate },
            { "Coral", Coral },
            { "CornflowerBlue", CornflowerBlue },
            { "Cornsilk", Cornsilk },
            { "Crimson", Crimson },
            { "DarkBlue", DarkBlue },
            { "DarkCyan", DarkCyan },
            { "DarkGoldenrod", DarkGoldenrod },
            { "DarkGreen", DarkGreen },
            { "DarkKhaki", DarkKhaki },
            { "DarkMagenta", DarkMagenta },
            { "DarkOliveGreen", DarkOliveGreen },
            { "DarkOrange", DarkOrange },
            { "DarkOrchid", DarkOrchid },
            { "DarkRed", DarkRed },
            { "DarkSalmon", DarkSalmon },
            { "DarkSeaGreen", DarkSeaGreen },
            { "DarkSlateBlue", DarkSlateBlue },
            { "DarkSlateGray", DarkSlateGray },
            { "DarkSlateGrey", DarkSlateGray },
            { "DarkTurquoise", DarkTurquoise },
            { "DarkViolet", DarkViolet },
            { "DeepPink", DeepPink },
            { "DeepSkyBlue", DeepSkyBlue },
            { "DodgerBlue", DodgerBlue },
            { "Firebrick", Firebrick },
            { "FloralWhite", FloralWhite },
            { "ForestGreen", ForestGreen },
            { "GhostWhite", GhostWhite },
            { "Gold", Gold },
            { "Goldenrod", Goldenrod },
            { "GreenYellow", GreenYellow },
            { "Honeydew", Honeydew },
            { "HotPink", HotPink },
            { "IndianRed", IndianRed },
            { "Indigo", Indigo },
            { "Ivory", Ivory },
            { "Khaki", Khaki },
            { "Lavender", Lavender },
            { "LavenderBlush", LavenderBlush },
            { "LawnGreen", LawnGreen },
            { "LemonChiffon", LemonChiffon },
            { "LightBlue", LightBlue },
            { "LightCoral", LightCoral },
            { "LightCyan", LightCyan },
            { "LightGoldenrodYellow", LightGoldenrodYellow },
            { "LightGreen", LightGreen },
            { "LightPink", LightPink },
            { "LightSalmon", LightSalmon },
            { "LightSeaGreen", LightSeaGreen },
            { "LightSkyBlue", LightSkyBlue },
            { "LightSlateGray", LightSlateGray },
            { "LightSlateGrey", LightSlateGray },
            { "LightSteelBlue", LightSteelBlue },
            { "LightYellow", LightYellow },
            { "LimeGreen", LimeGreen },
            { "Linen", Linen },
            { "MediumAquamarine", MediumAquamarine },
            { "MediumBlue", MediumBlue },
            { "MediumOrchid", MediumOrchid },
            { "MediumPurple", MediumPurple },
            { "MediumSeaGreen", MediumSeaGreen },
            { "MediumSlateBlue", MediumSlateBlue },
            { "MediumSpringGreen", MediumSpringGreen },
            { "MediumTurquoise", MediumTurquoise },
            { "MediumVioletRed", MediumVioletRed },
            { "MidnightBlue", MidnightBlue },
            { "MintCream", MintCream },
            { "MistyRose", MistyRose },
            { "Moccasin", Moccasin },
            { "NavajoWhite", NavajoWhite },
            { "OldLace", OldLace },
            { "OliveDrab", OliveDrab },
            { "Orange", Orange },
            { "OrangeRed", OrangeRed },
            { "Orchid", Orchid },
            { "PaleGoldenrod", PaleGoldenrod },
            { "PaleGreen", PaleGreen },
            { "PaleTurquoise", PaleTurquoise },
            { "PaleVioletRed", PaleVioletRed },
            { "PapayaWhip", PapayaWhip },
            { "PeachPuff", PeachPuff },
            { "Peru", Peru },
            { "Pink", Pink },
            { "Plum", Plum },
            { "PowderBlue", PowderBlue },
            { "RosyBrown", RosyBrown },
            { "RoyalBlue", RoyalBlue },
            { "SaddleBrown", SaddleBrown },
            { "Salmon", Salmon },
            { "SandyBrown", SandyBrown },
            { "SeaGreen", SeaGreen },
            { "SeaShell", SeaShell },
            { "Sienna", Sienna },
            { "SkyBlue", SkyBlue },
            { "SlateBlue", SlateBlue },
            { "SlateGray", SlateGray },
            { "SlateGrey", SlateGray },
            { "Snow", Snow },
            { "SpringGreen", SpringGreen },
            { "SteelBlue", SteelBlue },
            { "Tan", Tan },
            { "Thistle", Thistle },
            { "Tomato", Tomato },
            { "Transparent", Transparent },
            { "Turquoise", Turquoise },
            { "Violet", Violet },
            { "Wheat", Wheat },
            { "WhiteSmoke", WhiteSmoke },
            { "YellowGreen", YellowGreen },
        };

        public static DocxColor FromName(string name)
            => NamedColors.TryGetValue(name, out var color) ? color : DocxColor.Empty;

        public static bool TryFromName(string name, out DocxColor color)
            => NamedColors.TryGetValue(name, out color);

        public static bool IsValidName(string name)
            => NamedColors.ContainsKey(name);
    }
}
