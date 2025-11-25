// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using SkiaSharp;
using System;
using System.Collections.Generic;

namespace OpenXmlPowerTools
{
    /// <summary>
    /// Helper class providing color name mapping for SkiaSharp (which lacks Color.FromName).
    /// </summary>
    public static class ColorHelper
    {
        private static readonly Dictionary<string, SKColor> NamedColors = new(StringComparer.OrdinalIgnoreCase)
        {
            // Basic colors
            { "Black", SKColors.Black },
            { "White", SKColors.White },
            { "Red", SKColors.Red },
            { "Green", SKColors.Green },
            { "Blue", SKColors.Blue },
            { "Yellow", SKColors.Yellow },
            { "Cyan", SKColors.Cyan },
            { "Magenta", SKColors.Magenta },

            // Grays
            { "Gray", SKColors.Gray },
            { "Grey", SKColors.Gray },
            { "DarkGray", SKColors.DarkGray },
            { "DarkGrey", SKColors.DarkGray },
            { "LightGray", SKColors.LightGray },
            { "LightGrey", SKColors.LightGray },
            { "Silver", SKColors.Silver },

            // Standard HTML/CSS colors
            { "Aqua", SKColors.Aqua },
            { "Fuchsia", SKColors.Fuchsia },
            { "Lime", SKColors.Lime },
            { "Maroon", SKColors.Maroon },
            { "Navy", SKColors.Navy },
            { "Olive", SKColors.Olive },
            { "Purple", SKColors.Purple },
            { "Teal", SKColors.Teal },

            // Extended colors
            { "AliceBlue", SKColors.AliceBlue },
            { "AntiqueWhite", SKColors.AntiqueWhite },
            { "Aquamarine", SKColors.Aquamarine },
            { "Azure", SKColors.Azure },
            { "Beige", SKColors.Beige },
            { "Bisque", SKColors.Bisque },
            { "BlanchedAlmond", SKColors.BlanchedAlmond },
            { "BlueViolet", SKColors.BlueViolet },
            { "Brown", SKColors.Brown },
            { "BurlyWood", SKColors.BurlyWood },
            { "CadetBlue", SKColors.CadetBlue },
            { "Chartreuse", SKColors.Chartreuse },
            { "Chocolate", SKColors.Chocolate },
            { "Coral", SKColors.Coral },
            { "CornflowerBlue", SKColors.CornflowerBlue },
            { "Cornsilk", SKColors.Cornsilk },
            { "Crimson", SKColors.Crimson },
            { "DarkBlue", SKColors.DarkBlue },
            { "DarkCyan", SKColors.DarkCyan },
            { "DarkGoldenrod", SKColors.DarkGoldenrod },
            { "DarkGreen", SKColors.DarkGreen },
            { "DarkKhaki", SKColors.DarkKhaki },
            { "DarkMagenta", SKColors.DarkMagenta },
            { "DarkOliveGreen", SKColors.DarkOliveGreen },
            { "DarkOrange", SKColors.DarkOrange },
            { "DarkOrchid", SKColors.DarkOrchid },
            { "DarkRed", SKColors.DarkRed },
            { "DarkSalmon", SKColors.DarkSalmon },
            { "DarkSeaGreen", SKColors.DarkSeaGreen },
            { "DarkSlateBlue", SKColors.DarkSlateBlue },
            { "DarkSlateGray", SKColors.DarkSlateGray },
            { "DarkSlateGrey", SKColors.DarkSlateGray },
            { "DarkTurquoise", SKColors.DarkTurquoise },
            { "DarkViolet", SKColors.DarkViolet },
            { "DeepPink", SKColors.DeepPink },
            { "DeepSkyBlue", SKColors.DeepSkyBlue },
            { "DimGray", SKColors.DimGray },
            { "DimGrey", SKColors.DimGray },
            { "DodgerBlue", SKColors.DodgerBlue },
            { "Firebrick", SKColors.Firebrick },
            { "FloralWhite", SKColors.FloralWhite },
            { "ForestGreen", SKColors.ForestGreen },
            { "Gainsboro", SKColors.Gainsboro },
            { "GhostWhite", SKColors.GhostWhite },
            { "Gold", SKColors.Gold },
            { "Goldenrod", SKColors.Goldenrod },
            { "GreenYellow", SKColors.GreenYellow },
            { "Honeydew", SKColors.Honeydew },
            { "HotPink", SKColors.HotPink },
            { "IndianRed", SKColors.IndianRed },
            { "Indigo", SKColors.Indigo },
            { "Ivory", SKColors.Ivory },
            { "Khaki", SKColors.Khaki },
            { "Lavender", SKColors.Lavender },
            { "LavenderBlush", SKColors.LavenderBlush },
            { "LawnGreen", SKColors.LawnGreen },
            { "LemonChiffon", SKColors.LemonChiffon },
            { "LightBlue", SKColors.LightBlue },
            { "LightCoral", SKColors.LightCoral },
            { "LightCyan", SKColors.LightCyan },
            { "LightGoldenrodYellow", SKColors.LightGoldenrodYellow },
            { "LightGreen", SKColors.LightGreen },
            { "LightPink", SKColors.LightPink },
            { "LightSalmon", SKColors.LightSalmon },
            { "LightSeaGreen", SKColors.LightSeaGreen },
            { "LightSkyBlue", SKColors.LightSkyBlue },
            { "LightSlateGray", SKColors.LightSlateGray },
            { "LightSlateGrey", SKColors.LightSlateGray },
            { "LightSteelBlue", SKColors.LightSteelBlue },
            { "LightYellow", SKColors.LightYellow },
            { "LimeGreen", SKColors.LimeGreen },
            { "Linen", SKColors.Linen },
            { "MediumAquamarine", SKColors.MediumAquamarine },
            { "MediumBlue", SKColors.MediumBlue },
            { "MediumOrchid", SKColors.MediumOrchid },
            { "MediumPurple", SKColors.MediumPurple },
            { "MediumSeaGreen", SKColors.MediumSeaGreen },
            { "MediumSlateBlue", SKColors.MediumSlateBlue },
            { "MediumSpringGreen", SKColors.MediumSpringGreen },
            { "MediumTurquoise", SKColors.MediumTurquoise },
            { "MediumVioletRed", SKColors.MediumVioletRed },
            { "MidnightBlue", SKColors.MidnightBlue },
            { "MintCream", SKColors.MintCream },
            { "MistyRose", SKColors.MistyRose },
            { "Moccasin", SKColors.Moccasin },
            { "NavajoWhite", SKColors.NavajoWhite },
            { "OldLace", SKColors.OldLace },
            { "OliveDrab", SKColors.OliveDrab },
            { "Orange", SKColors.Orange },
            { "OrangeRed", SKColors.OrangeRed },
            { "Orchid", SKColors.Orchid },
            { "PaleGoldenrod", SKColors.PaleGoldenrod },
            { "PaleGreen", SKColors.PaleGreen },
            { "PaleTurquoise", SKColors.PaleTurquoise },
            { "PaleVioletRed", SKColors.PaleVioletRed },
            { "PapayaWhip", SKColors.PapayaWhip },
            { "PeachPuff", SKColors.PeachPuff },
            { "Peru", SKColors.Peru },
            { "Pink", SKColors.Pink },
            { "Plum", SKColors.Plum },
            { "PowderBlue", SKColors.PowderBlue },
            { "RosyBrown", SKColors.RosyBrown },
            { "RoyalBlue", SKColors.RoyalBlue },
            { "SaddleBrown", SKColors.SaddleBrown },
            { "Salmon", SKColors.Salmon },
            { "SandyBrown", SKColors.SandyBrown },
            { "SeaGreen", SKColors.SeaGreen },
            { "SeaShell", SKColors.SeaShell },
            { "Sienna", SKColors.Sienna },
            { "SkyBlue", SKColors.SkyBlue },
            { "SlateBlue", SKColors.SlateBlue },
            { "SlateGray", SKColors.SlateGray },
            { "SlateGrey", SKColors.SlateGray },
            { "Snow", SKColors.Snow },
            { "SpringGreen", SKColors.SpringGreen },
            { "SteelBlue", SKColors.SteelBlue },
            { "Tan", SKColors.Tan },
            { "Thistle", SKColors.Thistle },
            { "Tomato", SKColors.Tomato },
            { "Transparent", SKColors.Transparent },
            { "Turquoise", SKColors.Turquoise },
            { "Violet", SKColors.Violet },
            { "Wheat", SKColors.Wheat },
            { "WhiteSmoke", SKColors.WhiteSmoke },
            { "YellowGreen", SKColors.YellowGreen },
        };

        /// <summary>
        /// Gets an SKColor from a named color string.
        /// </summary>
        public static SKColor FromName(string name)
        {
            if (NamedColors.TryGetValue(name, out var color))
                return color;
            return SKColors.Empty;
        }

        /// <summary>
        /// Tries to get an SKColor from a named color string.
        /// </summary>
        public static bool TryFromName(string name, out SKColor color)
        {
            return NamedColors.TryGetValue(name, out color);
        }

        /// <summary>
        /// Checks if a color name is valid.
        /// </summary>
        public static bool IsValidName(string name)
        {
            return NamedColors.ContainsKey(name);
        }

        /// <summary>
        /// Creates an SKColor from ARGB components.
        /// </summary>
        public static SKColor FromArgb(int alpha, int red, int green, int blue)
        {
            return new SKColor((byte)red, (byte)green, (byte)blue, (byte)alpha);
        }

        /// <summary>
        /// Creates an SKColor from RGB components (fully opaque).
        /// </summary>
        public static SKColor FromArgb(int red, int green, int blue)
        {
            return new SKColor((byte)red, (byte)green, (byte)blue);
        }

        /// <summary>
        /// Creates an SKColor from a packed ARGB value.
        /// </summary>
        public static SKColor FromArgb(int argb)
        {
            return new SKColor((uint)argb);
        }
    }

    /// <summary>
    /// Extension methods for SKColor to provide System.Drawing.Color-like functionality.
    /// </summary>
    public static class SKColorExtensions
    {
        public static int ToArgb(this SKColor color)
        {
            return (color.Alpha << 24) | (color.Red << 16) | (color.Green << 8) | color.Blue;
        }

        public static byte GetA(this SKColor color) => color.Alpha;
        public static byte GetR(this SKColor color) => color.Red;
        public static byte GetG(this SKColor color) => color.Green;
        public static byte GetB(this SKColor color) => color.Blue;
    }
}
