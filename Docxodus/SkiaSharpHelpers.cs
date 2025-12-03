// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Docxodus
{
    /// <summary>
    /// Helper class providing color name mapping.
    /// This class delegates to DocxColors for platform-independent color handling.
    /// Maintained for backward compatibility.
    /// </summary>
    public static class ColorHelper
    {
        /// <summary>
        /// Gets a DocxColor from a named color string.
        /// </summary>
        public static DocxColor FromName(string name)
        {
            return DocxColors.FromName(name);
        }

        /// <summary>
        /// Tries to get a DocxColor from a named color string.
        /// </summary>
        public static bool TryFromName(string name, out DocxColor color)
        {
            return DocxColors.TryFromName(name, out color);
        }

        /// <summary>
        /// Checks if a color name is valid.
        /// </summary>
        public static bool IsValidName(string name)
        {
            return DocxColors.IsValidName(name);
        }

        /// <summary>
        /// Creates a DocxColor from ARGB components.
        /// </summary>
        public static DocxColor FromArgb(int alpha, int red, int green, int blue)
        {
            return DocxColor.FromArgb((byte)alpha, (byte)red, (byte)green, (byte)blue);
        }

        /// <summary>
        /// Creates a DocxColor from RGB components (fully opaque).
        /// </summary>
        public static DocxColor FromArgb(int red, int green, int blue)
        {
            return DocxColor.FromArgb((byte)red, (byte)green, (byte)blue);
        }

        /// <summary>
        /// Creates a DocxColor from a packed ARGB value.
        /// </summary>
        public static DocxColor FromArgb(int argb)
        {
            return DocxColor.FromArgb(argb);
        }
    }
}
