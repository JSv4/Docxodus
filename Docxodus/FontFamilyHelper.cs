#nullable enable
// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;

namespace Docxodus
{
    /// <summary>
    /// Platform-independent font family enumeration.
    /// Returns empty set for WASM (browser handles font fallback).
    /// Uses SkiaSharp for .NET builds when available.
    /// </summary>
    internal static class FontFamilyHelper
    {
        private static HashSet<string>? _knownFamilies;
        private static readonly HashSet<string> _unknownFonts = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Gets the set of known font families available on the system.
        /// Returns empty set for WASM builds (browser handles font fallback).
        /// </summary>
        public static HashSet<string> KnownFamilies
        {
            get
            {
                if (_knownFamilies == null)
                {
                    _knownFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
#if !WASM_BUILD
                    try
                    {
                        var families = SkiaSharp.SKFontManager.Default.FontFamilies;
                        foreach (var fam in families)
                            _knownFamilies.Add(fam);
                    }
                    catch
                    {
                        // SkiaSharp not available or failed, return empty set
                    }
#endif
                }
                return _knownFamilies;
            }
        }

        /// <summary>
        /// Gets the set of fonts that have been marked as unknown/unavailable.
        /// </summary>
        public static HashSet<string> UnknownFonts => _unknownFonts;

        /// <summary>
        /// Checks if a font family is available on the system.
        /// Always returns true for WASM (browser handles font fallback).
        /// </summary>
        public static bool IsFontAvailable(string fontName)
        {
#if WASM_BUILD
            return true; // Browser handles font fallback
#else
            if (string.IsNullOrEmpty(fontName))
                return false;
            return KnownFamilies.Contains(fontName);
#endif
        }

        /// <summary>
        /// Marks a font as unknown/unavailable to avoid repeated lookups.
        /// </summary>
        public static void MarkAsUnknown(string fontName)
        {
            if (!string.IsNullOrEmpty(fontName))
                _unknownFonts.Add(fontName);
        }

        /// <summary>
        /// Checks if a font has been marked as unknown.
        /// </summary>
        public static bool IsMarkedUnknown(string fontName)
        {
            return !string.IsNullOrEmpty(fontName) && _unknownFonts.Contains(fontName);
        }
    }
}
