#nullable enable
// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;

namespace Docxodus
{
    /// <summary>
    /// Platform-independent font family enumeration.
    /// Returns empty set for WASM (browser handles font fallback).
    /// Uses SkiaSharp for .NET builds when available.
    /// Thread-safe for concurrent document conversions.
    /// </summary>
    internal static class FontFamilyHelper
    {
        // Thread-safe lazy initialization for known font families
        private static readonly Lazy<HashSet<string>> _knownFamilies = new Lazy<HashSet<string>>(() =>
        {
            var families = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
#if !WASM_BUILD
            try
            {
                foreach (var fam in SkiaSharp.SKFontManager.Default.FontFamilies)
                    families.Add(fam);
            }
            catch
            {
                // SkiaSharp not available or failed, return empty set
            }
#endif
            return families;
        }, LazyThreadSafetyMode.ExecutionAndPublication);

        // Thread-safe cache for unknown fonts (using byte as dummy value since ConcurrentHashSet doesn't exist)
        private static readonly ConcurrentDictionary<string, byte> _unknownFonts =
            new ConcurrentDictionary<string, byte>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Gets the set of known font families available on the system.
        /// Returns empty set for WASM builds (browser handles font fallback).
        /// </summary>
        public static HashSet<string> KnownFamilies => _knownFamilies.Value;

        /// <summary>
        /// Gets the collection of fonts that have been marked as unknown/unavailable.
        /// </summary>
        public static ICollection<string> UnknownFonts => _unknownFonts.Keys;

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
        /// Thread-safe.
        /// </summary>
        public static void MarkAsUnknown(string fontName)
        {
            if (!string.IsNullOrEmpty(fontName))
                _unknownFonts.TryAdd(fontName, 0);
        }

        /// <summary>
        /// Checks if a font has been marked as unknown.
        /// Thread-safe.
        /// </summary>
        public static bool IsMarkedUnknown(string fontName)
        {
            return !string.IsNullOrEmpty(fontName) && _unknownFonts.ContainsKey(fontName);
        }

        /// <summary>
        /// Clears the unknown fonts cache.
        /// Useful for long-running processes to free memory.
        /// </summary>
        public static void ClearUnknownFontsCache()
        {
            _unknownFonts.Clear();
        }
    }
}
