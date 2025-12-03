# SkiaSharp Removal Plan for WASM

## Overview

This document outlines the plan to remove SkiaSharp dependency from the WASM build, reducing bundle size from ~50MB to ~5-10MB.

## Current SkiaSharp Usage Summary

| Category | Files | Lines Changed | Difficulty |
|----------|-------|---------------|------------|
| Color (SKColor) | 5 files | ~300 lines | Easy |
| Font Enumeration | 4 files | ~50 lines | Easy |
| Image Processing | 3 files | ~150 lines | Medium |
| Text Measurement | 1 file | ~50 lines | Medium |
| **Total** | **10 files** | **~550 lines** | |

---

## Phase 1: Create DocxColor Struct

### New File: `Docxodus/DocxColor.cs`

```csharp
#nullable enable
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

        public static DocxColor FromArgb(int argb)
            => new((byte)(argb >> 16), (byte)(argb >> 8), (byte)argb, (byte)(argb >> 24));

        public int ToArgb()
            => (Alpha << 24) | (Red << 16) | (Green << 8) | Blue;

        public string ToHex() => $"#{Red:X2}{Green:X2}{Blue:X2}";
        public string ToHexWithAlpha() => $"#{Alpha:X2}{Red:X2}{Green:X2}{Blue:X2}";

        // Standard colors
        public static DocxColor Empty => new(0, 0, 0, 0);
        public static DocxColor Transparent => new(0, 0, 0, 0);
        public static DocxColor Black => new(0, 0, 0);
        public static DocxColor White => new(255, 255, 255);
        public static DocxColor Red => new(255, 0, 0);
        public static DocxColor Green => new(0, 128, 0);
        public static DocxColor Blue => new(0, 0, 255);
        public static DocxColor Yellow => new(255, 255, 0);
        public static DocxColor Cyan => new(0, 255, 255);
        public static DocxColor Magenta => new(255, 0, 255);
        public static DocxColor Gray => new(128, 128, 128);
        // ... (all 140+ named colors from SkiaSharpHelpers.cs)

        public bool Equals(DocxColor other)
            => Alpha == other.Alpha && Red == other.Red && Green == other.Green && Blue == other.Blue;

        public override bool Equals(object? obj) => obj is DocxColor c && Equals(c);
        public override int GetHashCode() => ToArgb();
        public static bool operator ==(DocxColor left, DocxColor right) => left.Equals(right);
        public static bool operator !=(DocxColor left, DocxColor right) => !left.Equals(right);
    }

    /// <summary>
    /// Color name lookup replacing ColorHelper/ColorParser.
    /// </summary>
    public static class DocxColors
    {
        private static readonly Dictionary<string, DocxColor> NamedColors = new(StringComparer.OrdinalIgnoreCase)
        {
            { "Black", DocxColor.Black },
            { "White", DocxColor.White },
            // ... all named colors
        };

        public static DocxColor FromName(string name)
            => NamedColors.TryGetValue(name, out var color) ? color : DocxColor.Empty;

        public static bool TryFromName(string name, out DocxColor color)
            => NamedColors.TryGetValue(name, out color);

        public static bool IsValidName(string name)
            => NamedColors.ContainsKey(name);
    }
}
```

### Files to Update

| File | Changes |
|------|---------|
| `SkiaSharpHelpers.cs` | Replace `SKColor` → `DocxColor`, remove SkiaSharp using |
| `ColorParser.cs` | Replace `SKColor` → `DocxColor`, remove SkiaSharp using |
| `HtmlToWmlCssParser.cs` | Replace `SKColor` → `DocxColor` in ~10 locations |
| `WmlComparer.cs` | Replace `SKColor` → `DocxColor` in 2 locations |

---

## Phase 2: Conditional Font Enumeration

### Strategy
Use `#if` directives to provide different implementations for WASM vs .NET.

### File: `Docxodus/FontFamilyHelper.cs` (New)

```csharp
#nullable enable
namespace Docxodus
{
    /// <summary>
    /// Platform-independent font family enumeration.
    /// Returns empty set for WASM (browser handles font fallback).
    /// </summary>
    internal static class FontFamilyHelper
    {
        private static HashSet<string>? _knownFamilies;

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
                        // SkiaSharp not available, return empty set
                    }
#endif
                }
                return _knownFamilies;
            }
        }

        public static bool IsFontAvailable(string fontName)
        {
#if WASM_BUILD
            return true; // Browser handles fallback
#else
            return KnownFamilies.Contains(fontName);
#endif
        }
    }
}
```

### Files to Update

| File | Line | Change |
|------|------|--------|
| `WmlToHtmlConverter.cs` | 5653 | Use `FontFamilyHelper.KnownFamilies` |
| `HtmlToWmlConverterCore.cs` | 1871 | Use `FontFamilyHelper.KnownFamilies` |
| `SmlToHtmlConverter.cs` | 239 | Use `FontFamilyHelper.KnownFamilies` |
| `PtOpenXmlUtil.cs` | 755 | Use `FontFamilyHelper.KnownFamilies` |

---

## Phase 3: Image Header Dimension Parser

### New File: `Docxodus/ImageHeaderParser.cs`

```csharp
#nullable enable
namespace Docxodus
{
    /// <summary>
    /// Parses image dimensions from file headers without decoding.
    /// Supports PNG, JPEG, GIF, BMP, WebP formats.
    /// </summary>
    public static class ImageHeaderParser
    {
        public static (int Width, int Height)? GetDimensions(byte[] bytes)
        {
            if (bytes == null || bytes.Length < 24)
                return null;

            // PNG: 89 50 4E 47 0D 0A 1A 0A
            if (bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
                return GetPngDimensions(bytes);

            // JPEG: FF D8 FF
            if (bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
                return GetJpegDimensions(bytes);

            // GIF: 47 49 46 38 (GIF8)
            if (bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x38)
                return GetGifDimensions(bytes);

            // BMP: 42 4D (BM)
            if (bytes[0] == 0x42 && bytes[1] == 0x4D)
                return GetBmpDimensions(bytes);

            // WebP: 52 49 46 46 ... 57 45 42 50 (RIFF...WEBP)
            if (bytes[0] == 0x52 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x46 &&
                bytes.Length > 15 && bytes[8] == 0x57 && bytes[9] == 0x45 && bytes[10] == 0x42 && bytes[11] == 0x50)
                return GetWebPDimensions(bytes);

            return null;
        }

        private static (int, int)? GetPngDimensions(byte[] bytes)
        {
            // IHDR chunk starts at byte 8, dimensions at bytes 16-23 (big-endian)
            if (bytes.Length < 24) return null;
            int width = (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
            int height = (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23];
            return (width, height);
        }

        private static (int, int)? GetJpegDimensions(byte[] bytes)
        {
            // Scan for SOF0 (0xFFC0) or SOF2 (0xFFC2) marker
            int i = 2;
            while (i < bytes.Length - 9)
            {
                if (bytes[i] != 0xFF)
                {
                    i++;
                    continue;
                }

                byte marker = bytes[i + 1];

                // SOF0, SOF1, SOF2, SOF3 markers contain dimensions
                if (marker >= 0xC0 && marker <= 0xC3)
                {
                    int height = (bytes[i + 5] << 8) | bytes[i + 6];
                    int width = (bytes[i + 7] << 8) | bytes[i + 8];
                    return (width, height);
                }

                // Skip to next marker
                if (marker == 0xD8 || marker == 0xD9 || marker == 0x01 || (marker >= 0xD0 && marker <= 0xD7))
                {
                    i += 2;
                }
                else
                {
                    int length = (bytes[i + 2] << 8) | bytes[i + 3];
                    i += 2 + length;
                }
            }
            return null;
        }

        private static (int, int)? GetGifDimensions(byte[] bytes)
        {
            // Dimensions at bytes 6-9 (little-endian)
            if (bytes.Length < 10) return null;
            int width = bytes[6] | (bytes[7] << 8);
            int height = bytes[8] | (bytes[9] << 8);
            return (width, height);
        }

        private static (int, int)? GetBmpDimensions(byte[] bytes)
        {
            // Dimensions at bytes 18-25 (little-endian, signed)
            if (bytes.Length < 26) return null;
            int width = bytes[18] | (bytes[19] << 8) | (bytes[20] << 16) | (bytes[21] << 24);
            int height = bytes[22] | (bytes[23] << 8) | (bytes[24] << 16) | (bytes[25] << 24);
            return (width, Math.Abs(height)); // Height can be negative
        }

        private static (int, int)? GetWebPDimensions(byte[] bytes)
        {
            // WebP has multiple formats, check for VP8/VP8L/VP8X
            if (bytes.Length < 30) return null;

            // Simple lossy WebP (VP8)
            if (bytes[12] == 0x56 && bytes[13] == 0x50 && bytes[14] == 0x38 && bytes[15] == 0x20)
            {
                // Dimensions at bytes 26-29
                int width = (bytes[26] | (bytes[27] << 8)) & 0x3FFF;
                int height = (bytes[28] | (bytes[29] << 8)) & 0x3FFF;
                return (width, height);
            }

            // Lossless WebP (VP8L)
            if (bytes[12] == 0x56 && bytes[13] == 0x50 && bytes[14] == 0x38 && bytes[15] == 0x4C)
            {
                // Dimensions encoded in bytes 21-24
                int b0 = bytes[21], b1 = bytes[22], b2 = bytes[23], b3 = bytes[24];
                int width = 1 + (b0 | ((b1 & 0x3F) << 8));
                int height = 1 + (((b1 & 0xC0) >> 6) | (b2 << 2) | ((b3 & 0x0F) << 10));
                return (width, height);
            }

            return null;
        }
    }
}
```

### Files to Update

| File | Change |
|------|--------|
| `HtmlToWmlConverterCore.cs` | Use `ImageHeaderParser.GetDimensions()` instead of `bmp.Width/Height` |
| `WmlToHtmlConverter.cs` | Already handled (uses document markup for dimensions) |

---

## Phase 4: Update HtmlToWmlConverterCore

### Current Code (Lines 2261-2295)
```csharp
SKBitmap bmp = null;
if (srcAttribute.StartsWith("data:"))
{
    // ...
    bmp = SKBitmap.Decode(ba);
}
else
{
    ba = File.ReadAllBytes(imagePath);
    bmp = SKBitmap.Decode(ba);
}
```

### New Code
```csharp
#if WASM_BUILD
// For WASM, parse dimensions from header without decoding
var dimensions = ImageHeaderParser.GetDimensions(ba);
if (dimensions == null) return null;
int imageWidth = dimensions.Value.Width;
int imageHeight = dimensions.Value.Height;
#else
using var bmp = SkiaSharp.SKBitmap.Decode(ba);
if (bmp == null) return null;
int imageWidth = bmp.Width;
int imageHeight = bmp.Height;
#endif
```

### Methods to Update
- `TransformImageToWml()` - Line 2261
- `GetImageSizeInEmus()` - Line 2536
- `GetImageExtent()` - Line 2570
- `GetGraphicForImage()` - Line 2603

---

## Phase 5: Update MetricsGetter

### Current Usage
```csharp
private static readonly Lazy<SKPaint> MeasurePaint = new(() => new SKPaint { IsAntialias = true });

private static int _getTextWidth(string fontName, bool bold, bool italic, decimal sz, string text)
{
    using var typeface = SKTypeface.FromFamilyName(fontName, weight, width, slant);
    paint.Typeface = typeface;
    return (int)paint.MeasureText(text);
}
```

### WASM Strategy
For WASM, text measurement is not critical. Options:
1. **Return 0** - Skip measurement entirely
2. **Estimate** - Use average character width (0.6 × font size)
3. **Skip metrics** - Return null/empty for text-related metrics

### New Code
```csharp
private static int _getTextWidth(string fontName, bool bold, bool italic, decimal sz, string text)
{
#if WASM_BUILD
    // Estimate: average character width is ~60% of font size
    return (int)(text.Length * (float)sz * 0.6f);
#else
    // Original SkiaSharp implementation
    var weight = bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal;
    var slant = italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright;
    using var typeface = SKTypeface.FromFamilyName(fontName, weight, SKFontStyleWidth.Normal, slant);
    var paint = MeasurePaint.Value;
    paint.Typeface = typeface;
    paint.TextSize = (float)sz / 2f;
    return (int)paint.MeasureText(text);
#endif
}
```

---

## Phase 6: Configure Conditional Compilation

### File: `Docxodus/Docxodus.csproj`

```xml
<PropertyGroup>
  <!-- Existing properties -->
</PropertyGroup>

<!-- SkiaSharp only for non-WASM builds -->
<ItemGroup Condition="'$(WASM_BUILD)' != 'true'">
  <PackageReference Include="SkiaSharp" Version="2.88.9" />
  <PackageReference Include="SkiaSharp.NativeAssets.Linux.NoDependencies" Version="2.88.9" />
</ItemGroup>

<!-- Define WASM_BUILD for conditional compilation -->
<PropertyGroup Condition="'$(WASM_BUILD)' == 'true'">
  <DefineConstants>$(DefineConstants);WASM_BUILD</DefineConstants>
</PropertyGroup>
```

### File: `wasm/DocxodusWasm/DocxodusWasm.csproj`

```xml
<PropertyGroup>
  <!-- Pass WASM_BUILD to referenced projects -->
  <WASM_BUILD>true</WASM_BUILD>
</PropertyGroup>

<!-- Reference Docxodus with WASM_BUILD property -->
<ItemGroup>
  <ProjectReference Include="../../Docxodus/Docxodus.csproj"
                    Properties="WASM_BUILD=true" />
</ItemGroup>
```

---

## Implementation Order

| Step | Task | Est. Time | Risk |
|------|------|-----------|------|
| 1 | Create `DocxColor.cs` with all named colors | 2 hours | Low |
| 2 | Create `FontFamilyHelper.cs` | 30 min | Low |
| 3 | Create `ImageHeaderParser.cs` | 1 hour | Medium |
| 4 | Update `SkiaSharpHelpers.cs` → use `DocxColor` | 30 min | Low |
| 5 | Update `ColorParser.cs` → use `DocxColor` | 15 min | Low |
| 6 | Update `HtmlToWmlCssParser.cs` → use `DocxColor` | 1 hour | Low |
| 7 | Update `WmlComparer.cs` → use `DocxColor` | 30 min | Low |
| 8 | Update font enumeration in 4 files | 30 min | Low |
| 9 | Update `HtmlToWmlConverterCore.cs` for WASM | 2 hours | Medium |
| 10 | Update `MetricsGetter.cs` for WASM | 30 min | Low |
| 11 | Configure csproj files | 30 min | Low |
| 12 | Test and fix issues | 2 hours | Medium |
| **Total** | | **~11 hours** | |

---

## Testing Plan

### Unit Tests
1. `DocxColor` - All named colors, FromArgb, ToArgb, ToHex
2. `ImageHeaderParser` - PNG, JPEG, GIF, BMP, WebP dimensions
3. `FontFamilyHelper` - Empty for WASM, populated for .NET

### Integration Tests
1. DOCX → HTML conversion with images
2. HTML → DOCX conversion with images (if used)
3. Document comparison with highlights
4. All existing tests pass

### Bundle Size Verification
```bash
# Before
du -sh npm/dist/wasm/  # ~50MB

# After (expected)
du -sh npm/dist/wasm/  # ~5-10MB
```

---

## Rollback Plan

If issues arise:
1. All changes are behind `#if WASM_BUILD` conditionals
2. .NET builds continue to use SkiaSharp unchanged
3. Remove `WASM_BUILD=true` from DocxodusWasm.csproj to revert

---

## Future Considerations

### For Full SkiaSharp Removal (.NET too)
1. **ImageSharp** - Cross-platform image processing (~2MB)
2. **SixLabors.Fonts** - Font metrics without native deps
3. **SkiaSharp.Managed** - IL-only Skia (~4MB, no native)

### Performance Notes
- Image header parsing is O(1) for most formats
- Estimated text width is ~10x faster than actual measurement
- Font enumeration skip eliminates startup cost
