#nullable enable
// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

namespace Docxodus
{
    /// <summary>
    /// Parses image dimensions from file headers without decoding the full image.
    /// Supports PNG, JPEG, GIF, BMP, WebP, and TIFF formats.
    /// This enables image handling in WASM without SkiaSharp dependency.
    /// </summary>
    public static class ImageHeaderParser
    {
        /// <summary>
        /// Gets image dimensions by parsing the file header.
        /// Returns null if the format is not recognized or dimensions cannot be determined.
        /// </summary>
        /// <param name="bytes">The image file bytes</param>
        /// <returns>Tuple of (Width, Height) or null if parsing fails</returns>
        public static (int Width, int Height)? GetDimensions(byte[] bytes)
        {
            if (bytes == null || bytes.Length < 10)
                return null;

            // PNG: 89 50 4E 47 0D 0A 1A 0A
            if (bytes.Length >= 24 &&
                bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47 &&
                bytes[4] == 0x0D && bytes[5] == 0x0A && bytes[6] == 0x1A && bytes[7] == 0x0A)
            {
                return GetPngDimensions(bytes);
            }

            // JPEG: FF D8 FF
            if (bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
            {
                return GetJpegDimensions(bytes);
            }

            // GIF: 47 49 46 38 (GIF8)
            if (bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x38)
            {
                return GetGifDimensions(bytes);
            }

            // BMP: 42 4D (BM)
            if (bytes[0] == 0x42 && bytes[1] == 0x4D && bytes.Length >= 26)
            {
                return GetBmpDimensions(bytes);
            }

            // WebP: 52 49 46 46 ... 57 45 42 50 (RIFF...WEBP)
            if (bytes.Length > 15 &&
                bytes[0] == 0x52 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x46 &&
                bytes[8] == 0x57 && bytes[9] == 0x45 && bytes[10] == 0x42 && bytes[11] == 0x50)
            {
                return GetWebPDimensions(bytes);
            }

            // TIFF: 49 49 2A 00 (little-endian) or 4D 4D 00 2A (big-endian)
            if (bytes.Length >= 8 &&
                ((bytes[0] == 0x49 && bytes[1] == 0x49 && bytes[2] == 0x2A && bytes[3] == 0x00) ||
                 (bytes[0] == 0x4D && bytes[1] == 0x4D && bytes[2] == 0x00 && bytes[3] == 0x2A)))
            {
                return GetTiffDimensions(bytes);
            }

            return null;
        }

        /// <summary>
        /// Detects the image format from file header bytes.
        /// </summary>
        public static string? DetectFormat(byte[] bytes)
        {
            if (bytes == null || bytes.Length < 4)
                return null;

            if (bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47)
                return "png";

            if (bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
                return "jpeg";

            if (bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x38)
                return "gif";

            if (bytes[0] == 0x42 && bytes[1] == 0x4D)
                return "bmp";

            if (bytes.Length > 11 &&
                bytes[0] == 0x52 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x46 &&
                bytes[8] == 0x57 && bytes[9] == 0x45 && bytes[10] == 0x42 && bytes[11] == 0x50)
                return "webp";

            if ((bytes[0] == 0x49 && bytes[1] == 0x49 && bytes[2] == 0x2A && bytes[3] == 0x00) ||
                (bytes[0] == 0x4D && bytes[1] == 0x4D && bytes[2] == 0x00 && bytes[3] == 0x2A))
                return "tiff";

            return null;
        }

        private static (int, int)? GetPngDimensions(byte[] bytes)
        {
            // IHDR chunk: offset 8 (chunk length) + 4 (type) = 12
            // Dimensions at bytes 16-23 (big-endian)
            if (bytes.Length < 24) return null;

            int width = (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
            int height = (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23];

            if (width <= 0 || height <= 0 || width > 65535 || height > 65535)
                return null;

            return (width, height);
        }

        private static (int, int)? GetJpegDimensions(byte[] bytes)
        {
            // Scan for SOF0 (0xFFC0), SOF1 (0xFFC1), SOF2 (0xFFC2), or SOF3 (0xFFC3) markers
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

                    if (width > 0 && height > 0)
                        return (width, height);
                }

                // Skip marker
                if (marker == 0xD8 || marker == 0xD9 || marker == 0x01)
                {
                    // Standalone markers
                    i += 2;
                }
                else if (marker >= 0xD0 && marker <= 0xD7)
                {
                    // RST markers (no length)
                    i += 2;
                }
                else if (marker == 0x00)
                {
                    // Stuffed byte, skip
                    i += 1;
                }
                else if (i + 3 < bytes.Length)
                {
                    // Read segment length and skip
                    int length = (bytes[i + 2] << 8) | bytes[i + 3];
                    if (length < 2)
                        break; // Invalid length
                    i += 2 + length;
                }
                else
                {
                    break;
                }
            }
            return null;
        }

        private static (int, int)? GetGifDimensions(byte[] bytes)
        {
            // Logical screen dimensions at bytes 6-9 (little-endian)
            if (bytes.Length < 10) return null;

            int width = bytes[6] | (bytes[7] << 8);
            int height = bytes[8] | (bytes[9] << 8);

            if (width <= 0 || height <= 0)
                return null;

            return (width, height);
        }

        private static (int, int)? GetBmpDimensions(byte[] bytes)
        {
            // DIB header starts at offset 14
            // Dimensions at bytes 18-25 (little-endian, signed for height)
            if (bytes.Length < 26) return null;

            int width = bytes[18] | (bytes[19] << 8) | (bytes[20] << 16) | (bytes[21] << 24);
            int height = bytes[22] | (bytes[23] << 8) | (bytes[24] << 16) | (bytes[25] << 24);

            // Height can be negative (top-down bitmap)
            height = Math.Abs(height);

            if (width <= 0 || height <= 0)
                return null;

            return (width, height);
        }

        private static (int, int)? GetWebPDimensions(byte[] bytes)
        {
            if (bytes.Length < 30) return null;

            // Check for VP8 (lossy), VP8L (lossless), or VP8X (extended)
            // Format identifier starts at byte 12

            // VP8 (lossy): "VP8 " (note the space)
            if (bytes.Length >= 30 &&
                bytes[12] == 0x56 && bytes[13] == 0x50 && bytes[14] == 0x38 && bytes[15] == 0x20)
            {
                // Frame header at offset 23 (after VP8 bitstream header)
                // Check for frame tag
                if (bytes.Length < 30) return null;

                // Width and height are at offset 26-29, but need to parse VP8 frame header
                // Simplified: look for dimensions after keyframe signature
                int offset = 23;
                if (offset + 6 < bytes.Length)
                {
                    // Check for keyframe (0x9D 0x01 0x2A)
                    if (bytes[offset] == 0x9D && bytes[offset + 1] == 0x01 && bytes[offset + 2] == 0x2A)
                    {
                        int width = (bytes[offset + 3] | (bytes[offset + 4] << 8)) & 0x3FFF;
                        int height = (bytes[offset + 5] | (bytes[offset + 6] << 8)) & 0x3FFF;
                        if (width > 0 && height > 0)
                            return (width, height);
                    }
                }
            }

            // VP8L (lossless): "VP8L"
            if (bytes.Length >= 25 &&
                bytes[12] == 0x56 && bytes[13] == 0x50 && bytes[14] == 0x38 && bytes[15] == 0x4C)
            {
                // Signature byte at offset 20 should be 0x2F
                if (bytes[20] != 0x2F) return null;

                // Dimensions are encoded in bytes 21-24
                int b0 = bytes[21], b1 = bytes[22], b2 = bytes[23], b3 = bytes[24];
                int width = 1 + ((b0) | ((b1 & 0x3F) << 8));
                int height = 1 + (((b1 & 0xC0) >> 6) | (b2 << 2) | ((b3 & 0x0F) << 10));

                if (width > 0 && height > 0)
                    return (width, height);
            }

            // VP8X (extended): "VP8X"
            if (bytes.Length >= 30 &&
                bytes[12] == 0x56 && bytes[13] == 0x50 && bytes[14] == 0x38 && bytes[15] == 0x58)
            {
                // Canvas size at offset 24-29 (24-bit values, little-endian, +1)
                int width = 1 + (bytes[24] | (bytes[25] << 8) | (bytes[26] << 16));
                int height = 1 + (bytes[27] | (bytes[28] << 8) | (bytes[29] << 16));

                if (width > 0 && height > 0)
                    return (width, height);
            }

            return null;
        }

        private static (int, int)? GetTiffDimensions(byte[] bytes)
        {
            if (bytes.Length < 8) return null;

            bool isLittleEndian = bytes[0] == 0x49; // 'I' = little-endian, 'M' = big-endian

            // Read IFD offset (bytes 4-7)
            int ifdOffset = isLittleEndian
                ? bytes[4] | (bytes[5] << 8) | (bytes[6] << 16) | (bytes[7] << 24)
                : (bytes[4] << 24) | (bytes[5] << 16) | (bytes[6] << 8) | bytes[7];

            if (ifdOffset < 0 || ifdOffset + 2 >= bytes.Length)
                return null;

            // Read number of directory entries
            int numEntries = isLittleEndian
                ? bytes[ifdOffset] | (bytes[ifdOffset + 1] << 8)
                : (bytes[ifdOffset] << 8) | bytes[ifdOffset + 1];

            int width = 0, height = 0;

            // Each entry is 12 bytes
            for (int i = 0; i < numEntries && ifdOffset + 2 + (i + 1) * 12 <= bytes.Length; i++)
            {
                int entryOffset = ifdOffset + 2 + i * 12;

                int tag = isLittleEndian
                    ? bytes[entryOffset] | (bytes[entryOffset + 1] << 8)
                    : (bytes[entryOffset] << 8) | bytes[entryOffset + 1];

                // Tag 256 = ImageWidth, Tag 257 = ImageLength (height)
                if (tag == 256 || tag == 257)
                {
                    int type = isLittleEndian
                        ? bytes[entryOffset + 2] | (bytes[entryOffset + 3] << 8)
                        : (bytes[entryOffset + 2] << 8) | bytes[entryOffset + 3];

                    int value;
                    if (type == 3) // SHORT (2 bytes)
                    {
                        value = isLittleEndian
                            ? bytes[entryOffset + 8] | (bytes[entryOffset + 9] << 8)
                            : (bytes[entryOffset + 8] << 8) | bytes[entryOffset + 9];
                    }
                    else // LONG (4 bytes)
                    {
                        value = isLittleEndian
                            ? bytes[entryOffset + 8] | (bytes[entryOffset + 9] << 8) |
                              (bytes[entryOffset + 10] << 16) | (bytes[entryOffset + 11] << 24)
                            : (bytes[entryOffset + 8] << 24) | (bytes[entryOffset + 9] << 16) |
                              (bytes[entryOffset + 10] << 8) | bytes[entryOffset + 11];
                    }

                    if (tag == 256) width = value;
                    else height = value;
                }

                if (width > 0 && height > 0)
                    return (width, height);
            }

            if (width > 0 && height > 0)
                return (width, height);

            return null;
        }
    }
}
