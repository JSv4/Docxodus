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
            if (bytes == null)
                return null;

            // PNG: 89 50 4E 47 0D 0A 1A 0A
            if (bytes.Length >= 24 &&
                bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47 &&
                bytes[4] == 0x0D && bytes[5] == 0x0A && bytes[6] == 0x1A && bytes[7] == 0x0A)
            {
                return GetPngDimensions(bytes);
            }

            // JPEG: FF D8 FF
            if (bytes.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
            {
                return GetJpegDimensions(bytes);
            }

            // GIF: 47 49 46 38 (GIF8)
            if (HasGifSignature(bytes))
            {
                return GetGifDimensions(bytes);
            }

            // BMP: 42 4D (BM)
            if (bytes.Length >= 26 && bytes[0] == 0x42 && bytes[1] == 0x4D)
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

            if (bytes.Length >= 8 && bytes[0] == 0x89 && bytes[1] == 0x50
                && bytes[2] == 0x4E && bytes[3] == 0x47 && bytes[4] == 0x0D
                && bytes[5] == 0x0A && bytes[6] == 0x1A && bytes[7] == 0x0A)
                return "png";

            if (bytes.Length >= 3 && bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF)
                return "jpeg";

            if (HasGifSignature(bytes))
                return "gif";

            if (bytes.Length >= 2 && bytes[0] == 0x42 && bytes[1] == 0x4D)
                return "bmp";

            if (bytes.Length > 11 &&
                bytes[0] == 0x52 && bytes[1] == 0x49 && bytes[2] == 0x46 && bytes[3] == 0x46 &&
                bytes[8] == 0x57 && bytes[9] == 0x45 && bytes[10] == 0x42 && bytes[11] == 0x50)
                return "webp";

            if (bytes.Length >= 4 && ((bytes[0] == 0x49 && bytes[1] == 0x49 && bytes[2] == 0x2A && bytes[3] == 0x00) ||
                (bytes[0] == 0x4D && bytes[1] == 0x4D && bytes[2] == 0x00 && bytes[3] == 0x2A)))
                return "tiff";

            return null;
        }

        private static (int, int)? GetPngDimensions(byte[] bytes)
        {
            // IHDR chunk: offset 8 (chunk length) + 4 (type) = 12
            // Dimensions at bytes 16-23 (big-endian)
            if (bytes.Length < 24) return null;
            if (bytes[8] != 0 || bytes[9] != 0 || bytes[10] != 0 || bytes[11] != 13
                || bytes[12] != (byte)'I' || bytes[13] != (byte)'H'
                || bytes[14] != (byte)'D' || bytes[15] != (byte)'R') return null;

            int width = (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
            int height = (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23];

            if (width <= 0 || height <= 0)
                return null;

            return (width, height);
        }

        private static (int, int)? GetJpegDimensions(byte[] bytes)
        {
            int i = 2;
            while (i < bytes.Length)
            {
                while (i < bytes.Length && bytes[i] != 0xFF) i++;
                while (i < bytes.Length && bytes[i] == 0xFF) i++;
                if (i >= bytes.Length) return null;
                byte marker = bytes[i++];
                if (marker == 0x00) continue;
                if (marker == 0xD9 || marker == 0xDA) return null;
                if (marker == 0xD8 || marker == 0x01 || marker is >= 0xD0 and <= 0xD7) continue;
                if (i + 2 > bytes.Length) return null;
                int length = (bytes[i] << 8) | bytes[i + 1];
                if (length < 2 || length > bytes.Length - i) return null;
                bool isStartOfFrame = marker is >= 0xC0 and <= 0xCF
                    && marker is not (0xC4 or 0xC8 or 0xCC);
                if (isStartOfFrame)
                {
                    if (length < 8 || i + 7 >= bytes.Length) return null;
                    int height = (bytes[i + 3] << 8) | bytes[i + 4];
                    int width = (bytes[i + 5] << 8) | bytes[i + 6];
                    return width > 0 && height > 0 ? (width, height) : null;
                }
                i += length;
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

        private static bool HasGifSignature(byte[] bytes) => bytes.Length >= 6
            && bytes[0] == (byte)'G' && bytes[1] == (byte)'I' && bytes[2] == (byte)'F'
            && bytes[3] == (byte)'8' && (bytes[4] == (byte)'7' || bytes[4] == (byte)'9')
            && bytes[5] == (byte)'a';

        private static (int, int)? GetBmpDimensions(byte[] bytes)
        {
            // DIB header starts at offset 14
            // Dimensions at bytes 18-25 (little-endian, signed for height)
            if (bytes.Length < 26) return null;

            uint dibSize = (uint)(bytes[14] | (bytes[15] << 8) | (bytes[16] << 16) | (bytes[17] << 24));
            if (dibSize < 40 || dibSize > bytes.Length - 14) return null;

            int width = bytes[18] | (bytes[19] << 8) | (bytes[20] << 16) | (bytes[21] << 24);
            int height = bytes[22] | (bytes[23] << 8) | (bytes[24] << 16) | (bytes[25] << 24);

            // Height can be negative (top-down bitmap)
            if (height == int.MinValue) return null;
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
            uint ifdOffsetValue = ReadUInt32(bytes, 4, isLittleEndian);

            if (ifdOffsetValue > int.MaxValue || ifdOffsetValue > (uint)(bytes.Length - 2))
                return null;
            int ifdOffset = (int)ifdOffsetValue;

            // Read number of directory entries
            int numEntries = isLittleEndian
                ? bytes[ifdOffset] | (bytes[ifdOffset + 1] << 8)
                : (bytes[ifdOffset] << 8) | bytes[ifdOffset + 1];

            int width = 0, height = 0;

            // Each entry is 12 bytes
            long entriesEnd = (long)ifdOffset + 2L + (long)numEntries * 12L;
            if (entriesEnd > bytes.Length) return null;
            for (int i = 0; i < numEntries; i++)
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
                    uint count = ReadUInt32(bytes, entryOffset + 4, isLittleEndian);
                    if (count != 1 || type is not (3 or 4)) continue;

                    uint rawValue;
                    if (type == 3) // SHORT (2 bytes)
                    {
                        rawValue = (uint)(isLittleEndian
                            ? bytes[entryOffset + 8] | (bytes[entryOffset + 9] << 8)
                            : (bytes[entryOffset + 8] << 8) | bytes[entryOffset + 9]);
                    }
                    else rawValue = ReadUInt32(bytes, entryOffset + 8, isLittleEndian);
                    if (rawValue == 0 || rawValue > int.MaxValue) return null;
                    int value = (int)rawValue;

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

        private static uint ReadUInt32(byte[] bytes, int offset, bool littleEndian) => littleEndian
            ? (uint)(bytes[offset] | (bytes[offset + 1] << 8)
                | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24))
            : (uint)((bytes[offset] << 24) | (bytes[offset + 1] << 16)
                | (bytes[offset + 2] << 8) | bytes[offset + 3]);
    }
}
