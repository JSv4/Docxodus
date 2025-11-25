// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using SkiaSharp;

namespace OpenXmlPowerTools
{
    public static class ColorParser
    {
        public static SKColor FromName(string name)
        {
            return ColorHelper.FromName(name);
        }

        public static bool TryFromName(string name, out SKColor color)
        {
            return ColorHelper.TryFromName(name, out color);
        }

        public static bool IsValidName(string name)
        {
            return ColorHelper.IsValidName(name);
        }
    }
}
