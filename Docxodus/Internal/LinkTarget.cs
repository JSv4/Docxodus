// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

namespace Docxodus.Internal;

/// <summary>
/// Which link targets from a document the HTML converter writes into an <c>href</c> (issue #1880). A link the
/// document makes keeps its target unless a browser would run it as script or open it as inline content.
/// </summary>
internal static class LinkTarget
{
    /// <summary>
    /// Whether <paramref name="target"/> is a <c>javascript:</c>, <c>vbscript:</c> or <c>data:</c> URL, read as a
    /// browser's URL parser reads it: without the spaces and control characters around it, and without tabs and line
    /// breaks anywhere in it, and with the scheme in any case. Every other target, relative ones included, is not.
    /// </summary>
    public static bool IsScript(string target)
    {
        // The longest scheme asked about is "javascript"; a longer name is not one of them.
        Span<char> scheme = stackalloc char[10];
        var length = 0;
        var i = 0;
        while (i < target.Length && target[i] <= ' ') i++;
        for (; i < target.Length; i++)
        {
            var c = target[i];
            if (c is '\t' or '\n' or '\r')
                continue;
            if (c == ':')
            {
                ReadOnlySpan<char> name = scheme[..length];
                return name.Equals("javascript", StringComparison.OrdinalIgnoreCase)
                    || name.Equals("vbscript", StringComparison.OrdinalIgnoreCase)
                    || name.Equals("data", StringComparison.OrdinalIgnoreCase);
            }
            // Only ASCII letters can spell these names; anything else before the colon (a digit, '/', '.', a
            // look-alike letter) means the target is not one of them.
            if (!char.IsAsciiLetter(c) || length == scheme.Length)
                return false;
            scheme[length++] = c;
        }
        return false;
    }
}
