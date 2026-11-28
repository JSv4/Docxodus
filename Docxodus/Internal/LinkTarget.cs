// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text;

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
        var start = 0;
        var end = target.Length;
        while (start < end && target[start] <= ' ') start++;
        while (end > start && target[end - 1] <= ' ') end--;
        var scheme = new StringBuilder();
        for (var i = start; i < end; i++)
        {
            var c = target[i];
            if (c is '\t' or '\n' or '\r')
                continue;
            if (c == ':')
                return scheme.ToString() is var name
                    && (name.Equals("javascript", StringComparison.OrdinalIgnoreCase)
                        || name.Equals("vbscript", StringComparison.OrdinalIgnoreCase)
                        || name.Equals("data", StringComparison.OrdinalIgnoreCase));
            // A scheme is a letter, then letters, digits, '+', '-' or '.'; anything else makes the target relative.
            if (!(char.IsAsciiLetter(c) || (scheme.Length > 0 && (char.IsAsciiDigit(c) || c is '+' or '-' or '.'))))
                return false;
            scheme.Append(c);
        }
        return false;
    }
}
