// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text;

namespace Docxodus.Internal;

/// <summary>
/// Reads where a <c>HYPERLINK</c> field instruction points. The IR and the HTML converter both read it here, so a
/// field's link has one target whichever of them reads it. A field nested in the code is the exception: the HTML
/// converter passes the code with each nested field replaced by its cached result, as Word links (issue #1875),
/// and the IR passes the code as written.
/// </summary>
internal static class HyperlinkInstruction
{
    /// <summary>
    /// The instruction's target in the convention of the <c>w:hyperlink</c> element form: the address as written,
    /// a <c>\l</c> location as <c>#location</c>, or <c>address#location</c> when both are given. Null when the
    /// instruction is not a <c>HYPERLINK</c> or names neither.
    /// </summary>
    /// <remarks>
    /// A switch is an unquoted backslash and the character after it. A switch that takes a value (<c>\l \o \t</c>, and
    /// the general <c>\* \# \@</c>) takes the value glued to it (<c>\l"Target"</c>, <c>\*MERGEFORMAT</c>) or else the
    /// next operand, so a screen tip or a format is never read as the address; <c>\m</c>, <c>\n</c> and <c>\h</c> take
    /// none. A quoted operand, or one starting with two backslashes, is an operand, not a switch. In a quoted operand a
    /// backslash escapes the character after it, as Word writes a path (<c>"\\\\server\\share"</c>) and as
    /// <c>DocxSession</c>'s field tokenizer reads it. An empty value names nothing.
    /// </remarks>
    public static string? Target(string instruction)
    {
        var s = instruction.Trim();
        if (!s.StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase))
            return null;
        string? address = null, location = null;
        char? pending = null;
        int i = "HYPERLINK".Length;
        while (i < s.Length)
        {
            while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
            if (i >= s.Length)
                break;
            if (s[i] == '\\' && i + 1 < s.Length && s[i + 1] != '\\' && !char.IsWhiteSpace(s[i + 1]))
            {
                char name = char.ToLowerInvariant(s[i + 1]);
                bool takesValue = name is 'l' or 'o' or 't' or '*' or '#' or '@';
                i += 2;
                pending = takesValue ? name : null;
                if (i < s.Length && !char.IsWhiteSpace(s[i]))
                {
                    var glued = ReadOperand(s, ref i);
                    if (takesValue) Assign(name, glued);
                    pending = null;
                }
                continue;
            }
            if (s[i] == '\\' && (i + 1 == s.Length || char.IsWhiteSpace(s[i + 1])))
            {
                i++; // a lone backslash names no switch
                continue;
            }
            var operand = ReadOperand(s, ref i);
            if (pending is { } owner) Assign(owner, operand);
            else if (operand.Length > 0) address ??= operand;
            pending = null;
        }
        if (address is null && location is null)
            return null;
        return location is null ? address : (address ?? "") + "#" + location;

        void Assign(char name, string value)
        {
            if (name == 'l' && value.Length > 0) location ??= value;
        }
    }

    /// <summary>The quoted (without its quotes, double or single as the field tokenizer takes them, escapes read) or
    /// bare operand at <paramref name="i"/>, which moves past it.</summary>
    private static string ReadOperand(string s, ref int i)
    {
        if (s[i] is '"' or '\'')
        {
            char quote = s[i];
            var quoted = new StringBuilder();
            for (i++; i < s.Length && s[i] != quote; i++)
            {
                if (s[i] == '\\' && i + 1 < s.Length) i++;
                quoted.Append(s[i]);
            }
            if (i < s.Length) i++;
            return quoted.ToString();
        }
        int end = i;
        while (end < s.Length && !char.IsWhiteSpace(s[end])) end++;
        var bare = s.Substring(i, end - i);
        i = end;
        return bare;
    }
}
