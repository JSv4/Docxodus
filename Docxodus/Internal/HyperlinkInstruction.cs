// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

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

    /// <summary>
    /// The target of a code made of pieces, each the field's own code or (<c>Result</c>) a nested field's cached result,
    /// as <see cref="FieldRetriever.EvaluatedInstrText"/> reads them. A result is read literally, as one stretch of the
    /// operand it is part of: its spaces, quotes and backslashes are text, not the code's syntax, so a merged
    /// <c>C:\Docs\a.pdf</c> in <c>"{ MERGEFIELD Path }"</c> is the address <c>C:\Docs\a.pdf</c> (issue #1875). An
    /// empty target names nothing.
    /// </summary>
    public static string? Target(IReadOnlyList<(string Text, bool Result)> pieces)
    {
        // Each result stands in the code as a placeholder that no tokenizer rule touches, and is put back in the target.
        var code = new StringBuilder();
        var results = new List<string>();
        foreach (var (text, result) in pieces)
        {
            if (!result)
            {
                code.Append(text);
                continue;
            }
            code.Append(PlaceholderStart).Append(results.Count.ToString(CultureInfo.InvariantCulture)).Append(PlaceholderEnd);
            results.Add(text);
        }
        var target = Target(code.ToString());
        if (target is null || results.Count == 0)
            return target;
        target = Placeholder.Replace(target, m => int.TryParse(m.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture,
            out var index) && index < results.Count ? results[index] : m.Value);
        return target.Length == 0 ? null : target;
    }

    private const char PlaceholderStart = '\uE000', PlaceholderEnd = '\uE001';

    private static readonly Regex Placeholder = new("\uE000([0-9]+)\uE001", RegexOptions.CultureInvariant);

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
