// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>
/// The library's one line-break convention: <c>"\n"</c> on every platform, so the same input
/// gives the same bytes on Windows, macOS and Linux (issue #1659).
/// </summary>
/// <remarks>
/// .NET takes its default line break from <see cref="System.Environment.NewLine"/>, which is
/// <c>"\r\n"</c> on Windows. <c>StringBuilder.AppendLine</c>, <c>XNode.ToString</c>,
/// <c>XDocument.Save</c> and <c>XmlWriterSettings</c> (through <c>NewLineChars</c>, which also
/// rewrites every line break inside serialized text), and indented JSON (through
/// <c>NewLine</c>) all inherit it. Left alone, the markdown projection, anchor offsets, HTML,
/// IR snapshots and the hashes over them all differed by OS. <c>Docxodus/BannedSymbols.txt</c>
/// bans those APIs in the library, so every such write goes through this class instead.
/// </remarks>
internal static class LineEndings
{
    /// <summary>The line break the library writes.</summary>
    public const string Lf = "\n";

    /// <summary><see cref="StringBuilder.AppendLine()"/>, ending the line with <see cref="Lf"/>.</summary>
    public static StringBuilder AppendLf(this StringBuilder sb) => sb.Append('\n');

    /// <summary><see cref="StringBuilder.AppendLine(string)"/>, ending the line with <see cref="Lf"/>.</summary>
    public static StringBuilder AppendLf(this StringBuilder sb, string? value) => sb.Append(value).Append('\n');

    /// <summary>
    /// <see cref="StringBuilder.AppendLine(ref StringBuilder.AppendInterpolatedStringHandler)"/>,
    /// ending the line with <see cref="Lf"/>. The handler has already written the text into
    /// <paramref name="sb"/>, so no intermediate string is allocated.
    /// </summary>
    public static StringBuilder AppendLf(
        this StringBuilder sb,
        [InterpolatedStringHandlerArgument(nameof(sb))] ref StringBuilder.AppendInterpolatedStringHandler value) =>
        sb.Append('\n');

    /// <summary>
    /// New <see cref="XmlWriterSettings"/> with <see cref="Lf"/> line breaks, otherwise the .NET
    /// defaults. Library code creates writer settings only through here.
    /// </summary>
    public static XmlWriterSettings CreateXmlWriterSettings(SaveOptions options = SaveOptions.DisableFormatting)
    {
#pragma warning disable RS0030 // The one sanctioned construction: it fixes NewLineChars.
        var settings = new XmlWriterSettings { NewLineChars = Lf };
#pragma warning restore RS0030
        if ((options & SaveOptions.DisableFormatting) == 0) settings.Indent = true;
        if ((options & SaveOptions.OmitDuplicateNamespaces) != 0) settings.NamespaceHandling |= NamespaceHandling.OmitDuplicates;
        return settings;
    }

    /// <summary><see cref="XNode.ToString(SaveOptions)"/> with <see cref="Lf"/> line breaks.</summary>
    public static string ToXmlString(this XNode node, SaveOptions options = SaveOptions.None)
    {
        var settings = CreateXmlWriterSettings(options);
        settings.OmitXmlDeclaration = true;
        if (node is XText) settings.ConformanceLevel = ConformanceLevel.Fragment;
        using var text = new StringWriter(CultureInfo.InvariantCulture);
        using (var writer = XmlWriter.Create(text, settings))
        {
            if (node is XDocument document)
            {
                foreach (var child in document.Nodes()) child.WriteTo(writer);
            }
            else
            {
                node.WriteTo(writer);
            }
        }

        return text.ToString();
    }

    /// <summary><see cref="XDocument.Save(Stream, SaveOptions)"/> with <see cref="Lf"/> line breaks.</summary>
    public static void SaveLf(this XDocument document, Stream stream, SaveOptions options = SaveOptions.None)
    {
        var settings = CreateXmlWriterSettings(options);
        if (!string.IsNullOrEmpty(document.Declaration?.Encoding))
        {
            try
            {
                settings.Encoding = Encoding.GetEncoding(document.Declaration.Encoding);
            }
            catch (System.ArgumentException)
            {
                // XDocument.Save ignores an encoding it does not know; so does this.
            }
        }

        using var writer = XmlWriter.Create(stream, settings);
        document.Save(writer);
    }

    /// <summary><see cref="XDocument.Save(TextWriter, SaveOptions)"/> with <see cref="Lf"/> line breaks.</summary>
    public static void SaveLf(this XDocument document, TextWriter textWriter, SaveOptions options = SaveOptions.None)
    {
        using var writer = XmlWriter.Create(textWriter, CreateXmlWriterSettings(options));
        document.Save(writer);
    }

    /// <summary>Sets JSON indentation, always with <see cref="Lf"/> line breaks.</summary>
    public static JsonSerializerOptions Indent(this JsonSerializerOptions options, bool indented)
    {
#pragma warning disable RS0030 // The one sanctioned setter: it pairs indentation with NewLine.
        options.WriteIndented = indented;
#pragma warning restore RS0030
        options.NewLine = Lf;
        return options;
    }

    /// <summary>Returns <paramref name="options"/> indented or not, always with <see cref="Lf"/> line breaks.</summary>
    public static JsonWriterOptions Indent(this JsonWriterOptions options, bool indented)
    {
#pragma warning disable RS0030 // The one sanctioned setter: it pairs indentation with NewLine.
        options.Indented = indented;
#pragma warning restore RS0030
        options.NewLine = Lf;
        return options;
    }
}
