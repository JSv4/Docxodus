// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace Docxodus;

/// <summary>
/// Open-time repair of package shapes the Open XML stack rejects. It runs once, in
/// the shared reading funnel (<see cref="StrictOoxmlNormalizer.NormalizeForReading"/>), so every
/// pipeline that opens a package sees the repaired form; a package needing no repair is returned
/// as the same array.
/// </summary>
internal static partial class PackageOpenRepair
{
    private static readonly XNamespace PackageRelationships =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    // An external relationship target System.Uri cannot parse — a multi-recipient
    // "mailto:a@x;b@y" is the wild-caught shape (issue #1388) — is data the library never resolves,
    // but System.IO.Packaging parses every target into a Uri and throws on it. The SDK's own
    // workaround swaps in a random rewritten:// URI that is put back only when the package that
    // swapped it is disposed, so any output taken before that, or any relationship copied into
    // another package, carries the placeholder. Instead the target travels under this scheme with
    // its spelling percent-escaped: parseable, deterministic for both sides of a comparison, and
    // undone by DecodeTarget wherever a target leaves the package model — ZipPackageOutputNormalizer
    // for package bytes, TargetText for readers. Writers take a caller's target through TargetUri.
    private const string VerbatimScheme = "x-docxodus-verbatim:";
    private static readonly byte[] VerbatimSchemeBytes = Encoding.ASCII.GetBytes(VerbatimScheme);

    /// <summary>Returns <paramref name="packageBytes"/> repaired, or the same array when it needs no
    /// repair. A damaged ZIP throws <see cref="InvalidDataException"/> for the caller to handle.</summary>
    internal static byte[] Repair(byte[] packageBytes)
    {
        // The common package needs nothing, so it is first checked from scans that build no tree. A
        // package that needs a repair is planned again from its parsed parts, so a scan that misreads
        // can cost a parse or miss a repair but never misplace an edit.
        var repaired = Repairs(packageBytes, exact: false) is null ? null : Repairs(packageBytes, exact: true);
        if (repaired is null)
            return packageBytes;

        using var ms = new MemoryStream();
        ms.Write(packageBytes, 0, packageBytes.Length);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var (name, rels) in repaired)
            {
                using var stream = zip.GetEntry(name)!.Open();
                stream.SetLength(0);
                // As UTF-8 whatever the part declared, which is what the output side reads back.
                using var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false) });
                rels.Save(writer);
            }
        }
        return ms.ToArray();
    }

    /// <summary>The repaired relationships parts. Without <paramref name="exact"/> this only answers
    /// whether anything needs repair: null when nothing does, otherwise an empty map.</summary>
    private static Dictionary<string, XDocument>? Repairs(byte[] packageBytes, bool exact)
    {
        Dictionary<string, XDocument>? repaired = null;
        using var source = new ZipArchive(new MemoryStream(packageBytes, writable: false), ZipArchiveMode.Read);
        foreach (var entry in source.Entries)
        {
            if (!entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                continue;
            string xml;
            using (var reader = new StreamReader(entry.Open()))
                xml = reader.ReadToEnd();
            var tree = exact ? Parse(xml) : null;
            var relationships = tree is null ? (exact ? null : Scan(xml)) : Read(tree);
            if (relationships is null && !exact && (tree = Parse(xml)) is not null)
                relationships = Read(tree); // a part the scan cannot read is checked from its tree
            if (relationships is null)
                continue; // left for the open that follows to report
            var edits = Plan(relationships);
            if (edits.Count == 0)
                continue;
            repaired ??= new Dictionary<string, XDocument>();
            if (!exact)
                return repaired;
            Apply(tree!, edits);
            repaired[entry.FullName] = tree!;
        }
        return repaired;
    }

    /// <summary>One relationship as the repair plans over it, in document order.</summary>
    private sealed record Relationship(string? Id, string? Type, string? Target, bool External);

    /// <summary>A change to the relationship at <paramref name="Index"/>.</summary>
    private sealed record Edit(int Index, string? Target = null);

    private static List<Edit> Plan(List<Relationship> relationships)
    {
        var edits = new List<Edit>();
        for (var i = 0; i < relationships.Count; i++)
        {
            if (relationships[i] is { External: true, Target: { } target } && EncodeTarget(target) is var encoded
                && encoded != target)
                edits.Add(new Edit(i, Target: encoded));
        }
        return edits;
    }

    private static void Apply(XDocument rels, List<Edit> edits)
    {
        var elements = rels.Root!.Elements(PackageRelationships + "Relationship").ToList();
        foreach (var edit in edits)
        {
            if (edit.Target is { } target)
                elements[edit.Index].SetAttributeValue("Target", target);
        }
    }

    private static XDocument? Parse(string xml)
    {
        try
        {
            var tree = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            return tree.Root?.Name == PackageRelationships + "Relationships" ? tree : null;
        }
        catch (XmlException)
        {
            return null;
        }
    }

    private static List<Relationship> Read(XDocument rels) =>
        rels.Root!.Elements(PackageRelationships + "Relationship").Select(r => new Relationship(
            (string?)r.Attribute("Id"), (string?)r.Attribute("Type"), (string?)r.Attribute("Target"),
            (string?)r.Attribute("TargetMode") == "External")).ToList();

    // Relationships parts are flat, so a scan reads them exactly unless the part uses character
    // references, comments, a DTD, CDATA or a prefixed element name; for those it returns null.
    private static List<Relationship>? Scan(string xml)
    {
        if (xml.Contains('&') || xml.Contains("<!", StringComparison.Ordinal)
            || xml.Contains(":Relationship", StringComparison.Ordinal)
            || !xml.Contains("<Relationships xmlns=\"" + PackageRelationships.NamespaceName + "\"", StringComparison.Ordinal))
            return null;
        var relationships = new List<Relationship>();
        for (var start = xml.IndexOf("<Relationship", StringComparison.Ordinal); start >= 0;
             start = xml.IndexOf("<Relationship", start + 1, StringComparison.Ordinal))
        {
            if (start + 13 >= xml.Length || !char.IsWhiteSpace(xml[start + 13]))
                continue; // <Relationships, or an element with no attributes
            var end = xml.IndexOf('>', start);
            if (end < 0)
                return null;
            var element = xml.AsSpan(start, end - start);
            relationships.Add(new Relationship(
                Attribute(element, "Id"), Attribute(element, "Type"), Attribute(element, "Target"),
                Attribute(element, "TargetMode") == "External"));
        }
        return relationships;
    }

    private static string? Attribute(ReadOnlySpan<char> element, string name)
    {
        for (var at = element.IndexOf(name); at >= 0; at = NextIndex(element, name, at))
        {
            var rest = element[(at + name.Length)..].TrimStart();
            if (!char.IsWhiteSpace(element[at - 1]) || rest.Length < 2 || rest[0] != '=')
                continue;
            rest = rest[1..].TrimStart();
            if (rest.Length == 0 || (rest[0] != '"' && rest[0] != '\''))
                return null;
            var close = rest[1..].IndexOf(rest[0]);
            return close < 0 ? null : rest.Slice(1, close).ToString();
        }
        return null;
    }

    private static int NextIndex(ReadOnlySpan<char> element, string name, int at)
    {
        var next = element[(at + 1)..].IndexOf(name);
        return next < 0 ? -1 : at + 1 + next;
    }

    /// <summary>An external target as the package model carries it: unchanged when System.Uri parses
    /// it (the SDK's own test), otherwise encoded. A target under the scheme is encoded too unless it
    /// is already exactly what encoding writes, which keeps reading idempotent: bytes the funnel
    /// produced read back unchanged. Only a target spelled as such an encoding is not its own text.</summary>
    internal static string EncodeTarget(string target) =>
        target.Length == 0 || (Uri.TryCreate(target, UriKind.RelativeOrAbsolute, out _)
            && (!target.StartsWith(VerbatimScheme, StringComparison.Ordinal) || IsEncoding(target)))
            ? target
            : VerbatimScheme + Uri.EscapeDataString(target);

    // An empty target is left for System.IO.Packaging to reject, as it always has, so an empty payload
    // is never an encoding.
    private static bool IsEncoding(string target)
    {
        var payload = target[VerbatimScheme.Length..];
        return payload.Length > 0 && Uri.EscapeDataString(Uri.UnescapeDataString(payload)) == payload;
    }

    /// <summary>Whether XML can carry <paramref name="text"/>; a target that it cannot is refused.</summary>
    internal static bool IsXmlText(string text)
    {
        for (var i = 0; i < text.Length; i++)
        {
            if (XmlConvert.IsXmlChar(text[i]))
                continue;
            if (i + 1 < text.Length && XmlConvert.IsXmlSurrogatePair(text[i + 1], text[i]))
                i++;
            else
                return false;
        }
        return true;
    }

    /// <summary>The inverse of <see cref="EncodeTarget"/>; any other target comes back unchanged.</summary>
    internal static string DecodeTarget(string target) =>
        target.Length > VerbatimScheme.Length && target.StartsWith(VerbatimScheme, StringComparison.Ordinal)
            ? Uri.UnescapeDataString(target[VerbatimScheme.Length..])
            : target;

    /// <summary>The relationship <see cref="Uri"/> for a target a caller supplies.</summary>
    internal static Uri TargetUri(string target) => new(EncodeTarget(target), UriKind.RelativeOrAbsolute);

    /// <summary>The target a relationship's <see cref="Uri"/> stands for, as the package spells it.
    /// Readers that surface a target use this rather than <c>Uri.ToString()</c>.</summary>
    internal static string TargetText(Uri target) =>
        target.OriginalString.StartsWith(VerbatimScheme, StringComparison.Ordinal)
            ? DecodeTarget(target.OriginalString)
            : target.ToString();

    /// <summary>Restores every encoded target in one <c>.rels</c> entry's bytes to its original
    /// spelling; an entry with none comes back as the same array.</summary>
    internal static byte[] RestoreTargets(byte[] relsBytes)
    {
        if (relsBytes.AsSpan().IndexOf(VerbatimSchemeBytes) < 0)
            return relsBytes;
        var bom = relsBytes.AsSpan().StartsWith(Encoding.UTF8.Preamble) ? Encoding.UTF8.Preamble.Length : 0;
        // An encoded value is plain ASCII; the decoded one is written as XML writes an attribute, so a
        // tab or line break comes back as itself rather than as the space a raw one would read as.
        var xml = EncodedTarget().Replace(Encoding.UTF8.GetString(relsBytes, bom, relsBytes.Length - bom), m =>
            new XAttribute("Target", DecodeTarget(m.Groups[2].Value)).ToString());
        return relsBytes[..bom].Concat(Encoding.UTF8.GetBytes(xml)).ToArray();
    }


    [GeneratedRegex("\\bTarget\\s*=\\s*([\"'])(" + VerbatimScheme + "[A-Za-z0-9%._~-]+)\\1")]
    private static partial Regex EncodedTarget();
}
