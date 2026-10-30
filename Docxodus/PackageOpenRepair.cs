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
using DocumentFormat.OpenXml.Packaging;

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

    private static readonly XNamespace ContentTypes =
        "http://schemas.openxmlformats.org/package/2006/content-types";

    private const string OfficeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string FontRelationship = OfficeRelationships + "font";
    private const string ImageRelationship = OfficeRelationships + "image";
    private const string ThumbnailRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";
    private const string WordprocessingMl = "application/vnd.openxmlformats-officedocument.wordprocessingml.";

    // The content type the Open XML SDK requires of a part each relationship type reaches, for the
    // part kinds a WordprocessingML package carries whose content type is fixed (the SDK refuses to
    // open the package when it differs). A part a relationship reaches by any other type is not
    // retyped.
    internal static readonly IReadOnlyDictionary<string, string> FixedContentTypes = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        [OfficeRelationships + "styles"] = WordprocessingMl + "styles+xml",
        ["http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects"] = "application/vnd.ms-word.stylesWithEffects+xml",
        [OfficeRelationships + "settings"] = WordprocessingMl + "settings+xml",
        [OfficeRelationships + "webSettings"] = WordprocessingMl + "webSettings+xml",
        [OfficeRelationships + "fontTable"] = WordprocessingMl + "fontTable+xml",
        [OfficeRelationships + "numbering"] = WordprocessingMl + "numbering+xml",
        [OfficeRelationships + "footnotes"] = WordprocessingMl + "footnotes+xml",
        [OfficeRelationships + "endnotes"] = WordprocessingMl + "endnotes+xml",
        [OfficeRelationships + "comments"] = WordprocessingMl + "comments+xml",
        ["http://schemas.microsoft.com/office/2011/relationships/commentsExtended"] = WordprocessingMl + "commentsExtended+xml",
        ["http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"] = WordprocessingMl + "commentsIds+xml",
        ["http://schemas.microsoft.com/office/2011/relationships/people"] = WordprocessingMl + "people+xml",
        [OfficeRelationships + "header"] = WordprocessingMl + "header+xml",
        [OfficeRelationships + "footer"] = WordprocessingMl + "footer+xml",
        [OfficeRelationships + "glossaryDocument"] = WordprocessingMl + "document.glossary+xml",
        [OfficeRelationships + "theme"] = "application/vnd.openxmlformats-officedocument.theme+xml",
        [OfficeRelationships + "customXmlProps"] = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
        [OfficeRelationships + "extended-properties"] = "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        [OfficeRelationships + "custom-properties"] = "application/vnd.openxmlformats-officedocument.custom-properties+xml",
        ["http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"] = "application/vnd.openxmlformats-package.core-properties+xml",
    };

    private static readonly PartTypeInfo[] FontTypes = { FontPartType.FontTtf, FontPartType.FontOdttf, FontPartType.FontData };

    private static readonly PartTypeInfo[] ImageTypes =
    {
        ImagePartType.Bmp, ImagePartType.Gif, ImagePartType.Png, ImagePartType.Tif, ImagePartType.Tiff, ImagePartType.Icon,
        ImagePartType.Pcx, ImagePartType.Jpeg, ImagePartType.Jp2, ImagePartType.Emf, ImagePartType.Wmf, ImagePartType.Svg,
    };

    // Image extensions Word writes that the SDK's part types do not name.
    private static readonly Dictionary<string, string> ImageExtensionAliases = new(StringComparer.OrdinalIgnoreCase)
    {
        [".jpeg"] = ImagePartType.Jpeg.ContentType,
        [".jpe"] = ImagePartType.Jpeg.ContentType,
        [".emz"] = "image/x-emz",
        [".wmz"] = "image/x-wmz",
        [".wdp"] = "image/vnd.ms-photo",
    };

    /// <summary>Returns <paramref name="packageBytes"/> repaired, or the same array when it needs no
    /// repair. A damaged ZIP throws <see cref="InvalidDataException"/> for the caller to handle.</summary>
    /// <remarks>
    /// A repair is made only where the package itself says what the repaired form is; anything else is
    /// left for the open that follows to reject as it always has, so a repair never trades a failure for
    /// lost content. For each relationships part:
    /// <list type="bullet">
    /// <item>An unparseable external target is encoded (see <see cref="VerbatimScheme"/>).</item>
    /// <item>Relationships sharing an Id, when at most one of them is of a kind referenced by Id: that one
    /// keeps the Id (the last one when all are reached by type) and the others get fresh Ids. When two
    /// or more are referenced by Id, nothing says which reference means which, and they are left.</item>
    /// <item>A relationship to a part the zip does not hold is dropped.</item>
    /// <item>A part a relationship reaches with no content type, or with only the generic XML type where
    /// the relationship requires a specific one, gets an Override with the type the relationship implies
    /// (by its type, or for a font or image by the part's extension). A part with no content type and
    /// none implied is left.</item>
    /// </list>
    /// Hyperlink and media relationships are references the SDK does not load as parts, so the last
    /// two never touch them.
    /// </remarks>
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
        var package = new PackageFacts(source, exact);
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
            var edits = Plan(relationships, SourcePartOf(entry.FullName), package);
            if (edits.Count == 0)
                continue;
            repaired ??= new Dictionary<string, XDocument>();
            if (!exact)
                return repaired;
            Apply(tree!, edits);
            repaired[entry.FullName] = tree!;
        }
        if (package.ChangesContentTypes)
        {
            repaired ??= new Dictionary<string, XDocument>();
            if (exact)
                repaired[package.TypesEntryName!] = package.ChangedContentTypes();
        }
        return repaired;
    }

    /// <summary>One relationship as the repair plans over it, in document order.</summary>
    private sealed record Relationship(string? Id, string? Type, string? Target, string? Mode)
    {
        // System.IO.Packaging parses the mode as an enum (surrounding whitespace ignored), so "0" or
        // "Internal, Internal" is internal too; those are left alone here, which is how main treated them.
        public bool External => Mode?.Trim() == "External";

        public bool Internal => Mode is null || Mode.Trim() == "Internal";
    }

    /// <summary>A change to the relationship at <paramref name="Index"/>.</summary>
    private sealed record Edit(int Index, string? Target = null, string? Id = null, bool Remove = false);

    private static List<Edit> Plan(List<Relationship> relationships, string? sourcePart, PackageFacts package)
    {
        var edits = new List<Edit>();
        for (var i = 0; i < relationships.Count; i++)
        {
            if (relationships[i] is { External: true, Target: { } target } && EncodeTarget(target) is var encoded
                && encoded != target)
                edits.Add(new Edit(i, Target: encoded));
        }
        var unresolved = RenumberDuplicateIds(relationships, edits);
        if (sourcePart is not null && package.HasContentTypes)
        {
            for (var i = 0; i < relationships.Count; i++)
            {
                if (!unresolved.Contains(i) && RepairPartRelationship(relationships[i], sourcePart, package))
                    edits.Add(new Edit(i, Remove: true));
            }
        }
        return edits;
    }

    // Single-instance kinds the SDK reaches by relationship type, so their Id names nothing in markup.
    // Header and footer have fixed content types but are referenced by r:id from w:sectPr.
    private static readonly HashSet<string> ReachedByType = new(StringComparer.Ordinal)
    {
        OfficeRelationships + "officeDocument", OfficeRelationships + "styles", OfficeRelationships + "settings",
        OfficeRelationships + "webSettings", OfficeRelationships + "fontTable", OfficeRelationships + "numbering",
        OfficeRelationships + "footnotes", OfficeRelationships + "endnotes", OfficeRelationships + "comments",
        OfficeRelationships + "glossaryDocument", OfficeRelationships + "theme", OfficeRelationships + "customXml",
        OfficeRelationships + "customXmlProps", OfficeRelationships + "extended-properties",
        OfficeRelationships + "custom-properties",
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        ThumbnailRelationship,
        "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects",
        "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
        "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds",
        "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible",
        "http://schemas.microsoft.com/office/2011/relationships/people",
        "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
        "http://schemas.microsoft.com/office/2006/relationships/keyMapCustomizations",
    };

    // The kinds whose relationship may be dropped when its part is absent: the reference left in
    // markup (w:embedRegular, r:embed, the package thumbnail) is known to be harmless on every front
    // door and in LibreOffice. For any other kind the package does not say what the repaired form is.
    private static readonly HashSet<string> DroppableWhenAbsent = new(StringComparer.Ordinal)
    {
        FontRelationship, ImageRelationship, ThumbnailRelationship,
    };

    /// <summary>Plans the renumbering of shared Ids and returns the relationships of every group left
    /// alone, which no other repair may touch either: dropping or retyping one member would bind its
    /// Id to the other without evidence.</summary>
    private static HashSet<int> RenumberDuplicateIds(List<Relationship> relationships, List<Edit> edits)
    {
        var unresolved = new HashSet<int>();
        var used = new HashSet<string>(relationships.Select(r => r.Id ?? ""), StringComparer.Ordinal);
        if (used.Count == relationships.Count)
            return unresolved;
        foreach (var group in relationships.Select((r, i) => (r, i)).Where(x => x.r.Id is not null)
                     .GroupBy(x => x.r.Id!, StringComparer.Ordinal).Where(g => g.Count() > 1))
        {
            var byId = group.Where(x => !ReachedByType.Contains(x.r.Type ?? "")).ToList();
            if (byId.Count > 1)
            {
                unresolved.UnionWith(group.Select(x => x.i));
                continue;
            }
            var keeper = byId.Count == 1 ? byId[0].i : group.Last().i;
            foreach (var (_, i) in group.Where(x => x.i != keeper))
            {
                var n = 1;
                string fresh;
                while (!used.Add(fresh = $"{group.Key}_{n}"))
                    n++;
                edits.Add(new Edit(i, Id: fresh));
            }
        }
        return unresolved;
    }

    /// <summary>Plans the content-type repair for a part relationship; returns true when the
    /// relationship is to be dropped because its part is absent and its kind allows that.</summary>
    private static bool RepairPartRelationship(Relationship relationship, string sourcePart, PackageFacts package)
    {
        if (!relationship.Internal || relationship.Target is not { } target
            || ResolvePart(sourcePart, target) is not { } part)
            return false;
        var type = relationship.Type ?? "";
        if (package.IsAmbiguous(part))
            return false;
        if (package.EntryFor(part) is not { } entry)
        {
            if (!DroppableWhenAbsent.Contains(type))
                return false;
            package.RemoveOverride(part);
            return true;
        }
        var declared = package.ContentTypeOf(entry);
        var fixedType = FixedContentTypes.GetValueOrDefault(type);
        if ((declared is null || (fixedType is not null && IsGenericXml(declared)))
            && (fixedType ?? ContentTypeByExtension(type, entry)) is { } implied)
            package.AddOverride(entry, implied);
        return false;
    }

    private static void Apply(XDocument rels, List<Edit> edits)
    {
        var elements = rels.Root!.Elements(PackageRelationships + "Relationship").ToList();
        foreach (var edit in edits)
        {
            if (edit.Target is { } target)
                elements[edit.Index].SetAttributeValue("Target", target);
            if (edit.Id is { } id)
                elements[edit.Index].SetAttributeValue("Id", id);
            if (edit.Remove)
                elements[edit.Index].Remove();
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
            (string?)r.Attribute("TargetMode"))).ToList();

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
                Attribute(element, "TargetMode")));
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

    // "word/_rels/document.xml.rels" belongs to /word/document.xml; "_rels/.rels" to the package.
    private static string? SourcePartOf(string relsEntry)
    {
        var slash = relsEntry.LastIndexOf('/');
        var folder = relsEntry[..(slash + 1)];
        if (!folder.EndsWith("_rels/", StringComparison.OrdinalIgnoreCase))
            return null;
        return "/" + folder[..^"_rels/".Length] + relsEntry[(slash + 1)..^".rels".Length];
    }

    // Resolved exactly as System.IO.Packaging and the SDK resolve a relationship target, dot segments,
    // fragments, queries, backslashes and surrounding spaces included.
    private static string? ResolvePart(string sourcePart, string target)
    {
        try
        {
            // A result that is not itself a valid part name ("media/", ".") names no part.
            return PartName(System.IO.Packaging.PackUriHelper.ResolvePartUri(
                new Uri(sourcePart, UriKind.Relative), new Uri(target, UriKind.RelativeOrAbsolute)).OriginalString);
        }
        catch (Exception e) when (e is ArgumentException or UriFormatException or InvalidOperationException)
        {
            return null; // a target that cannot be resolved keeps the diagnostics it had
        }
    }

    /// <summary>A part name as System.IO.Packaging forms it, or null when it forms none.</summary>
    private static string? PartName(string name)
    {
        try
        {
            return System.IO.Packaging.PackUriHelper.CreatePartUri(new Uri(name, UriKind.Relative)).OriginalString;
        }
        catch (Exception e) when (e is ArgumentException or UriFormatException)
        {
            return null;
        }
    }

    private static string? ContentTypeByExtension(string relationshipType, string part)
    {
        var extension = Path.GetExtension(part);
        var byExtension = relationshipType == FontRelationship ? FontTypes
            : relationshipType is ImageRelationship or ThumbnailRelationship ? ImageTypes
            : Array.Empty<PartTypeInfo>();
        foreach (var type in byExtension)
        {
            if (string.Equals(type.Extension, extension, StringComparison.OrdinalIgnoreCase))
                return type.ContentType;
        }
        return relationshipType is ImageRelationship or ThumbnailRelationship ? ImageExtensionAliases.GetValueOrDefault(extension) : null;
    }

    private static bool IsGenericXml(string contentType) =>
        contentType.Equals("application/xml", StringComparison.OrdinalIgnoreCase)
        || contentType.Equals("text/xml", StringComparison.OrdinalIgnoreCase);

    /// <summary>What the repair needs to know of the package: which parts the zip holds, which
    /// content types <c>[Content_Types].xml</c> declares, and the Overrides the plan adds.</summary>
    private sealed class PackageFacts
    {
        // The part names the zip entries hold, as System.IO.Packaging names them: an interleaved part
        // stored as name/[0].piece ... name/[n].last.piece is one part, and a target finds its part
        // without regard to case. Each maps to the entry's own spelling, which is what an Override must
        // match: the packaging layer compares an Override's PartName with it case and all (escapes aside).
        private readonly Dictionary<string, string> _parts = new(StringComparer.OrdinalIgnoreCase);
        // Names more than one entry forms (case variants, or a whole entry beside pieces): which entry
        // System.IO.Packaging takes depends on their types and order, so the repair leaves them alone.
        private readonly HashSet<string> _ambiguous = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _defaults = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _overrides = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _overrideSpellings = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _added = new(StringComparer.Ordinal);
        private readonly HashSet<string> _removed = new(StringComparer.OrdinalIgnoreCase);
        private readonly string? _typesXml;

        public PackageFacts(ZipArchive source, bool exact)
        {
            foreach (var entry in source.Entries)
            {
                if (PartName("/" + WithoutPieceSuffix(entry.FullName)) is { } part)
                {
                    if (!_parts.TryAdd(part, part))
                        _ambiguous.Add(part);
                }
            }
            var types = source.Entries.FirstOrDefault(e =>
                string.Equals(e.FullName, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase));
            if (types is null)
                return;
            using (var reader = new StreamReader(types.Open()))
                _typesXml = reader.ReadToEnd();
            TypesEntryName = types.FullName;
            HasContentTypes = ReadContentTypes(_typesXml, exact);
            foreach (var (spelling, contentType) in _overrides.ToList())
            {
                var part = OverrideKey(spelling);
                _overrides.Remove(spelling);
                _overrides.TryAdd(part, contentType);
                _overrideSpellings.TryAdd(part, spelling);
            }
        }

        private static string OverrideKey(string name) =>
            EscapeHex().Replace(PartName(name) ?? name, m => m.Value.ToUpperInvariant());

        /// <summary>The part name, as its zip entry spells it, of the entry holding <paramref name="part"/>
        /// (typed or not), or null when no entry does.</summary>
        public string? EntryFor(string part) => _parts.GetValueOrDefault(part);

        public bool IsAmbiguous(string part) => _ambiguous.Contains(part);


        /// <summary>False without a readable content-types part: then nothing tells a typed part from
        /// an untyped one, and no part relationship is repaired.</summary>
        public bool HasContentTypes { get; }

        public string? TypesEntryName { get; }

        private static void ScanDeclarations(string xml, string tag, string key, Dictionary<string, string> into)
        {
            for (var start = xml.IndexOf(tag, StringComparison.Ordinal); start >= 0;
                 start = xml.IndexOf(tag, start + 1, StringComparison.Ordinal))
            {
                if (start + tag.Length >= xml.Length || !char.IsWhiteSpace(xml[start + tag.Length]))
                    continue;
                var end = xml.IndexOf('>', start);
                var element = xml.AsSpan(start, (end < 0 ? xml.Length : end) - start);
                if (Attribute(element, key) is { } name && Attribute(element, "ContentType") is { } contentType)
                    into.TryAdd(name, contentType);
            }
        }

        public string? ContentTypeOf(string entry) =>
            _added.GetValueOrDefault(entry) ?? _overrides.GetValueOrDefault(OverrideKey(entry))
            ?? _defaults.GetValueOrDefault(Path.GetExtension(entry).TrimStart('.'));

        public void AddOverride(string entry, string contentType) => _added[entry] = contentType;

        /// <summary>Drops the Override of a part the package lacks, which would otherwise outlive it.</summary>
        public void RemoveOverride(string part)
        {
            var key = OverrideKey(part);
            foreach (var (name, spelling) in _overrideSpellings)
            {
                if (string.Equals(name, key, StringComparison.OrdinalIgnoreCase))
                    _removed.Add(spelling);
            }
        }

        public bool ChangesContentTypes => _added.Count > 0 || _removed.Count > 0;

        /// <summary>The content-types part with the planned changes; only called when the exact read
        /// of the part succeeded.</summary>
        public XDocument ChangedContentTypes()
        {
            var types = XDocument.Parse(_typesXml!, LoadOptions.PreserveWhitespace);
            types.Root!.Elements(ContentTypes + "Override")
                .Where(o => (string?)o.Attribute("PartName") is { } name && _removed.Contains(name)).Remove();
            foreach (var (part, contentType) in _added)
            {
                // An Override naming the part in another case is the same name to the packaging layer, so it
                // is rewritten under the entry's spelling rather than joined by a second one.
                var existing = types.Root!.Elements(ContentTypes + "Override").FirstOrDefault(o =>
                    (string?)o.Attribute("PartName") is { } name
                    && string.Equals(OverrideKey(name), OverrideKey(part), StringComparison.OrdinalIgnoreCase));
                if (existing is not null)
                {
                    existing.SetAttributeValue("PartName", part);
                    existing.SetAttributeValue("ContentType", contentType);
                }
                else
                    types.Root.Add(new XElement(ContentTypes + "Override",
                        new XAttribute("PartName", part), new XAttribute("ContentType", contentType)));
            }
            return types;
        }

        private bool ReadContentTypes(string xml, bool exact)
        {
            if (!exact && !xml.Contains('&') && !xml.Contains("<!", StringComparison.Ordinal)
                && xml.Contains("<Types xmlns=\"" + ContentTypes.NamespaceName + "\"", StringComparison.Ordinal))
            {
                ScanDeclarations(xml, "<Default", "Extension", _defaults);
                ScanDeclarations(xml, "<Override", "PartName", _overrides);
                return true;
            }
            XDocument types;
            try
            {
                types = XDocument.Parse(xml);
            }
            catch (XmlException)
            {
                return false;
            }
            if (types.Root?.Name != ContentTypes + "Types")
                return false;
            foreach (var element in types.Root.Elements())
            {
                var contentType = (string?)element.Attribute("ContentType");
                if (contentType is null)
                    continue;
                if (element.Name == ContentTypes + "Default" && (string?)element.Attribute("Extension") is { } extension)
                    _defaults.TryAdd(extension, contentType);
                else if (element.Name == ContentTypes + "Override" && (string?)element.Attribute("PartName") is { } part)
                    _overrides.TryAdd(part, contentType);
            }
            return true;
        }
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


    // System.IO.Packaging compares the hex digits of an escaped non-ASCII byte without case, and an
    // escaped ASCII character with it.
    [GeneratedRegex("%[89a-f][0-9a-f]", RegexOptions.IgnoreCase)]
    private static partial Regex EscapeHex();

    /// <summary>The name of the part an interleaved piece entry (<c>name/[n].piece</c>,
    /// <c>name/[n].last.piece</c>) belongs to, as System.IO.Packaging assembles it; any other entry
    /// name unchanged.</summary>
    internal static string WithoutPieceSuffix(string entryName) => InterleavedPiece().Replace(entryName, "");

    [GeneratedRegex("/\\[\\d+\\](\\.last)?\\.piece$", RegexOptions.IgnoreCase)]
    private static partial Regex InterleavedPiece();

    [GeneratedRegex("\\bTarget\\s*=\\s*([\"'])(" + VerbatimScheme + "[A-Za-z0-9%._~-]+)\\1")]
    private static partial Regex EncodedTarget();
}
