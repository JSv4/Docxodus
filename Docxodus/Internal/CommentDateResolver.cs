using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>Resolve modern comment timestamps once per comment-set read.</summary>
internal sealed class CommentDateResolver
{
    private static readonly XNamespace Cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
    private static readonly XNamespace Cex = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
    private readonly Dictionary<string, XElement> _ids;
    private readonly Dictionary<string, XElement> _dates;

    internal CommentDateResolver(MainDocumentPart main)
    {
        var ids = UniqueEntries(main.WordprocessingCommentsIdsPart?.GetXDocument().Root,
            Cid + "commentId", Cid + "paraId");
        _ids = ids.Values.Where(entry => !string.IsNullOrEmpty((string?)entry.Attribute(Cid + "durableId")))
            .GroupBy(entry => (string)entry.Attribute(Cid + "durableId")!, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() == 1).Select(group => group.Single())
            .ToDictionary(entry => (string)entry.Attribute(Cid + "paraId")!, StringComparer.OrdinalIgnoreCase);
        _dates = UniqueEntries(main.WordCommentsExtensiblePart?.GetXDocument().Root,
            Cex + "commentExtensible", Cex + "durableId");
    }

    internal string? Resolve(XElement comment)
    {
        // MS-DOCX CT_CommentId keys the last paragraph, not the first or the w:id.
        var paraId = (string?)comment.Elements(W.p).LastOrDefault()?.Attribute(W14.paraId);
        if (paraId != null && _ids.TryGetValue(paraId, out var id))
        {
            var durableId = (string?)id.Attribute(Cid + "durableId");
            if (durableId != null && _dates.TryGetValue(durableId, out var entry) &&
                TryParseUtc((string?)entry.Attribute(Cex + "dateUtc"), out var utc))
                return utc.ToString("yyyy-MM-dd'T'HH:mm:ss.FFFFFFF'Z'", CultureInfo.InvariantCulture);
        }
        return (string?)comment.Attribute(W.date);
    }

    internal static bool TryParseUtc(string? value, out DateTimeOffset utc)
    {
        utc = default;
        if (string.IsNullOrEmpty(value)) return false;
        var timeStart = value.IndexOf('T');
        if (timeStart < 0) return false;
        // dateUtc declares UTC even when its XML dateTime omits the optional zone.
        var hasZone = value.EndsWith("Z", StringComparison.Ordinal) ||
            value.IndexOf('+', timeStart) >= 0 || value.IndexOf('-', timeStart) >= 0;
        try
        {
            utc = XmlConvert.ToDateTimeOffset(hasZone ? value : value + "Z").ToUniversalTime();
            return true;
        }
        catch (FormatException)
        {
            return false;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }

    private static Dictionary<string, XElement> UniqueEntries(XElement? root, XName element, XName key) =>
        (root?.Elements(element) ?? Enumerable.Empty<XElement>())
            .Where(entry => !string.IsNullOrEmpty((string?)entry.Attribute(key)))
            .GroupBy(entry => (string)entry.Attribute(key)!, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() == 1)
            .ToDictionary(group => group.Key, group => group.Single(), StringComparer.OrdinalIgnoreCase);
}
