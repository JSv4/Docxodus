// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Makes every bookmark id in a document unique, keeping each <c>w:bookmarkStart</c> paired with its
/// <c>w:bookmarkEnd</c> (issue #840).
/// <para>A comparison keeps content from both documents — a deleted original paragraph beside the inserted
/// revised table — and both documents number their bookmarks from 0, so their ids collide. The id is the only
/// link from an end to its start, so the pairing has to be worked out before an id can change. A start and an
/// end cloned from the same document share its revision side: the original's markers sit in deleted content,
/// the revised document's in inserted content, and a bookmark both kept is bare. Each end therefore closes the
/// earliest open start of its id on its own side; failing that, one it can share with a bookmark both kept (an
/// unchanged start, or any start when the end is unchanged), so a deleted end never closes an inserted start;
/// and failing that, the earliest open start of its id.</para>
/// <para>Of the bookmarks that share an id, one keeps it and the others take fresh ids above every id in the
/// document; a bookmark the comparison left untouched keeps its id first. ECMA-376 makes an id unique per document,
/// so every story shares one id space, but two untouched bookmarks in different stories already
/// shared their id in the source document and are left alone. A bookmark inside math or a drawing is renumbered
/// like any other: the IR leaves bookmark ids out of that content's hash (<see cref="Ir.IrHasher"/>).</para>
/// <para>An end whose start is gone marks nothing and is dropped.</para>
/// </summary>
internal static class BookmarkIds
{
    private enum Side
    {
        Unchanged,
        Deleted,
        Inserted,
    }

    private sealed record Bookmark(XElement? Start, XElement? End)
    {
        public IEnumerable<XElement> Markers => new[] { Start, End }.OfType<XElement>();
    }

    /// <summary>Renumber the bookmarks of every story of <paramref name="main"/> (body, headers, footers,
    /// footnotes, endnotes, comments) whose id another bookmark already uses, and write back each part that
    /// changed. Returns whether any did.</summary>
    internal static bool MakeUnique(MainDocumentPart main)
    {
        var parts = new List<OpenXmlPart> { main };
        parts.AddRange(main.HeaderParts);
        parts.AddRange(main.FooterParts);
        parts.AddRange(new OpenXmlPart?[] { main.FootnotesPart, main.EndnotesPart, main.WordprocessingCommentsPart }
            .OfType<OpenXmlPart>());

        var changed = MakeUnique(parts.Select(part => part.GetXDocument().Root).OfType<XElement>().ToList());
        foreach (var part in parts.Where(part => part.GetXDocument().Root is { } root && changed.Contains(root)))
            part.PutXDocument();
        return changed.Count > 0;
    }

    /// <summary>Renumber the bookmarks under <paramref name="storyRoots"/>, which share one id space; a bookmark
    /// pairs only within its own root. Returns the roots that changed.</summary>
    internal static HashSet<XElement> MakeUnique(IReadOnlyList<XElement> storyRoots)
    {
        var changed = new HashSet<XElement>();
        var bookmarks = new List<(XElement Root, Bookmark Bookmark)>();
        foreach (var root in storyRoots)
        {
            var paired = Pair(root);
            var unclosed = paired.Where(b => b.End is null).Select(IdOf).ToHashSet();
            foreach (var bookmark in paired)
            {
                // An end whose start is gone (a renderer dropped the start's content) marks nothing, and its id
                // would still collide. An end is kept while a start of its id is still open, as when a source
                // wrote the end first.
                if (bookmark.Start is null && !unclosed.Contains(IdOf(bookmark)))
                {
                    var wrapper = bookmark.End!.Parent;
                    bookmark.End.Remove();
                    if (wrapper is { Parent: not null, HasElements: false } && (wrapper.Name == W.ins || wrapper.Name == W.del))
                        wrapper.Remove();
                    changed.Add(root);
                    continue;
                }
                bookmarks.Add((root, bookmark));
            }
        }

        var duplicated = bookmarks.GroupBy(b => IdOf(b.Bookmark)).Where(g => g.Count() > 1).ToList();
        if (duplicated.Count == 0)
            return changed;

        int next = bookmarks.Select(b => int.TryParse(IdOf(b.Bookmark), out var id) ? id : 0)
            .DefaultIfEmpty(0).Max() + 1;
        foreach (var group in duplicated)
        {
            // Bookmarks the comparison left untouched keep their ids first. Two of those in different stories
            // already shared the id in the source document, so they stay as they are; any other two collide.
            var kept = new List<(XElement Root, bool Unchanged)>();
            foreach (var (root, bookmark) in group.OrderBy(b => IsUnchanged(b.Bookmark) ? 0 : 1))
            {
                var unchanged = IsUnchanged(bookmark);
                if (!kept.Any(k => k.Root == root || !(k.Unchanged && unchanged)))
                {
                    kept.Add((root, unchanged));
                    continue;
                }

                var fresh = (next++).ToString(System.Globalization.CultureInfo.InvariantCulture);
                foreach (var marker in bookmark.Markers)
                    marker.SetAttributeValue(W.id, fresh);
                changed.Add(root);
            }
        }
        return changed;
    }

    /// <summary>The bookmarks under <paramref name="root"/> in document order of their first marker; a start
    /// without an end, or an end without a start, is a bookmark of its own.</summary>
    private static List<Bookmark> Pair(XElement root)
    {
        var bookmarks = new List<Bookmark>();
        var open = new Dictionary<string, List<(XElement Start, int Index)>>();
        foreach (var marker in root.Descendants().Where(e => e.Name == W.bookmarkStart || e.Name == W.bookmarkEnd))
        {
            var id = (string?)marker.Attribute(W.id) ?? string.Empty;
            if (marker.Name == W.bookmarkStart)
            {
                if (!open.TryGetValue(id, out var starts))
                    open[id] = starts = new List<(XElement, int)>();
                starts.Add((marker, bookmarks.Count));
                bookmarks.Add(new Bookmark(marker, null));
                continue;
            }

            if (open.TryGetValue(id, out var candidates) && candidates.Count > 0)
            {
                var side = SideOf(marker);
                var match = candidates.FindIndex(c => SideOf(c.Start) == side);
                // A bookmark both documents kept can be split: its start in unchanged content, its end in
                // content one side changed. It still never pairs a deleted end with an inserted start.
                if (match < 0)
                    match = candidates.FindIndex(c => SideOf(c.Start) == Side.Unchanged || side == Side.Unchanged);
                if (match < 0)
                    match = 0;
                var (start, index) = candidates[match];
                candidates.RemoveAt(match);
                bookmarks[index] = new Bookmark(start, marker);
            }
            else
            {
                bookmarks.Add(new Bookmark(null, marker));
            }
        }
        return bookmarks;
    }

    private static string IdOf(Bookmark bookmark) => IdOf(bookmark.Markers.First());

    private static string IdOf(XElement marker) => (string?)marker.Attribute(W.id) ?? string.Empty;

    private static bool IsUnchanged(Bookmark bookmark) => bookmark.Markers.All(m => SideOf(m) == Side.Unchanged);

    /// <summary>The revision side of the innermost tracked container of <paramref name="marker"/>: a run-level
    /// wrapper, or a paragraph mark, row or cell marked as inserted or deleted.</summary>
    private static Side SideOf(XElement marker)
    {
        foreach (var ancestor in marker.Ancestors())
        {
            var name = ancestor.Name;
            if (name == W.del || name == W.moveFrom)
                return Side.Deleted;
            if (name == W.ins || name == W.moveTo)
                return Side.Inserted;
            var mark = name == W.p ? ancestor.Element(W.pPr)?.Element(W.rPr)
                : name == W.tr ? ancestor.Element(W.trPr)
                : null;
            if (mark?.Element(W.del) is not null || mark?.Element(W.moveFrom) is not null)
                return Side.Deleted;
            if (mark?.Element(W.ins) is not null || mark?.Element(W.moveTo) is not null)
                return Side.Inserted;
            if (name == W.tc && ancestor.Element(W.tcPr) is { } cell)
            {
                if (cell.Element(W.cellDel) is not null)
                    return Side.Deleted;
                if (cell.Element(W.cellIns) is not null)
                    return Side.Inserted;
            }
        }
        return Side.Unchanged;
    }
}
