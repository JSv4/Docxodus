using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.Ir.Diff;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

public sealed partial class DocxSession
{
    private sealed record RetainedControl(XElement Element, int Start, int Length);

    // Markdown cannot carry an SDT's binding, citation instructions or other metadata. Match
    // unchanged displayed characters/formatting, then keep the native envelope at its new position.
    private List<RetainedControl>? RetainedInlineControls(XElement paragraph, XElement replacement)
    {
        var candidates = paragraph.Descendants(W.sdt)
            .Where(e => ReferenceEquals(e.Ancestors(W.p).FirstOrDefault(), paragraph)
                && !IsInsideNestedRemoval(e, paragraph))
            .Select(e => e.Ancestors(W.hyperlink).LastOrDefault(link => InlineChildTextLength(link) == 0) ?? e)
            .Distinct().ToList();
        if (candidates.Count == 0) return new();
        var owner = OwnedPartRelationships.FindOwner(_doc!, paragraph)?.Part;
        var left = RunTextMap.Build(paragraph);
        var right = RunTextMap.Build(replacement);
        var ranges = candidates.Select(control =>
        {
            var runs = left.Segments.Where(s => s.Run.Ancestors().Contains(control)).ToList();
            int start = runs.Count == 0 ? MarkerOffset(paragraph, control) : runs[0].StartOffsetInBlock;
            return new RetainedControl(control, start, runs.Count == 0 ? 0 : runs[^1].EndOffsetInBlock - start);
        }).ToList();
        var leftKeys = RewriteKeys(left, owner);
        var rightKeys = RewriteKeys(right, owner);
        var positions = Enumerable.Repeat(-1, left.FlatText.Length).ToArray();
        void Align(int ls, int le, int rs, int re)
        {
            foreach (var (a, b) in IrTokenDiffer.BoundedAnchors(leftKeys[ls..le], rightKeys[rs..re],
                Enumerable.Repeat(1, le - ls).ToArray())) positions[ls + a] = rs + b;
        }
        // A long unchanged control region can exceed the bounded LCS without containing a
        // rare character. Anchor its complete, unique displayed sequence before aligning outside it.
        int first = ranges.Min(r => r.Start), last = ranges.Max(r => r.Start + r.Length);
        var region = leftKeys.AsSpan(first, last - first);
        int regionAt = region.Length > 0 && UniqueRewriteMatch(leftKeys, region) == first
            ? UniqueRewriteMatch(rightKeys, region) : -1;
        if (regionAt >= 0)
        {
            Align(0, first, 0, regionAt);
            for (int i = first; i < last; i++) positions[i] = regionAt + i - first;
            Align(last, leftKeys.Length, regionAt + region.Length, rightKeys.Length);
        }
        else Align(0, leftKeys.Length, 0, rightKeys.Length);

        var kept = new List<RetainedControl>();
        foreach (var range in ranges)
        {
            var control = range.Element;
            if (control.Ancestors().Any(a => kept.Any(k => ReferenceEquals(k.Element, a)))) continue;
            int start = range.Start, length = range.Length;
            int at = length > 0 ? positions[start]
                : start > 0 && start == positions.Length ? rightKeys.Length
                : start > 0 && positions[start - 1] >= 0 ? positions[start - 1] + 1
                : positions.Length > 0 && positions[start] >= 0 ? positions[start] : start == 0 ? 0 : -1;
            if (at < 0 || Enumerable.Range(0, length).Any(i => positions[start + i] != at + i))
            {
                if (length == 0) return null;
                var pattern = leftKeys.AsSpan(start, length);
                at = UniqueRewriteMatch(rightKeys, pattern);
                if (at == -1) continue; // The control's displayed content or formatting was replaced.
                if (at < 0 || UniqueRewriteMatch(leftKeys, pattern) != start) return null;
            }
            if (kept.Count > 0 && at < kept[^1].Start + kept[^1].Length) return null;
            kept.Add(new(control, at, length));
        }
        return kept;
    }

    private static int UniqueRewriteMatch(string[] keys, ReadOnlySpan<string> pattern)
    {
        // Prefix matching keeps near-matches in long repetitive controls linear too.
        var prefix = new int[pattern.Length];
        for (int i = 1, j = 0; i < pattern.Length; i++)
        {
            while (j > 0 && pattern[i] != pattern[j]) j = prefix[j - 1];
            if (pattern[i] == pattern[j]) j++;
            prefix[i] = j;
        }
        int found = -1;
        for (int i = 0, j = 0; i < keys.Length; i++)
        {
            while (j > 0 && keys[i] != pattern[j]) j = prefix[j - 1];
            if (keys[i] == pattern[j]) j++;
            if (j != pattern.Length) continue;
            if (found >= 0) return -2;
            found = i - j + 1;
            j = prefix[j - 1];
        }
        return found;
    }

    private static string[] RewriteKeys(RunTextMap.Map map, OpenXmlPart? owner)
    {
        var keys = new string[map.FlatText.Length];
        foreach (var segment in map.Segments)
        {
            var run = segment.Run;
            var props = run.Element(W.rPr);
            var link = run.Ancestors(W.hyperlink).FirstOrDefault();
            var href = RewriteHref(link, owner);
            var format = href is not null ? ":" + href : WmlToMarkdownConverter.IsCodeRun(props) ? "c:"
                : string.Concat(WmlToMarkdownConverter.HasToggle(props, W.b) ? "b" : "",
                WmlToMarkdownConverter.HasToggle(props, W.i) ? "i" : "",
                WmlToMarkdownConverter.HasToggle(props, W.strike) ? "s" : "", ":");
            for (int i = segment.StartOffsetInBlock; i < segment.EndOffsetInBlock; i++)
                keys[i] = map.FlatText[i] + "\u0001" + (map.FlatText[i] == '\n' ? ":" : format);
        }
        return keys;
    }

    private static string? RewriteHref(XElement? link, OpenXmlPart? owner)
    {
        var relationship = link?.Attribute(R.id) is { } id
            ? owner?.HyperlinkRelationships.FirstOrDefault(r => r.Id == id.Value) : null;
        return (string?)link?.Attribute(MarkdownPayloadParser.HrefAttr)
            ?? (relationship is null ? null : PackageOpenRepair.TargetText(relationship.Uri))
            ?? (link?.Attribute(W.anchor) is { } anchor ? "#" + anchor.Value : null);
    }

    private void RestoreInlineControl(XElement replacement, RetainedControl control)
    {
        SplitRunsAtOffset(replacement, control.Start, NextRevisionId);
        SplitRunsAtOffset(replacement, control.Start + control.Length, NextRevisionId);
        var map = RunTextMap.Build(replacement);
        var pieces = RunTextMap.ResolveRange(map, control.Start, control.Length);
        // Keep an enclosing link outside the SDT; a link owned by the SDT stays inside it.
        var originalLink = control.Element.Ancestors(W.hyperlink).FirstOrDefault();
        var href = RewriteHref(originalLink, OwnedPartRelationships.FindOwner(_doc!, control.Element)?.Part);
        var parent = originalLink is null ? replacement : pieces.Count > 0
            ? pieces[0].Item1.Run.Ancestors(W.hyperlink).FirstOrDefault(link =>
                pieces[^1].Item1.Run.Ancestors().Contains(link)) ?? replacement
            : map.Segments.Where(s => s.StartOffsetInBlock == control.Start || s.EndOffsetInBlock == control.Start)
                .Select(s => s.Run.Ancestors(W.hyperlink).FirstOrDefault())
                .FirstOrDefault(link => link is not null && RewriteHref(link, null) == href) ?? replacement;
        int origin = ReferenceEquals(parent, replacement) ? 0 : RunTextMap.Build(replacement).Segments
            .First(s => s.Run.Ancestors().Contains(parent)).StartOffsetInBlock;
        int start = control.Start - origin, end = start + control.Length;
        SplitInlineContainersAtOffset(parent, start, NextRevisionId);
        SplitInlineContainersAtOffset(parent, end, NextRevisionId);
        int consumed = 0;
        var removed = new List<XElement>();
        XElement? before = null;
        foreach (var child in parent.Elements().Where(IsInlineChild))
        {
            int length = InlineChildTextLength(child);
            if (consumed == start && before is null) before = child;
            if (length > 0 && consumed >= start && consumed + length <= end) removed.Add(child);
            consumed += length;
        }
        var restored = control.Element;
        var emptyLinks = control.Length == 0 && ReferenceEquals(parent, replacement)
            ? restored.Ancestors(W.hyperlink).ToArray() : Array.Empty<XElement>();
        restored.Remove();
        // Empty links have no markdown spelling or replacement run to receive their control.
        foreach (var link in emptyLinks)
            restored = new XElement(W.hyperlink, link.Attributes().Where(a => a.Name != PtOpenXml.Unid), restored);
        if (before is null) parent.Add(restored);
        else before.AddBeforeSelf(restored);
        foreach (var child in removed) child.Remove();
    }
}
