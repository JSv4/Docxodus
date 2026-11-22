// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Docxodus.Internal;

namespace Docxodus;

/// <summary>Options for <see cref="DocxSession.UpdateFields"/>.</summary>
public sealed record UpdateFieldsOptions
{
    /// <summary>Only the fields that start in these paragraphs (a selection: Word's F9); null updates
    /// every field in the body (Ctrl+A, F9) and, the cached results then being current, clears the
    /// document's request that Word update its fields on open.</summary>
    public DocumentRange? Range { get; init; }

    /// <summary>The page each paragraph is on, as the caller's layout shows it, keyed by the
    /// paragraph's anchor id or its unid: the page numbers of a table of contents and of
    /// <c>PAGEREF</c> fields. The session has no page layout of its own. A paragraph missing from
    /// the map leaves a <c>PAGEREF</c> as it was and gives a contents entry page <c>1</c>.</summary>
    public IReadOnlyDictionary<string, string>? Pages { get; init; }

    /// <summary>Only tables of contents (Word's Update Table); other fields are left as they are.</summary>
    public bool TablesOfContentsOnly { get; init; }
}

public sealed partial class DocxSession
{
    private const string ReferenceSourceNotFound = "Error! Reference source not found.";
    private const string NoTocEntries = "No table of contents entries found.";

    /// <summary>
    /// Word's Update Field (F9) over the body: a table of contents is built again from the headings
    /// (its <c>\o</c> levels, <c>\h</c> hyperlinks; each entry a <c>TOC n</c> paragraph with the
    /// heading's number and text and a <c>PAGEREF</c> to a hidden <c>_Toc</c> bookmark on it);
    /// <c>REF</c> fields take their bookmark's current text or number (<c>\r</c>, <c>\n</c>,
    /// <c>\w</c>, <c>\p</c>); <c>PAGEREF</c> fields their bookmark's page from
    /// <see cref="UpdateFieldsOptions.Pages"/>; <c>SEQ</c> fields are numbered again in document
    /// order. Other fields (tables of figures and authorities among them) are left as they are. One
    /// undo step. Refused while changes are tracked: Word regenerates these results wholesale, which
    /// no revision mark can take back.
    /// </summary>
    public EditResult UpdateFields(UpdateFieldsOptions? options = null)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported("UpdateFields", null);
        var o = options ?? new UpdateFieldsOptions();
        var main = _doc!.MainDocumentPart;
        var body = main?.GetXDocument().Root?.Element(W.body);
        if (main is null || body is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no main document body");

        HashSet<XElement>? inRange = null;
        if (o.Range is { } range)
        {
            var ends = ValidateDocumentRange(range);
            if (ends.Error is not null) return ends.Error;
            var all = body.Descendants(W.p).ToList();
            int first = all.IndexOf(ends.StartParagraph), last = all.IndexOf(ends.EndParagraph);
            if (first < 0 || last < first)
                return EditResult.Fail(EditErrorCode.AnchorWrongKind, "UpdateFields takes a range in the body");
            inRange = new HashSet<XElement>(all.Skip(first).Take(last - first + 1));
        }

        var pages = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (key, value) in o.Pages ?? new Dictionary<string, string>())
            pages[key.Contains(':') ? key[(key.LastIndexOf(':') + 1)..] : key] = value;
        string? PageOf(XElement paragraph) =>
            (string?)paragraph.Attribute(PtOpenXml.Unid) is { } unid && pages.TryGetValue(unid, out var page) ? page : null;

        // (no field this updates: no edit, and no undo step, as when printing a document without any)
        var updatable = o.TablesOfContentsOnly ? new[] { "TOC" } : new[] { "TOC", "REF", "PAGEREF", "SEQ" };
        if (!ScanFields(body).Any(f => updatable.Contains(f.Kind) && (inRange is null || inRange.Contains(f.Paragraph)
                || (f.Kind == "TOC" && f.Paragraphs().Any(inRange.Contains)))))
            return new EditResult { Success = true };

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // 1. Tables of contents, rebuilt (their entries' PAGEREFs come out current).
            foreach (var toc in ScanFields(body).Where(f => f.Kind == "TOC" && f.Complex && f.Depth == 0).ToList())
            {
                if (inRange is not null && !toc.Paragraphs().Any(inRange.Contains)) continue;
                RebuildTableOfContents(toc, PageOf);
            }
            InvalidateProjectionCache();
            if (o.TablesOfContentsOnly) return new EditResult { Success = true };

            // 2. SEQ numbers, counted over the whole body even when only part is updated.
            var counters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var fields = ScanFields(body);
            foreach (var field in fields.Where(f => f.Kind == "SEQ"))
            {
                var result = SequenceResult(field.Tokens, counters);
                if (result is null || (inRange is not null && !inRange.Contains(field.Paragraph))) continue;
                field.SetResult(result);
            }

            // 3. REF and PAGEREF.
            foreach (var field in fields.Where(f => f.Kind is "REF" or "PAGEREF"))
            {
                if (inRange is not null && !inRange.Contains(field.Paragraph)) continue;
                if (field.Tokens.Length < 2 || field.Tokens[1].StartsWith('\\')) continue;
                var bookmark = field.Tokens[1];
                var switches = new HashSet<string>(field.Tokens.Skip(2).Select(t => t.ToLowerInvariant()));
                string? result;
                if (ResolveBookmarkPair(bookmark, hyperlinkTarget: true, null, out var start, out _) is not null)
                    result = ReferenceSourceNotFound;
                else if (field.Kind == "PAGEREF")
                {
                    var target = start.Ancestors(W.p).FirstOrDefault()
                        ?? start.ElementsAfterSelf(W.p).FirstOrDefault();
                    var page = target is null ? null : PageOf(target);
                    if (page is null) continue;
                    result = page;
                    if (switches.Contains("\\p"))
                        result = PageOf(field.Paragraph) == page
                            ? (XNode.DocumentOrderComparer.Compare(start, field.Begin) < 0 ? "above" : "below")
                            : "on page " + page;
                }
                else
                {
                    string? numberSwitch = switches.Contains("\\w") ? "\\w" : switches.Contains("\\n") ? "\\n"
                        : switches.Contains("\\r") ? "\\r" : null;
                    result = ReferenceResult(bookmark, numberSwitch, switches.Contains("\\p"), field.Begin, start);
                }
                field.SetResult(result);
            }

            if (inRange is null) ClearFieldUpdateOnOpen(main);
            InvalidateProjectionCache();
            return new EditResult { Success = true };
        }
        catch (Exception ex)
        {
            return FailInternal(ex);
        }
    }

    /// <summary>A <c>REF</c> field's result, as Word renders it now: the bookmarked text, or with a
    /// number switch the target's number (<c>\n</c> and <c>\w</c> without its trailing period),
    /// then with <c>\p</c> the word above / below.</summary>
    private string ReferenceResult(string bookmarkName, string? numberSwitch, bool position, XElement field, XElement bookmarkStart)
    {
        string cached;
        if (numberSwitch is not null)
        {
            var prefix = ResolveBookmarkAnchors(bookmarkName)
                .Select(covered => covered.AutoNumberPrefix)
                .FirstOrDefault(value => !string.IsNullOrEmpty(value));
            cached = prefix ?? "0";
            if (numberSwitch != "\\r" && cached.Length > 1) cached = cached.TrimEnd('.');
        }
        else
        {
            cached = BookmarkedText(bookmarkName);
        }
        if (position && ReferenceEquals(field.AncestorsAndSelf().Last(), bookmarkStart.AncestorsAndSelf().Last()))
        {
            var word = XNode.DocumentOrderComparer.Compare(bookmarkStart, field) < 0 ? "above" : "below";
            cached = numberSwitch is not null ? cached + " " + word : word;
        }
        return cached;
    }

    /// <summary>The next value of a <c>SEQ</c> field (<c>\r n</c> resets, <c>\c</c> repeats,
    /// <c>\h</c> hides; <c>\* ARABIC/alphabetic/ALPHABETIC/roman/ROMAN</c>), or null for one
    /// without an identifier.</summary>
    private static string? SequenceResult(string[] tokens, Dictionary<string, int> counters)
    {
        if (tokens.Length < 2 || tokens[1].StartsWith('\\')) return null;
        var id = tokens[1];
        counters.TryGetValue(id, out var value);
        string? format = null;
        bool hidden = false, repeat = false;
        int? reset = null;
        for (int i = 2; i < tokens.Length; i++)
        {
            var t = tokens[i];
            if (t.Equals("\\h", StringComparison.OrdinalIgnoreCase)) hidden = true;
            else if (t.Equals("\\c", StringComparison.OrdinalIgnoreCase)) repeat = true;
            else if (t.Equals("\\r", StringComparison.OrdinalIgnoreCase) && i + 1 < tokens.Length
                && int.TryParse(tokens[i + 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var r)) { reset = r; i++; }
            else if (t == "\\*" && i + 1 < tokens.Length) { format = tokens[i + 1]; i++; }
        }
        value = reset ?? (repeat ? Math.Max(value, 1) : value + 1);
        counters[id] = value;
        if (hidden) return string.Empty;
        return format switch
        {
            "alphabetic" => Letters(value).ToLowerInvariant(),
            "ALPHABETIC" => Letters(value),
            "roman" => Roman(value).ToLowerInvariant(),
            "ROMAN" => Roman(value),
            _ => value.ToString(CultureInfo.InvariantCulture),
        };

        static string Letters(int n) => n <= 0 ? "0" : new string((char)('A' + (n - 1) % 26), (n - 1) / 26 + 1);
        static string Roman(int n)
        {
            if (n <= 0 || n >= 4000) return n.ToString(CultureInfo.InvariantCulture);
            var sb = new StringBuilder();
            foreach (var (v, s) in new[] { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
                         (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") })
                while (n >= v) { sb.Append(s); n -= v; }
            return sb.ToString();
        }
    }

    // ─── Fields as found in the body ────────────────────────────────────────

    private sealed class FoundField
    {
        public required string Kind;
        public required string[] Tokens;
        /// <summary>The instruction as written.</summary>
        public string Instr = "";
        public required bool Complex;
        public int Depth;
        /// <summary>The <c>w:fldSimple</c>, or the run holding the begin <c>w:fldChar</c>.</summary>
        public required XElement Begin;
        public XElement? Separate;
        public XElement? End;
        public XElement Paragraph => Begin.AncestorsAndSelf(W.p).First();

        /// <summary>The paragraphs from the field's begin to its end.</summary>
        public IEnumerable<XElement> Paragraphs()
        {
            var last = (End ?? Begin).AncestorsAndSelf(W.p).First();
            var p = Paragraph;
            yield return p;
            if (ReferenceEquals(p, last)) yield break;
            foreach (var next in p.ElementsAfterSelf())
            {
                if (next.Name == W.p) yield return next;
                if (ReferenceEquals(next, last)) yield break;
            }
        }

        /// <summary>Replace the cached result with one run of <paramref name="text"/>, keeping the
        /// first result run's formatting; the field's dirty mark goes.</summary>
        public void SetResult(string text)
        {
            if (!Complex)
            {
                var keep = Begin.Elements(W.r).FirstOrDefault()?.Element(W.rPr);
                foreach (var child in Begin.Nodes().ToList()) child.Remove();
                Begin.Add(ResultRun(keep, text));
                Begin.Attribute(W.dirty)?.Remove();
                return;
            }
            if (End is null) return;
            Begin.Element(W.fldChar)?.Attribute(W.dirty)?.Remove();
            if (Separate is null)
            {
                Separate = new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate")));
                End.AddBeforeSelf(Separate, ResultRun(Begin.Element(W.rPr), text));
                return;
            }
            // (a result over paragraphs is not a REF / PAGEREF / SEQ Word writes: left alone)
            if (!ReferenceEquals(Separate.Parent, End.Parent)) return;
            var between = Separate.NodesAfterSelf().TakeWhile(n => !ReferenceEquals(n, End)).ToList();
            var rPr = between.OfType<XElement>().Where(e => e.Name == W.r).Select(r => r.Element(W.rPr)).FirstOrDefault(r => r is not null)
                ?? Begin.Element(W.rPr);
            foreach (var node in between) node.Remove();
            Separate.AddAfterSelf(ResultRun(rPr, text));
        }

        private static XElement ResultRun(XElement? rPr, string text)
        {
            var run = new XElement(W.r, rPr is null ? null : new XElement(rPr), TextContent(text, W.t));
            UnidHelper.AssignToSelfAndDescendants(run);
            return run;
        }
    }

    /// <summary>Every field in <paramref name="root"/> in document order, simple and complex, with
    /// its kind (the first word of the instruction, upper-case) and instruction tokens. Fields in
    /// deleted text are left out.</summary>
    private static List<FoundField> ScanFields(XElement root)
    {
        var found = new List<(FoundField Field, StringBuilder? Instr)>();
        var open = new Stack<(FoundField Field, StringBuilder Instr)>();
        foreach (var e in root.Descendants())
        {
            if (e.Ancestors(W.del).Any()) continue;
            if (e.Name == W.fldSimple)
            {
                var tokens = Tokens((string?)e.Attribute(W.instr) ?? "");
                if (tokens.Length > 0)
                    found.Add((new FoundField { Kind = tokens[0].ToUpperInvariant(), Tokens = tokens, Complex = false, Begin = e, Instr = (string?)e.Attribute(W.instr) ?? "" }, null));
            }
            else if (e.Name == W.fldChar && e.Parent?.Name == W.r)
            {
                var type = (string?)e.Attribute(W.fldCharType);
                if (type == "begin")
                {
                    var field = new FoundField { Kind = "", Tokens = Array.Empty<string>(), Complex = true, Begin = e.Parent!, Depth = open.Count };
                    open.Push((field, new StringBuilder()));
                }
                else if (type == "separate" && open.Count > 0 && open.Peek().Field.Separate is null)
                    open.Peek().Field.Separate = e.Parent;
                else if (type == "end" && open.Count > 0)
                {
                    var (field, instr) = open.Pop();
                    field.End = e.Parent;
                    var tokens = Tokens(instr.ToString());
                    if (tokens.Length > 0)
                    {
                        field.Kind = tokens[0].ToUpperInvariant();
                        field.Tokens = tokens;
                        field.Instr = instr.ToString();
                        found.Add((field, instr));
                    }
                }
            }
            else if (e.Name == W.instrText && open.Count > 0 && open.Peek().Field.Separate is null)
                open.Peek().Instr.Append(e.Value);
        }
        return found.Select(f => f.Field).OrderBy(f => f.Begin, Comparer<XElement>.Create(XNode.DocumentOrderComparer.Compare)).ToList();
    }

    /// <summary>An instruction's words; a quoted argument is one word, without its quotes.</summary>
    private static string[] Tokens(string instruction) =>
        Regex.Matches(instruction, "\"([^\"]*)\"|(\\S+)")
            .Select(m => m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value)
            .ToArray();

    // ─── Table of contents ──────────────────────────────────────────────────

    private void RebuildTableOfContents(FoundField toc, Func<XElement, string?> pageOf)
    {
        if (toc.End is null || toc.Separate is null) return;
        var paragraphs = toc.Paragraphs().ToList();
        var first = paragraphs[0];
        var last = paragraphs[^1];
        if (!ReferenceEquals(first.Parent, last.Parent)) return;

        // Levels: \o "a-b" (else 1-9); \h makes each entry a link; \n leaves the page numbers out.
        int from = 1, to = 9;
        bool links = false, noPages = false;
        for (int i = 1; i < toc.Tokens.Length; i++)
        {
            var t = toc.Tokens[i];
            if (t.Equals("\\o", StringComparison.OrdinalIgnoreCase) && i + 1 < toc.Tokens.Length && !toc.Tokens[i + 1].StartsWith('\\'))
            {
                var m = Regex.Match(toc.Tokens[++i], @"^(\d)\s*-\s*(\d)$");
                if (m.Success) { from = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture); to = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture); }
            }
            else if (t.Equals("\\h", StringComparison.OrdinalIgnoreCase)) links = true;
            else if (t.Equals("\\n", StringComparison.OrdinalIgnoreCase)) noPages = true;
        }

        var tocSet = new HashSet<XElement>(paragraphs);
        var headings = ListOutline()
            .Where(h => h.Level >= from && h.Level <= to && h.Text.Trim().Length > 0)
            .Select(h => (Entry: h, Target: FindAnchor(h.AnchorId)))
            .Where(h => h.Target?.Resolve(_doc!) is { } p && !tocSet.Contains(p))
            .ToList();

        // A hidden _Toc bookmark over each heading's text (one there already is kept).
        var existing = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var b in ListBookmarks())
            if (b.Name.StartsWith("_Toc", StringComparison.Ordinal) && b.IsPaired && b.IsValid && b.Range is { } r
                && r.StartAnchorId == r.EndAnchorId && r.StartOffset == 0 && !existing.ContainsKey(r.StartAnchorId))
                existing[r.StartAnchorId] = b.Name + "|" + r.EndOffset.ToString(CultureInfo.InvariantCulture);
        var entries = new List<(int Level, string? Number, string Text, string Bookmark, string Page)>();
        foreach (var (entry, target) in headings)
        {
            var paragraph = target!.Resolve(_doc!)!;
            int length = RunTextMap.Build(paragraph).FlatText.Length;
            string? name = existing.TryGetValue(entry.AnchorId, out var known) && known.EndsWith("|" + length.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)
                ? known[..known.IndexOf('|')] : null;
            if (name is null)
            {
                do name = "_Toc" + Random.Shared.Next(100_000_000, 1_000_000_000).ToString(CultureInfo.InvariantCulture);
                while (BookmarkStarts(name).Count != 0);
                var ends = ValidateDocumentRange(new DocumentRange(entry.AnchorId, 0, entry.AnchorId, length));
                if (ends.Error is not null) continue;
                InsertRangeMarkers(ends, NextGlobalBookmarkId(), name);
            }
            var number = GetAnchorInfo(entry.AnchorId)?.AutoNumberPrefix;
            entries.Add((entry.Level, string.IsNullOrEmpty(number) ? null : number, entry.Text, name, pageOf(paragraph) ?? "1"));
        }

        // The entries' styles: TOC 1-9, indented by level, as Word's.
        foreach (var level in entries.Select(e => e.Level).Distinct()) EnsureTocLevelStyle(level);

        var tabs = first.Element(W.pPr)?.Element(W.tabs);
        var before = first.Elements().Where(n => n.Name != W.pPr).TakeWhile(n => !ReferenceEquals(n, toc.Begin) && !n.DescendantsAndSelf().Contains(toc.Begin)).ToList();
        var after = last.Elements().SkipWhile(n => !ReferenceEquals(n, toc.End) && !n.DescendantsAndSelf().Contains(toc.End)).Skip(1).ToList();
        var fieldStart = new[]
        {
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"))),
            new XElement(W.r, new XElement(W.instrText, new XAttribute(XNamespace.Xml + "space", "preserve"),
                toc.Instr)),
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate"))),
        };
        var fieldEnd = new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end")));

        var built = new List<XElement>();
        if (entries.Count == 0)
        {
            built.Add(new XElement(W.p, first.Element(W.pPr) is { } pPr ? new XElement(pPr) : null,
                before.Select(n => new XElement(n)), fieldStart,
                new XElement(W.r, new XElement(W.rPr, new XElement(W.b)), new XElement(W.t, NoTocEntries)),
                fieldEnd, after.Select(n => new XElement(n))));
        }
        else
        {
            for (int i = 0; i < entries.Count; i++)
            {
                var (level, number, text, bookmark, page) = entries[i];
                var content = new List<XElement>();
                if (number is not null) { content.Add(TextRun(number)); content.Add(new XElement(W.r, new XElement(W.tab))); }
                content.Add(TextRun(text));
                if (!noPages)
                {
                    content.Add(new XElement(W.r, new XElement(W.tab)));
                    content.Add(new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"))));
                    content.Add(new XElement(W.r, new XElement(W.instrText, new XAttribute(XNamespace.Xml + "space", "preserve"), $" PAGEREF {bookmark} \\h ")));
                    content.Add(new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate"))));
                    content.Add(TextRun(page));
                    content.Add(new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))));
                }
                var paragraph = new XElement(W.p,
                    new XElement(W.pPr,
                        new XElement(W.pStyle, new XAttribute(W.val, "TOC" + level.ToString(CultureInfo.InvariantCulture))),
                        tabs is null ? null : new XElement(tabs),
                        new XElement(W.rPr, new XElement(W.noProof))));
                if (i == 0) paragraph.Add(before.Select(n => new XElement(n)), fieldStart);
                if (links)
                    paragraph.Add(new XElement(W.hyperlink, new XAttribute(W.anchor, bookmark), new XAttribute(W.history, "1"), content));
                else paragraph.Add(content);
                if (i == entries.Count - 1) paragraph.Add(fieldEnd, after.Select(n => new XElement(n)));
                built.Add(paragraph);
            }
        }
        foreach (var p in built) UnidHelper.AssignToSelfAndDescendants(p);
        first.AddBeforeSelf(built);
        foreach (var p in paragraphs) p.Remove();

        static XElement TextRun(string text) => new XElement(W.r, TextContent(text, W.t));
    }

    /// <summary>Word's <c>TOC n</c> paragraph style (entries of level n), made when the document has
    /// none: Normal, indented 11 pt a level.</summary>
    private void EnsureTocLevelStyle(int level)
    {
        var id = "TOC" + level.ToString(CultureInfo.InvariantCulture);
        StyleFactory.EnsureReferenceFieldStyles(_doc!, id, "toc " + level.ToString(CultureInfo.InvariantCulture), withHeading: false);
        if (level <= 1) return;
        var part = _doc!.MainDocumentPart!.StyleDefinitionsPart;
        var style = part?.GetXDocument().Root?.Elements(W.style).FirstOrDefault(s => (string?)s.Attribute(W.styleId) == id);
        var pPr = style?.Element(W.pPr);
        // (only a style just made here, still without an indent: a document's own is left as it is)
        if (style is null || pPr is null || pPr.Element(W.ind) is not null || style.Element(W.autoRedefine) is null) return;
        pPr.Add(new XElement(W.ind, new XAttribute(W.left, ((level - 1) * 220).ToString(CultureInfo.InvariantCulture))));
        part!.PutXDocument();
    }

    /// <summary>The cached results are current: the document no longer asks Word to update its
    /// fields when it is opened.</summary>
    private static void ClearFieldUpdateOnOpen(DocumentFormat.OpenXml.Packaging.MainDocumentPart main)
    {
        var part = main.DocumentSettingsPart;
        var update = part?.GetXDocument().Root?.Element(W.updateFields);
        if (update is null) return;
        update.Remove();
        part!.PutXDocument();
    }
}
