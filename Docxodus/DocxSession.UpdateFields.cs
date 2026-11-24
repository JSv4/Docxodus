// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;

namespace Docxodus;

/// <summary>Options for <see cref="DocxSession.UpdateFields"/>.</summary>
public sealed record UpdateFieldsOptions
{
    /// <summary>Only the fields that start in these paragraphs (a selection: Word's F9); null updates
    /// every field in the body (Ctrl+A, F9).</summary>
    public DocumentRange? Range { get; init; }

    /// <summary>The page each paragraph is on, as the caller's layout shows it, keyed by the
    /// paragraph's anchor id or its unid: the page numbers of a table of contents and of
    /// <c>PAGEREF</c> fields. The session has no page layout of its own. A paragraph missing from
    /// the map keeps the page its fields show; a contents entry that had none shows <c>1</c> and is
    /// marked dirty, so Word updates it.</summary>
    public IReadOnlyDictionary<string, string>? Pages { get; init; }

    /// <summary>Only tables of contents (Word's Update Table); other fields are left as they are.</summary>
    public bool TablesOfContentsOnly { get; init; }
}

public sealed partial class DocxSession
{
    private const string ReferenceSourceNotFound = "Error! Reference source not found.";
    private const string BookmarkNotDefined = "Error! Bookmark not defined.";
    private const string NoTocEntries = "No table of contents entries found.";

    /// <summary>Fields with no cached result that can go stale: Word computes them as it lays the page
    /// out, or they only mark or accept something.</summary>
    private static readonly HashSet<string> FieldsThatNeedNoUpdate = new(StringComparer.Ordinal)
    {
        "HYPERLINK", "PAGE", "TC", "XE", "TA", "RD", "FORMTEXT", "FORMCHECKBOX", "FORMDROPDOWN",
        "MACROBUTTON", "GOTOBUTTON", "SYMBOL", "ADVANCE", "PRIVATE",
    };

    /// <summary>
    /// Word's Update Field (F9) over the body. A table of contents is built again from the headings
    /// (its <c>\o</c> levels and <c>\u</c> outline levels, <c>\h</c> hyperlinks; each entry a
    /// <c>TOC n</c> paragraph with the heading's number and text and a <c>PAGEREF</c> to a hidden
    /// <c>_Toc</c> bookmark on it); <c>REF</c> fields take their bookmark's current text or number
    /// (<c>\r</c>, <c>\n</c>, <c>\w</c>, <c>\t</c>, <c>\p</c>); <c>PAGEREF</c> fields their bookmark's
    /// page from <see cref="UpdateFieldsOptions.Pages"/>; <c>SEQ</c> fields are numbered again in
    /// document order. One undo step, or none when nothing changes.
    /// <para>A field whose result the engine cannot work out exactly keeps the result it has: a locked
    /// field, another kind of field (a table of figures or authorities among them), a switch not listed
    /// here, a result that holds tracked changes, comments or one end of a range. The document's
    /// request that Word update its fields on open is cleared only when no such field is left in any
    /// story. Refused while changes are tracked: Word regenerates these results wholesale, which no
    /// revision mark can take back.</para>
    /// </summary>
    public EditResult UpdateFields(UpdateFieldsOptions? options = null)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported("UpdateFields", "");
        var o = options ?? new UpdateFieldsOptions();
        var main = _doc!.MainDocumentPart;
        var body = main?.GetXDocument().Root?.Element(W.body);
        if (main is null || body is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no main document body");

        HashSet<XElement>? inRange = null;
        if (o.Range is { } range)
        {
            // (only the range's paragraphs matter, so not ValidateDocumentRange: a bookmark's rules
            // refuse a paragraph holding a field's fldChar, which is every contents entry)
            // (a block holding paragraphs, a content control or a table, stands for its first and last)
            var startBlock = FindAnchor(range.StartAnchorId)?.Resolve(_doc!);
            var endBlock = FindAnchor(range.EndAnchorId)?.Resolve(_doc!);
            var startParagraph = startBlock?.Name == W.p ? startBlock : startBlock?.Descendants(W.p).FirstOrDefault();
            var endParagraph = endBlock?.Name == W.p ? endBlock : endBlock?.Descendants(W.p).LastOrDefault();
            if (startParagraph is null || endParagraph is null)
                return EditResult.Fail(EditErrorCode.AnchorNotFound, "range endpoint not found", range.StartAnchorId);
            var all = body.Descendants(W.p).ToList();
            int first = all.IndexOf(startParagraph), last = all.IndexOf(endParagraph);
            if (first < 0 || last < first)
                return EditResult.Fail(EditErrorCode.AnchorWrongKind, "UpdateFields takes a range in the body");
            inRange = new HashSet<XElement>(all.Skip(first).Take(last - first + 1));
        }

        var pages = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var (key, value) in o.Pages ?? new Dictionary<string, string>())
        {
            if (XmlText.ValidatePayload(value, key) is { } textError) return EditResult.Fail(textError);
            pages[key.Contains(':') ? key[(key.LastIndexOf(':') + 1)..] : key] = value;
        }

        // (no field this updates: no edit, and no undo step, as when printing a document without any)
        var updatable = o.TablesOfContentsOnly ? new[] { "TOC" } : new[] { "TOC", "REF", "PAGEREF", "SEQ" };
        if (!ScanFields(body).Any(f => updatable.Contains(f.Kind) && f.StartsIn(inRange)))
            return new EditResult { Success = true };

        // (the undo step is taken just before the first change: one taken and dropped again would have
        // emptied the redo history and pushed out the oldest undo step)
        var update = new FieldUpdate(pages, () => _history.RecordPreOp(TakeSnapshot()));
        try
        {
            void References(string kind)
            {
                foreach (var field in ScanFields(body).Where(f => f.Kind == kind && f.StartsIn(inRange)))
                    if (field.Begin.Ancestors(W.body).Any() && ReferenceResult(field, update) is { } result)
                        update.Set(field, result);
            }
            // The headings' own fields first: a table of contents copies their text. Page references last:
            // its entries bring their own.
            if (!o.TablesOfContentsOnly)
            {
                UpdateSequences(body, ScanFields(body), inRange, update);
                References("REF");
                if (update.Changed) InvalidateProjectionCache();
            }
            foreach (var toc in ScanFields(body).Where(f => f.Kind == "TOC" && f.StartsIn(inRange)).ToList())
                RebuildTableOfContents(body, toc, update);
            if (!o.TablesOfContentsOnly)
            {
                References("PAGEREF");
                update.CopyToFallbacks(body);
                if (inRange is null && !FieldsLeft(main, body, update)) ClearFieldUpdateOnOpen(main, update);
            }

            if (!update.Changed) return new EditResult { Success = true };
            InvalidateProjectionCache();
            var anchors = BodyAnchors();
            Anchor[] Anchors(IEnumerable<XElement> paragraphs) => paragraphs.Distinct()
                .Where(p => p.Ancestors(W.body).Any()).Select(p => AnchorIn(anchors, p)).OfType<Anchor>().ToArray();
            var created = Anchors(update.Created);
            var modified = Anchors(update.Modified.Where(p => !update.Created.Contains(p)));
            return new EditResult
            {
                Success = true,
                Created = created,
                Removed = update.Removed.ToArray(),
                Modified = modified,
                Patch = created.Concat(modified).Select(a => FindAnchor(a.Id)).FirstOrDefault(t => t is not null) is { } touched
                    ? PatchFor(touched) : null,
            };
        }
        catch (Exception ex)
        {
            // (nothing was changed before the undo step, so without one there is nothing to take back)
            return update.Changed ? FailInternal(ex) : EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    /// <summary>What one <see cref="UpdateFields"/> call has done so far.</summary>
    private sealed class FieldUpdate
    {
        private readonly IReadOnlyDictionary<string, string> _pages;
        private readonly Action _recordUndoStep;
        private readonly HashSet<XElement> _current = new();

        public FieldUpdate(IReadOnlyDictionary<string, string> pages, Action recordUndoStep)
        {
            _pages = pages;
            _recordUndoStep = recordUndoStep;
        }

        public bool Changed { get; private set; }

        /// <summary>The document is about to change: the first time, take the undo step.</summary>
        public void Changing()
        {
            if (Changed) return;
            _recordUndoStep();
            Changed = true;
        }

        public readonly HashSet<XElement> Created = new();
        public readonly List<XElement> Modified = new();
        public readonly List<Anchor> Removed = new();

        public string? PageOf(XElement paragraph) =>
            (string?)paragraph.Attribute(PtOpenXml.Unid) is { } unid && _pages.TryGetValue(unid, out var page) ? page : null;

        /// <summary>The field's result is known to be what Word would show now.</summary>
        public bool IsCurrent(FoundField field) => _current.Contains(field.Begin);
        public void MarkCurrent(XElement begin) => _current.Add(begin);

        public void Set(FoundField field, string result)
        {
            if (field.SetResult(result, Changing) is not { } changed) return;
            MarkCurrent(field.Begin);
            if (changed) Modified.Add(field.Paragraph);
        }

        /// <summary>A drawing's fallback holds a copy of its text: give each field in it the result its
        /// counterpart took.</summary>
        public void CopyToFallbacks(XElement body)
        {
            foreach (var alternate in body.Descendants(MC.AlternateContent))
            {
                if (alternate.Element(MC.Choice) is not { } choice || alternate.Element(MC.Fallback) is not { } fallback) continue;
                var from = ScanFields(choice);
                var to = ScanFields(fallback);
                if (from.Count == 0 || !from.Select(f => f.Instr).SequenceEqual(to.Select(f => f.Instr))) continue;
                foreach (var (source, copy) in from.Zip(to))
                    if (IsCurrent(source)) copy.SetResult(source.ResultText(), Changing);
            }
        }
    }

    /// <summary>Whether any story still holds a field that only Word can bring up to date.</summary>
    private bool FieldsLeft(MainDocumentPart main, XElement body, FieldUpdate update) =>
        ScanFields(body).Any(f => !update.IsCurrent(f) && NeedsUpdate(f))
        || OwnedPartRelationships.StoryParts(_doc!).Any(story => story.Part != main
            && story.Part.GetXDocument().Root is { } root && ScanFields(root).Any(NeedsUpdate));

    private static bool NeedsUpdate(FoundField field) => !field.Locked && !FieldsThatNeedNoUpdate.Contains(field.Kind);

    // ─── REF and PAGEREF ────────────────────────────────────────────────────

    /// <summary>What a <c>REF</c> or <c>PAGEREF</c> field shows now, or null when that is not known.</summary>
    private string? ReferenceResult(FoundField field, FieldUpdate update)
    {
        if (field.Locked || field.Tokens.Length < 2 || field.Tokens[1].StartsWith('\\')) return null;
        var bookmark = field.Tokens[1];
        var switches = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 2; i < field.Tokens.Length; i++)
        {
            var token = field.Tokens[i].ToLowerInvariant();
            if (token == "\\*" && i + 1 < field.Tokens.Length && IsKeepFormat(field.Tokens[i + 1])) { i++; continue; }
            if (token is not ("\\h" or "\\p") && (field.Kind == "PAGEREF" || token is not ("\\n" or "\\r" or "\\w" or "\\t"))) return null;
            switches.Add(token);
        }

        if (ResolveBookmarkPair(bookmark, hyperlinkTarget: true, null, out var start, out var end) is not null)
            // (bookmark names match in any case, and a malformed pair may still be one to Word)
            return OwnedPartRelationships.StoryParts(_doc!).Any(story => story.Part.GetXDocument().Descendants(W.bookmarkStart)
                    .Any(b => string.Equals((string?)b.Attribute(W.name), bookmark, StringComparison.OrdinalIgnoreCase)))
                ? null
                : field.Kind == "PAGEREF" ? BookmarkNotDefined : ReferenceSourceNotFound;

        bool sameStory = ReferenceEquals(field.Begin.AncestorsAndSelf().Last(), start.AncestorsAndSelf().Last());
        string? position = !switches.Contains("\\p") ? null : !sameStory ? ""
            : XNode.DocumentOrderComparer.Compare(start, field.Begin) < 0 ? "above" : "below";

        if (field.Kind == "PAGEREF")
        {
            var target = start.Ancestors(W.p).FirstOrDefault() ?? start.ElementsAfterSelf(W.p).FirstOrDefault();
            if ((target is null ? null : update.PageOf(target)) is not { } page) return null;
            if (position is null) return page;
            if (update.PageOf(field.Paragraph) is not { } here) return null;
            return here != page ? "on page " + page : position.Length > 0 ? position : null;
        }

        if (position is "") return null;
        string? numberSwitch = new[] { ReferenceNumber.NoContext, ReferenceNumber.FullContext, ReferenceNumber.RelativeContext }
            .FirstOrDefault(switches.Contains);
        if (numberSwitch is null) return position ?? TextBetween(start, end);
        var number = ReferenceNumberOf(bookmark, field.Paragraph, numberSwitch, switches.Contains("\\t"));
        return number is null || position is null ? number : number + " " + position;
    }

    /// <summary>The text a bookmark inside one paragraph covers, or null for one that covers more (other
    /// paragraphs, a picture, a note, an equation) or nothing.</summary>
    private static string? TextBetween(XElement start, XElement end)
    {
        var paragraph = start.Ancestors(W.p).FirstOrDefault();
        if (paragraph is null || !ReferenceEquals(paragraph, end.Ancestors(W.p).FirstOrDefault())) return null;
        var covered = paragraph.Descendants().SkipWhile(e => !ReferenceEquals(e, start)).TakeWhile(e => !ReferenceEquals(e, end));
        if (covered.Any(e => e.Name == W.drawing || e.Name == W.pict || e.Name == W._object || e.Name == MC.AlternateContent
                || e.Name == W.footnoteReference || e.Name == W.endnoteReference || e.Name == M.oMath))
            return null;
        int from = MarkerOffset(paragraph, start), to = MarkerOffset(paragraph, end);
        return to > from ? RunTextMap.Build(paragraph).FlatText[from..to] : null;
    }

    /// <summary>The general format switches that only keep formatting.</summary>
    private static bool IsKeepFormat(string format) =>
        format.Equals("MERGEFORMAT", StringComparison.OrdinalIgnoreCase) || format.Equals("CHARFORMAT", StringComparison.OrdinalIgnoreCase);

    // ─── SEQ ────────────────────────────────────────────────────────────────

    /// <summary>Number the <c>SEQ</c> fields again, counting over the whole body even when only part is
    /// updated.</summary>
    private void UpdateSequences(XElement body, IReadOnlyList<FoundField> fields, HashSet<XElement>? inRange, FieldUpdate update)
    {
        var sequences = fields.Where(f => f.Kind == "SEQ").ToList();
        if (sequences.Count == 0) return;
        var place = new Dictionary<XElement, int>();
        foreach (var e in body.Descendants()) place[e] = place.Count;
        var levels = new OutlineLevels(_doc!.MainDocumentPart!);
        var headings = body.Descendants(W.p).Select(p => (Place: place[p], Level: levels.BuiltInHeading(p) + 1))
            .Where(h => h.Level is not null).ToList();
        var counters = new Dictionary<string, (int Value, int Place)>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in sequences)
        {
            // (an identifier is required, and one followed by a bookmark repeats that bookmark's number)
            if (field.Tokens.Length < 2 || field.Tokens[1].StartsWith('\\')
                || (field.Tokens.Length > 2 && !field.Tokens[2].StartsWith('\\'))) continue;
            var id = field.Tokens[1];
            counters.TryGetValue(id, out var counter);
            int value = counter.Value;
            string format = "Arabic";
            bool hidden = false, repeat = false, known = true;
            int? reset = null;
            for (int i = 2; i < field.Tokens.Length; i++)
            {
                var token = field.Tokens[i].ToLowerInvariant();
                int? Number(int at) => int.TryParse(token.Length > 2 ? token[2..] : at < field.Tokens.Length ? field.Tokens[at] : "",
                    NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : null;
                if (token == "\\h") hidden = true;
                else if (token == "\\c") repeat = true;
                else if (token == "\\n") continue;
                else if (token == "\\*" && i + 1 < field.Tokens.Length)
                {
                    if (!IsKeepFormat(field.Tokens[++i])) format = field.Tokens[i];
                }
                else if (token.StartsWith("\\r", StringComparison.Ordinal) && Number(i + 1) is { } restart)
                {
                    reset = restart;
                    if (token.Length == 2) i++;
                }
                else if (token.StartsWith("\\s", StringComparison.Ordinal) && Number(i + 1) is { } level)
                {
                    if (headings.Any(h => h.Level <= level && h.Place > counter.Place && h.Place < place[field.Begin])) value = 0;
                    if (token.Length == 2) i++;
                }
                else known = false;
            }
            value = reset ?? (repeat ? Math.Max(value, 1) : value + 1);
            counters[id] = (value, place[field.Begin]);
            if (!known || field.Locked || !field.StartsIn(inRange) || FormatSequence(value, format) is not { } shown) continue;
            update.Set(field, hidden ? string.Empty : shown);
        }
    }

    /// <summary>A sequence number in a <c>\*</c> format, whose first letter's case Word follows; null for
    /// a format this does not write.</summary>
    private static string? FormatSequence(int value, string format)
    {
        string Cased(string text) => char.IsLower(format[0]) ? text.ToLowerInvariant() : text;
        return format.ToLowerInvariant() switch
        {
            "arabic" => value.ToString(CultureInfo.InvariantCulture),
            "alphabetic" when value > 0 => Cased(new string((char)('A' + (value - 1) % 26), (value - 1) / 26 + 1)),
            "roman" when value is > 0 and < 4000 => Cased(Roman(value)),
            _ => null,
        };

        static string Roman(int n)
        {
            var sb = new StringBuilder();
            foreach (var (v, s) in new[] { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
                         (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") })
                while (n >= v) { sb.Append(s); n -= v; }
            return sb.ToString();
        }
    }

    // ─── Fields as found in a story ─────────────────────────────────────────

    private sealed class FoundField
    {
        public string Kind = "";
        public string[] Tokens = Array.Empty<string>();
        /// <summary>The instruction as written.</summary>
        public string Instr = "";
        public required bool Complex;
        /// <summary>The <c>w:fldSimple</c>, or the run holding the begin <c>w:fldChar</c>.</summary>
        public required XElement Begin;
        public XElement? Separate;
        public XElement? End;
        public XElement Paragraph => Begin.AncestorsAndSelf(W.p).First();

        public bool Locked => WordprocessingMLUtil.OnOffValue(
            (string?)(Complex ? Begin.Element(W.fldChar) : Begin)?.Attribute(W.fldLock)) == true;

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

        /// <summary>Whether the field starts in <paramref name="paragraphs"/> (null: anywhere); a table of
        /// contents also when one of its entries is there.</summary>
        public bool StartsIn(HashSet<XElement>? paragraphs) => paragraphs is null || paragraphs.Contains(Paragraph)
            || (Kind == "TOC" && Paragraphs().Any(paragraphs.Contains));

        /// <summary>The nodes of the cached result, or null when it runs over paragraphs.</summary>
        private List<XNode>? Result() => !Complex ? Begin.Nodes().ToList()
            : Separate is null ? new List<XNode>()
            : End is null || !ReferenceEquals(Separate.Parent, End.Parent) ? null
            : Separate.NodesAfterSelf().TakeWhile(n => !ReferenceEquals(n, End)).ToList();

        public string ResultText() => string.Concat((Result() ?? new List<XNode>()).OfType<XElement>().Select(VisibleText));

        /// <summary>Replace the cached result with one run of <paramref name="text"/>, keeping the
        /// first result run's formatting; the field's dirty mark goes. True when that changed the field,
        /// false when it already shows the text, null when the result is not one to replace.
        /// <paramref name="changing"/> is called before anything changes.</summary>
        public bool? SetResult(string text, Action changing)
        {
            if (Complex && End is null) return null;
            if (Result() is not { } result || result.OfType<XElement>().Any(HoldsMoreThanText)) return null;
            var dirty = (Complex ? Begin.Element(W.fldChar) : Begin)?.Attribute(W.dirty);
            bool shown = ResultText() == text;
            if (shown && dirty is null) return false;
            changing();
            dirty?.Remove();
            if (shown) return true;
            var rPr = result.OfType<XElement>().SelectMany(e => e.DescendantsAndSelf(W.r)).Select(r => r.Element(W.rPr)).FirstOrDefault(r => r is not null)
                ?? (Complex ? Begin.Element(W.rPr) : null);
            foreach (var node in result) node.Remove();
            var run = new XElement(W.r, rPr is null ? null : new XElement(rPr), TextContent(text, W.t));
            UnidHelper.AssignToSelfAndDescendants(run);
            if (!Complex) Begin.Add(run);
            else if (Separate is null)
            {
                Separate = new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate")));
                UnidHelper.AssignToSelfAndDescendants(Separate);
                End!.AddBeforeSelf(Separate, run);
            }
            else Separate.AddAfterSelf(run);
            return true;
        }
    }

    /// <summary>Content a regenerated field result would take with it: a revision, a comment, a note, or
    /// a marker whose other end may be elsewhere.</summary>
    private static bool HoldsMoreThanText(XElement content) => content.DescendantsAndSelf().Any(e =>
        e.Name == W.ins || e.Name == W.del || e.Name == W.moveFrom || e.Name == W.moveTo
        || e.Name == W.rPrChange || e.Name == W.pPrChange || e.Name == W.sectPrChange
        || e.Name == W.commentReference || e.Name == W.footnoteReference || e.Name == W.endnoteReference
        || CrossBlockRangePairs.Any(pair => e.Name == pair.Start || e.Name == pair.End));

    /// <summary>Every field in <paramref name="root"/> in document order, simple and complex, with
    /// its kind (the first word of the instruction, upper-case) and instruction tokens. Fields in
    /// removed text are left out, and so is a drawing's fallback, which repeats its text.</summary>
    private static List<FoundField> ScanFields(XElement root)
    {
        var found = new List<FoundField>();
        var open = new Stack<(FoundField Field, StringBuilder Instr)>();
        void Add(FoundField field, string instr)
        {
            field.Instr = instr;
            field.Tokens = Regex.Matches(instr, "\"([^\"]*)\"|(\\S+)")
                .Select(m => m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value).ToArray();
            if (field.Tokens.Length == 0) return;
            field.Kind = field.Tokens[0].ToUpperInvariant();
            found.Add(field);
        }
        foreach (var e in root.Descendants())
        {
            if (e.Name != W.fldSimple && e.Name != W.fldChar && e.Name != W.instrText) continue;
            if (e.Ancestors().TakeWhile(a => !ReferenceEquals(a, root))
                .Any(a => a.Name == W.del || a.Name == W.moveFrom || a.Name == MC.Fallback)) continue;
            if (e.Name == W.fldSimple)
                Add(new FoundField { Complex = false, Begin = e }, (string?)e.Attribute(W.instr) ?? "");
            else if (e.Name == W.instrText)
            {
                if (open.Count > 0 && open.Peek().Field.Separate is null) open.Peek().Instr.Append(e.Value);
            }
            else if (e.Parent?.Name == W.r)
            {
                var type = (string?)e.Attribute(W.fldCharType);
                if (type == "begin") open.Push((new FoundField { Complex = true, Begin = e.Parent }, new StringBuilder()));
                else if (type == "separate" && open.Count > 0) open.Peek().Field.Separate ??= e.Parent;
                else if (type == "end" && open.Count > 0)
                {
                    var (field, instr) = open.Pop();
                    field.End = e.Parent;
                    Add(field, instr.ToString());
                }
            }
        }
        return found.OrderBy(f => f.Begin, Comparer<XElement>.Create(XNode.DocumentOrderComparer.Compare)).ToList();
    }

    // ─── Table of contents ──────────────────────────────────────────────────

    /// <summary>What a <c>TOC</c> instruction asks for, when it asks only for headings.</summary>
    private sealed record TocSwitches(int From, int To, bool OutlineLevels, bool Links, int NoPagesFrom, int NoPagesTo)
    {
        /// <summary>Null for an instruction with a switch that selects other entries or lays them out
        /// otherwise (<c>\c</c> figures, <c>\t</c> styles, <c>\b</c> bookmark, <c>\f</c>, <c>\p</c>…).</summary>
        public static TocSwitches? Parse(string[] tokens)
        {
            var result = new TocSwitches(1, 9, false, false, 0, -1);
            for (int i = 1; i < tokens.Length; i++)
            {
                bool hasArgument = i + 1 < tokens.Length && !tokens[i + 1].StartsWith('\\');
                (int From, int To)? levels = !hasArgument ? (1, 9)
                    : Regex.Match(tokens[i + 1], @"^([1-9])\s*-\s*([1-9])$") is { Success: true } m
                        ? (m.Groups[1].Value[0] - '0', m.Groups[2].Value[0] - '0') : null;
                switch (tokens[i].ToLowerInvariant())
                {
                    case "\\o" when levels is { } l: result = result with { From = l.From, To = l.To }; break;
                    case "\\n" when levels is { } l: result = result with { NoPagesFrom = l.From, NoPagesTo = l.To }; break;
                    case "\\h" when !hasArgument: result = result with { Links = true }; break;
                    case "\\u" when !hasArgument: result = result with { OutlineLevels = true }; break;
                    case "\\z" or "\\w" or "\\x" when !hasArgument: break;
                    case "\\*" when hasArgument && IsKeepFormat(tokens[i + 1]): break;
                    default: return null;
                }
                if (hasArgument) i++;
            }
            return result;
        }
    }

    private void RebuildTableOfContents(XElement body, FoundField toc, FieldUpdate update)
    {
        if (toc.Locked || !toc.Complex || toc.End is null || toc.Separate is null || TocSwitches.Parse(toc.Tokens) is not { } switches) return;
        var paragraphs = toc.Paragraphs().ToList();
        XElement first = paragraphs[0], last = paragraphs[^1];
        // Only a table that is whole paragraphs, one after another: nothing else goes with them.
        if (!ReferenceEquals(toc.Begin.Parent, first) || !ReferenceEquals(toc.End.Parent, last) || !ReferenceEquals(first.Parent, last.Parent)
            || first.ElementsAfterSelf().TakeWhile(e => !ReferenceEquals(e, last) && !ReferenceEquals(first, last)).Any(e => e.Name != W.p)
            || paragraphs.SkipLast(1).Any(p => p.Element(W.pPr)?.Element(W.sectPr) is not null))
            return;
        var before = first.Elements().Where(e => e.Name != W.pPr).TakeWhile(e => !ReferenceEquals(e, toc.Begin)).ToList();
        var after = toc.End.ElementsAfterSelf().ToList();
        var content = paragraphs.SelectMany(p => p.Elements()).Where(e => !before.Contains(e) && !after.Contains(e)).ToList();
        if (content.Any(HoldsMoreThanText)) return;

        update.Changing();
        var anchors = BodyAnchors();
        var inTable = new HashSet<XElement>(paragraphs);
        var levels = new OutlineLevels(_doc!.MainDocumentPart!);
        var shownPages = ScanFields(body).Where(f => f.Kind == "PAGEREF" && f.Tokens.Length > 1 && inTable.Contains(f.Paragraph))
            .GroupBy(f => f.Tokens[1]).ToDictionary(g => g.Key, g => g.First().ResultText());
        var bookmarks = ListBookmarks();
        var entries = new List<(int Level, string? Number, string Text, string Bookmark, string? Page, bool IsCurrent)>();
        foreach (var heading in body.Descendants(W.p).Where(p => !inTable.Contains(p)
                     && !p.Ancestors().Any(a => a.Name == W.txbxContent || a.Name == MC.Fallback)).ToList())
        {
            int? level = (levels.BuiltInHeading(heading) ?? (switches.OutlineLevels ? levels.Of(heading) : null)) + 1;
            if (level is null || level < switches.From || level > switches.To) continue;
            var text = Regex.Replace(ParagraphText(heading), @"\s+", " ").Trim();
            if (text.Length == 0 || AnchorIn(anchors, heading) is not { } anchor) continue;
            // A hidden _Toc bookmark over the heading's text; one there already is kept, first the one the
            // table used (Word leaves a heading several over time).
            var bookmark = bookmarks.FirstOrDefault(b => shownPages.ContainsKey(b.Name) && b.IsPaired && b.IsValid
                    && b.Range?.StartAnchorId == anchor.Id)?.Name
                ?? ParagraphBookmark(bookmarks, anchor.Id, heading, "_Toc");
            if (bookmark is null)
            {
                bookmark = NewHiddenBookmarkName("_Toc");
                BookmarkWholeParagraph(heading, bookmark);
                update.Modified.Add(heading);
            }
            var page = update.PageOf(heading);
            entries.Add((level.Value, ListNumberResolver.Resolve(heading, _doc!), text, bookmark,
                page ?? shownPages.GetValueOrDefault(bookmark), page is not null));
        }

        // The page numbers' right tab as the table had it, else (unless the entry style sets one) Word's,
        // just inside the right margin.
        var pageTab = first.Element(W.pPr)?.Element(W.tabs)?.Elements(W.tab)
            .LastOrDefault(t => (string?)t.Attribute(W.val) is "right" or "end");
        var begin = new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin")));
        var opening = new[]
        {
            begin,
            new XElement(W.r, new XElement(W.instrText, new XAttribute(XNamespace.Xml + "space", "preserve"), toc.Instr)),
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate"))),
        };
        var built = new List<XElement>();
        if (entries.Count == 0)
            built.Add(new XElement(W.p, WithoutSection(first.Element(W.pPr)),
                new XElement(W.r, new XElement(W.rPr, new XElement(W.b)), TextContent(NoTocEntries, W.t))));
        foreach (var (level, number, text, bookmark, page, isCurrent) in entries)
        {
            var runs = new List<XElement>();
            if (number is not null) { runs.Add(TextRun(number)); runs.Add(new XElement(W.r, new XElement(W.tab))); }
            runs.Add(TextRun(text));
            bool showsPage = level < switches.NoPagesFrom || level > switches.NoPagesTo;
            if (showsPage)
            {
                runs.Add(new XElement(W.r, new XElement(W.tab)));
                runs.Add(new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"),
                    page is null ? new XAttribute(W.dirty, "true") : null)));
                runs.Add(new XElement(W.r, new XElement(W.instrText, new XAttribute(XNamespace.Xml + "space", "preserve"), $" PAGEREF {bookmark} \\h ")));
                runs.Add(new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate"))));
                runs.Add(TextRun(page ?? "1"));
                runs.Add(new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))));
            }
            var style = TocLevelStyle(level);
            var tabs = new XElement(W.tabs);
            if (number is not null)
                tabs.Add(new XElement(W.tab, new XAttribute(W.val, "left"), new XAttribute(W.pos,
                    ((level - 1) * 220 + Math.Max(440, (number.Length + 1) * 120)).ToString(CultureInfo.InvariantCulture))));
            if (showsPage)
                tabs.Add(pageTab is not null ? new XElement(pageTab)
                    : StyleHasRightTab(style) ? null : PageNumberTab(first));
            built.Add(new XElement(W.p,
                new XElement(W.pPr, new XElement(W.pStyle, new XAttribute(W.val, style)),
                    tabs.HasElements ? tabs : null, new XElement(W.rPr, new XElement(W.noProof))),
                switches.Links
                    ? new XElement(W.hyperlink, new XAttribute(W.anchor, bookmark), new XAttribute(W.history, "1"), runs)
                    : runs));
        }
        foreach (var p in built) UnidHelper.AssignToSelfAndDescendants(p);

        // The paragraph that closes the field stays when it holds no entry: Word keeps it, and it is
        // where the document's text (and often a section break) goes on.
        bool keepsClosing = !ReferenceEquals(first, last)
            && !toc.End.ElementsBeforeSelf().Any(e => e.Name != W.pPr && VisibleText(e).Length > 0);
        update.Removed.AddRange(paragraphs.Where(p => !keepsClosing || !ReferenceEquals(p, last)).Select(p => AnchorIn(anchors, p)).OfType<Anchor>());
        foreach (var e in before) e.Remove();
        if (keepsClosing)
            foreach (var e in toc.End.ElementsBeforeSelf().Where(e => e.Name != W.pPr).ToList()) e.Remove();
        else
        {
            toc.End.Remove();
            foreach (var e in after) e.Remove();
            built[^1].Add(toc.End, after);
            if (last.Element(W.pPr)?.Element(W.sectPr) is { } section)
            {
                section.Remove();
                var pPr = built[^1].Element(W.pPr);
                if (pPr is null) built[^1].AddFirst(pPr = new XElement(W.pPr));
                SetPPrChildInOrder(pPr, section);
            }
        }
        if (built[0].Element(W.pPr) is { } openingPPr) openingPPr.AddAfterSelf(before, opening);
        else built[0].AddFirst(before, opening);
        foreach (var run in opening) UnidHelper.AssignToSelfAndDescendants(run);
        first.AddBeforeSelf(built);
        foreach (var p in paragraphs)
            if (!keepsClosing || !ReferenceEquals(p, last)) p.Remove();

        InvalidateProjectionCache();
        update.MarkCurrent(begin);
        foreach (var p in built) update.Created.Add(p);
        if (keepsClosing) update.Modified.Add(last);

        static XElement TextRun(string text) => new(W.r, TextContent(text, W.t));
        static XElement? WithoutSection(XElement? pPr)
        {
            if (pPr is null) return null;
            var copy = new XElement(pPr);
            copy.Elements(W.sectPr).Remove();
            return copy;
        }
    }

    /// <summary>The style of a contents entry at <paramref name="level"/>: the document's own (found by
    /// its built-in name, so a localized id counts), else Word's, indented 11 pt a level.</summary>
    private string TocLevelStyle(int level)
    {
        var number = level.ToString(CultureInfo.InvariantCulture);
        IEnumerable<XElement> Styles() => _doc!.MainDocumentPart!.StyleDefinitionsPart?.GetXDocument().Root?
            .Elements(W.style).Where(s => (string?)s.Attribute(W.type) == "paragraph") ?? Enumerable.Empty<XElement>();
        if (Styles().FirstOrDefault(s => string.Equals((string?)s.Element(W.name)?.Attribute(W.val), "toc " + number, StringComparison.OrdinalIgnoreCase))
                ?.Attribute(W.styleId)?.Value is { } own)
            return own;
        var id = "TOC" + number;
        if (Styles().Any(s => (string?)s.Attribute(W.styleId) == id)) return id;
        StyleFactory.EnsureReferenceFieldStyles(_doc!, id, "toc " + number, withHeading: false);
        if (level > 1 && Styles().FirstOrDefault(s => (string?)s.Attribute(W.styleId) == id)?.Element(W.pPr) is { } pPr)
        {
            pPr.Add(new XElement(W.ind, new XAttribute(W.left, ((level - 1) * 220).ToString(CultureInfo.InvariantCulture))));
            _doc!.MainDocumentPart!.StyleDefinitionsPart!.PutXDocument();
        }
        return id;
    }

    private bool StyleHasRightTab(string styleId) => FormattingAssembler.ResolveEffectiveTabs(_doc!,
            new XElement(W.p, new XElement(W.pPr, new XElement(W.pStyle, new XAttribute(W.val, styleId)))), withDirect: false)
        ?.Elements(W.tab).Any(t => (string?)t.Attribute(W.val) is "right" or "end") == true;

    /// <summary>The dotted right tab Word gives a contents entry: half a point inside the text width of
    /// the section the table is in.</summary>
    private XElement PageNumberTab(XElement paragraph)
    {
        int width = AnchorForElement(paragraph) is { } anchor && FindAnchor(anchor.Id) is { } target
            && BlockMetadataOps.GetSectionInfo(_doc!, target) is { } section
            ? section.PageWidthTwips - section.MarginLeftTwips - section.MarginRightTwips : 9360;
        return new XElement(W.tab, new XAttribute(W.val, "right"), new XAttribute(W.leader, "dot"),
            new XAttribute(W.pos, Math.Max(width - 10, 720).ToString(CultureInfo.InvariantCulture)));
    }

    /// <summary>The cached results are current: the document no longer asks Word to update its
    /// fields when it is opened.</summary>
    private static void ClearFieldUpdateOnOpen(MainDocumentPart main, FieldUpdate update)
    {
        var part = main.DocumentSettingsPart;
        if (part?.GetXDocument().Root?.Element(W.updateFields) is not { } request) return;
        update.Changing();
        request.Remove();
        part.PutXDocument();
    }
}
