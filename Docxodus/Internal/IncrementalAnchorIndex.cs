// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// The session's lookup anchor index, kept up to date block by block (issue #1022). It holds
/// exactly what <see cref="WmlToMarkdownConverter.BuildAnchorIndexOnly"/> returns — the same keys,
/// entries and document order — but after an edit it re-indexes only the top-level blocks whose
/// trees changed, as reported by each part's <see cref="PartChangeTracker"/>, instead of walking
/// the whole document again.
/// </summary>
/// <remarks>
/// <para><b>Layout.</b> Per projected scope, an ordered list of segments, one per element child of
/// the scope's block container (<c>w:body</c>, or the part root). A segment holds that block's
/// entries in walk order, so iterating scopes, then segments, then entries is the full index's
/// insertion order — the document order callers such as the query ops depend on. A dictionary
/// beside it answers lookups.</para>
/// <para><b>When it gives up.</b> <see cref="TryRefresh"/> returns false, and the session rebuilds
/// from scratch, whenever a block-local update could differ from a full rebuild:</para>
/// <list type="bullet">
/// <item>the scope list changed (a header, footer, notes or comments part came or went, or its
/// cached tree was replaced, as undo, redo and rollback do);</item>
/// <item>anything outside the blocks of a scope changed, or the styles part changed at all (a
/// paragraph's kind reads its style chain);</item>
/// <item>a changed block has no Unid of its own (its Unid would derive from its siblings), or
/// holds or held a content control (whose identity is decided story-wide);</item>
/// <item>a changed block's entry collides with an id already indexed (the full walk keeps the first
/// occurrence, which a block-local update cannot know).</item>
/// </list>
/// <para>It is never built for a non-<see cref="AnchorIdRendering.FullUnid"/> rendering, whose alias
/// keys are derived from every anchor in a bucket, nor when the full walk itself finds duplicate
/// ids or an addressable element outside the block containers.</para>
/// </remarks>
internal sealed class IncrementalAnchorIndex : IReadOnlyDictionary<string, AnchorTarget>
{
    private sealed class Segment
    {
        internal Segment(XElement block, string scopeName, AnchorTarget[] entries, bool hasContentControl)
        {
            Block = block;
            Entries = entries;
            HasContentControl = hasContentControl;
            BlockId = WmlToMarkdownConverter.BlockAnchorId(block, scopeName);
            NoteId = scopeName is "fn" or "en" ? (string?)block.Attribute(W.id) : null;
        }

        internal XElement Block { get; }

        /// <summary>The block's anchor id when it was indexed: a patch reports it as removed when
        /// the block goes or its id changes.</summary>
        internal string? BlockId { get; }

        /// <summary>A note block's <c>w:id</c> when it was indexed. References render a note's
        /// label through its id, so a note gained, lost or renumbered can change blocks that cite
        /// it without touching them.</summary>
        internal string? NoteId { get; }

        internal AnchorTarget[] Entries { get; }

        internal bool HasContentControl { get; }

        internal LinkedListNode<Segment>? Node { get; set; }
    }

    private sealed class Scope
    {
        internal required string Name { get; init; }
        internal required OpenXmlPart Part { get; init; }
        internal required string PartUri { get; init; }
        internal required XDocument Document { get; init; }
        internal required XElement Container { get; init; }
        internal required PartChangeTracker Tracker { get; init; }
        internal LinkedList<Segment> Segments { get; } = new();
        internal Dictionary<XElement, Segment> SegmentOf { get; } = new(ReferenceEqualityComparer.Instance);

        /// <summary>Blocks re-indexed since the last <see cref="TakePatchChanges"/>.</summary>
        internal HashSet<XElement> PendingChanged { get; } = new(ReferenceEqualityComparer.Instance);

        /// <summary>Ids of blocks whose old entries were dropped since the last <see cref="TakePatchChanges"/>.</summary>
        internal List<string> PendingRemoved { get; } = new();
    }

    private readonly WmlToMarkdownConverterSettings _settings;
    private readonly WordprocessingDocument _document;
    private readonly List<Scope> _scopes = new();
    private readonly Dictionary<string, AnchorTarget> _byKey = new(StringComparer.Ordinal);
    private XDocument? _styles;
    private int _version;

    /// <summary>Test seam: when set, the next refresh skips re-indexing the first changed block it
    /// meets, so a test can prove the differential oracle detects a stale entry. Per thread.</summary>
    [ThreadStatic]
    internal static bool SkipOneBlockForTests;

    /// <summary>Diagnostics for tests, per thread: refreshes that succeeded, and refreshes that gave
    /// up and left the session to rebuild.</summary>
    [ThreadStatic]
    internal static int RefreshesForTests;

    [ThreadStatic]
    internal static int FallbacksForTests;

    private IncrementalAnchorIndex(WordprocessingDocument document, WmlToMarkdownConverterSettings settings)
    {
        _document = document;
        _settings = settings;
    }

    /// <summary>Build the index over the whole document, or return null when this document or
    /// rendering cannot be maintained incrementally (the caller then uses the plain full index).
    /// Assigns Unids exactly as the full index build does.</summary>
    /// <param name="trackPatchChanges">Record changed blocks for <see cref="TakePatchChanges"/>. Off
    /// when the session emits no patches, so nothing accumulates that nobody takes.</param>
    internal static IncrementalAnchorIndex? TryBuild(WordprocessingDocument document, WmlToMarkdownConverterSettings settings,
        bool trackPatchChanges = false)
    {
        if (settings.AnchorIdRendering != AnchorIdRendering.FullUnid) return null;
        var main = document.MainDocumentPart
            ?? throw new InvalidOperationException("Document has no MainDocumentPart.");
        var index = new IncrementalAnchorIndex(document, settings) { _trackPatchChanges = trackPatchChanges };
        foreach (var (name, part) in WmlToMarkdownConverter.ProjectedScopes(main, settings))
        {
            var xml = part.GetXDocument();
            var root = xml.Root!;
            UnidHelper.AssignToAllElementsDeterministic(root);
            if (root.Annotation<OpenXmlPart>() == null) root.AddAnnotation(part);
            // Trust the tracker from here on: everything before this point is read in full below.
            var tracker = PartChangeTracker.For(xml);
            tracker.ForIndex.Activate();
            var container = PartChangeTracker.ContainerOf(xml)!;
            var scope = new Scope
            {
                Name = name,
                Part = part,
                PartUri = part.Uri.ToString(),
                Document = xml,
                Container = container,
                Tracker = tracker,
            };
            index._scopes.Add(scope);
            if (HasShellAnchor(root, container, name, settings))
            {
                index.Detach();
                return null;
            }
            foreach (var block in container.Elements())
            {
                var segment = index.BuildSegment(scope, block);
                if (segment is null || !index.Add(scope, segment, after: scope.Segments.Last))
                {
                    index.Detach();
                    return null;
                }
            }
        }
        index.WatchStyles(main);
        return index;
    }

    /// <summary>Fold every change the trackers recorded since the last build or refresh into the
    /// index. False means the index can no longer be trusted and must be rebuilt from scratch; it
    /// may then be partly updated and must be discarded.</summary>
    internal bool TryRefresh(WordprocessingDocument document)
    {
        // Every reopen path resets the session's caches today; this keeps an index built over a
        // package that is no longer the session's from ever passing its own checks.
        if (ReferenceEquals(document, _document) && RefreshCore())
        {
            RefreshesForTests++;
            return true;
        }
        FallbacksForTests++;
        Detach();
        return false;
    }

    /// <summary>Stop the trackers recording for this index: it is being discarded.</summary>
    internal void Detach()
    {
        foreach (var scope in _scopes) scope.Tracker.ForIndex.Deactivate();
        if (_styles is not null) PartChangeTracker.Existing(_styles)?.ForIndex.Deactivate();
    }

    private bool RefreshCore()
    {
        var main = _document.MainDocumentPart;
        if (main is null) return false;
        var scopes = WmlToMarkdownConverter.ProjectedScopes(main, _settings);
        if (scopes.Count != _scopes.Count) return false;
        for (int i = 0; i < scopes.Count; i++)
        {
            var scope = _scopes[i];
            if (scopes[i].Name != scope.Name || !ReferenceEquals(scopes[i].Part, scope.Part)
                || !ReferenceEquals(scope.Part.GetXDocument(), scope.Document)
                || !ReferenceEquals(PartChangeTracker.ContainerOf(scope.Document), scope.Container)
                || scope.Tracker.ForIndex.ShellChanged)
                return false;
        }
        var styles = main.StyleDefinitionsPart?.GetXDocument();
        if (!ReferenceEquals(styles, _styles) || (styles is not null && PartChangeTracker.For(styles).ForIndex.Any))
            return false;

        foreach (var scope in _scopes)
        {
            var changes = scope.Tracker.ForIndex;
            if (changes.Blocks.Count == 0) continue;
            var changed = changes.Blocks.OfType<XElement>().ToList();
            // Drop every changed block's old segment first, so a block that moved is re-placed and
            // the placement walk below never stops at a stale segment.
            var attached = new List<XElement>(changed.Count);
            bool notes = scope.Name is "fn" or "en";
            foreach (var block in changed)
            {
                bool isAttached = ReferenceEquals(block.Parent, scope.Container);
                if (scope.SegmentOf.TryGetValue(block, out var old))
                {
                    if (old.HasContentControl) return false;
                    if (old.BlockId is not null && _trackPatchChanges) scope.PendingRemoved.Add(old.BlockId);
                    if (notes && (!isAttached || (string?)block.Attribute(W.id) != old.NoteId))
                        _patchNeedsFullDocument = true;
                    Remove(scope, old);
                }
                else if (notes && isAttached)
                {
                    _patchNeedsFullDocument = true;
                }
                if (isAttached)
                {
                    attached.Add(block);
                    if (_trackPatchChanges) scope.PendingChanged.Add(block);
                }
                else
                {
                    scope.PendingChanged.Remove(block);
                }
            }
            foreach (var block in attached)
            {
                if (SkipOneBlockForTests)
                {
                    SkipOneBlockForTests = false;
                    continue;
                }
                if (block.Attribute(PtOpenXml.Unid) is null) return false;
                if (block.DescendantsAndSelf(W.sdt).Any()) return false;
                UnidHelper.AssignWithinBlock(block);
                var segment = BuildSegment(scope, block);
                if (segment is null || !Add(scope, segment, PlacementAfter(scope, block))) return false;
            }
            // Assigning Unids above reported those blocks again; they are now indexed.
            changes.Clear();
        }
        _version++;
        return true;
    }

    /// <summary>Set when a change since the last patch can alter the markdown of a block it did
    /// not touch (a note gained, lost or renumbered changes the label of every reference to it).</summary>
    private bool _patchNeedsFullDocument;

    private bool _trackPatchChanges;

    /// <summary>
    /// The top-level blocks changed since the last call (or since the build), for a block-scoped
    /// <see cref="MarkdownPatch"/>: each still-present changed block with its scope, in document
    /// order, and the ids of blocks that are gone (or whose id changed). Returns false when the
    /// changes cannot be expressed block by block; either way the pending changes are consumed.
    /// </summary>
    /// <summary>Whether changes are waiting for <see cref="TakePatchChanges"/>.</summary>
    internal bool HasPendingPatchChanges =>
        _patchNeedsFullDocument || _scopes.Any(s => s.PendingChanged.Count > 0 || s.PendingRemoved.Count > 0);

    internal bool TakePatchChanges(out List<(string Scope, XElement Block)> changed, out List<string> removed)
    {
        changed = new List<(string, XElement)>();
        removed = new List<string>();
        var seenRemoved = new HashSet<string>(StringComparer.Ordinal);
        bool exact = !_patchNeedsFullDocument;
        _patchNeedsFullDocument = false;
        foreach (var scope in _scopes)
        {
            if (scope.PendingChanged.Count == 0 && scope.PendingRemoved.Count == 0) continue;
            var present = new HashSet<string>(StringComparer.Ordinal);
            var blocks = scope.PendingChanged.Where(b => ReferenceEquals(b.Parent, scope.Container)).ToList();
            blocks.Sort(XNode.DocumentOrderComparer);
            foreach (var block in blocks)
            {
                changed.Add((scope.Name, block));
                if (WmlToMarkdownConverter.BlockAnchorId(block, scope.Name) is { } id) present.Add(id);
            }
            foreach (var id in scope.PendingRemoved)
                if (!present.Contains(id) && seenRemoved.Add(id)) removed.Add(id);
            scope.PendingChanged.Clear();
            scope.PendingRemoved.Clear();
        }
        return exact;
    }

    private void WatchStyles(MainDocumentPart main)
    {
        _styles = main.StyleDefinitionsPart?.GetXDocument();
        if (_styles is not null) PartChangeTracker.For(_styles).ForIndex.Activate();
    }

    /// <summary>The segment a newly placed block follows: that of its nearest preceding sibling
    /// element that is indexed. Every unchanged element child of the container is indexed, and
    /// blocks are placed one at a time, so this keeps segment order equal to document order
    /// whatever order the changed blocks arrive in.</summary>
    private static LinkedListNode<Segment>? PlacementAfter(Scope scope, XElement block)
    {
        for (var node = block.PreviousNode; node is not null; node = node.PreviousNode)
            if (node is XElement sibling && scope.SegmentOf.TryGetValue(sibling, out var segment))
                return segment.Node;
        return null;
    }

    /// <summary>The entries of one block, or null when it repeats an id within itself.</summary>
    private Segment? BuildSegment(Scope scope, XElement block)
    {
        bool hasContentControl = false;
        if (WmlToMarkdownConverter.IsSkippedNoteBlock(block, scope.Name))
            return new Segment(block, scope.Name, Array.Empty<AnchorTarget>(), block.DescendantsAndSelf(W.sdt).Any());

        List<AnchorTarget>? entries = null;
        foreach (var el in block.DescendantsAndSelf())
        {
            if (el.Name == W.sdt) hasContentControl = true;
            var id = WmlToMarkdownConverter.AnchorKeyFor(el, scope.Name, _settings, out var kind, out var unid);
            if (id is null) continue;
            entries ??= new List<AnchorTarget>();
            entries.Add(WmlToMarkdownConverter.CreateAnchorTarget(
                el, id, kind!, scope.Name, unid!, scope.PartUri, enrich: false, _document));
        }
        return new Segment(block, scope.Name, entries?.ToArray() ?? Array.Empty<AnchorTarget>(), hasContentControl);
    }

    private bool Add(Scope scope, Segment segment, LinkedListNode<Segment>? after)
    {
        for (int i = 0; i < segment.Entries.Length; i++)
        {
            if (_byKey.TryAdd(segment.Entries[i].Anchor.Id, segment.Entries[i])) continue;
            // A repeated id: the full walk keeps the first occurrence in document order, which only
            // a full walk can decide. Undo this segment's keys and give up.
            for (int j = 0; j < i; j++) _byKey.Remove(segment.Entries[j].Anchor.Id);
            return false;
        }
        segment.Node = after is null ? scope.Segments.AddFirst(segment) : scope.Segments.AddAfter(after, segment);
        scope.SegmentOf[segment.Block] = segment;
        return true;
    }

    private void Remove(Scope scope, Segment segment)
    {
        foreach (var entry in segment.Entries) _byKey.Remove(entry.Anchor.Id);
        scope.Segments.Remove(segment.Node!);
        scope.SegmentOf.Remove(segment.Block);
    }

    /// <summary>Whether any element outside the scope's blocks would be indexed by the full walk.
    /// Such an entry would precede or follow the blocks' entries in an order the segments cannot
    /// represent, so the document is left to the full index.</summary>
    private static bool HasShellAnchor(XElement root, XElement container, string scopeName,
        WmlToMarkdownConverterSettings settings)
    {
        if (WmlToMarkdownConverter.AnchorKeyFor(root, scopeName, settings, out _, out _) is not null) return true;
        if (ReferenceEquals(root, container)) return false;
        if (WmlToMarkdownConverter.AnchorKeyFor(container, scopeName, settings, out _, out _) is not null) return true;
        // The container is a child of the root (w:body under w:document); its siblings are shell.
        foreach (var sibling in root.Elements())
        {
            if (ReferenceEquals(sibling, container)) continue;
            foreach (var el in sibling.DescendantsAndSelf())
                if (WmlToMarkdownConverter.AnchorKeyFor(el, scopeName, settings, out _, out _) is not null) return true;
        }
        return false;
    }

    // ── IReadOnlyDictionary ──────────────────────────────────────────────

    public AnchorTarget this[string key] => _byKey[key];

    public IEnumerable<string> Keys
    {
        get
        {
            foreach (var pair in this) yield return pair.Key;
        }
    }

    public IEnumerable<AnchorTarget> Values
    {
        get
        {
            foreach (var pair in this) yield return pair.Value;
        }
    }

    public int Count => _byKey.Count;

    public bool ContainsKey(string key) => _byKey.ContainsKey(key);

    public bool TryGetValue(string key, [MaybeNullWhen(false)] out AnchorTarget value) =>
        _byKey.TryGetValue(key, out value);

    /// <summary>Entries in document order. Throws if the index is refreshed mid-enumeration, as a
    /// dictionary does when modified, so a caller holding the index across an edit cannot silently
    /// read a mixture of two states.</summary>
    public IEnumerator<KeyValuePair<string, AnchorTarget>> GetEnumerator()
    {
        int version = _version;
        foreach (var scope in _scopes)
        {
            foreach (var segment in scope.Segments)
            {
                foreach (var entry in segment.Entries)
                {
                    if (version != _version)
                        throw new InvalidOperationException("The anchor index was refreshed during enumeration.");
                    yield return new KeyValuePair<string, AnchorTarget>(entry.Anchor.Id, entry);
                }
            }
        }
        if (version != _version)
            throw new InvalidOperationException("The anchor index was refreshed during enumeration.");
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
