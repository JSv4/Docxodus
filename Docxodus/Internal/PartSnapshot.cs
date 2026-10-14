// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>A payload the undo ring counts once, together with the payloads inside it, which it
/// also counts once each however many groups hold them.</summary>
internal interface ISharedPayloadGroup
{
    int MemberCount { get; }

    /// <summary>Member <paramref name="index"/> and its size. Indexed rather than enumerated so
    /// the ring's per-edit recount allocates no iterator per group.</summary>
    (object Payload, long Bytes) Member(int index);
}

/// <summary>
/// An immutable deep copy of one top-level block of a part (issue #1022). Never attached to any
/// tree and never mutated, so any number of snapshots can hold the same instance.
/// </summary>
internal sealed class FrozenBlock
{
    private long? _bytes;

    internal FrozenBlock(XNode node) => Node = node;

    internal XNode Node { get; }

    internal long Bytes => _bytes ??= XmlMemoryEstimator.Estimate(Node);
}

/// <summary>
/// A run of consecutive frozen blocks. Snapshots hold their blocks in chunks rather than one flat
/// array so that a snapshot taken after a one-block edit can reuse every chunk the edit did not
/// touch: its own allocation is the touched chunk plus one reference per chunk.
/// </summary>
internal sealed class FrozenChunk : ISharedPayloadGroup
{
    internal FrozenChunk(FrozenBlock[] blocks) => Blocks = blocks;

    internal FrozenBlock[] Blocks { get; }

    /// <summary>The chunk's own array; its blocks are counted separately.</summary>
    internal long OwnBytes => 32 + (8L * Blocks.Length);

    public int MemberCount => Blocks.Length;

    public (object Payload, long Bytes) Member(int index) => (Blocks[index], Blocks[index].Bytes);
}

/// <summary>
/// One part's XML as captured by an undo snapshot: the part's shell (its tree with the block
/// container emptied) plus its blocks, frozen. Consecutive snapshots share every block, chunk and
/// shell that did not change between them, so a snapshot costs what the edit touched rather than
/// what the document holds. <see cref="Materialize"/> rebuilds an independent tree.
/// </summary>
internal sealed class PartSnapshot
{
    private readonly long _shellBytes;

    private PartSnapshot(string partUri, XDocument shell, long shellBytes, FrozenChunk[] chunks)
    {
        PartUri = partUri;
        Shell = shell;
        _shellBytes = shellBytes;
        Chunks = chunks;
    }

    internal string PartUri { get; }

    /// <summary>The part's tree without its blocks. Shared between snapshots while unchanged;
    /// never mutated.</summary>
    internal XDocument Shell { get; }

    internal FrozenChunk[] Chunks { get; }

    /// <summary>Everything this snapshot references, for the undo ring's shared accounting: the
    /// shell and each chunk (whose blocks the ring counts once each).</summary>
    internal IEnumerable<(object Payload, long Bytes)> Payloads
    {
        get
        {
            yield return (Shell, _shellBytes);
            foreach (var chunk in Chunks) yield return (chunk, chunk.OwnBytes);
        }
    }

    /// <summary>A new, independent tree equal to the part at capture time. The caller owns it.</summary>
    internal XDocument Materialize() => Materialize(adopt: null);

    /// <summary>
    /// <see cref="Materialize"/>, then make the new tree <paramref name="part"/>'s cached XML and
    /// teach its snapshot cache that each new block equals the frozen block it was copied from,
    /// and that this snapshot already describes it. The first snapshot after an undo then reuses
    /// this one instead of copying the whole part again.
    /// </summary>
    internal XDocument MaterializeInto(OpenXmlPart part, bool flushToStream)
    {
        var state = new PartSnapshotCache.State();
        var document = Materialize(state);
        if (flushToStream) part.PutXDocument(document);
        else part.SetXDocumentCache(document);
        // Attach after the build: the tree was assembled from these exact blocks, so nothing the
        // tracker could have reported is missing from the seeded state.
        PartChangeTracker.For(document).ForSnapshot.Activate();
        state.Last = this;
        state.Shell = Shell;
        foreach (var chunk in Chunks)
            if (chunk.Blocks.Length > 0) state.ChunkByFirst[chunk.Blocks[0]] = chunk;
        document.AddAnnotation(state);
        return document;
    }

    private XDocument Materialize(PartSnapshotCache.State? adopt)
    {
        var document = new XDocument(Shell);
        var container = PartChangeTracker.ContainerOf(document);
        if (container is null) return document;
        foreach (var chunk in Chunks)
        {
            foreach (var block in chunk.Blocks)
            {
                var live = PartSnapshotCache.CloneNode(block.Node);
                container.Add(live);
                adopt?.FrozenByLive.Add(live, block);
            }
        }
        return document;
    }

    /// <summary>Capture <paramref name="document"/>, reusing <paramref name="state"/>'s frozen blocks
    /// for every block it has not seen change.</summary>
    internal static PartSnapshot Capture(string partUri, XDocument document, PartSnapshotCache.State state,
        bool shellChanged)
    {
        var container = PartChangeTracker.ContainerOf(document);
        XDocument shell;
        if (!shellChanged && state.Shell is { } previousShell)
        {
            shell = previousShell;
        }
        else
        {
            shell = BuildShell(document, container);
            state.Shell = shell;
        }

        var previous = state.Last;
        var chunks = new List<FrozenChunk>(previous?.Chunks.Length ?? 4);
        if (container is not null)
        {
            var chunkByFirst = state.ChunkByFirst;
            var pending = state.Scratch;
            pending.Clear();
            for (var node = container.FirstNode; node is not null; node = node.NextNode)
            {
                if (!state.FrozenByLive.TryGetValue(node, out var frozen))
                {
                    frozen = new FrozenBlock(PartSnapshotCache.CloneNode(node));
                    state.FrozenByLive[node] = frozen;
                }
                pending.Add(frozen);
                // Chunk boundaries are chosen by the frozen block itself (its identity hash), not by
                // position, so inserting, removing or editing a block moves only its own chunk's
                // boundaries and every other chunk keeps matching the previous snapshot's. Frozen
                // identity, not live: a restored tree is new live nodes seeded with the same frozen
                // blocks, so its first snapshot cuts the same chunks as the snapshot it came from.
                if (pending.Count >= PartSnapshotCache.MaxChunk
                    || (RuntimeHelpers.GetHashCode(frozen) & PartSnapshotCache.BoundaryMask) == 0)
                    Flush();
            }
            Flush();
            pending.Clear();

            void Flush()
            {
                if (pending.Count == 0) return;
                if (chunkByFirst.TryGetValue(pending[0], out var reuse) && SameBlocks(reuse.Blocks, pending))
                {
                    chunks.Add(reuse);
                }
                else
                {
                    var chunk = new FrozenChunk(pending.ToArray());
                    chunkByFirst[pending[0]] = chunk;
                    chunks.Add(chunk);
                }
                pending.Clear();
            }
        }

        // A chunk start that moved leaves its old entry behind; once stale entries outnumber live
        // ones, rebuild the map from the chunks in use (amortized, so a snapshot stays cheap).
        if (state.ChunkByFirst.Count > 2 * chunks.Count + 64)
        {
            state.ChunkByFirst.Clear();
            foreach (var chunk in chunks)
                if (chunk.Blocks.Length > 0) state.ChunkByFirst[chunk.Blocks[0]] = chunk;
        }

        var snapshot = new PartSnapshot(partUri, shell,
            ReferenceEquals(shell, previous?.Shell) ? previous!._shellBytes : XmlMemoryEstimator.Estimate(shell),
            chunks.ToArray());
        state.Last = snapshot;
        return snapshot;
    }

    private static bool SameBlocks(FrozenBlock[] blocks, List<FrozenBlock> pending)
    {
        if (blocks.Length != pending.Count) return false;
        for (int i = 0; i < blocks.Length; i++)
            if (!ReferenceEquals(blocks[i], pending[i])) return false;
        return true;
    }

    private static XDocument BuildShell(XDocument document, XElement? container)
    {
        var shell = new XDocument(document.Declaration is { } d ? new XDeclaration(d) : null);
        foreach (var node in document.Nodes())
            shell.Add(node is XElement root ? CopyWithoutBlocks(root, container) : PartSnapshotCache.CloneNode(node));
        return shell;
    }

    /// <summary>A copy of <paramref name="element"/> in which <paramref name="container"/>'s
    /// children are left out (the container itself is kept, empty, with its attributes).</summary>
    private static XElement CopyWithoutBlocks(XElement element, XElement? container)
    {
        if (ReferenceEquals(element, container)) return new XElement(element.Name, element.Attributes());
        var copy = new XElement(element.Name, element.Attributes());
        foreach (var node in element.Nodes())
            copy.Add(node is XElement child ? CopyWithoutBlocks(child, container) : PartSnapshotCache.CloneNode(node));
        return copy;
    }
}

/// <summary>
/// Per-part snapshot reuse (issue #1022). For each live part tree the session keeps a frozen copy
/// of every block, keyed by the live block, and evicts a block's copy when the tree's
/// <see cref="PartChangeTracker"/> reports it changed. A snapshot is then the frozen copies in
/// order: unchanged blocks cost nothing, and an untouched part returns its previous snapshot whole.
/// </summary>
internal static class PartSnapshotCache
{
    /// <summary>Expected chunk length is <c>BoundaryMask + 1</c> blocks.</summary>
    internal const int BoundaryMask = 31;

    internal const int MaxChunk = 256;

    /// <summary>The snapshot cache for one live part tree, stored as an annotation on it. It holds a
    /// frozen copy of every block of the part, which is the price of not copying them per edit.</summary>
    internal sealed class State
    {
        internal Dictionary<XNode, FrozenBlock> FrozenByLive { get; } = new(ReferenceEqualityComparer.Instance);

        /// <summary>The latest chunk starting at each frozen block, so a snapshot can reuse a chunk
        /// the previous one built. Keyed only by blocks still in <see cref="FrozenByLive"/>.</summary>
        internal Dictionary<FrozenBlock, FrozenChunk> ChunkByFirst { get; } = new(ReferenceEqualityComparer.Instance);

        /// <summary>Forget a live block's frozen copy, and any chunk keyed by it.</summary>
        internal void Evict(XNode live)
        {
            if (FrozenByLive.Remove(live, out var frozen)) ChunkByFirst.Remove(frozen);
        }

        internal void Reset()
        {
            FrozenByLive.Clear();
            ChunkByFirst.Clear();
            Last = null;
        }

        internal PartSnapshot? Last { get; set; }

        internal XDocument? Shell { get; set; }

        internal List<FrozenBlock> Scratch { get; } = new();
    }

    /// <summary>Test seam: when set, the cache skips evicting the first changed block it is told
    /// about, so a test can prove the snapshot oracle detects a stale frozen block. Per thread, so
    /// a test cannot disturb a session running concurrently on another.</summary>
    [ThreadStatic]
    internal static bool SkipOneEvictionForTests;

    /// <summary>Snapshot <paramref name="part"/>'s current XML, sharing every block, chunk and shell
    /// unchanged since the previous snapshot of the same tree.</summary>
    internal static PartSnapshot Take(OpenXmlPart part)
    {
        var document = part.GetXDocument();
        var partUri = part.Uri.ToString();
        var state = document.Annotation<State>();
        if (state is null)
        {
            // First sight of this tree: start trusting its tracker from now on, and copy it all.
            var fresh = new State();
            var tracker = PartChangeTracker.For(document);
            tracker.ForSnapshot.Activate();
            document.AddAnnotation(fresh);
            return PartSnapshot.Capture(partUri, document, fresh, shellChanged: true);
        }

        var changes = PartChangeTracker.For(document).ForSnapshot;
        if (!changes.Active)
        {
            // A cache whose tracker stopped recording cannot know what changed: start over.
            changes.Activate();
            state.Reset();
            return PartSnapshot.Capture(partUri, document, state, shellChanged: true);
        }
        var last = state.Last;
        bool declarationChanged = last is not null && !SameDeclaration(last.Shell.Declaration, document.Declaration);
        if (last is not null && !changes.Any && !declarationChanged
            && string.Equals(last.PartUri, partUri, StringComparison.Ordinal))
            return last;

        bool shellChanged = changes.ShellChanged || declarationChanged || last is null;
        if (shellChanged)
        {
            // The container itself may have been replaced; keyed copies of its old children
            // would never be hit again.
            state.Reset();
        }
        else
        {
            bool skip = SkipOneEvictionForTests;
            foreach (var block in changes.Blocks)
            {
                if (skip) { skip = false; SkipOneEvictionForTests = false; continue; }
                state.Evict(block);
            }
        }
        changes.Clear();
        return PartSnapshot.Capture(partUri, document, state, shellChanged);
    }

    private static bool SameDeclaration(XDeclaration? a, XDeclaration? b) =>
        a is null
            ? b is null
            : b is not null && a.Version == b.Version && a.Encoding == b.Encoding && a.Standalone == b.Standalone;

    /// <summary>A deep copy of any node a part's block container can hold.</summary>
    internal static XNode CloneNode(XNode node) => node switch
    {
        XElement element => new XElement(element),
        XCData cdata => new XCData(cdata),
        XText text => new XText(text),
        XComment comment => new XComment(comment),
        XProcessingInstruction pi => new XProcessingInstruction(pi),
        XDocumentType type => new XDocumentType(type),
        _ => throw new InvalidOperationException($"unexpected node type {node.GetType().Name}"),
    };
}
