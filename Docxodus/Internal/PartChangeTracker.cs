// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>
/// Records which top-level blocks of one live part tree have changed, from the tree's own
/// LINQ-to-XML change events, so a session can repair its caches block by block instead of
/// rebuilding them from the whole document after every edit (issue #1022).
/// </summary>
/// <remarks>
/// <para><b>Blocks and the shell.</b> A part's <em>block container</em> is <c>w:body</c> for the
/// main document and the root element for every other part (<c>w:hdr</c>, <c>w:footnotes</c>,
/// <c>w:comments</c>, <c>w:styles</c>, …). Its child nodes are the blocks. Everything else — the
/// root element's own attributes, the container's attributes, nodes beside the container, the
/// container itself — is the <em>shell</em>. A change inside a block marks that block; a change
/// anywhere in the shell marks the shell, and consumers treat a shell change as "rebuild this
/// part from scratch".</para>
/// <para><b>Why events rather than per-op declarations.</b> Ops have side effects outside the
/// block they address (range markers and note references that cross blocks, styles a markdown
/// payload creates, Save's strip-and-restore of every Unid). An event is raised for every one of
/// them by LINQ to XML itself, so no op can forget to report a change, and a new op needs no
/// bookkeeping.</para>
/// <para><b>Lifetime.</b> One tracker per <see cref="XDocument"/> instance, stored as an annotation
/// on it. When a part's cached tree is replaced (undo, redo, rollback, package reopen), the new
/// tree has no tracker, and every consumer sees a tree it has no record of and starts over. A
/// clone (<c>new XDocument(doc)</c>) does not copy annotations, so snapshots never carry one.</para>
/// <para>Each consumer reads and clears its own <see cref="Channel"/>, so the snapshot cache and
/// the anchor index can consume changes at different moments.</para>
/// </remarks>
internal sealed class PartChangeTracker
{
    /// <summary>The changes one consumer has not yet processed.</summary>
    internal sealed class Channel
    {
        /// <summary>Container children (by live node) changed, added or removed since the last
        /// <see cref="Clear"/>. A removed block appears here too; its <c>Parent</c> is then no
        /// longer the container.</summary>
        internal HashSet<XNode> Blocks { get; } = new(ReferenceEqualityComparer.Instance);

        /// <summary>Something outside every block changed.</summary>
        internal bool ShellChanged { get; set; }

        internal bool Any => ShellChanged || Blocks.Count > 0;

        /// <summary>Whether a consumer reads this channel. An inactive channel records nothing, so
        /// a tree nobody is maintaining a cache for does not accumulate references to every block
        /// it ever changed or removed.</summary>
        internal bool Active { get; private set; }

        /// <summary>Start recording, from a clean slate: the consumer has just read the tree in full.</summary>
        internal void Activate()
        {
            Clear();
            Active = true;
        }

        /// <summary>Stop recording and drop what was recorded.</summary>
        internal void Deactivate()
        {
            Active = false;
            Clear();
        }

        internal void Clear()
        {
            Blocks.Clear();
            ShellChanged = false;
        }

        internal void AddBlock(XNode block)
        {
            if (Active) Blocks.Add(block);
        }

        internal void MarkShell()
        {
            if (Active) ShellChanged = true;
        }
    }

    private readonly XDocument _document;

    private PartChangeTracker(XDocument document)
    {
        _document = document;
        document.Changing += OnChanging;
        document.Changed += OnChanged;
    }

    /// <summary>Changes not yet folded into the session's snapshot cache.</summary>
    internal Channel ForSnapshot { get; } = new();

    /// <summary>Changes not yet folded into the session's anchor index.</summary>
    internal Channel ForIndex { get; } = new();

    /// <summary>The tracker for <paramref name="document"/>, attached on first use. A tracker
    /// attached now has seen nothing that happened before, so a consumer must only trust it
    /// from the moment it first read the tree in full.</summary>
    internal static PartChangeTracker For(XDocument document)
    {
        var tracker = document.Annotation<PartChangeTracker>();
        if (tracker is null)
        {
            tracker = new PartChangeTracker(document);
            document.AddAnnotation(tracker);
        }
        return tracker;
    }

    /// <summary>The tracker already attached to <paramref name="document"/>, or null.</summary>
    internal static PartChangeTracker? Existing(XDocument document) => document.Annotation<PartChangeTracker>();

    /// <summary>The block container of a part tree: <c>w:body</c> under the root when present,
    /// otherwise the root itself. Null for a tree with no root.</summary>
    internal static XElement? ContainerOf(XDocument document)
    {
        var root = document.Root;
        if (root is null) return null;
        return root.Element(W.body) ?? root;
    }

    // A removal is reported while the node is still attached (Changing), so its block can be
    // found; every other change is reported once the node is in place (Changed), because an
    // added node has no parent yet when Changing fires.
    private void OnChanging(object? sender, XObjectChangeEventArgs e)
    {
        if (e.ObjectChange == XObjectChange.Remove) Mark(sender as XObject);
    }

    private void OnChanged(object? sender, XObjectChangeEventArgs e)
    {
        if (e.ObjectChange == XObjectChange.Remove) return;
        // Renaming the root or one of its children can change which element is the block
        // container (a renamed w:body), so it is a shell change however it looks afterwards.
        if (e.ObjectChange == XObjectChange.Name && sender is XElement renamed
            && (renamed.Parent is null || ReferenceEquals(renamed.Parent, _document.Root)))
        {
            MarkShell();
            return;
        }
        Mark(sender as XObject);
    }

    private void Mark(XObject? changed)
    {
        XNode? node = changed switch
        {
            XAttribute attribute => attribute.Parent,
            XNode n => n,
            _ => null,
        };
        var container = ContainerOf(_document);
        if (node is null || container is null || ReferenceEquals(node, container))
        {
            MarkShell();
            return;
        }

        var current = node;
        while (true)
        {
            var parent = current.Parent;
            if (ReferenceEquals(parent, container))
            {
                ForSnapshot.AddBlock(current);
                ForIndex.AddBlock(current);
                return;
            }
            if (parent is null)
            {
                MarkShell();
                return;
            }
            current = parent;
        }
    }

    private void MarkShell()
    {
        ForSnapshot.MarkShell();
        ForIndex.MarkShell();
    }
}
