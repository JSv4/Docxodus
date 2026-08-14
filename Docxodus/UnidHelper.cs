#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Buffers;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;

namespace Docxodus;

/// <summary>
/// Shared helpers for the <c>PtOpenXml.Unid</c> stable-id attribute. The Unid is a 32-char
/// hex string. Two assignment strategies coexist:
/// <list type="bullet">
/// <item><see cref="AssignToAllElements"/> uses random Guids — the legacy behavior that
/// <see cref="WmlComparer"/> relies on for its comparison heuristics. <b>Do not change
/// this call site to the deterministic path.</b> WmlComparer's matching algorithm
/// assumes Unids are content-independent within each version it compares; making
/// them content-addressable causes same-content-but-different-content elements in
/// the two versions (e.g. two distinct images that happen to share a tag-name
/// signature) to be matched by Unid instead of by content, which inflates the
/// detected revision count. The split keeps each consumer pointed at the scheme
/// its algorithm expects.</item>
/// <item><see cref="AssignToAllElementsDeterministic"/> uses content-addressable hashes
/// keyed on element content + structural position, so the same document content
/// produces the same Unids across sessions. Used by <see cref="WmlToMarkdownConverter"/>
/// so an anchor id captured in one <see cref="DocxSession"/> still resolves in a
/// fresh session opened over the same bytes.</item>
/// </list>
/// </summary>
/// <remarks>
/// <para>
/// The deterministic scheme hashes <c>parent_unid : tag_name : content_sig : dup_index</c>
/// — where <c>dup_index</c> is the count of preceding siblings with the same
/// (tag, content_sig). Properties:
/// </para>
/// <list type="bullet">
/// <item>Two opens of the same bytes produce identical Unids on every element.</item>
/// <item>Editing a paragraph's text changes that paragraph's Unid; siblings stay stable.</item>
/// <item>Inserting a unique-content paragraph anywhere does not shift any other Unid.</item>
/// <item>Inserting/editing a duplicate-content paragraph between duplicates shifts the
/// <c>dup_index</c> of later duplicates of the same content (the only rough edge in the scheme).</item>
/// </list>
/// </remarks>
internal static class UnidHelper
{
    /// <summary>Random 32-char hex Unid. Used by the legacy bulk-assign path and by
    /// <see cref="AssignToSelfAndDescendants"/> on freshly-inserted elements that
    /// don't yet have a parent.</summary>
    internal static string GenerateUnid() => Guid.NewGuid().ToString().Replace("-", "");

    /// <summary>
    /// Random-Guid assignment. Assigns a <c>PtOpenXml.Unid</c> attribute to
    /// <paramref name="contentParent"/> (if it is a footnote/endnote root) and to
    /// every descendant that does not already have one. This is the path
    /// <see cref="WmlComparer"/> uses — its matching heuristics expect Unids to be
    /// distinct across siblings regardless of content.
    /// </summary>
    internal static void AssignToAllElements(XElement contentParent)
    {
        if (contentParent.Name == W.footnote || contentParent.Name == W.endnote)
        {
            if (contentParent.Attribute(PtOpenXml.Unid) == null)
            {
                contentParent.Add(new XAttribute(PtOpenXml.Unid, GenerateUnid()));
            }
        }

        foreach (var d in contentParent.Descendants())
        {
            if (d.Attribute(PtOpenXml.Unid) == null)
            {
                d.Add(new XAttribute(PtOpenXml.Unid, GenerateUnid()));
            }
        }
    }

    /// <summary>
    /// Content-addressable assignment. Identical to <see cref="AssignToAllElements"/>
    /// in shape (assigns <c>PtOpenXml.Unid</c> on the root if it's a footnote/endnote
    /// and on every descendant that does not already have one), but the values are
    /// derived deterministically from element content + structural position so the
    /// same bytes produce the same Unids across sessions.
    /// </summary>
    /// <returns><c>true</c> when at least one Unid was assigned — callers use this to
    /// skip persistence work (part flushes) on the no-op passes that dominate per-op
    /// index rebuilds.</returns>
    internal static bool AssignToAllElementsDeterministic(XElement contentParent)
    {
        bool assignedRoot = false;
        if (contentParent.Name == W.footnote || contentParent.Name == W.endnote)
        {
            if (contentParent.Attribute(PtOpenXml.Unid) == null)
            {
                var noteId = (string?)contentParent.Attribute(W.id) ?? string.Empty;
                contentParent.Add(new XAttribute(PtOpenXml.Unid,
                    DeriveUnid(rootSeed: contentParent.Name.LocalName, tag: "id", sig: noteId, dupIndex: 0)));
                assignedRoot = true;
            }
        }

        // Prune the recursion to subtrees that actually contain an element missing a
        // Unid. After the first assignment over a document, the common case (an edit
        // touched one block) leaves 99%+ of the tree fully assigned — recursing into
        // it only to recompute content signatures for dup-index bookkeeping nobody
        // needs is the dominant cost of every per-op index rebuild (36 ms of a 74 ms
        // rebuild on a 15k-element document). `live` = every element whose subtree
        // (self or descendants) holds a missing Unid; parents outside it are skipped
        // wholesale. Assigned values are IDENTICAL to the unpruned walk: dup-index
        // bookkeeping only ever influences an assignment within a parent that has a
        // missing child, and such parents are always in `live`.
        HashSet<XElement>? live = null;
        foreach (var el in contentParent.DescendantsAndSelf())
        {
            if (el.Attribute(PtOpenXml.Unid) is not null || el == contentParent) continue;
            live ??= new HashSet<XElement>();
            for (XElement? a = el.Parent; a is not null; a = a.Parent)
            {
                if (!live.Add(a)) break; // ancestors above are already marked
            }
        }
        if (live is null) return assignedRoot; // fully assigned — nothing to do

        var parentUnid = (string?)contentParent.Attribute(PtOpenXml.Unid) ?? contentParent.Name.LocalName;
        AssignDescendantsDeterministic(contentParent, parentUnid, live);
        return true;
    }

    /// <summary>
    /// Return an element's existing Unid, or derive the exact deterministic value that
    /// <see cref="AssignToAllElementsDeterministic"/> would assign without changing the XML tree.
    /// Read-only inspection fallbacks use this when an element lies outside the normal projected
    /// walk. The derivation follows the ancestor chain and counts same-signature preceding siblings,
    /// matching <see cref="AssignDescendantsDeterministic"/>.
    /// </summary>
    internal static string ReadOrDeriveUnid(XElement element)
    {
        ArgumentNullException.ThrowIfNull(element);
        if ((string?)element.Attribute(PtOpenXml.Unid) is { Length: > 0 } existing)
            return existing;

        var chain = element.AncestorsAndSelf().Reverse().ToArray();
        var root = chain[0];
        string parentUnid;
        if ((string?)root.Attribute(PtOpenXml.Unid) is { Length: > 0 } rootUnid)
        {
            parentUnid = rootUnid;
        }
        else if (root.Name == W.footnote || root.Name == W.endnote)
        {
            var noteId = (string?)root.Attribute(W.id) ?? string.Empty;
            parentUnid = DeriveUnid(root.Name.LocalName, "id", noteId, 0);
        }
        else
        {
            // Scope roots such as w:document/w:hdr/w:ftr are seeds, not assigned descendants.
            parentUnid = root.Name.LocalName;
        }

        for (int i = 1; i < chain.Length; i++)
        {
            var current = chain[i];
            if ((string?)current.Attribute(PtOpenXml.Unid) is { Length: > 0 } currentUnid)
            {
                parentUnid = currentUnid;
                continue;
            }

            var signature = ContentSignature(current);
            int duplicateIndex = 0;
            foreach (var preceding in current.ElementsBeforeSelf())
            {
                if (preceding.Name == current.Name
                    && string.Equals(ContentSignature(preceding), signature, StringComparison.Ordinal))
                    duplicateIndex++;
            }
            parentUnid = DeriveUnid(parentUnid, current.Name.LocalName, signature, duplicateIndex);
        }

        return parentUnid;
    }

    /// <summary>
    /// Like <see cref="AssignToAllElements"/> but also assigns to the root element
    /// itself (regardless of element name). Used for freshly-built block elements
    /// inserted into a document by <c>DocxSession</c>. Uses the random Unid path
    /// because the inserted root often isn't yet attached to a parent at call time;
    /// once saved and reopened, the deterministic projector path will re-derive a
    /// stable Unid for the same slot.
    /// </summary>
    internal static void AssignToSelfAndDescendants(XElement root)
    {
        if (root.Attribute(PtOpenXml.Unid) == null)
            root.Add(new XAttribute(PtOpenXml.Unid, GenerateUnid()));
        foreach (var d in root.Descendants())
        {
            if (d.Attribute(PtOpenXml.Unid) == null)
                d.Add(new XAttribute(PtOpenXml.Unid, GenerateUnid()));
        }
    }

    /// <summary>
    /// Content hash of an element's full subtree — names, attributes (minus the
    /// <c>PtOpenXml.Unid</c> bookkeeping) and text — for change DETECTION, not identity.
    /// In-session a block element keeps its Unid attribute across edits (ops rebuild
    /// its children, not the element), so a renderer diffing by Unid alone would keep
    /// stale DOM for any in-place edit; this hash moves whenever anything inside the
    /// unit changes. One tree walk, no serialization allocation beyond the builder.
    /// </summary>
    internal static string ContentHash(XElement element)
    {
        var sb = new StringBuilder(1024);
        void Walk(XElement e)
        {
            sb.Append('<').Append(e.Name.LocalName);
            // A note's w:id is EXCLUDED — on the reference AND on the definition:
            // inserting a footnote shifts the id of every later citation and its
            // definition (ids ascend in reference order), yet their rendered content is
            // unchanged — marker/list numbering and hrefs are POSITION-derived chrome a
            // renumber pass owns, not content. Hashing the id would make one note insert
            // look like a change to every later citing block and every note definition.
            bool isNoteRef = e.Name == W.footnoteReference || e.Name == W.endnoteReference
                || e.Name == W.footnote || e.Name == W.endnote;
            foreach (var a in e.Attributes())
            {
                if (a.Name == PtOpenXml.Unid || a.IsNamespaceDeclaration) continue;
                if (isNoteRef && a.Name == W.id) continue;
                sb.Append(' ').Append(a.Name.LocalName).Append('=').Append(a.Value);
            }
            sb.Append('>');
            foreach (var n in e.Nodes())
            {
                if (n is XElement ce) Walk(ce);
                else if (n is XText t) sb.Append(t.Value);
            }
            sb.Append("</").Append('>');
        }
        Walk(element);
        return ShortHash(sb.ToString(), hexChars: 16);
    }

    // ─── Deterministic derivation internals ──────────────────────────────

    private static void AssignDescendantsDeterministic(XElement parent, string parentUnid, HashSet<XElement> live)
    {
        // Signature/dup-index bookkeeping is only needed when THIS parent has a child
        // to assign — for fully-assigned parents (the overwhelming majority after the
        // first pass) it would be pure waste, so it is gated on a cheap presence scan.
        bool anyMissing = false;
        foreach (var c in parent.Elements())
        {
            if (c.Attribute(PtOpenXml.Unid) == null) { anyMissing = true; break; }
        }

        var dup = anyMissing ? new Dictionary<(string Tag, string Sig), int>() : null;
        foreach (var child in parent.Elements())
        {
            if (child.Attribute(PtOpenXml.Unid) == null)
            {
                var sig = ContentSignature(child);
                var key = (child.Name.LocalName, sig);
                dup!.TryGetValue(key, out var dupIndex);
                dup[key] = dupIndex + 1;
                child.Add(new XAttribute(PtOpenXml.Unid,
                    DeriveUnid(rootSeed: parentUnid, tag: child.Name.LocalName, sig: sig, dupIndex: dupIndex)));
            }
            else if (anyMissing)
            {
                // Pre-existing Unid (persisted across save, or freshly-inserted via
                // AssignToSelfAndDescendants). Still count it for dup-index of its
                // same-content siblings so unassigned later siblings get a
                // consistent index regardless of which subset already had Unids.
                var sig = ContentSignature(child);
                var key = (child.Name.LocalName, sig);
                dup!.TryGetValue(key, out var dupIndex);
                dup[key] = dupIndex + 1;
            }

            // Recurse only where something below still needs assignment: `live` holds
            // every ancestor of a missing element, so a child with any unassigned
            // descendant — including one just assigned above whose own children are
            // still missing — is a member; anything else has a fully-assigned subtree.
            if (live.Contains(child))
            {
                var childUnid = (string?)child.Attribute(PtOpenXml.Unid)!;
                AssignDescendantsDeterministic(child, childUnid, live);
            }
        }
    }

    /// <summary>
    /// Compact content signature for the element's identity within its parent.
    /// <para>
    /// Container elements (those that have block-level descendants like nested
    /// <c>w:p</c> or <c>w:tbl</c>) get a purely-structural signature — the tag
    /// names of their direct children. Including their descendants' text here
    /// would couple every child's Unid to every other child's content, so
    /// editing one paragraph would shift every other paragraph's Unid via the
    /// shared parent.
    /// </para>
    /// <para>
    /// Leaf-ish elements (paragraphs, runs, table cells) include their flat
    /// <c>w:t</c> text plus style id + numbering id + tag names of non-text
    /// descendants (so text-empty paragraphs holding distinct images / math /
    /// fields get distinct sigs).
    /// </para>
    /// </summary>
    private static string ContentSignature(XElement element)
    {
        // Container elements (those that contain block-level descendants like
        // nested w:p or w:tbl) collapse to a tag-name-only signature. Their
        // Unid is used as parent_unid for the blocks inside; we don't want
        // editing/inserting one paragraph to invalidate every other paragraph
        // by shifting their parent's signature.
        bool hasBlockDescendants = element.Descendants().Any(d => d.Name == W.p || d.Name == W.tbl);
        if (hasBlockDescendants)
        {
            return ShortHash(element.Name.LocalName, hexChars: 16);
        }

        var text = string.Concat(element.Descendants(W.t).Select(t => (string)t));
        var pPr = element.Element(W.pPr);
        var styleId = pPr?.Element(W.pStyle)?.Attribute(W.val)?.Value ?? string.Empty;
        var numId = pPr?.Element(W.numPr)?.Element(W.numId)?.Attribute(W.val)?.Value ?? string.Empty;
        var sb2 = new StringBuilder(text.Length + 64);
        sb2.Append(text).Append('|').Append(styleId).Append('|').Append(numId).Append('|');
        foreach (var d in element.Descendants())
        {
            if (d.Name == W.t) continue;
            sb2.Append(d.Name.LocalName).Append(',');
        }
        return ShortHash(sb2.ToString(), hexChars: 16);
    }

    private static string DeriveUnid(string rootSeed, string tag, string sig, int dupIndex)
    {
        var input = rootSeed + ":" + tag + ":" + sig + ":" + dupIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return ShortHash(input, hexChars: 32);
    }

    /// <remarks>
    /// Allocation-free apart from the returned string. Same bytes in, same lowercase-hex digest
    /// prefix out as the <c>SHA256.Create()</c> + <c>GetBytes</c> + <c>StringBuilder</c> version
    /// it replaces. This runs once per block for every render plan and once per element for every
    /// Unid derivation — hundreds of calls per interactive operation — so the per-call garbage was
    /// worth removing even though it does not dominate on a JIT-ed desktop runtime.
    /// </remarks>
    internal static string ShortHash(string input, int hexChars)
    {
        Span<byte> digest = stackalloc byte[32];
        Span<byte> inline = stackalloc byte[1024];
        int maxBytes = Encoding.UTF8.GetMaxByteCount(input.Length);
        byte[]? rented = maxBytes <= inline.Length ? null : ArrayPool<byte>.Shared.Rent(maxBytes);
        Span<byte> buffer = rented is null ? inline : rented;

        try
        {
            int written = Encoding.UTF8.GetBytes(input, buffer);
            SHA256.HashData(buffer[..written], digest);
            return Convert.ToHexStringLower(digest[..(hexChars / 2)]);
        }
        finally
        {
            if (rented is not null) ArrayPool<byte>.Shared.Return(rented);
        }
    }
}
