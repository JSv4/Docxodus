#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
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
    internal static void AssignToAllElementsDeterministic(XElement contentParent)
    {
        if (contentParent.Name == W.footnote || contentParent.Name == W.endnote)
        {
            if (contentParent.Attribute(PtOpenXml.Unid) == null)
            {
                var noteId = (string?)contentParent.Attribute(W.id) ?? string.Empty;
                contentParent.Add(new XAttribute(PtOpenXml.Unid,
                    DeriveUnid(rootSeed: contentParent.Name.LocalName, tag: "id", sig: noteId, dupIndex: 0)));
            }
        }

        var parentUnid = (string?)contentParent.Attribute(PtOpenXml.Unid) ?? contentParent.Name.LocalName;
        AssignDescendantsDeterministic(contentParent, parentUnid);
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

    // ─── Deterministic derivation internals ──────────────────────────────

    private static void AssignDescendantsDeterministic(XElement parent, string parentUnid)
    {
        var dup = new Dictionary<(string Tag, string Sig), int>();
        foreach (var child in parent.Elements())
        {
            if (child.Attribute(PtOpenXml.Unid) == null)
            {
                var sig = ContentSignature(child);
                var key = (child.Name.LocalName, sig);
                dup.TryGetValue(key, out var dupIndex);
                dup[key] = dupIndex + 1;
                child.Add(new XAttribute(PtOpenXml.Unid,
                    DeriveUnid(rootSeed: parentUnid, tag: child.Name.LocalName, sig: sig, dupIndex: dupIndex)));
            }
            else
            {
                // Pre-existing Unid (persisted across save, or freshly-inserted via
                // AssignToSelfAndDescendants). Still count it for dup-index of its
                // same-content siblings so unassigned later siblings get a
                // consistent index regardless of which subset already had Unids.
                var sig = ContentSignature(child);
                var key = (child.Name.LocalName, sig);
                dup.TryGetValue(key, out var dupIndex);
                dup[key] = dupIndex + 1;
            }

            var childUnid = (string?)child.Attribute(PtOpenXml.Unid)!;
            AssignDescendantsDeterministic(child, childUnid);
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

    private static string ShortHash(string input, int hexChars)
    {
        using var sha = SHA256.Create();
        var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(input));
        var byteCount = hexChars / 2;
        var sb = new StringBuilder(hexChars);
        for (int i = 0; i < byteCount; i++)
            sb.Append(bytes[i].ToString("x2", System.Globalization.CultureInfo.InvariantCulture));
        return sb.ToString();
    }
}
