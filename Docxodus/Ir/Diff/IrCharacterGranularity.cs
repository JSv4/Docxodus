#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Character-level change granularity — Word Compare's "Show changes at: Character level" radio
/// (<see cref="IrDiffSettings.ChangeGranularity"/>). A pure REFINEMENT of already-produced word-level
/// output: the tokenizer, aligner and edit script stay word-grained (character-level ALIGNMENT would
/// wreck block similarity, move detection and the content-hash correspondence), and only the rendered
/// revision — markup or revision list — is narrowed to the characters that actually differ.
/// </summary>
/// <remarks>
/// <para><b>Why a post-pass is exactly right here, not a shortcut.</b> Take a <c>w:ins</c>/<c>w:del</c>
/// pair over the same word, e.g. ins <c>color</c> beside del <c>colour</c>. Accept keeps the ins text and
/// drops the del; reject does the reverse. Moving a character that both texts SHARE out of both wrappers
/// into a plain run therefore changes neither result — the character was going to survive either way. So
/// this transform preserves <c>accept ≡ right</c> and <c>reject ≡ left</c> by construction, for any pair,
/// however it arose (an ordinary word replacement, or a move simplified to del/ins).</para>
/// <para><b>Scope guard.</b> The refinement applies only to a del/ins sibling pair where each wrapper holds
/// exactly ONE run whose sole content is one text element, and the two runs' <c>w:rPr</c> are equal. The rPr
/// requirement is a correctness condition, not convenience: when the word's FORMAT changed too, the shared
/// characters are not really unchanged, and lifting them into one plain run would have to pick a side and
/// would silently drop the format change. Any richer pair is left word-level.</para>
/// </remarks>
internal static class IrCharacterGranularity
{
    /// <summary>
    /// Lengths of the longest common prefix and, over what remains after it, the longest common suffix of
    /// <paramref name="a"/> and <paramref name="b"/>, each backed off to a TEXT-ELEMENT (grapheme-cluster)
    /// boundary in BOTH strings. The two never overlap.
    /// </summary>
    /// <remarks>
    /// The back-off is not cosmetic. Comparing UTF-16 code units alone splits an astral-plane character —
    /// two emoji or two CJK Ext-B ideographs from the same block share a high surrogate, so the raw prefix
    /// stops between the surrogates and the produced runs carry LONE surrogates, which are not writable as
    /// XML (the serializer throws). Combining marks have a milder version of the same problem: cutting
    /// <c>café</c>/<c>cafe</c> between <c>e</c> and U+0301 would mark a bare floating accent. Checking the
    /// boundary in BOTH strings mirrors <c>IrRevisionRenderer.TrimCommonWordAffixes</c>'s word-boundary rule.
    /// </remarks>
    public static (int Prefix, int Suffix) CommonAffixes(string a, string b)
    {
        int n = Math.Min(a.Length, b.Length);

        int prefix = 0;
        while (prefix < n && a[prefix] == b[prefix])
            prefix++;
        var aStarts = TextElementStarts(a);
        var bStarts = TextElementStarts(b);
        while (prefix > 0 && !(aStarts.Contains(prefix) && bStarts.Contains(prefix)))
            prefix--;

        int remaining = n - prefix;
        int suffix = 0;
        while (suffix < remaining && a[a.Length - 1 - suffix] == b[b.Length - 1 - suffix])
            suffix++;
        while (suffix > 0 &&
               !(aStarts.Contains(a.Length - suffix) && bStarts.Contains(b.Length - suffix)))
            suffix--;

        return (prefix, suffix);
    }

    /// <summary>
    /// Every index at which a text element (grapheme cluster) starts, plus the string's end — the positions
    /// a cut may legally land on.
    /// </summary>
    private static HashSet<int> TextElementStarts(string s)
    {
        var starts = new HashSet<int>(StringInfo.ParseCombiningCharacters(s)) { s.Length };
        return starts;
    }

    /// <summary>
    /// Narrow every eligible <c>w:ins</c>/<c>w:del</c> sibling pair in one part root to the characters that
    /// differ, lifting the shared prefix/suffix into plain runs around it. Returns true when anything changed.
    /// </summary>
    public static bool RefineInRoot(XElement root)
    {
        bool changed = false;
        foreach (var ins in root.Descendants(W.ins).ToList())
        {
            if (ins.Parent is null)
                continue; // consumed by an earlier refinement

            var next = ins.ElementsAfterSelf().FirstOrDefault();
            var previous = ins.ElementsBeforeSelf().LastOrDefault();
            var del = next?.Name == W.del ? next
                : previous?.Name == W.del ? previous
                : null;
            if (del is null)
                continue;

            bool insFirst = ReferenceEquals(del, next);
            changed |= TryRefinePair(
                first: insFirst ? ins : del,
                second: insFirst ? del : ins,
                ins: ins,
                del: del);
        }
        return changed;
    }

    private static bool TryRefinePair(XElement first, XElement second, XElement ins, XElement del)
    {
        if (SoleRun(ins) is not { } insRun || SoleRun(del) is not { } delRun)
            return false;
        if (SoleTextElement(insRun, W.t) is not { } insText ||
            SoleTextElement(delRun, W.delText) is not { } delText)
            return false;

        var insProps = insRun.Element(W.rPr);
        if (!RunPropsEqual(insProps, delRun.Element(W.rPr)))
            return false;

        // Same author AND date. Every pair this engine produces carries the settings' single author/date, so
        // the condition is free for our own output — and it excludes pairs we did not produce: input revisions
        // carried through under PreserveInputRevisions, and the per-reviewer wrappers a multi-author document
        // can place side by side. Lifting a shared character out of a pair authored by two different people
        // would silently drop one author's claim to it.
        if (!AttributeEqual(ins, del, W.author) || !AttributeEqual(ins, del, W.date))
            return false;

        string inserted = insText.Value;
        string deleted = delText.Value;
        var (prefix, suffix) = CommonAffixes(inserted, deleted);
        if (prefix == 0 && suffix == 0)
            return false;

        // Equal texts (a pair whose change is elsewhere — a retargeted hyperlink keeps its display text) are
        // not a character-level change. Refining would empty BOTH wrappers and erase the revision entirely,
        // losing the author, the date and the fact that anything changed.
        if (prefix + suffix == inserted.Length && prefix + suffix == deleted.Length)
            return false;

        if (prefix > 0)
            first.AddBeforeSelf(PlainRun(insProps, inserted.Substring(0, prefix)));
        if (suffix > 0)
            second.AddAfterSelf(PlainRun(insProps, inserted.Substring(inserted.Length - suffix)));

        insText.Value = inserted.Substring(prefix, inserted.Length - prefix - suffix);
        delText.Value = deleted.Substring(prefix, deleted.Length - prefix - suffix);
        Preserve(insText);
        Preserve(delText);

        // A side whose differing middle is empty is no longer a revision at all (a pure insertion or pure
        // deletion of characters inside the word).
        if (insText.Value.Length == 0)
            ins.Remove();
        if (delText.Value.Length == 0)
            del.Remove();
        return true;
    }

    /// <summary>The wrapper's single <c>w:r</c> child, or null when it holds anything else.</summary>
    private static XElement? SoleRun(XElement wrapper)
    {
        var children = wrapper.Elements().ToList();
        return children.Count == 1 && children[0].Name == W.r ? children[0] : null;
    }

    /// <summary>The run's single text element of <paramref name="name"/>, ignoring its <c>w:rPr</c>.</summary>
    private static XElement? SoleTextElement(XElement run, XName name)
    {
        var content = run.Elements().Where(e => e.Name != W.rPr).ToList();
        return content.Count == 1 && content[0].Name == name ? content[0] : null;
    }

    private static bool AttributeEqual(XElement a, XElement b, XName name) =>
        (string?)a.Attribute(name) == (string?)b.Attribute(name);

    private static bool RunPropsEqual(XElement? a, XElement? b) =>
        (a is null && b is null) || (a is not null && b is not null && XNode.DeepEquals(a, b));

    private static XElement PlainRun(XElement? runProps, string text)
    {
        var t = new XElement(W.t, text);
        Preserve(t);
        return runProps is null
            ? new XElement(W.r, t)
            : new XElement(W.r, new XElement(runProps), t);
    }

    /// <summary>A mid-word slice can begin or end with a space, so every produced text keeps it.</summary>
    private static void Preserve(XElement text) =>
        text.SetAttributeValue(XNamespace.Xml + "space", "preserve");

    /// <summary>
    /// Narrow adjacent opposite-kind revision pairs from the same block to the characters that differ,
    /// dropping a side whose differing middle is empty. The revision-list twin of
    /// <see cref="RefineInRoot"/>, so <c>Compare</c> and <c>GetRevisions</c> report the same grain.
    /// </summary>
    public static void RefineRevisions(List<IrRevision> revisions)
    {
        for (int i = 0; i + 1 < revisions.Count; i++)
        {
            var a = revisions[i];
            var b = revisions[i + 1];
            if (!IsRefinablePair(a, b))
                continue;

            var deleted = a.Type == IrRevisionType.Deleted ? a : b;
            var inserted = ReferenceEquals(deleted, a) ? b : a;
            var (prefix, suffix) = CommonAffixes(inserted.Text, deleted.Text);
            if (prefix == 0 && suffix == 0)
                continue;

            string newDeleted = Middle(deleted.Text, prefix, suffix);
            string newInserted = Middle(inserted.Text, prefix, suffix);

            // Equal texts are not a character-level change — see the markup twin. Dropping both would erase
            // a real change (a retargeted hyperlink keeps its display text) from the revision list while the
            // markup still carries it, so the two surfaces would disagree.
            if (newDeleted.Length == 0 && newInserted.Length == 0)
                continue;

            revisions[i] = a with { Text = ReferenceEquals(a, deleted) ? newDeleted : newInserted };
            revisions[i + 1] = b with { Text = ReferenceEquals(b, deleted) ? newDeleted : newInserted };

            // Remove the emptied side (at most one, per the guard above) and rescan from the survivor, so a
            // three-revision run (del, ins, del) still gets its second pair considered.
            if (revisions[i + 1].Text.Length == 0)
                revisions.RemoveAt(i + 1);
            else if (revisions[i].Text.Length == 0)
                revisions.RemoveAt(i);
            i--;
        }
    }

    /// <summary>
    /// A pair is refinable when it is an Inserted/Deleted pair over the SAME block on both sides (so the two
    /// really are the two halves of one word replacement) and neither half is part of a move. (The move test
    /// is deliberately explicit rather than left to the type and anchor tests that happen to also exclude
    /// moves today: never narrowing across a relocation is the intent.)
    /// </summary>
    private static bool IsRefinablePair(IrRevision a, IrRevision b) =>
        a.MoveGroupId is null && b.MoveGroupId is null &&
        ((a.Type == IrRevisionType.Inserted && b.Type == IrRevisionType.Deleted) ||
         (a.Type == IrRevisionType.Deleted && b.Type == IrRevisionType.Inserted)) &&
        a.LeftAnchor is not null && a.LeftAnchor == b.LeftAnchor &&
        a.RightAnchor is not null && a.RightAnchor == b.RightAnchor;

    private static string Middle(string s, int prefix, int suffix) =>
        s.Substring(prefix, s.Length - prefix - suffix);
}
