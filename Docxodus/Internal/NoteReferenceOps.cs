#nullable enable

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// The single owner of Word's note-lifecycle rule: <b>a footnote/endnote definition exists exactly
/// as long as something still cites it.</b> Word deletes a note when its last reference goes, and
/// every path in this library that can carry a reference away has to do the same or the note
/// survives as a reference-less husk that still ships its full text inside the package.
/// </summary>
/// <remarks>
/// Two callers, one rule. <see cref="DocxSession"/> applies it after a structural delete or a
/// revision resolution (issues #516, #591); <see cref="RevisionProcessor"/> applies it after a
/// stateless accept/reject, which is the path every non-.NET transport reaches through
/// <c>DocxDiffOps</c> (issues #614, #631). Every caller captures <see cref="ReferencedNoteIds"/>
/// before the mutation and prunes after it — via <see cref="PruneOrphanedNotes"/>, except the
/// stateless accept, whose <see cref="PruneNotesEmptiedByAccept"/> adds a guard the comparison
/// contract needs.
///
/// The PART itself always stays. Word never prunes a notes part — the RP050-Deleted-Footnote
/// oracle keeps its separator-only <c>footnotes.xml</c> after accepting the only note's deletion,
/// and a separator-only part is exactly what Word ships in every fresh document.
///
/// The rule is scoped strictly to ids referenced <em>before</em> and unreferenced <em>after</em>
/// one operation: a note that was already dangling on the way in is pre-existing document state
/// and is left alone.
/// </remarks>
internal static class NoteReferenceOps
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>The parts a <c>w:footnoteReference</c>/<c>w:endnoteReference</c> can legally live
    /// in: the body, the running header/footer stories, and the note stories themselves (a note may
    /// cite another note).</summary>
    internal static IEnumerable<OpenXmlPart> ReferenceHostParts(MainDocumentPart main)
    {
        yield return main;
        foreach (var header in main.HeaderParts) yield return header;
        foreach (var footer in main.FooterParts) yield return footer;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart;
    }

    /// <summary>Footnote/endnote ids cited from anywhere in the package. Every host story counts,
    /// not just the body — a note cited from a header as well as the body must survive the body
    /// citation going away.</summary>
    internal static (HashSet<int> Footnotes, HashSet<int> Endnotes) ReferencedNoteIds(
        MainDocumentPart? main)
    {
        var footnotes = new HashSet<int>();
        var endnotes = new HashSet<int>();
        if (main is null) return (footnotes, endnotes);
        foreach (var part in ReferenceHostParts(main))
        {
            var root = part.GetXDocument().Root;
            if (root is null) continue;
            Collect(root, W + "footnoteReference", footnotes);
            Collect(root, W + "endnoteReference", endnotes);
        }

        return (footnotes, endnotes);

        static void Collect(XElement root, XName referenceName, HashSet<int> into)
        {
            foreach (var reference in root.Descendants(referenceName))
                if (int.TryParse((string?)reference.Attribute(W + "id"), out var id)) into.Add(id);
        }
    }

    /// <summary>Remove every note definition whose last citation disappeared between
    /// <paramref name="before"/> and now.</summary>
    /// <returns>The removed note elements with their part uri, for removed-anchor reporting.</returns>
    internal static List<(XElement Note, string PartUri)> PruneOrphanedNotes(
        MainDocumentPart? main, (HashSet<int> Footnotes, HashSet<int> Endnotes) before) =>
        Prune(main, before, onlyBlockless: false);

    /// <summary>
    /// The accept-side variant of the rule (issue #631): remove a note definition whose last
    /// citation disappeared between <paramref name="before"/> and now AND which is left without
    /// block content. Accepting can only strip a note bare when every block was deletion-marked —
    /// an unmarked paragraph survives its runs — so a bare, newly-uncited definition is the
    /// redline saying "this whole note is deleted", and accepting must take the definition with
    /// it or the accepted package ships a note the counterpart deleted. The blockless guard is
    /// what distinguishes that from a counterpart's own reference-less husk: a definition the
    /// counterpart kept arrives with its content unmarked, keeps that content through the accept,
    /// and stays — so <c>Accept(Compare(l, r)) ≡ r</c> holds for both note stores.
    /// </summary>
    internal static List<(XElement Note, string PartUri)> PruneNotesEmptiedByAccept(
        MainDocumentPart? main, (HashSet<int> Footnotes, HashSet<int> Endnotes) before) =>
        Prune(main, before, onlyBlockless: true);

    private static List<(XElement Note, string PartUri)> Prune(
        MainDocumentPart? main, (HashSet<int> Footnotes, HashSet<int> Endnotes) before,
        bool onlyBlockless)
    {
        var removed = new List<(XElement, string)>();
        if (main is null) return removed;
        var after = ReferencedNoteIds(main);
        PruneStory(main.FootnotesPart, W + "footnote", before.Footnotes, after.Footnotes, removed, onlyBlockless);
        PruneStory(main.EndnotesPart, W + "endnote", before.Endnotes, after.Endnotes, removed, onlyBlockless);
        return removed;
    }

    /// <summary>Whether a note definition has no block-level content left (no <c>w:p</c>,
    /// <c>w:tbl</c> or <c>w:sdt</c> children).</summary>
    private static bool IsBlockless(XElement note) =>
        !note.Elements().Any(e =>
            e.Name == W + "p" || e.Name == W + "tbl" || e.Name == W + "sdt");

    /// <summary>Word's reserved scaffolding notes, which are never citation targets and are never
    /// pruned.</summary>
    internal static bool IsSeparatorNote(XElement note) =>
        (string?)note.Attribute(W + "type") is "separator" or "continuationSeparator";

    private static void PruneStory(
        OpenXmlPart? part,
        XName noteName,
        HashSet<int> referencedBefore,
        HashSet<int> referencedAfter,
        List<(XElement, string)> removed,
        bool onlyBlockless)
    {
        var root = part?.GetXDocument().Root;
        if (part is null || root is null) return;
        var orphaned = referencedBefore.Except(referencedAfter).ToHashSet();
        if (orphaned.Count == 0) return;

        var partUri = part.Uri.ToString();
        bool changed = false;
        foreach (var note in root.Elements(noteName).ToList())
        {
            if (!int.TryParse((string?)note.Attribute(W + "id"), out var id)
                || !orphaned.Contains(id) || IsSeparatorNote(note))
                continue;
            if (onlyBlockless && !IsBlockless(note))
                continue;
            note.Remove();
            removed.Add((note, partUri));
            changed = true;
        }

        if (changed) part.PutXDocument();
    }
}
