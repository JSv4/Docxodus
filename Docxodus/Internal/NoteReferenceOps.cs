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
/// <c>DocxDiffOps</c> (issues #614, #631, #636). Every caller captures
/// <see cref="ReferencedNoteIds"/> before the mutation and prunes after it — the session's
/// editorial resolves via <see cref="PruneOrphanedNotes"/>, both stateless resolutions via
/// <see cref="PruneNotesEmptiedByResolution"/>, whose guard the comparison contracts need.
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

        // No notes parts means nothing a prune could ever remove, and every caller captures these
        // ids only to prune against them afterwards — skip the package-wide scans instead of
        // walking the body and every header/footer twice per resolution on note-free documents.
        if (main is null || (main.FootnotesPart is null && main.EndnotesPart is null))
            return (footnotes, endnotes);
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
        Prune(main, before, onlyEmptied: false);

    /// <summary>
    /// The guarded variant of the rule for the stateless resolutions (issues #631, #636): remove a
    /// note definition whose last citation disappeared between <paramref name="before"/> and now
    /// AND which the resolution left without block content. A resolution can only strip a note
    /// bare when every block was revision-marked in its direction — an unmarked paragraph survives
    /// its runs — so a bare, newly-uncited definition is the redline saying "this whole note goes
    /// with its citation": a wholly-deleted note on accept (every block in <c>w:del</c>), a
    /// wholly-inserted note on reject (every block in <c>w:ins</c>). The emptied guard is what
    /// distinguishes both from a document's own reference-less husk, which arrives with its
    /// content unmarked, keeps it through the resolution, and stays — so
    /// <c>Accept(Compare(l, r)) ≡ r</c> and <c>Reject(Compare(l, r)) ≡ l</c> both hold for the
    /// note stores. (Redlines whose producer left an inserted definition's TEXT unmarked —
    /// sessions from the brief #614-era — reject to an orphaned husk instead; see CHANGELOG.)
    /// <para>"Emptied" is scaffolding-aware, not merely blockless: <see cref="WmlComparer"/>
    /// re-synthesizes an UNMARKED <c>w:footnoteRef</c> marker run into every inserted definition
    /// it renders, so rejecting one of its redlines strips the definition to a marker-only
    /// paragraph rather than to nothing. A definition left with no text, tables, SDTs or other
    /// run content beyond note-reference-mark scaffolding is emptied.</para>
    /// <para>Two interlocks worth knowing. A still-CITED note a resolution strips bare is left
    /// bare on purpose: a cited <c>w:footnote</c> with no block child is schema-degenerate, but it
    /// is a shape the corpus genuinely contains (<c>WC064-Footnote</c> ships one, and
    /// <c>WC063-Footnote-Mod</c>'s counterpart note is childless), so reproducing it is what
    /// <c>Accept(Compare(l, r)) ≡ r</c> means — any "repair" here broke those round-trips when it
    /// was tried. And the session's editorial resolves stay on the UNGUARDED
    /// <see cref="PruneOrphanedNotes"/> deliberately: the session resolver keeps the last block
    /// of a note alive (<c>RevisionOps</c>' surviving-block guard), so a session-emptied
    /// definition can never satisfy this predicate — unifying the session onto the guarded prune
    /// would silently stop pruning session-rejected recorded notes.</para>
    /// </summary>
    internal static List<(XElement Note, string PartUri)> PruneNotesEmptiedByResolution(
        MainDocumentPart? main, (HashSet<int> Footnotes, HashSet<int> Endnotes) before) =>
        Prune(main, before, onlyEmptied: true);

    private static List<(XElement Note, string PartUri)> Prune(
        MainDocumentPart? main, (HashSet<int> Footnotes, HashSet<int> Endnotes) before,
        bool onlyEmptied)
    {
        var removed = new List<(XElement, string)>();
        if (main is null) return removed;
        var after = ReferencedNoteIds(main);
        PruneStory(main.FootnotesPart, W + "footnote", before.Footnotes, after.Footnotes, removed, onlyEmptied);
        PruneStory(main.EndnotesPart, W + "endnote", before.Endnotes, after.Endnotes, removed, onlyEmptied);
        return removed;
    }

    /// <summary>Whether a note definition was stripped to scaffolding: no block children at all,
    /// or only paragraphs whose remaining run content is note-reference-mark scaffolding
    /// (<c>w:footnoteRef</c>/<c>w:endnoteRef</c>) — the marker run <see cref="WmlComparer"/>
    /// re-synthesizes UNMARKED into inserted definitions, which would otherwise shield them from
    /// the guarded prune. Tables, SDTs, any text, and any other run content mean real content
    /// survived. A direct-children walk is sufficient HERE — and only here — because this runs
    /// after <see cref="RevisionProcessor"/>'s resolution, whose block-content coalescing pass
    /// (<c>AcceptDeletedAndMoveFromParagraphMarksTransform</c>, applied to every note via
    /// <c>BlockLevelContentContainers</c>) promotes block content out of wrappers like
    /// <c>mc:AlternateContent</c> to direct children. It is deliberately narrower than
    /// <c>RevisionProcessor.BlockLevelElements</c> and <c>IsBlockContentElement</c>, which
    /// classify pre-resolution trees that still carry revision wrappers; if the pipeline ever
    /// starts leaving a new wrapper shape inside notes, this predicate must learn it too.</summary>
    private static bool IsEmptiedToScaffolding(XElement note)
    {
        foreach (var block in note.Elements())
        {
            if (block.Name == W + "tbl" || block.Name == W + "sdt") return false;
            if (block.Name != W + "p") continue;
            foreach (var child in block.Elements())
            {
                if (child.Name == W + "pPr"
                    || child.Name == W + "bookmarkStart" || child.Name == W + "bookmarkEnd"
                    || child.Name == W + "proofErr") continue;
                if (child.Name != W + "r") return false;
                if (child.Elements().Any(rc =>
                        rc.Name != W + "rPr"
                        && rc.Name != W + "footnoteRef" && rc.Name != W + "endnoteRef"))
                    return false;
            }
        }

        return true;
    }

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
        bool onlyEmptied)
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
            if (onlyEmptied && !IsEmptiedToScaffolding(note))
                continue;
            note.Remove();
            removed.Add((note, partUri));
            changed = true;
        }

        if (changed) part.PutXDocument();
    }
}
