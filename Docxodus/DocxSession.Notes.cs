// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.Verification;
using DocumentFormat.OpenXml.Packaging;
using GridCell = Docxodus.Internal.TableGridCell;

namespace Docxodus;

public sealed partial class DocxSession
{
    /// <summary>
    /// The document's footnotes (or endnotes) in citation order — see
    /// <see cref="NoteListEntry"/>. References are collected from the main body in
    /// document order (Word disallows note references anywhere else this API can
    /// author them); a reference whose definition is missing is skipped.
    /// </summary>
    public IReadOnlyList<NoteListEntry> ListNotes(bool endnotes = false)
    {
        ThrowIfDisposed();
        _ = AnchorIndex(); // guarantees Unids on the note parts

        var result = new List<NoteListEntry>();
        var main = _doc!.MainDocumentPart;
        var bodyRoot = main?.GetXDocument().Root;
        var notesRoot = endnotes
            ? main?.EndnotesPart?.GetXDocument().Root
            : main?.FootnotesPart?.GetXDocument().Root;
        if (bodyRoot is null || notesRoot is null) return result;

        var refName = endnotes ? W.endnoteReference : W.footnoteReference;
        var defName = endnotes ? W.endnote : W.footnote;
        var kindScope = endnotes ? "en" : "fn";

        var defsById = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var def in notesRoot.Elements(defName))
        {
            var defId = (string?)def.Attribute(W.id);
            if (defId is not null) defsById[defId] = def;
        }

        foreach (var r in bodyRoot.Descendants(refName))
        {
            var id = (string?)r.Attribute(W.id);
            if (id is null || !defsById.TryGetValue(id, out var def)) continue;
            var unid = (string?)def.Attribute(PtOpenXml.Unid);
            if (unid is null) continue;
            result.Add(new NoteListEntry(id, $"{kindScope}:{kindScope}:{unid}", result.Count + 1));
        }
        return result;
    }

    // ─── Footnotes / endnotes ─────────────────────────────────────────────────
    //
    // Author a note: write the definition into the FootnotesPart/EndnotesPart (creating the part
    // and Word's two reserved separator notes when it doesn't exist yet) and cite it from a body
    // paragraph at a character offset. Undo/redo of the part creation is handled by the note-part
    // reconcile in RestoreSnapshot, mirroring the header/footer one above.
    //
    // Editing and deleting an authored note need no new op: ReplaceText already accepts the note's
    // p:fn/p:en paragraph anchor, and DeleteBlock already removes an fn/en definition together with
    // every reference to it anywhere in the package (DS140/DS141). InsertFootnote/InsertEndnote
    // were the only missing verb.

    /// <summary>
    /// Create a <b>footnote</b> whose body is <paramref name="markdownPayload"/> and cite it from the
    /// paragraph named by <paramref name="anchorId"/> at <paramref name="characterOffset"/> characters
    /// into that paragraph's text (0 = before all text, text length = after all of it). Creates the
    /// <c>FootnotesPart</c>, Word's two reserved separator notes, the <c>FootnoteText</c>/
    /// <c>FootnoteReference</c> styles, and the <c>w:footnotePr</c> settings declaration if the document
    /// has no footnotes yet; otherwise the existing part is reused and only a fresh definition is added.
    /// The note id is allocated above every id already used in the package, so a document with
    /// non-contiguous note ids can't collide. Returns the created note anchors — the definition
    /// (kind <c>fn</c>) and its paragraphs (kind <c>p</c>, scope <c>fn</c>) — in
    /// <see cref="EditResult.Created"/>, so a caller can immediately edit the note with
    /// <see cref="ReplaceText"/> or remove it with <see cref="DeleteBlock"/>.
    /// </summary>
    /// <remarks>
    /// Body paragraphs only. Word does not allow a note reference inside a header/footer story or
    /// inside another note, and the projection's <c>fn</c>/<c>en</c> scopes are note <em>definitions</em>,
    /// not legal citation hosts — a non-body anchor is rejected with
    /// <see cref="EditErrorCode.AnchorWrongKind"/> rather than silently producing a document Word repairs.
    /// </remarks>
    public EditResult InsertFootnote(string anchorId, int characterOffset, string markdownPayload)
        => InsertNote(isFootnote: true, anchorId, characterOffset, markdownPayload);

    /// <summary>
    /// Create an <b>endnote</b> and cite it from a body paragraph. Behaves exactly like
    /// <see cref="InsertFootnote"/> but writes into the <c>EndnotesPart</c>, emits a
    /// <c>w:endnoteReference</c>, and uses the <c>EndnoteText</c>/<c>EndnoteReference</c> styles;
    /// the created definition anchor has kind <c>en</c> and its paragraphs scope <c>en</c>.
    /// </summary>
    public EditResult InsertEndnote(string anchorId, int characterOffset, string markdownPayload)
        => InsertNote(isFootnote: false, anchorId, characterOffset, markdownPayload);

    private EditResult InsertNote(bool isFootnote, string anchorId, int characterOffset, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var opName = isFootnote ? "InsertFootnote" : "InsertEndnote";

        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"{opName} requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);
        if (target.Anchor.Scope != "body")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"{opName} requires a body paragraph anchor; Word does not allow a note reference in the " +
                $"'{target.Anchor.Scope}' story", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        var main = _doc!.MainDocumentPart;
        if (main is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no main document part", anchorId);

        var totalText = ParagraphText(element);
        if (characterOffset < 0 || characterOffset > totalText.Length)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"offset {characterOffset} out of [0, {totalText.Length}]", anchorId);
        if (HasUnsupportedInlineInsertionBoundary(element, characterOffset))
            return EditResult.Fail(EditErrorCode.UnsupportedInlineBoundary,
                $"{opName} offset falls inside a revision or unsupported inline container; " +
                "choose a boundary before or after that container", anchorId);

        // Parse the note body BEFORE snapshotting so a malformed payload is a clean no-op
        // (no part created, no undo entry pushed).
        var paras = new List<XElement>();
        if (!string.IsNullOrEmpty(markdownPayload))
        {
            var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
            if (!parsed.Success)
                return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, anchorId);
            foreach (var block in parsed.Blocks)
                paras.Add(BuildParagraphFromParsedBlock(block));
        }
        if (paras.Count == 0) paras.Add(new XElement(W.p));
        if (ValidatePendingHyperlinks(paras, anchorId) is { } linkError)
            return linkError;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var part = EnsureNotePart(main, isFootnote);
            Internal.StyleFactory.EnsureNoteStyles(_doc!, isFootnote);

            var noteName = isFootnote ? W.footnote : W.endnote;
            var root = part.GetXDocument().Root!;

            // Body-side citation FIRST, carrying a placeholder id — the note's real id is chosen
            // from where this citation lands among the existing ones (see NextNoteIdInReferenceOrder).
            // Split before inserting so no run or hyperlink straddles the offset: the same offset
            // mechanism SplitParagraph and ApplyFormat use, not a second walker.
            SplitRunsAtOffset(element, characterOffset);
            SplitInlineContainersAtOffset(element, characterOffset);
            var refRun = BuildNoteReferenceRun(isFootnote, NoteIdPlaceholder);
            UnidHelper.AssignToSelfAndDescendants(refRun);

            // Under recording, the CITATION is the reversible unit and the definition follows it —
            // and the definition's content is marked inserted too (issue #636, revisiting #614's
            // unmarked-definition choice). Marking is what Word writes, what the diff engine's own
            // redline carries for a wholly-inserted note, and what lets the STATELESS reject read
            // its way to the right answer: rejecting empties the definition, and the guarded
            // note-lifecycle prune (NoteReferenceOps.PruneNotesEmptiedByResolution) takes a
            // definition it both orphaned and emptied — without eating a baseline-owned husk whose
            // citation the redline merely inserted, which is the mirror of the #631 accept bug
            // that the old unguarded reject prune had.
            // One user action, one revision identity: the citation envelope and every definition
            // paragraph share a single author/date pair, or RevisionOps.BuildGroups (exact-date
            // equality) splinters one insert into per-paragraph revision groups.
            var recording = _trackedChanges == TrackedChangeMode.RenderInline;
            var recordingAuthor = _revisionAuthor ?? "docxodus";
            var recordingDate = recording ? NextTrackedFormatRevisionDate() : string.Empty;
            InsertInlineAtOffset(element, characterOffset,
                recording
                    ? CreateRevisionEnvelope(W.ins, recordingAuthor, recordingDate, refRun)
                    : refRun);

            var id = NextNoteIdInReferenceOrder(main, root, noteName);
            refRun.Descendants(isFootnote ? W.footnoteReference : W.endnoteReference).First()
                .SetAttributeValue(W.id, id.ToString(System.Globalization.CultureInfo.InvariantCulture));

            ApplyNoteBodyStyle(paras, isFootnote);
            if (recording)
                foreach (var p in paras)
                    MarkParagraphContentAndMark(p, W.ins, recordingAuthor, recordingDate);
            var note = new XElement(noteName,
                new XAttribute(W.id, id.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                paras);
            root.Add(note);
            UnidHelper.AssignToSelfAndDescendants(note);
            foreach (var p in paras) PromoteHyperlinkRelationships(p);
            part.PutXDocument();

            InvalidateProjectionCache();

            var created = new List<Anchor>();
            var notePartUri = part.Uri.ToString();
            if (AnchorForUnid((string?)note.Attribute(PtOpenXml.Unid), notePartUri) is { } noteAnchor)
                created.Add(noteAnchor);
            foreach (var p in note.Elements(W.p))
                if (AnchorForUnid((string?)p.Attribute(PtOpenXml.Unid), notePartUri) is { } pa)
                    created.Add(pa);

            return new EditResult
            {
                Success = true,
                Created = created,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            RollbackFailedOp();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>
    /// Find-or-create the footnotes/endnotes part. A part created here is seeded with the two
    /// notes Word reserves for page-rendering scaffolding (<c>type="separator"</c> at id -1 and
    /// <c>type="continuationSeparator"</c> at id 0) and declared in <c>settings.xml</c>, which is
    /// what Word itself writes for a document's first note — a part holding only user notes opens,
    /// but has no separator rule above the note area.
    /// </summary>
    private static OpenXmlPart EnsureNotePart(MainDocumentPart main, bool isFootnote)
    {
        OpenXmlPart? existing = isFootnote ? main.FootnotesPart : main.EndnotesPart;
        if (existing is not null) return existing;

        var part = isFootnote ? main.AddNewPart<FootnotesPart>() : (OpenXmlPart)main.AddNewPart<EndnotesPart>();
        part.PutXDocument(new XDocument(BuildNotePartRoot(isFootnote)));
        DeclareNoteSeparatorsInSettings(main, isFootnote);
        return part;
    }

    /// <summary>A fresh <c>w:footnotes</c>/<c>w:endnotes</c> root holding only Word's two reserved notes.</summary>
    private static XElement BuildNotePartRoot(bool isFootnote)
    {
        var noteName = isFootnote ? W.footnote : W.endnote;
        // The separator marks are shared between footnotes and endnotes (there is no w:endnoteSeparator).
        XElement Reserved(string type, int id, XName mark) =>
            new XElement(noteName,
                new XAttribute(W.type, type),
                new XAttribute(W.id, id.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                new XElement(W.p,
                    new XElement(W.pPr,
                        new XElement(W.spacing,
                            new XAttribute(W.after, "0"),
                            new XAttribute(W.line, "240"),
                            new XAttribute(W.lineRule, "auto"))),
                    new XElement(W.r, new XElement(mark))));

        return new XElement(isFootnote ? W.footnotes : W.endnotes,
            new XAttribute(XNamespace.Xmlns + "w", W.w),
            new XAttribute(XNamespace.Xmlns + "r", R.r),
            Reserved("separator", -1, W.separator),
            Reserved("continuationSeparator", 0, W.continuationSeparator));
    }

    /// <summary>
    /// Declare the reserved separator notes in <c>settings.xml</c> (<c>w:footnotePr</c>/
    /// <c>w:endnotePr</c>), the form Word writes. Uses the shared schema-slot insert so the
    /// settings part is never wholesale reordered (see <c>EnsureSettingsChildInOrder</c>); an
    /// existing <c>w:footnotePr</c> carrying the document's own numbering options is left alone.
    /// </summary>
    private static void DeclareNoteSeparatorsInSettings(MainDocumentPart main, bool isFootnote)
    {
        var settingsPart = main.DocumentSettingsPart ?? main.AddNewPart<DocumentSettingsPart>();
        var xDoc = settingsPart.GetXDocument();
        var root = xDoc.Root;
        if (root is null)
        {
            root = new XElement(W.settings, new XAttribute(XNamespace.Xmlns + "w", W.w));
            xDoc.Add(root);
        }

        var noteName = isFootnote ? W.footnote : W.endnote;
        var pr = new XElement(isFootnote ? W.footnotePr : W.endnotePr,
            new XElement(noteName, new XAttribute(W.id, "-1")),
            new XElement(noteName, new XAttribute(W.id, "0")));
        if (WordprocessingMLUtil.EnsureSettingsChildInOrder(root, pr))
            settingsPart.PutXDocument();
    }

    /// <summary>
    /// The id for a note whose citation (marked by <see cref="NoteIdPlaceholder"/>) is already in
    /// the body — chosen so ids stay ascending in <em>reference</em> order, shifting the notes cited
    /// after it up by one when necessary. Returns the new note's id.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Ascending-in-reference-order is an invariant every Word-authored document holds (verified
    /// across the TestFiles corpus — including documents whose ids have gaps, e.g. 17/21/26 — and
    /// the 94-footnote NVCA model certificate). Renderers rely on it: LibreOffice numbers the body
    /// markers by citation position but pairs them against the <em>id-sorted</em> definition list,
    /// so a first-cited note holding the highest id renders the WRONG note text — the marker reads
    /// "1" and points at somebody else's footnote. Nothing errors; the document is simply wrong.
    /// </para>
    /// <para>
    /// Appending <c>max(id) + 1</c> is therefore correct only when the citation follows every
    /// existing one — the common case, and the one that costs nothing here. Otherwise the new note
    /// takes the smallest id cited after it and everything at or above that shifts up, which leaves
    /// the notes cited *earlier* untouched. Taking the minimum of the following ids (rather than the
    /// first) keeps this correct even if the input document already violated the invariant.
    /// </para>
    /// </remarks>
    private static int NextNoteIdInReferenceOrder(MainDocumentPart main, XElement notePartRoot, XName noteName)
    {
        var refName = noteName == W.footnote ? W.footnoteReference : W.endnoteReference;

        // Citations in document order; the placeholder marks where the new one landed.
        var ordered = main.GetXDocument().Root!.Descendants(refName)
            .Select(r => int.TryParse((string?)r.Attribute(W.id), out var v) ? v : 0)
            .ToList();
        var following = ordered.SkipWhile(v => v != NoteIdPlaceholder).Skip(1)
            .Where(v => v != NoteIdPlaceholder).ToList();

        if (following.Count > 0)
        {
            var slot = following.Min();
            ShiftNoteIdsAtOrAbove(main, notePartRoot, noteName, slot);
            return slot;
        }

        // Cited after every existing note — appending keeps ids ascending, so no shift is needed.
        // Scan definitions AND references (a note may be cited from a header/footer or another
        // note) so a document with gaps can't alias an existing definition.
        int max = 0;
        foreach (var n in notePartRoot.Elements(noteName))
            if (int.TryParse((string?)n.Attribute(W.id), out var id) && id > max) max = id;
        foreach (var part in NoteReferenceHostParts(main))
        {
            var root = part.GetXDocument().Root;
            if (root is null) continue;
            foreach (var r in root.Descendants(refName))
                if (int.TryParse((string?)r.Attribute(W.id), out var id) && id > max) max = id;
        }
        return max + 1;
    }

    /// <summary>
    /// Renumber every user note with id &gt;= <paramref name="threshold"/> up by one — both the
    /// definitions and every reference to them anywhere in the package — to open that id for a
    /// newly inserted note. Word-reserved notes (any <c>w:type</c>: separator,
    /// continuationSeparator, continuationNotice) are never touched; their ids sit below any user
    /// id, so shifting upward can't collide with them.
    /// </summary>
    private static void ShiftNoteIdsAtOrAbove(
        MainDocumentPart main, XElement notePartRoot, XName noteName, int threshold)
    {
        static void Bump(XElement el)
        {
            var v = int.Parse((string)el.Attribute(W.id)!, System.Globalization.CultureInfo.InvariantCulture);
            el.SetAttributeValue(W.id, (v + 1).ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        foreach (var n in notePartRoot.Elements(noteName))
        {
            if (n.Attribute(W.type) is not null) continue; // Word-reserved scaffolding
            if (int.TryParse((string?)n.Attribute(W.id), out var v) && v >= threshold) Bump(n);
        }

        var refName = noteName == W.footnote ? W.footnoteReference : W.endnoteReference;
        foreach (var part in NoteReferenceHostParts(main))
        {
            var root = part.GetXDocument().Root;
            if (root is null) continue;
            bool any = false;
            foreach (var r in root.Descendants(refName).ToList())
                if (int.TryParse((string?)r.Attribute(W.id), out var v) && v >= threshold) { Bump(r); any = true; }
            // The main part is flushed by Save with the rest of the edit; peer parts (headers,
            // footers, the other note part) are written back here, as RemoveCrossReferences does.
            if (any && !ReferenceEquals(part, main)) part.PutXDocument();
        }
    }

    /// <summary>Every part whose XML can carry a note reference, for id-collision scanning.</summary>
    private static IEnumerable<OpenXmlPart> NoteReferenceHostParts(MainDocumentPart main) =>
        Internal.NoteReferenceOps.ReferenceHostParts(main);

    /// <summary>
    /// Shape the note body the way Word does: every paragraph gets the note-text style (unless the
    /// payload already set one) and the first paragraph opens with the auto-number mark
    /// (<c>w:footnoteRef</c>/<c>w:endnoteRef</c>) followed by a separating space, so the note reads
    /// "1 Text" rather than "1Text".
    /// </summary>
    private static void ApplyNoteBodyStyle(List<XElement> paras, bool isFootnote)
    {
        var styleId = isFootnote ? "FootnoteText" : "EndnoteText";
        foreach (var p in paras)
        {
            var pPr = p.Element(W.pPr);
            if (pPr is null) { pPr = new XElement(W.pPr); p.AddFirst(pPr); }
            if (pPr.Element(W.pStyle) is null)
                pPr.AddFirst(new XElement(W.pStyle, new XAttribute(W.val, styleId)));
        }

        // The loop above guarantees a w:pPr on every paragraph, so the mark + its separating
        // space go straight after the first paragraph's — schema-ordered, ahead of the payload.
        var firstPPr = paras[0].Element(W.pPr)!;
        firstPPr.AddAfterSelf(
            new XElement(W.r,
                new XElement(W.rPr,
                    new XElement(W.rStyle, new XAttribute(W.val, NoteReferenceStyleId(isFootnote)))),
                new XElement(isFootnote ? W.footnoteRef : W.endnoteRef)),
            new XElement(W.r,
                new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), " ")));
    }

    /// <summary>The body-side citation run: a superscript-styled <c>w:footnoteReference</c>/
    /// <c>w:endnoteReference</c> pointing at <paramref name="id"/>.</summary>
    private static XElement BuildNoteReferenceRun(bool isFootnote, int id) =>
        new XElement(W.r,
            new XElement(W.rPr,
                new XElement(W.rStyle, new XAttribute(W.val, NoteReferenceStyleId(isFootnote)))),
            new XElement(isFootnote ? W.footnoteReference : W.endnoteReference,
                new XAttribute(W.id, id.ToString(System.Globalization.CultureInfo.InvariantCulture))));

    private static string NoteReferenceStyleId(bool isFootnote) =>
        isFootnote ? "FootnoteReference" : "EndnoteReference";
}
