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
    // ─── Reference fields: TOC / TOF / TOA (issue #607) ───────────────────────
    //
    // Narrowing the library to the DOCX toolchain removed ReferenceAdder, and with it the only way
    // Docxodus could CREATE a reference field. The CHANGELOG's workaround — hand-build the three
    // fldChars and the switch string through Raw.InsertXml — is a genuine regression in the level of
    // abstraction: the caller has to get w:instrText syntax right, nothing validates the result, and
    // a malformed field renders as NOTHING in Word. Silently.
    //
    // These are session ops rather than a resurrected static: anchor-addressed, undoable, and
    // consistent with the rest of the mutation API. The switches are modelled as typed options
    // rather than taken verbatim, which is the whole point — the old API's switch string pushed the
    // same OOXML knowledge back onto the caller.
    //
    // Marking ENTRIES (TC / TA fields) is deliberately not here; it is a larger job and a separate
    // issue. What these produce is the table, which is what a document is missing when it has none.

    /// <summary>
    /// Insert a <b>table of contents</b> before or after the block named by <paramref name="anchorId"/>.
    /// The field is written dirty and the document asks for a field update on open, so Word
    /// paginates and fills the table itself — the library never ships a cached result that is stale
    /// the moment anything above it moves.
    /// </summary>
    /// <remarks>
    /// <para>The table is wrapped in the <c>w:sdt</c> content control Word puts around one (gallery
    /// "Table of Contents"), which is what gives it the "Update Table" control in Word's UI. Word's
    /// <c>TOCHeading</c> and <c>TOC1</c> styles are find-or-created; a document that already defines
    /// them keeps its own.</para>
    /// <para>Body blocks only. A reference field in a header/footer or note story is not a shape
    /// Word produces, so a non-body anchor is rejected rather than silently written.</para>
    /// </remarks>
    /// <returns>The created block anchors — the content control and its paragraphs — in
    /// <see cref="EditResult.Created"/>.</returns>
    public EditResult InsertTableOfContents(
        string anchorId, Position pos, TableOfContentsOptions? options = null)
    {
        var o = options ?? new TableOfContentsOptions();
        var levels = Internal.ReferenceFieldOps.NormalizeLevels(o.Levels, out var levelError);
        if (levels is null)
            return EditResult.Fail(EditErrorCode.InvalidReferenceField, levelError!, anchorId);
        if (o.RightTabPos <= 0)
            return EditResult.Fail(EditErrorCode.InvalidReferenceField,
                $"RightTabPos must be positive twips; got {o.RightTabPos}", anchorId);

        return InsertReferenceField(anchorId, pos, "InsertTableOfContents",
            entryStyleId: Internal.ReferenceFieldOps.TocEntryStyleId,
            entryStyleName: "toc 1",
            withHeadingStyle: true,
            instruction: Internal.ReferenceFieldOps.TocInstruction(
                levels, o.Hyperlinks, o.HideTabAndPageNumbersInWeb, o.UseOutlineLevels),
            rightTabPos: o.RightTabPos,
            title: string.IsNullOrEmpty(o.Title) ? null : o.Title,
            wrapInContentControl: true);
    }

    /// <summary>
    /// Insert a <b>table of figures</b> — the captions carrying <see cref="TableOfFiguresOptions.CaptionLabel"/>
    /// and their page numbers. Behaves exactly like <see cref="InsertTableOfContents"/> otherwise,
    /// except that Word writes a table of figures as a bare paragraph rather than inside a content
    /// control, so this does too.
    /// </summary>
    public EditResult InsertTableOfFigures(
        string anchorId, Position pos, TableOfFiguresOptions? options = null)
    {
        var o = options ?? new TableOfFiguresOptions();
        if (string.IsNullOrWhiteSpace(o.CaptionLabel))
            return EditResult.Fail(EditErrorCode.InvalidReferenceField,
                "CaptionLabel must name the caption label to list (e.g. \"Figure\")", anchorId);
        if (o.RightTabPos <= 0)
            return EditResult.Fail(EditErrorCode.InvalidReferenceField,
                $"RightTabPos must be positive twips; got {o.RightTabPos}", anchorId);

        return InsertReferenceField(anchorId, pos, "InsertTableOfFigures",
            entryStyleId: Internal.ReferenceFieldOps.TofEntryStyleId,
            entryStyleName: "table of figures",
            withHeadingStyle: false,
            instruction: Internal.ReferenceFieldOps.TofInstruction(o.CaptionLabel.Trim(), o.Hyperlinks),
            rightTabPos: o.RightTabPos,
            title: null,
            wrapInContentControl: false);
    }

    /// <summary>
    /// Insert a <b>table of authorities</b> — the cases, statutes or other authorities marked in the
    /// document, grouped by <see cref="TableOfAuthoritiesOptions.Category"/>. Table stakes for the
    /// legal documents this library targets.
    /// </summary>
    /// <remarks>The table lists entries the document has MARKED with <c>TA</c> fields. A document
    /// with no marked citations produces a table that is correct and empty; marking entries is a
    /// separate capability.</remarks>
    public EditResult InsertTableOfAuthorities(
        string anchorId, Position pos, TableOfAuthoritiesOptions? options = null)
    {
        var o = options ?? new TableOfAuthoritiesOptions();
        if (!Enum.IsDefined(typeof(AuthorityCategory), o.Category))
            return EditResult.Fail(EditErrorCode.InvalidReferenceField,
                $"unknown authority category {(int)o.Category}", anchorId);
        if (o.RightTabPos <= 0)
            return EditResult.Fail(EditErrorCode.InvalidReferenceField,
                $"RightTabPos must be positive twips; got {o.RightTabPos}", anchorId);

        return InsertReferenceField(anchorId, pos, "InsertTableOfAuthorities",
            entryStyleId: Internal.ReferenceFieldOps.ToaEntryStyleId,
            entryStyleName: "table of authorities",
            withHeadingStyle: false,
            instruction: Internal.ReferenceFieldOps.ToaInstruction(
                (int)o.Category, o.Hyperlinks, o.EntryPageSeparator),
            rightTabPos: o.RightTabPos,
            title: null,
            wrapInContentControl: false);
    }

    /// <summary>The one insertion path all three reference-field ops take.</summary>
    private EditResult InsertReferenceField(
        string anchorId, Position pos, string opName,
        string entryStyleId, string entryStyleName, bool withHeadingStyle,
        string instruction, int rightTabPos, string? title, bool wrapInContentControl)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);

        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Scope != "body")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"{opName} requires a body anchor; Word does not generate a reference table in the "
                + $"'{target.Anchor.Scope}' story", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        var main = _doc!.MainDocumentPart;
        if (main is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no main document part", anchorId);

        // A generated table is regenerated wholesale by Word on every field update, so there is no
        // reversible way to redline it: refuse under recording rather than write a mark that
        // rejecting cannot take back (the shape #614 established).
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported(opName, anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            Internal.StyleFactory.EnsureReferenceFieldStyles(
                _doc!, entryStyleId, entryStyleName, withHeadingStyle);

            var blocks = new List<XElement>();
            if (title is not null) blocks.Add(Internal.ReferenceFieldOps.TitleParagraph(title));
            blocks.Add(Internal.ReferenceFieldOps.FieldParagraph(entryStyleId, instruction, rightTabPos));

            var inserted = wrapInContentControl
                ? new[] { Internal.ReferenceFieldOps.TableOfContentsControl(blocks) }
                : blocks.ToArray();
            foreach (var block in inserted) UnidHelper.AssignToSelfAndDescendants(block);

            if (pos == Position.Before)
            {
                foreach (var block in inserted) element.AddBeforeSelf(block);
            }
            else
            {
                XElement after = element;
                foreach (var block in inserted) { after.AddAfterSelf(block); after = block; }
            }

            Internal.ReferenceFieldOps.RequestFieldUpdateOnOpen(main);
            InvalidateProjectionCache();

            var created = new List<Anchor>();
            foreach (var block in inserted)
                foreach (var node in block.DescendantsAndSelf())
                    if ((string?)node.Attribute(PtOpenXml.Unid) is { } unid
                        && AnchorForUnid(unid, target.PartUri) is { } anchor)
                        created.Add(anchor);

            return new EditResult
            {
                Success = true,
                Created = created,
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }
}
