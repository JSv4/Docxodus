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
    // ─── Headers / footers / page-number fields ───────────────────────────────
    //
    // Author the per-section running header/footer stories (which live in their own OOXML
    // parts, outside the body) and page-number fields. SetHeaderText/SetFooterText are
    // addressed by ANY body block in the target section — the governing w:sectPr is resolved
    // the same way GetSectionInfo resolves it (a mid-document section break, else the body's
    // trailing sectPr, creating one if the body has none). The created header/footer paragraph
    // anchors come back in EditResult.Created with a hdr{N}/ftr{N} scope, so a page-number field
    // can then be inserted into them with InsertPageNumberField. Undo/redo of the part creation
    // is handled by the header/footer reconcile in RestoreSnapshot.

    /// <summary>
    /// Set the running <b>header</b> story for the section that owns <paramref name="anchorId"/>
    /// (any body block in that section) to <paramref name="markdownPayload"/>. Creates the header
    /// part, its relationship, and the <c>w:headerReference</c> on the section if the story of the
    /// requested <paramref name="kind"/> does not exist yet; otherwise replaces that story's content.
    /// An empty payload yields a single empty header paragraph. <see cref="HeaderFooterKind.First"/>
    /// sets the section's <c>w:titlePg</c>; <see cref="HeaderFooterKind.Even"/> sets
    /// <c>w:evenAndOddHeaders</c> in the settings part. Returns the created header-paragraph anchors
    /// (scope <c>hdr{N}</c>) in <see cref="EditResult.Created"/>.
    /// </summary>
    public EditResult SetHeaderText(string anchorId, HeaderFooterKind kind, string markdownPayload)
        => SetHeaderFooterText(isHeader: true, anchorId, kind, markdownPayload);

    /// <summary>
    /// Set the running <b>footer</b> story for the section that owns <paramref name="anchorId"/>.
    /// Behaves exactly like <see cref="SetHeaderText"/> but for the footer part / <c>w:footerReference</c>;
    /// the created footer-paragraph anchors (scope <c>ftr{N}</c>) come back in
    /// <see cref="EditResult.Created"/> — insert a page number into one with
    /// <see cref="InsertPageNumberField"/>.
    /// </summary>
    public EditResult SetFooterText(string anchorId, HeaderFooterKind kind, string markdownPayload)
        => SetHeaderFooterText(isHeader: false, anchorId, kind, markdownPayload);

    /// <summary>
    /// Ensure Word will actually RENDER the <paramref name="kind"/> header/footer stories of the
    /// section that owns <paramref name="anchorId"/> (any body block, resolved as
    /// <see cref="GetSectionInfo"/> resolves it): sets <c>w:titlePg</c> for
    /// <see cref="HeaderFooterKind.First"/> and the document-global <c>w:evenAndOddHeaders</c>
    /// for <see cref="HeaderFooterKind.Even"/>. <see cref="HeaderFooterKind.Default"/> needs no
    /// flag and is a successful no-op. Idempotent.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see cref="SetHeaderText"/>/<see cref="SetFooterText"/> set these flags as a side effect of
    /// writing content, which covers authoring a story from scratch. It does NOT cover a document
    /// that already carries a first/even reference with the flag absent — Word writes exactly that
    /// when "Different first page" / "Different odd &amp; even pages" is turned back off, leaving
    /// the part behind. Editing such a story through the anchor-addressed text ops then produces a
    /// document whose header content is present but invisible. An editor offering a
    /// first/even story selector needs this as its own operation, because the flags belong to the
    /// SECTION, not to a content write.
    /// </para>
    /// <para>
    /// See <see cref="SetHeaderText"/> for the <c>w:evenAndOddHeaders</c> caveat: it is
    /// document-global and governs footers too, so enabling it without an even FOOTER means even
    /// pages show no footer at all.
    /// </para>
    /// </remarks>
    public EditResult EnsureHeaderFooterVisible(string anchorId, HeaderFooterKind kind)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Scope != "body")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "EnsureHeaderFooterVisible requires a body block anchor (the section the header/footer belongs to)",
                anchorId);
        if (kind == HeaderFooterKind.Default)
            return new EditResult { Success = true, Modified = new[] { target.Anchor } };

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        var main = _doc!.MainDocumentPart;
        if (main is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no main document part", anchorId);

        var sectPr = Internal.BlockMetadataOps.FindGoverningSectPr(element);
        if (sectPr is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no governing section properties", anchorId);

        // Already set? Return BEFORE snapshotting. A UI that calls this on every kind selection
        // (the editor's band does) would otherwise push a no-op snapshot per click into the
        // bounded undo ring and evict the user's real history. The test is the flag's VALUE, not
        // its presence: Word writes <w:titlePg w:val="0"/> when the box is cleared, and
        // SectionInfo reads it the same way — an enable that saw "present" as "on" left the
        // checkbox unable to stick.
        var settingsRootForCheck = main.DocumentSettingsPart?.GetXDocument().Root;
        bool alreadySet = kind == HeaderFooterKind.First
            ? WordprocessingMLUtil.GetBoolProp(sectPr, W.titlePg)
            : settingsRootForCheck is not null
                && WordprocessingMLUtil.GetBoolProp(settingsRootForCheck, W.evenAndOddHeaders);
        if (alreadySet)
            return new EditResult { Success = true, Modified = new[] { target.Anchor } };

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (kind == HeaderFooterKind.First) InsertSectPrTitlePg(sectPr);
            else WordprocessingMLUtil.EnsureEvenAndOddHeaders(main);
            InvalidateProjectionCache();
            return new EditResult { Success = true, Modified = new[] { target.Anchor } };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>
    /// Word's "Different first page" / "Different odd &amp; even pages" checkboxes as one verb:
    /// switch the <paramref name="kind"/> story of the section owning <paramref name="anchorId"/>
    /// (any body block in it) on or off. <c>enabled</c> is exactly
    /// <see cref="EnsureHeaderFooterVisible"/>; <c>!enabled</c> removes the flag that selects the
    /// story — <c>w:titlePg</c> from the governing <c>w:sectPr</c> for <see cref="HeaderFooterKind.First"/>,
    /// the document-global <c>w:evenAndOddHeaders</c> from the settings part for
    /// <see cref="HeaderFooterKind.Even"/>. The story PARTS and their references are left in place,
    /// which is what Word does when the checkbox is cleared: the content survives, it just stops
    /// being selected, and re-enabling brings it straight back.
    /// </summary>
    /// <remarks>
    /// <see cref="HeaderFooterKind.Default"/> has no flag, so disabling it is refused with
    /// <see cref="EditErrorCode.InvalidPageSetup"/> (enabling it is the same successful no-op
    /// <see cref="EnsureHeaderFooterVisible"/> returns). A flag already in the requested state is a
    /// successful no-op that records NO undo snapshot — a checkbox handler that fires on every
    /// render must not evict the user's real history from the bounded ring. Read the current state
    /// from <see cref="SectionInfo.TitlePage"/> / <see cref="SectionInfo.EvenAndOddHeaders"/>.
    /// </remarks>
    public EditResult SetHeaderFooterKindEnabled(string anchorId, HeaderFooterKind kind, bool enabled)
    {
        if (enabled) return EnsureHeaderFooterVisible(anchorId, kind);

        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Scope != "body")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "SetHeaderFooterKindEnabled requires a body block anchor (the section the header/footer belongs to)",
                anchorId);
        if (kind == HeaderFooterKind.Default)
            return EditResult.Fail(EditErrorCode.InvalidPageSetup,
                "the default header/footer story has no on/off flag; only First (w:titlePg) and Even (w:evenAndOddHeaders) can be disabled",
                anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        var main = _doc!.MainDocumentPart;
        if (main is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no main document part", anchorId);

        var sectPr = Internal.BlockMetadataOps.FindGoverningSectPr(element);
        var settingsPart = main.DocumentSettingsPart;
        var settingsRoot = settingsPart?.GetXDocument().Root;

        // Already off? Return BEFORE snapshotting (same reasoning as EnsureHeaderFooterVisible).
        // The element is removed whatever its w:val, so "present" is the test, not "on".
        bool alreadyOff = kind == HeaderFooterKind.First
            ? sectPr?.Element(W.titlePg) is null
            : settingsRoot?.Element(W.evenAndOddHeaders) is null;
        if (alreadyOff)
            return new EditResult { Success = true, Modified = new[] { target.Anchor } };

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (kind == HeaderFooterKind.First)
            {
                sectPr!.Elements(W.titlePg).Remove();
            }
            else
            {
                settingsRoot!.Elements(W.evenAndOddHeaders).Remove();
                settingsPart!.PutXDocument();
            }
            InvalidateProjectionCache();
            return new EditResult { Success = true, Modified = new[] { target.Anchor } };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    private EditResult SetHeaderFooterText(bool isHeader, string anchorId, HeaderFooterKind kind, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Scope != "body")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "SetHeaderText/SetFooterText require a body block anchor (the section the header/footer belongs to)", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        var body = element.AncestorsAndSelf(W.body).FirstOrDefault();
        if (body is null)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "anchor is not in the document body", anchorId);
        var main = _doc!.MainDocumentPart;
        if (main is null)
            return EditResult.Fail(EditErrorCode.InternalError, "no main document part", anchorId);

        // Parse the payload into paragraphs (empty payload ⇒ one empty paragraph), then apply the
        // built-in Header/Footer style so the paragraphs inherit Word's centre/right tab stops.
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
        ApplyHeaderFooterStyle(paras, isHeader);

        // Resolve an existing same-kind story before snapshotting so replacing its root cannot
        // silently remove one end of a cross-boundary bookmark or a still-targeted bookmark.
        var currentSectPr = Internal.BlockMetadataOps.FindGoverningSectPr(element);
        if (currentSectPr is not null)
        {
            var currentRefName = isHeader ? W.headerReference : W.footerReference;
            var currentType = HeaderFooterTypeValue(kind);
            var currentRef = currentSectPr.Elements(currentRefName)
                .FirstOrDefault(r => (string?)r.Attribute(W.type) == currentType);
            OpenXmlPart? currentPart = null;
            if ((string?)currentRef?.Attribute(R.id) is { } currentRid)
                foreach (var pp in main.Parts)
                    if (pp.RelationshipId == currentRid) { currentPart = pp.OpenXmlPart; break; }
            bool currentTypeMatches = isHeader ? currentPart is HeaderPart : currentPart is FooterPart;
            if (currentTypeMatches && currentPart?.GetXDocument().Root is { } oldRoot
                && ValidateBookmarkRemoval(new[] { oldRoot }, anchorId) is { } bookmarkError)
                return bookmarkError;
        }

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var sectPr = Internal.BlockMetadataOps.FindGoverningSectPr(element);
            if (sectPr is null)
            {
                // A body with no section properties at all — synthesize the document-final section.
                sectPr = new XElement(W.sectPr);
                body.Add(sectPr);
            }

            var refName = isHeader ? W.headerReference : W.footerReference;
            var typeVal = HeaderFooterTypeValue(kind);

            // Reuse the same-kind reference's part if it resolves to the right part type; otherwise
            // add a fresh part and reference (dropping any stale/mismatched same-kind reference).
            var existingRef = sectPr.Elements(refName)
                .FirstOrDefault(r => (string?)r.Attribute(W.type) == typeVal);
            OpenXmlPart? reuse = null;
            if (existingRef is not null && (string?)existingRef.Attribute(R.id) is { } rid)
                foreach (var pp in main.Parts)
                    if (pp.RelationshipId == rid) { reuse = pp.OpenXmlPart; break; }
            bool typeMatches = isHeader ? reuse is HeaderPart : reuse is FooterPart;

            OpenXmlPart part;
            var oldHyperlinkIds = new List<string>();
            if (reuse is not null && typeMatches)
            {
                part = reuse;
                oldHyperlinkIds.AddRange(part.GetXDocument().Descendants(W.hyperlink)
                    .Select(h => (string?)h.Attribute(R.id)).Where(id => !string.IsNullOrEmpty(id)).Cast<string>());
            }
            else
            {
                part = isHeader ? main.AddNewPart<HeaderPart>() : main.AddNewPart<FooterPart>();
                existingRef?.Remove();
                // Header/footer references lead the CT_SectPr sequence, so AddFirst is schema-ordered.
                sectPr.AddFirst(new XElement(refName,
                    new XAttribute(W.type, typeVal),
                    new XAttribute(R.id, main.GetIdOfPart(part))));
            }

            // Stamp Unids so the new paragraphs can be reported as anchors after re-projection.
            foreach (var p in paras) UnidHelper.AssignToSelfAndDescendants(p);

            var newRoot = new XElement(isHeader ? W.hdr : W.ftr,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                new XAttribute(XNamespace.Xmlns + "r", R.r),
                paras);
            part.PutXDocument(new XDocument(newRoot));
            foreach (var p in paras) PromoteHyperlinkRelationships(p);
            foreach (var relationshipId in oldHyperlinkIds)
                Internal.OwnedPartRelationships.DeleteReferenceRelationshipIfOrphaned(part, relationshipId, R.id);
            Internal.OwnedPartRelationships.SweepOrphanedImages(part);

            // Visibility flags so Word actually shows the First/Even stories.
            if (kind == HeaderFooterKind.First && !WordprocessingMLUtil.GetBoolProp(sectPr, W.titlePg))
                InsertSectPrTitlePg(sectPr);
            else if (kind == HeaderFooterKind.Even)
                WordprocessingMLUtil.EnsureEvenAndOddHeaders(main);

            InvalidateProjectionCache();
            var index = AnchorIndex();
            var created = new List<Anchor>();
            foreach (var p in paras)
            {
                var unid = (string?)p.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                var t = AnchorForUnid(unid, PartUriOf(p));
                if (t is { } a) created.Add(a);
            }

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

    /// <summary>
    /// Append a page-number field to the paragraph named by <paramref name="anchorId"/> — typically a
    /// header/footer paragraph (e.g. one returned by <see cref="SetFooterText"/>), though any paragraph
    /// is accepted. <see cref="PageNumberField.CurrentPage"/> emits a <c>PAGE</c> field,
    /// <see cref="PageNumberField.TotalPages"/> a <c>NUMPAGES</c> field, both as a native Word complex
    /// field (<c>fldChar</c>/<c>instrText</c>) with a cached result. Center the number by setting the
    /// paragraph alignment (<see cref="SetParagraphFormat"/>) or by relying on the Header/Footer style's
    /// centre tab. Returns the affected paragraph anchor in <see cref="EditResult.Modified"/>.
    /// </summary>
    /// <param name="format">
    /// Optional per-field number format, written as the field's <c>\*</c> general-formatting switch
    /// (<c>PAGE \* roman</c> → <c>i, ii, iii</c>). <c>null</c> — the default — emits a plain field,
    /// which is what Word inserts and what follows the SECTION's format
    /// (<see cref="SetPageNumbering"/>). Prefer the section setting for ordinary page numbering; a
    /// switch here OVERRIDES it for this one field and keeps overriding it if the section later
    /// changes. <see cref="NumberFormat.Bullet"/> is rejected.
    /// </param>
    public EditResult InsertPageNumberField(
        string anchorId,
        PageNumberField field = PageNumberField.CurrentPage,
        NumberFormat? format = null)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (format is { } f && !Internal.NumberFormats.IsPageNumberFormat(f))
            return EditResult.Fail(EditErrorCode.InvalidPageNumbering,
                $"{f} cannot format a page number", anchorId);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        if (element.Name != W.p)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                "InsertPageNumberField requires a paragraph anchor", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var runs = field == PageNumberField.PageOfTotal
                ? BuildPageOfTotalRuns(element, format)
                : BuildPageNumberFieldRuns(field, format);
            foreach (var r in runs)
            {
                UnidHelper.AssignToSelfAndDescendants(r);
                element.Add(r);
            }

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>Map <see cref="HeaderFooterKind"/> to the OOXML <c>w:type</c> token.</summary>
    private static string HeaderFooterTypeValue(HeaderFooterKind kind) => kind switch
    {
        HeaderFooterKind.First => "first",
        HeaderFooterKind.Even => "even",
        _ => "default",
    };

    /// <summary>Give every header/footer paragraph that carries no explicit <c>w:pStyle</c> the
    /// built-in Header/Footer style, so it inherits Word's centre-of-page and right-margin tab stops.</summary>
    private static void ApplyHeaderFooterStyle(List<XElement> paras, bool isHeader)
    {
        var styleId = isHeader ? "Header" : "Footer";
        foreach (var p in paras)
        {
            var pPr = p.Element(W.pPr);
            if (pPr is null) { pPr = new XElement(W.pPr); p.AddFirst(pPr); }
            if (pPr.Element(W.pStyle) is null)
                pPr.AddFirst(new XElement(W.pStyle, new XAttribute(W.val, styleId)));
        }
    }

    /// <summary>The runs of a native complex page-number field (PAGE / NUMPAGES) with a cached
    /// result — the form Word emits, so it renders and updates like a hand-authored field. With a
    /// <paramref name="format"/> the instruction carries the <c>\*</c> general-formatting switch and
    /// the cached result is page 1 rendered in that format, so a renderer that does not recompute
    /// fields shows a number consistent with the switch instead of always "1".</summary>
    private static XElement[] BuildPageNumberFieldRuns(PageNumberField field, NumberFormat? format)
    {
        var name = field == PageNumberField.TotalPages ? "NUMPAGES" : "PAGE";
        var instr = format is { } f && Internal.NumberFormats.ToFieldSwitch(f) is { } sw
            ? $" {name} \\* {sw} "
            : $" {name} ";
        var cached = format is { } cf ? Internal.NumberFormats.Render(1, cf) : "1";
        return new[]
        {
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"))),
            new XElement(W.r, new XElement(W.instrText, new XAttribute(XNamespace.Xml + "space", "preserve"), instr)),
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate"))),
            new XElement(W.r, new XElement(W.t, cached)),
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))),
        };
    }

    /// <summary>
    /// Word's "Page X of Y" gallery entry as runs: <c>Page </c>, a PAGE field, <c> of </c>, a
    /// NUMPAGES field. Exists because the pieces cannot be composed from the single-field op —
    /// text after a field cannot be appended without rewriting the field's result run. Every run
    /// (text and field alike) carries a copy of the paragraph's last run properties, minus any
    /// revision marker, so the composite takes the formatting the paragraph already has, as the
    /// gallery does. The literal text runs are <c>xml:space="preserve"</c> so their spaces survive.
    /// </summary>
    private static XElement[] BuildPageOfTotalRuns(XElement paragraph, NumberFormat? format)
    {
        var inherited = InlineRuns(paragraph).LastOrDefault(r => r.Element(W.rPr) is not null)?.Element(W.rPr);
        XElement? InheritedRPr()
        {
            if (inherited is null) return null;
            var copy = new XElement(inherited);
            copy.Elements(W.rPrChange).Remove();
            return copy.HasElements ? copy : null;
        }
        static XElement Text(XElement? rPr, string text) =>
            new XElement(W.r, rPr, new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), text));

        var runs = new List<XElement> { Text(InheritedRPr(), "Page ") };
        runs.AddRange(BuildPageNumberFieldRuns(PageNumberField.CurrentPage, format));
        runs.Add(Text(InheritedRPr(), " of "));
        runs.AddRange(BuildPageNumberFieldRuns(PageNumberField.TotalPages, format));
        foreach (var run in runs)
        {
            if (run.Element(W.rPr) is null && InheritedRPr() is { } rPr) run.AddFirst(rPr);
        }
        return runs.ToArray();
    }

    // ─── Section page numbering (w:pgNumType, issue #277) ─────────────────────
    //
    // The section-level half of page numbering: which number the section starts at and which format
    // its pages are numbered in. A plain PAGE field renders through this, so it — not the field —
    // is the normal place to say "front matter is i, ii, iii and the body restarts at 1".
    // Addressed by any body block in the target section, resolving the governing w:sectPr exactly
    // as GetSectionInfo does.

    /// <summary>
    /// Set the page-numbering properties (<c>w:pgNumType</c>) of the section that owns
    /// <paramref name="anchorId"/> (any body block in that section). Null fields on
    /// <paramref name="op"/> leave that attribute alone, so a caller can set the start without
    /// disturbing the format and vice versa. Creates the element, and a trailing <c>w:sectPr</c>,
    /// if absent. Idempotent.
    /// </summary>
    /// <remarks>
    /// Applying values the section already has is a successful no-op that does NOT consume undo
    /// history — a format dropdown firing on every selection must not evict the user's real edits
    /// from the bounded ring (same reasoning as <see cref="EnsureHeaderFooterVisible"/>).
    /// </remarks>
    public EditResult SetPageNumbering(string anchorId, PageNumberingOp op)
    {
        if (op.Start is { } s && s < 0)
            return EditResult.Fail(EditErrorCode.InvalidPageNumbering,
                "page-number start cannot be negative", anchorId);
        if (op.Format is { } f && !Internal.NumberFormats.IsPageNumberFormat(f))
            return EditResult.Fail(EditErrorCode.InvalidPageNumbering,
                $"{f} cannot format a page number", anchorId);

        return EditGoverningSectPr(anchorId, "SetPageNumbering", sectPr =>
        {
            var existing = sectPr.Element(W.pgNumType);
            var start = op.Start?.ToString(System.Globalization.CultureInfo.InvariantCulture);
            var fmt = op.Format is { } pf ? Internal.NumberFormats.ToOoxml(pf) : null;

            if (existing is not null
                && (start is null || (string?)existing.Attribute(W.start) == start)
                && (fmt is null || (string?)existing.Attribute(W.fmt) == fmt))
                return false;
            if (existing is null && start is null && fmt is null)
                return false;

            var pgNumType = existing;
            if (pgNumType is null)
            {
                pgNumType = new XElement(W.pgNumType);
                WordprocessingMLUtil.InsertSectPrChildInOrder(sectPr, pgNumType);
            }
            if (start is not null) pgNumType.SetAttributeValue(W.start, start);
            if (fmt is not null) pgNumType.SetAttributeValue(W.fmt, fmt);
            return true;
        });
    }

    /// <summary>
    /// Remove the page-numbering setup written by <see cref="SetPageNumbering"/> from the section
    /// that owns <paramref name="anchorId"/>: the section reverts to continuing the previous
    /// section's numbering in Word's default <c>1, 2, 3</c> format.
    /// </summary>
    /// <remarks>
    /// Narrowed to <c>w:start</c> and <c>w:fmt</c> — the chapter-numbering attributes
    /// (<c>w:chapStyle</c>/<c>w:chapSep</c>), which this surface never writes, are preserved, and
    /// the <c>w:pgNumType</c> element is removed only once nothing is left on it. A section with no
    /// page numbering to clear is a successful no-op that consumes no undo history.
    /// </remarks>
    public EditResult ClearPageNumbering(string anchorId) =>
        EditGoverningSectPr(anchorId, "ClearPageNumbering", sectPr =>
        {
            var pgNumType = sectPr.Element(W.pgNumType);
            if (pgNumType is null) return false;
            var start = pgNumType.Attribute(W.start);
            var fmt = pgNumType.Attribute(W.fmt);
            if (start is null && fmt is null) return false;

            start?.Remove();
            fmt?.Remove();
            // Only w:-namespaced attributes are CT_PageNumberType content; the element may also
            // carry pt: bookkeeping (Unid), which must not keep an otherwise-empty element alive.
            if (!pgNumType.Attributes().Any(a => a.Name.Namespace == W.w)) pgNumType.Remove();
            return true;
        });

    /// <summary>
    /// Set the page geometry (<c>w:pgSz</c> / <c>w:pgMar</c>) of the section that owns
    /// <paramref name="anchorId"/> (any body block in that section) — Word's <i>Page Setup</i>
    /// dialog. Null fields on <paramref name="op"/> leave that attribute alone; the elements are
    /// created in their <c>CT_SectPr</c> schema slots when absent, as is a trailing
    /// <c>w:sectPr</c> for a body that has none. Read back through <see cref="GetSectionInfo"/>.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see cref="PageSetupOp.Landscape"/> follows Word: <c>true</c> with no explicit size swaps a
    /// portrait-shaped page's width and height and writes <c>w:orient="landscape"</c>; <c>false</c>
    /// removes the attribute and swaps a landscape-shaped page back. With an explicit size the
    /// dimensions are written as given and only the attribute changes.
    /// </para>
    /// <para>
    /// Validation runs against the section's EFFECTIVE values after the op (the op merged over
    /// the current <c>w:pgSz</c>/<c>w:pgMar</c>, with Word's Letter/1-inch/0.5-inch defaults for
    /// absent attributes): width and height positive, margins and header/footer distances
    /// non-negative, <c>left + right &lt; width</c> and <c>top + bottom &lt; height</c>. A failure
    /// is <see cref="EditErrorCode.InvalidPageSetup"/>, leaves the document untouched and records no
    /// undo step. Values the section already has are a successful no-op that consumes no undo
    /// history either (same reasoning as <see cref="SetPageNumbering"/>).
    /// </para>
    /// </remarks>
    public EditResult SetPageSetup(string anchorId, PageSetupOp op)
    {
        ArgumentNullException.ThrowIfNull(op);
        if (op.PageWidthTwips is { } pw && pw <= 0)
            return EditResult.Fail(EditErrorCode.InvalidPageSetup, "page width must be positive", anchorId);
        if (op.PageHeightTwips is { } ph && ph <= 0)
            return EditResult.Fail(EditErrorCode.InvalidPageSetup, "page height must be positive", anchorId);
        foreach (var (name, value) in new[]
        {
            ("top margin", op.MarginTopTwips), ("bottom margin", op.MarginBottomTwips),
            ("left margin", op.MarginLeftTwips), ("right margin", op.MarginRightTwips),
            ("header distance", op.HeaderDistanceTwips), ("footer distance", op.FooterDistanceTwips),
        })
        {
            if (value is { } v && v < 0)
                return EditResult.Fail(EditErrorCode.InvalidPageSetup, $"{name} cannot be negative", anchorId);
        }

        // The cross-field checks need the section's CURRENT values, so they run inside the
        // mutator on the detached copy; a validation failure there throws and is mapped below,
        // before anything is snapshotted.
        EditResult result;
        try
        {
            result = EditGoverningSectPr(anchorId, "SetPageSetup", sectPr => ApplyPageSetup(sectPr, op));
        }
        catch (PageSetupException ex)
        {
            return EditResult.Fail(EditErrorCode.InvalidPageSetup, ex.Message, anchorId);
        }
        return result;
    }

    private sealed class PageSetupException : Exception
    {
        public PageSetupException(string message) : base(message) { }
    }

    /// <summary>The <see cref="SetPageSetup"/> mutator: returns false when the sectPr already
    /// carries the requested geometry, throws <see cref="PageSetupException"/> when the merged
    /// geometry is invalid, and otherwise writes it.</summary>
    private static bool ApplyPageSetup(XElement sectPr, PageSetupOp op)
    {
        var pgSz = sectPr.Element(W.pgSz);
        var pgMar = sectPr.Element(W.pgMar);
        static int Read(XElement? e, XName attr, int fallback) =>
            int.TryParse((string?)e?.Attribute(attr), System.Globalization.NumberStyles.Integer,
                System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : fallback;

        int curWidth = Read(pgSz, W._w, 12240);
        int curHeight = Read(pgSz, W.h, 15840);
        bool curLandscape = string.Equals((string?)pgSz?.Attribute(W.orient), "landscape", StringComparison.Ordinal);

        int width = op.PageWidthTwips ?? curWidth;
        int height = op.PageHeightTwips ?? curHeight;
        bool landscape = op.Landscape ?? curLandscape;
        bool explicitSize = op.PageWidthTwips is not null || op.PageHeightTwips is not null;
        if (op.Landscape is { } wantLandscape && !explicitSize)
        {
            // Word's orientation toggle rotates the sheet: only swap when the current shape
            // disagrees with the requested orientation, so a square or already-rotated page is
            // left alone.
            if (wantLandscape && width < height) (width, height) = (height, width);
            else if (!wantLandscape && width > height) (width, height) = (height, width);
        }

        int top = op.MarginTopTwips ?? Read(pgMar, W.top, 1440);
        int bottom = op.MarginBottomTwips ?? Read(pgMar, W.bottom, 1440);
        int left = op.MarginLeftTwips ?? Read(pgMar, W.left, 1440);
        int right = op.MarginRightTwips ?? Read(pgMar, W.right, 1440);
        int header = op.HeaderDistanceTwips
            ?? Read(pgMar, W.header, SectionInfo.DefaultHeaderFooterDistanceTwips);
        int footer = op.FooterDistanceTwips
            ?? Read(pgMar, W.footer, SectionInfo.DefaultHeaderFooterDistanceTwips);

        if (left + right >= width)
            throw new PageSetupException(
                $"left + right margins ({left} + {right} twips) leave no room on a {width}-twip-wide page");
        if (top + bottom >= height)
            throw new PageSetupException(
                $"top + bottom margins ({top} + {bottom} twips) leave no room on a {height}-twip-tall page");

        // w:pgSz is touched only when the op names size/orientation; w:pgMar only when it names a
        // margin. An op that touches neither is a no-op even on a section with no elements at all.
        bool touchesSize = explicitSize || op.Landscape is not null;
        bool touchesMargins = op.MarginTopTwips is not null || op.MarginBottomTwips is not null
            || op.MarginLeftTwips is not null || op.MarginRightTwips is not null
            || op.HeaderDistanceTwips is not null || op.FooterDistanceTwips is not null;

        static string S(int v) => v.ToString(System.Globalization.CultureInfo.InvariantCulture);
        bool changed = false;

        if (touchesSize)
        {
            var target = pgSz;
            if (target is null)
            {
                target = new XElement(W.pgSz);
                WordprocessingMLUtil.InsertSectPrChildInOrder(sectPr, target);
                changed = true;
            }
            changed |= SetAttr(target, W._w, S(width));
            changed |= SetAttr(target, W.h, S(height));
            changed |= SetAttr(target, W.orient, landscape ? "landscape" : null);
        }

        if (touchesMargins)
        {
            var target = pgMar;
            if (target is null)
            {
                target = new XElement(W.pgMar);
                WordprocessingMLUtil.InsertSectPrChildInOrder(sectPr, target);
                changed = true;
            }
            changed |= SetAttr(target, W.top, S(top));
            changed |= SetAttr(target, W.bottom, S(bottom));
            changed |= SetAttr(target, W.left, S(left));
            changed |= SetAttr(target, W.right, S(right));
            changed |= SetAttr(target, W.header, S(header));
            changed |= SetAttr(target, W.footer, S(footer));
        }

        return changed;
    }

    /// <summary>Set (or, for a null value, remove) an attribute; true when the element changed.</summary>
    private static bool SetAttr(XElement element, XName name, string? value)
    {
        var current = (string?)element.Attribute(name);
        if (string.Equals(current, value, StringComparison.Ordinal)) return false;
        element.SetAttributeValue(name, value);
        return true;
    }

    /// <summary>
    /// Shared body of the section-level verbs (page numbering, page setup): resolve a body anchor
    /// to its governing <c>w:sectPr</c> (synthesizing the document-final one if the body has none,
    /// as <see cref="SetHeaderText"/> does), then apply <paramref name="mutate"/>. The mutator
    /// returns false when the document already says what was asked, and NOTHING is snapshotted in
    /// that case.
    /// </summary>
    private EditResult EditGoverningSectPr(string anchorId, string opName, Func<XElement, bool> mutate)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Scope != "body")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"{opName} requires a body block anchor (the section it belongs to)",
                anchorId);
        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);
        var body = element.AncestorsAndSelf(W.body).FirstOrDefault();
        if (body is null)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "anchor is not in the document body", anchorId);

        var sectPr = Internal.BlockMetadataOps.FindGoverningSectPr(element);
        var succeeded = new EditResult { Success = true, Modified = new[] { target.Anchor } };

        // Decide on a detached COPY first. Returning before TakeSnapshot is what keeps a no-op out
        // of the bounded undo ring — and it also stops a no-op from synthesizing a sectPr that the
        // document did not ask for.
        if (!mutate(sectPr is null ? new XElement(W.sectPr) : new XElement(sectPr)))
            return succeeded;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (sectPr is null)
            {
                sectPr = new XElement(W.sectPr);
                body.Add(sectPr);
            }
            mutate(sectPr);
            InvalidateProjectionCache();
            return succeeded;
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>Turn the section's "Different first page" flag on. An existing
    /// <c>w:titlePg</c> loses its <c>w:val</c> (Word writes <c>w:val="0"</c> when the box is
    /// cleared, so presence is not "on"); otherwise one is inserted at its schema slot.</summary>
    private static void InsertSectPrTitlePg(XElement sectPr)
    {
        var existing = sectPr.Element(W.titlePg);
        if (existing is not null)
        {
            existing.Attributes(W.val).Remove();
            return;
        }
        WordprocessingMLUtil.InsertSectPrChildInOrder(sectPr, new XElement(W.titlePg));
    }
}
