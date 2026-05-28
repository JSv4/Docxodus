// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Anchor-addressed annotation mutations on an open <see cref="WordprocessingDocument"/>.
/// Shared backend for <see cref="DocxSession.AddAnnotation"/>,
/// <see cref="DocxSession.RemoveAnnotation"/>, <see cref="DocxSession.UpdateAnnotation"/>,
/// and <see cref="DocxSession.MoveAnnotation"/>.
/// </summary>
internal static class AnnotationOps
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static EditResult Add(
        WordprocessingDocument doc,
        AnchorTarget anchor,
        CharSpan? span,
        DocumentAnnotation annotation)
    {
        ArgumentNullException.ThrowIfNull(annotation);

        var block = anchor.Resolve(doc);
        if (block is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound,
                "element resolved null", anchor.Anchor.Id);

        // Resolve id (auto-generate or check for collision).
        var id = string.IsNullOrEmpty(annotation.Id) ? null : annotation.Id;
        if (id is null)
        {
            id = GenerateUniqueId(doc);
            if (id is null)
                return EditResult.Fail(EditErrorCode.DuplicateAnnotationId,
                    "auto-id collided 4 times", anchor.Anchor.Id);
        }
        else if (AnnotationsCustomXml.FindById(doc, id) is not null)
        {
            return EditResult.Fail(EditErrorCode.DuplicateAnnotationId,
                $"annotation id already exists: {id}", anchor.Anchor.Id);
        }

        // Build the run text map and resolve span.
        var map = RunTextMap.Build(block);
        int spanStart, spanLength;
        if (span.HasValue)
        {
            spanStart = span.Value.Start;
            spanLength = span.Value.Length;
            if (spanLength <= 0)
                return EditResult.Fail(EditErrorCode.EmptyAnnotationSpan,
                    "span length must be > 0", anchor.Anchor.Id);
            if (spanStart < 0 || spanStart + spanLength > map.FlatText.Length)
                return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                    $"span [{spanStart},{spanStart + spanLength}) outside block " +
                    $"of length {map.FlatText.Length}", anchor.Anchor.Id);
        }
        else
        {
            spanStart = 0;
            spanLength = map.FlatText.Length;
            if (spanLength == 0)
                return EditResult.Fail(EditErrorCode.EmptyAnnotationSpan,
                    "block has no inline runs to bookmark", anchor.Anchor.Id);
        }

        var annotatedText = map.FlatText.Substring(spanStart, spanLength);

        // Insert bookmarkStart/bookmarkEnd around the span.
        var bookmarkName = AnnotationManager.BookmarkPrefix + id;
        var bookmarkId = NextBookmarkId(block.Document!.Root!);

        var (startRunInsert, endRunInsert) = SplitRunsForSpan(block, spanStart, spanLength);

        var bookmarkStart = new XElement(W + "bookmarkStart",
            new XAttribute(W + "id", bookmarkId),
            new XAttribute(W + "name", bookmarkName));
        var bookmarkEnd = new XElement(W + "bookmarkEnd",
            new XAttribute(W + "id", bookmarkId));

        startRunInsert.AddBeforeSelf(bookmarkStart);
        endRunInsert.AddAfterSelf(bookmarkEnd);

        // Persist custom XML.
        annotation.Id = id;
        annotation.BookmarkName = bookmarkName;
        annotation.Created ??= DateTime.UtcNow;
        annotation.AnnotatedText = annotatedText;
        annotation.PageInfoStale = true;
        AnnotationsCustomXml.Write(doc, annotation);

        // Persist part XML.
        SavePart(doc, anchor.PartUri);

        return new EditResult
        {
            Success = true,
            AnnotationId = id,
            Modified = new[] { anchor.Anchor },
        };
    }

    public static EditResult Remove(
        WordprocessingDocument doc,
        string annotationId,
        Func<string, Anchor?>? canonicalize = null)
    {
        if (string.IsNullOrEmpty(annotationId))
            return EditResult.Fail(EditErrorCode.AnnotationNotFound,
                "annotation id required");

        var existing = AnnotationsCustomXml.FindById(doc, annotationId);
        if (existing is null)
            return EditResult.Fail(EditErrorCode.AnnotationNotFound,
                $"annotation not found: {annotationId}");

        var bookmarkName = existing.BookmarkName;
        Anchor? touchedBlock = null;
        if (!string.IsNullOrEmpty(bookmarkName))
        {
            touchedBlock = RemoveBookmarkPair(doc, bookmarkName!, canonicalize);
        }

        AnnotationsCustomXml.Remove(doc, annotationId);

        return new EditResult
        {
            Success = true,
            AnnotationId = annotationId,
            Modified = touchedBlock is null
                ? Array.Empty<Anchor>()
                : new[] { touchedBlock.Value },
        };
    }

    public static EditResult Update(
        WordprocessingDocument doc,
        string annotationId,
        AnnotationUpdate update)
    {
        ArgumentNullException.ThrowIfNull(update);

        var existing = AnnotationsCustomXml.FindById(doc, annotationId);
        if (existing is null)
            return EditResult.Fail(EditErrorCode.AnnotationNotFound,
                $"annotation not found: {annotationId}");

        if (update.LabelId is not null) existing.LabelId = update.LabelId;
        if (update.Label is not null) existing.Label = update.Label;
        if (update.Color is not null) existing.Color = update.Color;
        if (update.Author is not null) existing.Author = update.Author;
        if (update.MetadataPatch is not null)
        {
            existing.Metadata ??= new Dictionary<string, string>();
            foreach (var (key, value) in update.MetadataPatch)
            {
                if (value is null) existing.Metadata.Remove(key);
                else existing.Metadata[key] = value;
            }
        }

        AnnotationsCustomXml.Write(doc, existing);

        return new EditResult
        {
            Success = true,
            AnnotationId = annotationId,
        };
    }

    public static EditResult Move(
        WordprocessingDocument doc,
        string annotationId,
        AnchorTarget newAnchor,
        CharSpan? newSpan,
        Func<string, Anchor?>? canonicalize = null)
    {
        var existing = AnnotationsCustomXml.FindById(doc, annotationId);
        if (existing is null)
            return EditResult.Fail(EditErrorCode.AnnotationNotFound,
                $"annotation not found: {annotationId}");

        // Validate the new range BEFORE removing the old bookmark so we don't
        // strand the annotation.
        var newBlock = newAnchor.Resolve(doc);
        if (newBlock is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound,
                "element resolved null", newAnchor.Anchor.Id);

        var newMap = RunTextMap.Build(newBlock);
        int s, l;
        if (newSpan.HasValue)
        {
            s = newSpan.Value.Start;
            l = newSpan.Value.Length;
            if (l <= 0)
                return EditResult.Fail(EditErrorCode.EmptyAnnotationSpan,
                    "span length must be > 0", newAnchor.Anchor.Id);
            if (s < 0 || s + l > newMap.FlatText.Length)
                return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                    $"span [{s},{s + l}) outside block of length {newMap.FlatText.Length}",
                    newAnchor.Anchor.Id);
        }
        else
        {
            s = 0;
            l = newMap.FlatText.Length;
            if (l == 0)
                return EditResult.Fail(EditErrorCode.EmptyAnnotationSpan,
                    "block has no inline runs to bookmark", newAnchor.Anchor.Id);
        }

        var bookmarkName = existing.BookmarkName;
        Anchor? oldBlockAnchor = null;
        if (!string.IsNullOrEmpty(bookmarkName))
            oldBlockAnchor = RemoveBookmarkPair(doc, bookmarkName!, canonicalize);

        // Old bookmark removal may have invalidated the cached run map of the
        // new block when the old and new blocks are the same element. Rebuild.
        if (oldBlockAnchor is not null && oldBlockAnchor.Value.Id == newAnchor.Anchor.Id)
        {
            newBlock = newAnchor.Resolve(doc)!;
            newMap = RunTextMap.Build(newBlock);
            if (s + l > newMap.FlatText.Length)
                return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                    $"span [{s},{s + l}) outside block of length {newMap.FlatText.Length} " +
                    "(after old bookmark removal)", newAnchor.Anchor.Id);
        }

        // Reinsert with a fresh bookmark id at the new range.
        var bookmarkId = NextBookmarkId(newBlock.Document!.Root!);
        var (startRunInsert, endRunInsert) = SplitRunsForSpan(newBlock, s, l);
        startRunInsert.AddBeforeSelf(new XElement(W + "bookmarkStart",
            new XAttribute(W + "id", bookmarkId),
            new XAttribute(W + "name", bookmarkName!)));
        endRunInsert.AddAfterSelf(new XElement(W + "bookmarkEnd",
            new XAttribute(W + "id", bookmarkId)));

        existing.AnnotatedText = newMap.FlatText.Substring(s, l);
        existing.PageInfoStale = true;
        AnnotationsCustomXml.Write(doc, existing);

        SavePart(doc, newAnchor.PartUri);

        var modified = oldBlockAnchor is null || oldBlockAnchor.Value.Id == newAnchor.Anchor.Id
            ? new[] { newAnchor.Anchor }
            : new[] { oldBlockAnchor.Value, newAnchor.Anchor };

        return new EditResult
        {
            Success = true,
            AnnotationId = annotationId,
            Modified = modified,
        };
    }

    // ─── helpers ───────────────────────────────────────────────────────────

    private static string? GenerateUniqueId(WordprocessingDocument doc)
    {
        for (int i = 0; i < 4; i++)
        {
            var candidate = Guid.NewGuid().ToString("N").Substring(0, 16);
            if (AnnotationsCustomXml.FindById(doc, candidate) is null)
                return candidate;
        }
        return null;
    }

    private static int NextBookmarkId(XElement root)
    {
        var max = root.Descendants(W + "bookmarkStart")
            .Select(b => (int?)b.Attribute(W + "id"))
            .Where(v => v.HasValue)
            .Select(v => v!.Value)
            .DefaultIfEmpty(0)
            .Max();
        return max + 1;
    }

    /// <summary>
    /// Splits the runs at the span boundaries (when boundaries fall mid-run) so
    /// that <c>w:bookmarkStart</c> can be inserted before the run containing the
    /// span start and <c>w:bookmarkEnd</c> after the run containing the span end,
    /// with no other runs between them inside the span. Returns the start-side
    /// and end-side runs to insert before/after.
    ///
    /// Routes through <see cref="DocxSession.SplitRunsAtOffset"/>, which correctly
    /// handles multi-<c>w:t</c> runs and runs nested inside hyperlinks/sdts —
    /// boundaries always land on a run-element boundary after the splits run.
    /// </summary>
    private static (XElement startRun, XElement endRun) SplitRunsForSpan(
        XElement block, int start, int length)
    {
        // Split at the END boundary first — if start == end - 0 length would be illegal,
        // but length > 0 is guaranteed by the caller. Splitting end first keeps the
        // start offset valid (the left-of-start text is unchanged by an end-side split).
        int endOffset = start + length;
        DocxSession.SplitRunsAtOffset(block, endOffset);
        DocxSession.SplitRunsAtOffset(block, start);

        // After both splits, every run boundary in the span [start, end) falls on a
        // run-element boundary. Rebuild the text map and pick the first/last segment
        // whose flat-offset range falls inside [start, end).
        var map = RunTextMap.Build(block);
        XElement? startRun = null;
        XElement? endRun = null;
        foreach (var seg in map.Segments)
        {
            if (seg.EndOffsetInBlock <= start) continue;
            if (seg.StartOffsetInBlock >= endOffset) break;
            startRun ??= seg.Run;
            endRun = seg.Run;
        }

        // length > 0 with bounds-checked offsets means at least one segment overlaps,
        // unless the span is entirely zero-width text (impossible — RunTextMap excludes
        // empty runs). The bang is justified by the precondition.
        return (startRun!, endRun!);
    }

    /// <summary>
    /// Removes the bookmark pair matching <paramref name="bookmarkName"/> across every
    /// part we recognise, and returns the anchor of the enclosing block element so
    /// the session can report it as <see cref="EditResult.Modified"/>.
    ///
    /// The returned anchor is canonicalised against the live projection via
    /// <paramref name="canonicalize"/> when supplied — that's the only way to be
    /// sure the kind ("p" vs "h" vs "li") and scope ("hdr1" vs "hdr2" vs …) match
    /// what the next <see cref="DocxSession.Project"/> will emit. The local fallback
    /// (used only when no callback is supplied) reproduces the projector's logic
    /// inline so even the un-canonicalised result is correct against the standard
    /// layouts.
    /// </summary>
    private static Anchor? RemoveBookmarkPair(
        WordprocessingDocument doc,
        string bookmarkName,
        Func<string, Anchor?>? canonicalize)
    {
        Anchor? affectedBlock = null;
        foreach (var (part, scope) in EnumeratePartsWithScope(doc))
        {
            var root = part.GetXDocument().Root;
            if (root is null) continue;
            var start = root.Descendants(W + "bookmarkStart")
                .FirstOrDefault(b => (string?)b.Attribute(W + "name") == bookmarkName);
            if (start is null) continue;

            var id = (string?)start.Attribute(W + "id");
            var end = id is null
                ? null
                : root.Descendants(W + "bookmarkEnd")
                    .FirstOrDefault(b => (string?)b.Attribute(W + "id") == id);

            // Locate enclosing block for the Modified anchor.
            var enclosing = start.AncestorsAndSelf()
                .FirstOrDefault(e => (string?)e.Attribute(PtOpenXml.Unid) is not null);
            if (enclosing is not null)
            {
                var unid = (string)enclosing.Attribute(PtOpenXml.Unid)!;
                // Prefer the live projection — it's the authoritative source of
                // kind+scope and is guaranteed to match what the agent will see
                // after Project() on the next tick.
                Anchor? canonical = canonicalize?.Invoke(unid);
                if (canonical is not null)
                {
                    affectedBlock = canonical;
                }
                else
                {
                    // Fallback: classifier + scope iteration that mirrors
                    // WmlToMarkdownConverter.BuildAnchorIndex exactly.
                    var kind = DocxSession.ClassifyBlockKind(enclosing);
                    affectedBlock = new Anchor(
                        Id: $"{kind}:{scope}:{unid}",
                        Kind: kind,
                        Scope: scope,
                        Unid: unid);
                }
            }

            start.Remove();
            end?.Remove();
            part.PutXDocument();
            break;
        }
        return affectedBlock;
    }

    private static IEnumerable<OpenXmlPart> EnumerateParts(WordprocessingDocument doc)
    {
        foreach (var (part, _) in EnumeratePartsWithScope(doc)) yield return part;
    }

    /// <summary>
    /// Enumerate parts paired with the scope name the projector would assign.
    /// Header/footer scopes get a 1-based index ("hdr1", "hdr2", …) matching the
    /// loop in <see cref="WmlToMarkdownConverter.BuildAnchorIndex"/>.
    /// </summary>
    private static IEnumerable<(OpenXmlPart Part, string Scope)> EnumeratePartsWithScope(
        WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart;
        if (main is null) yield break;
        yield return (main, "body");
        int h = 1;
        foreach (var hp in main.HeaderParts) yield return (hp, $"hdr{h++}");
        int f = 1;
        foreach (var fp in main.FooterParts) yield return (fp, $"ftr{f++}");
        if (main.FootnotesPart is not null) yield return (main.FootnotesPart, "fn");
        if (main.EndnotesPart is not null) yield return (main.EndnotesPart, "en");
    }

    private static void SavePart(WordprocessingDocument doc, string partUri)
    {
        foreach (var part in EnumerateParts(doc))
        {
            if (part.Uri.ToString() == partUri)
            {
                part.PutXDocument();
                return;
            }
        }
    }
}
