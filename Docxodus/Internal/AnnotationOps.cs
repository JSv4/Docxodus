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

        var (startRunInsert, endRunInsert) = SplitRunsForSpan(map, spanStart, spanLength);

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

    public static EditResult Remove(WordprocessingDocument doc, string annotationId)
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
            touchedBlock = RemoveBookmarkPair(doc, bookmarkName!);
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
        CharSpan? newSpan)
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
            oldBlockAnchor = RemoveBookmarkPair(doc, bookmarkName!);

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
        var (startRunInsert, endRunInsert) = SplitRunsForSpan(newMap, s, l);
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
    /// with no other runs between them inside the span.
    /// Returns the start-side and end-side runs to insert before/after.
    /// </summary>
    private static (XElement startRun, XElement endRun) SplitRunsForSpan(
        RunTextMap.Map map, int start, int length)
    {
        var segments = RunTextMap.ResolveRange(map, start, length);
        // segments is guaranteed non-empty when length > 0 and bounds were checked.

        var first = segments[0];
        var last = segments[^1];

        // Handle the same-run case explicitly: the span starts AND ends inside a
        // single run. We have to split at most twice on the SAME run, and the
        // start/end run references must follow whatever split we did first.
        if (first.Segment.Run == last.Segment.Run)
        {
            var run = first.Segment.Run;
            var spanStartInRun = first.OffsetInRun;
            var spanEndInRun = first.OffsetInRun + first.Length; // last == first

            XElement startRunSingle = run;
            XElement endRunSingle = run;

            // Split off the trailing portion first (so the leading split's offset
            // remains valid against the unchanged left chunk).
            if (spanEndInRun < first.Segment.Length)
            {
                // Right half becomes a new sibling; we keep the left (run) as endRunSingle.
                _ = SplitRunAt(run, spanEndInRun, takeRightHalf: true);
                endRunSingle = run;
            }

            if (spanStartInRun > 0)
            {
                // Now split the (possibly shortened) left chunk at the span start.
                // The right half is the span content; both start and end reference it.
                var spanRun = SplitRunAt(run, spanStartInRun, takeRightHalf: true);
                startRunSingle = spanRun;
                endRunSingle = spanRun;
            }

            return (startRunSingle, endRunSingle);
        }

        // Span crosses run boundaries: at most one split on each end.
        XElement startRun = first.Segment.Run;
        if (first.OffsetInRun > 0)
        {
            startRun = SplitRunAt(startRun, first.OffsetInRun, takeRightHalf: true);
        }

        XElement endRun = last.Segment.Run;
        if (last.OffsetInRun + last.Length < last.Segment.Length)
        {
            endRun = SplitRunAt(endRun, last.OffsetInRun + last.Length, takeRightHalf: false);
        }

        return (startRun, endRun);
    }

    /// <summary>
    /// Splits a <c>w:r</c> element at <paramref name="offsetInRunText"/> characters
    /// from the start of its text. Returns either the original (left) run or the
    /// newly-inserted right-hand run, per <paramref name="takeRightHalf"/>.
    /// Only handles the common case of a run with a single <c>w:t</c> child;
    /// runs with multiple text fragments or other content are out of scope for
    /// v1 (the bookmark just sits at the closer of the two boundary positions).
    /// </summary>
    private static XElement SplitRunAt(XElement run, int offsetInRunText, bool takeRightHalf)
    {
        var text = run.Element(W + "t");
        if (text is null) return run;

        var full = text.Value;
        if (offsetInRunText <= 0 || offsetInRunText >= full.Length) return run;

        var leftText = full.Substring(0, offsetInRunText);
        var rightText = full.Substring(offsetInRunText);

        text.Value = leftText;
        if (string.IsNullOrEmpty(text.Attribute(XNamespace.Xml + "space")?.Value)
            && (leftText.StartsWith(' ') || leftText.EndsWith(' ')))
        {
            text.SetAttributeValue(XNamespace.Xml + "space", "preserve");
        }

        var rightRun = new XElement(run); // clone formatting + text
        var rightTextElement = rightRun.Element(W + "t")!;
        rightTextElement.Value = rightText;
        if (rightText.StartsWith(' ') || rightText.EndsWith(' '))
            rightTextElement.SetAttributeValue(XNamespace.Xml + "space", "preserve");

        run.AddAfterSelf(rightRun);
        return takeRightHalf ? rightRun : run;
    }

    private static Anchor? RemoveBookmarkPair(WordprocessingDocument doc, string bookmarkName)
    {
        Anchor? affectedBlock = null;
        foreach (var part in EnumerateParts(doc))
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
                var kind = KindOfElement(enclosing);
                var scope = ScopeOfPart(part);
                var unid = (string)enclosing.Attribute(PtOpenXml.Unid)!;
                affectedBlock = new Anchor(
                    Id: $"{kind}:{scope}:{unid}",
                    Kind: kind,
                    Scope: scope,
                    Unid: unid);
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
        var main = doc.MainDocumentPart;
        if (main is null) yield break;
        yield return main;
        foreach (var h in main.HeaderParts) yield return h;
        foreach (var f in main.FooterParts) yield return f;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart;
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

    private static string KindOfElement(XElement e)
    {
        if (e.Name == W + "p")
        {
            var pStyle = e.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val")?.Value;
            if (pStyle?.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) == true) return "h";
            if (e.Element(W + "pPr")?.Element(W + "numPr") is not null) return "li";
            return "p";
        }
        if (e.Name == W + "tbl") return "tbl";
        if (e.Name == W + "tr") return "tr";
        if (e.Name == W + "tc") return "tc";
        return "p";
    }

    private static string ScopeOfPart(OpenXmlPart part)
    {
        if (part is MainDocumentPart) return "body";
        if (part is HeaderPart) return "hdr";
        if (part is FooterPart) return "ftr";
        if (part is FootnotesPart) return "fn";
        if (part is EndnotesPart) return "en";
        return "body";
    }
}
