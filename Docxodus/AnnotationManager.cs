// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus
{
    /// <summary>
    /// Manages custom annotations in DOCX documents.
    /// Annotations are stored in a Custom XML Part and linked to document content via bookmarks.
    /// </summary>
    public static class AnnotationManager
    {
        /// <summary>
        /// XML namespace for Docxodus annotations.
        /// </summary>
        public const string AnnotationsNamespace = "http://docxodus.dev/annotations/v1";

        /// <summary>
        /// Prefix for annotation bookmark names.
        /// </summary>
        public const string BookmarkPrefix = "_Docxodus_Ann_";

        /// <summary>
        /// Content type for the annotations custom XML part.
        /// </summary>
        private const string AnnotationsContentType = "application/xml";

        private static readonly XNamespace Ann = AnnotationsNamespace;

        /// <summary>
        /// Add an annotation to a document.
        /// </summary>
        /// <param name="doc">The document to annotate.</param>
        /// <param name="annotation">The annotation to add.</param>
        /// <param name="range">Specifies the text range to annotate.</param>
        /// <returns>A new document with the annotation added.</returns>
        public static WmlDocument AddAnnotation(
            WmlDocument doc,
            DocumentAnnotation annotation,
            AnnotationRange range)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (annotation == null) throw new ArgumentNullException(nameof(annotation));
            if (range == null) throw new ArgumentNullException(nameof(range));
            if (string.IsNullOrEmpty(annotation.Id)) throw new ArgumentException("Annotation ID is required.", nameof(annotation));

            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            {
                using (WordprocessingDocument wordDoc = streamDoc.GetWordprocessingDocument())
                {
                    // Check for duplicate ID
                    var existingAnnotations = GetAnnotationsInternal(wordDoc);
                    if (existingAnnotations.Any(a => a.Id == annotation.Id))
                    {
                        throw new InvalidOperationException($"An annotation with ID '{annotation.Id}' already exists.");
                    }

                    // Set bookmark name if not already set
                    if (string.IsNullOrEmpty(annotation.BookmarkName))
                    {
                        annotation.BookmarkName = BookmarkPrefix + annotation.Id;
                    }

                    // Create or find the bookmark
                    if (!string.IsNullOrEmpty(range.ExistingBookmarkName))
                    {
                        // Use existing bookmark - verify it exists
                        var mainDoc = wordDoc.MainDocumentPart.GetXDocument();
                        var existingBookmark = mainDoc.Descendants(W.bookmarkStart)
                            .FirstOrDefault(b => (string)b.Attribute(W.name) == range.ExistingBookmarkName);

                        if (existingBookmark == null)
                        {
                            throw new InvalidOperationException($"Bookmark '{range.ExistingBookmarkName}' not found.");
                        }

                        // Update annotation to reference this bookmark
                        annotation.BookmarkName = range.ExistingBookmarkName;
                    }
                    else if (!string.IsNullOrEmpty(range.SearchText))
                    {
                        // Search for text and create bookmark
                        CreateBookmarkFromSearch(wordDoc, annotation.BookmarkName, range.SearchText, range.Occurrence);
                    }
                    else if (range.StartParagraphIndex.HasValue)
                    {
                        // Create bookmark from paragraph/run indices
                        CreateBookmarkFromIndices(wordDoc, annotation.BookmarkName, range);
                    }
                    else
                    {
                        throw new ArgumentException("Range must specify SearchText, ExistingBookmarkName, or paragraph indices.", nameof(range));
                    }

                    // Add to custom XML part
                    AddAnnotationToCustomXml(wordDoc, annotation);

                    wordDoc.MainDocumentPart.PutXDocument();
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        /// <summary>
        /// Remove an annotation by ID.
        /// </summary>
        /// <param name="doc">The document containing the annotation.</param>
        /// <param name="annotationId">The annotation ID to remove.</param>
        /// <returns>A new document with the annotation removed.</returns>
        public static WmlDocument RemoveAnnotation(WmlDocument doc, string annotationId)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (string.IsNullOrEmpty(annotationId)) throw new ArgumentNullException(nameof(annotationId));

            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            {
                using (WordprocessingDocument wordDoc = streamDoc.GetWordprocessingDocument())
                {
                    // Get the annotation to find the bookmark name
                    var annotation = GetAnnotationInternal(wordDoc, annotationId);
                    if (annotation == null)
                    {
                        // No annotation found - return unchanged
                        return doc;
                    }

                    // Remove bookmark from document (only if it's our annotation bookmark)
                    if (annotation.BookmarkName?.StartsWith(BookmarkPrefix) == true)
                    {
                        RemoveBookmark(wordDoc, annotation.BookmarkName);
                    }

                    // Remove from custom XML part
                    RemoveAnnotationFromCustomXml(wordDoc, annotationId);

                    wordDoc.MainDocumentPart.PutXDocument();
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        /// <summary>
        /// Get all annotations from a document.
        /// </summary>
        /// <param name="doc">The document to read.</param>
        /// <returns>List of all annotations in the document.</returns>
        public static List<DocumentAnnotation> GetAnnotations(WmlDocument doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            using (WordprocessingDocument wordDoc = streamDoc.GetWordprocessingDocument())
            {
                return GetAnnotationsInternal(wordDoc);
            }
        }

        /// <summary>
        /// Get a specific annotation by ID.
        /// </summary>
        /// <param name="doc">The document to read.</param>
        /// <param name="annotationId">The annotation ID.</param>
        /// <returns>The annotation, or null if not found.</returns>
        public static DocumentAnnotation GetAnnotation(WmlDocument doc, string annotationId)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (string.IsNullOrEmpty(annotationId)) throw new ArgumentNullException(nameof(annotationId));

            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            using (WordprocessingDocument wordDoc = streamDoc.GetWordprocessingDocument())
            {
                return GetAnnotationInternal(wordDoc, annotationId);
            }
        }

        /// <summary>
        /// Update an existing annotation's metadata (not range).
        /// </summary>
        /// <param name="doc">The document containing the annotation.</param>
        /// <param name="annotation">The annotation with updated values.</param>
        /// <returns>A new document with the annotation updated.</returns>
        public static WmlDocument UpdateAnnotation(WmlDocument doc, DocumentAnnotation annotation)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (annotation == null) throw new ArgumentNullException(nameof(annotation));
            if (string.IsNullOrEmpty(annotation.Id)) throw new ArgumentException("Annotation ID is required.", nameof(annotation));

            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            {
                using (WordprocessingDocument wordDoc = streamDoc.GetWordprocessingDocument())
                {
                    // Find existing annotation
                    var existing = GetAnnotationInternal(wordDoc, annotation.Id);
                    if (existing == null)
                    {
                        throw new InvalidOperationException($"Annotation with ID '{annotation.Id}' not found.");
                    }

                    // Preserve bookmark name from existing
                    annotation.BookmarkName = existing.BookmarkName;

                    // Remove old and add new
                    RemoveAnnotationFromCustomXml(wordDoc, annotation.Id);
                    AddAnnotationToCustomXml(wordDoc, annotation);
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        /// <summary>
        /// Update cached page span information for annotations.
        /// </summary>
        /// <param name="doc">The document to update.</param>
        /// <param name="pageSpans">Dictionary mapping annotation IDs to page spans.</param>
        /// <returns>A new document with updated page information.</returns>
        public static WmlDocument UpdateAnnotationPageSpans(
            WmlDocument doc,
            Dictionary<string, (int startPage, int endPage)> pageSpans)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (pageSpans == null || pageSpans.Count == 0) return doc;

            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            {
                using (WordprocessingDocument wordDoc = streamDoc.GetWordprocessingDocument())
                {
                    var customXmlPart = GetOrCreateAnnotationsCustomXmlPart(wordDoc);
                    var xdoc = customXmlPart.GetXDocument();

                    var root = xdoc.Root;
                    if (root == null) return doc;

                    foreach (var (annotationId, (startPage, endPage)) in pageSpans)
                    {
                        var annotationElement = root.Elements(Ann + "annotation")
                            .FirstOrDefault(a => (string)a.Attribute("id") == annotationId);

                        if (annotationElement != null)
                        {
                            // Remove existing pageSpan
                            annotationElement.Element(Ann + "pageSpan")?.Remove();

                            // Add updated pageSpan
                            annotationElement.Add(new XElement(Ann + "pageSpan",
                                new XAttribute("startPage", startPage),
                                new XAttribute("endPage", endPage),
                                new XAttribute("stale", "false"),
                                new XAttribute("computedAt", DateTime.UtcNow.ToString("o"))));
                        }
                    }

                    customXmlPart.PutXDocument();
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        /// <summary>
        /// Check if a document has any annotations.
        /// </summary>
        /// <param name="doc">The document to check.</param>
        /// <returns>True if the document has annotations.</returns>
        public static bool HasAnnotations(WmlDocument doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            using (WordprocessingDocument wordDoc = streamDoc.GetWordprocessingDocument())
            {
                var customXmlPart = FindAnnotationsCustomXmlPart(wordDoc);
                if (customXmlPart == null) return false;

                var xdoc = customXmlPart.GetXDocument();
                return xdoc.Root?.Elements(Ann + "annotation").Any() == true;
            }
        }

        /// <summary>
        /// Get the text content within an annotation's range.
        /// </summary>
        /// <param name="doc">The document containing the annotation.</param>
        /// <param name="annotationId">The annotation ID.</param>
        /// <returns>The text content, or null if not found.</returns>
        public static string GetAnnotatedText(WmlDocument doc, string annotationId)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));
            if (string.IsNullOrEmpty(annotationId)) throw new ArgumentNullException(nameof(annotationId));

            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
            using (WordprocessingDocument wordDoc = streamDoc.GetWordprocessingDocument())
            {
                var annotation = GetAnnotationInternal(wordDoc, annotationId);
                if (annotation == null) return null;

                return GetTextInBookmark(wordDoc, annotation.BookmarkName);
            }
        }

        #region Internal Methods

        private static List<DocumentAnnotation> GetAnnotationsInternal(WordprocessingDocument wordDoc)
        {
            var annotations = new List<DocumentAnnotation>();

            var customXmlPart = FindAnnotationsCustomXmlPart(wordDoc);
            if (customXmlPart == null) return annotations;

            var xdoc = customXmlPart.GetXDocument();
            var root = xdoc.Root;
            if (root == null) return annotations;

            foreach (var element in root.Elements(Ann + "annotation"))
            {
                var annotation = ParseAnnotationElement(element);
                if (annotation != null)
                {
                    // Optionally populate annotated text
                    annotation.AnnotatedText = GetTextInBookmark(wordDoc, annotation.BookmarkName);
                    annotations.Add(annotation);
                }
            }

            return annotations;
        }

        private static DocumentAnnotation GetAnnotationInternal(WordprocessingDocument wordDoc, string annotationId)
        {
            var customXmlPart = FindAnnotationsCustomXmlPart(wordDoc);
            if (customXmlPart == null) return null;

            var xdoc = customXmlPart.GetXDocument();
            var element = xdoc.Root?.Elements(Ann + "annotation")
                .FirstOrDefault(a => (string)a.Attribute("id") == annotationId);

            if (element == null) return null;

            var annotation = ParseAnnotationElement(element);
            if (annotation != null)
            {
                annotation.AnnotatedText = GetTextInBookmark(wordDoc, annotation.BookmarkName);
            }

            return annotation;
        }

        private static DocumentAnnotation ParseAnnotationElement(XElement element)
        {
            var id = (string)element.Attribute("id");
            if (string.IsNullOrEmpty(id)) return null;

            var annotation = new DocumentAnnotation
            {
                Id = id,
                LabelId = (string)element.Attribute("labelId"),
                Label = (string)element.Attribute("label"),
                Color = (string)element.Attribute("color"),
                Author = (string)element.Attribute("author"),
            };

            var createdStr = (string)element.Attribute("created");
            if (DateTime.TryParse(createdStr, out var created))
            {
                annotation.Created = created;
            }

            var rangeElement = element.Element(Ann + "range");
            if (rangeElement != null)
            {
                annotation.BookmarkName = (string)rangeElement.Attribute("bookmarkName");
            }

            var pageSpanElement = element.Element(Ann + "pageSpan");
            if (pageSpanElement != null)
            {
                var startPageStr = (string)pageSpanElement.Attribute("startPage");
                var endPageStr = (string)pageSpanElement.Attribute("endPage");
                var staleStr = (string)pageSpanElement.Attribute("stale");
                var computedAtStr = (string)pageSpanElement.Attribute("computedAt");

                if (int.TryParse(startPageStr, out var startPage))
                    annotation.StartPage = startPage;
                if (int.TryParse(endPageStr, out var endPage))
                    annotation.EndPage = endPage;

                annotation.PageInfoStale = staleStr?.ToLowerInvariant() == "true";

                if (DateTime.TryParse(computedAtStr, out var computedAt))
                    annotation.PageInfoComputedAt = computedAt;
            }

            var metadataElement = element.Element(Ann + "metadata");
            if (metadataElement != null)
            {
                foreach (var item in metadataElement.Elements(Ann + "item"))
                {
                    var key = (string)item.Attribute("key");
                    var value = item.Value;
                    if (!string.IsNullOrEmpty(key))
                    {
                        annotation.Metadata[key] = value;
                    }
                }
            }

            return annotation;
        }

        private static void AddAnnotationToCustomXml(WordprocessingDocument wordDoc, DocumentAnnotation annotation)
        {
            var customXmlPart = GetOrCreateAnnotationsCustomXmlPart(wordDoc);
            var xdoc = customXmlPart.GetXDocument();

            var annotationElement = new XElement(Ann + "annotation",
                new XAttribute("id", annotation.Id),
                new XAttribute("labelId", annotation.LabelId ?? ""),
                new XAttribute("label", annotation.Label ?? ""),
                new XAttribute("color", annotation.Color ?? "#FFFF00"));

            if (!string.IsNullOrEmpty(annotation.Author))
                annotationElement.Add(new XAttribute("author", annotation.Author));

            if (annotation.Created.HasValue)
                annotationElement.Add(new XAttribute("created", annotation.Created.Value.ToString("o")));

            // Range element
            annotationElement.Add(new XElement(Ann + "range",
                new XAttribute("bookmarkName", annotation.BookmarkName ?? "")));

            // Page span (if available)
            if (annotation.StartPage.HasValue && annotation.EndPage.HasValue)
            {
                annotationElement.Add(new XElement(Ann + "pageSpan",
                    new XAttribute("startPage", annotation.StartPage.Value),
                    new XAttribute("endPage", annotation.EndPage.Value),
                    new XAttribute("stale", annotation.PageInfoStale ? "true" : "false")));
            }

            // Metadata
            if (annotation.Metadata?.Count > 0)
            {
                var metadataElement = new XElement(Ann + "metadata");
                foreach (var (key, value) in annotation.Metadata)
                {
                    metadataElement.Add(new XElement(Ann + "item",
                        new XAttribute("key", key),
                        value ?? ""));
                }
                annotationElement.Add(metadataElement);
            }

            xdoc.Root.Add(annotationElement);
            customXmlPart.PutXDocument();
        }

        private static void RemoveAnnotationFromCustomXml(WordprocessingDocument wordDoc, string annotationId)
        {
            var customXmlPart = FindAnnotationsCustomXmlPart(wordDoc);
            if (customXmlPart == null) return;

            var xdoc = customXmlPart.GetXDocument();
            var element = xdoc.Root?.Elements(Ann + "annotation")
                .FirstOrDefault(a => (string)a.Attribute("id") == annotationId);

            if (element != null)
            {
                element.Remove();
                customXmlPart.PutXDocument();
            }
        }

        private static CustomXmlPart FindAnnotationsCustomXmlPart(WordprocessingDocument wordDoc)
        {
            foreach (var customXmlPart in wordDoc.MainDocumentPart.CustomXmlParts)
            {
                try
                {
                    var xdoc = customXmlPart.GetXDocument();
                    if (xdoc.Root?.Name.Namespace == Ann && xdoc.Root.Name.LocalName == "annotations")
                    {
                        return customXmlPart;
                    }
                }
                catch
                {
                    // Skip parts that can't be parsed as XML
                }
            }

            return null;
        }

        private static CustomXmlPart GetOrCreateAnnotationsCustomXmlPart(WordprocessingDocument wordDoc)
        {
            var existing = FindAnnotationsCustomXmlPart(wordDoc);
            if (existing != null) return existing;

            // Create new custom XML part
            var customXmlPart = wordDoc.MainDocumentPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);

            var xdoc = new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                new XElement(Ann + "annotations",
                    new XAttribute("version", "1.0")));

            customXmlPart.PutXDocument(xdoc);

            return customXmlPart;
        }

        private static void CreateBookmarkFromSearch(
            WordprocessingDocument wordDoc,
            string bookmarkName,
            string searchText,
            int occurrence)
        {
            var mainDoc = wordDoc.MainDocumentPart.GetXDocument();
            var body = mainDoc.Root?.Element(W.body);
            if (body == null)
                throw new InvalidOperationException("Document body not found.");

            // Find all text elements and locate the search text
            var textNodes = body.Descendants(W.t).ToList();
            var fullText = string.Concat(textNodes.Select(t => t.Value));

            int startIndex = -1;
            int currentOccurrence = 0;
            int searchPos = 0;

            while ((searchPos = fullText.IndexOf(searchText, searchPos, StringComparison.Ordinal)) >= 0)
            {
                currentOccurrence++;
                if (currentOccurrence == occurrence)
                {
                    startIndex = searchPos;
                    break;
                }
                searchPos++;
            }

            if (startIndex < 0)
            {
                throw new InvalidOperationException(
                    $"Could not find occurrence {occurrence} of text '{searchText}'.");
            }

            int endIndex = startIndex + searchText.Length;

            // Find which text elements contain the start and end positions
            int currentPos = 0;
            XElement startTextElement = null;
            XElement endTextElement = null;
            int startOffsetInElement = 0;
            int endOffsetInElement = 0;

            foreach (var textNode in textNodes)
            {
                var text = textNode.Value;
                var elementStart = currentPos;
                var elementEnd = currentPos + text.Length;

                if (startTextElement == null && startIndex >= elementStart && startIndex < elementEnd)
                {
                    startTextElement = textNode;
                    startOffsetInElement = startIndex - elementStart;
                }

                if (endIndex > elementStart && endIndex <= elementEnd)
                {
                    endTextElement = textNode;
                    endOffsetInElement = endIndex - elementStart;
                }

                currentPos = elementEnd;

                if (startTextElement != null && endTextElement != null)
                    break;
            }

            if (startTextElement == null || endTextElement == null)
            {
                throw new InvalidOperationException("Could not locate text elements for bookmark.");
            }

            // Generate unique bookmark ID
            var existingIds = mainDoc.Descendants(W.bookmarkStart)
                .Select(b => (int?)b.Attribute(W.id))
                .Where(id => id.HasValue)
                .Select(id => id.Value)
                .ToHashSet();

            int newId = 1;
            while (existingIds.Contains(newId)) newId++;

            // Insert bookmark markers
            // If start and end are in the same text element
            if (startTextElement == endTextElement)
            {
                InsertBookmarkInSameTextElement(startTextElement, bookmarkName, newId,
                    startOffsetInElement, endOffsetInElement);
            }
            else
            {
                // Different text elements
                InsertBookmarkStart(startTextElement, bookmarkName, newId, startOffsetInElement);
                InsertBookmarkEnd(endTextElement, bookmarkName, newId, endOffsetInElement);
            }

            wordDoc.MainDocumentPart.PutXDocument();
        }

        private static void InsertBookmarkInSameTextElement(
            XElement textElement,
            string bookmarkName,
            int bookmarkId,
            int startOffset,
            int endOffset)
        {
            var text = textElement.Value;
            var run = textElement.Parent;
            if (run?.Name != W.r) return;

            var beforeText = text.Substring(0, startOffset);
            var markedText = text.Substring(startOffset, endOffset - startOffset);
            var afterText = text.Substring(endOffset);

            var newElements = new List<object>();

            if (beforeText.Length > 0)
            {
                newElements.Add(new XElement(W.r,
                    run.Elements(W.rPr).Select(e => new XElement(e)),
                    new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), beforeText)));
            }

            newElements.Add(new XElement(W.bookmarkStart,
                new XAttribute(W.id, bookmarkId),
                new XAttribute(W.name, bookmarkName)));

            newElements.Add(new XElement(W.r,
                run.Elements(W.rPr).Select(e => new XElement(e)),
                new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), markedText)));

            newElements.Add(new XElement(W.bookmarkEnd,
                new XAttribute(W.id, bookmarkId)));

            if (afterText.Length > 0)
            {
                newElements.Add(new XElement(W.r,
                    run.Elements(W.rPr).Select(e => new XElement(e)),
                    new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), afterText)));
            }

            run.ReplaceWith(newElements);
        }

        private static void InsertBookmarkStart(
            XElement textElement,
            string bookmarkName,
            int bookmarkId,
            int offset)
        {
            var text = textElement.Value;
            var run = textElement.Parent;
            if (run?.Name != W.r) return;

            if (offset == 0)
            {
                // Insert bookmark start before the run
                run.AddBeforeSelf(new XElement(W.bookmarkStart,
                    new XAttribute(W.id, bookmarkId),
                    new XAttribute(W.name, bookmarkName)));
            }
            else
            {
                // Split the run
                var beforeText = text.Substring(0, offset);
                var afterText = text.Substring(offset);

                var beforeRun = new XElement(W.r,
                    run.Elements(W.rPr).Select(e => new XElement(e)),
                    new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), beforeText));

                var bookmarkStart = new XElement(W.bookmarkStart,
                    new XAttribute(W.id, bookmarkId),
                    new XAttribute(W.name, bookmarkName));

                var afterRun = new XElement(W.r,
                    run.Elements(W.rPr).Select(e => new XElement(e)),
                    new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), afterText));

                run.ReplaceWith(beforeRun, bookmarkStart, afterRun);
            }
        }

        private static void InsertBookmarkEnd(
            XElement textElement,
            string bookmarkName,
            int bookmarkId,
            int offset)
        {
            var text = textElement.Value;
            var run = textElement.Parent;
            if (run?.Name != W.r) return;

            if (offset == text.Length)
            {
                // Insert bookmark end after the run
                run.AddAfterSelf(new XElement(W.bookmarkEnd,
                    new XAttribute(W.id, bookmarkId)));
            }
            else
            {
                // Split the run
                var beforeText = text.Substring(0, offset);
                var afterText = text.Substring(offset);

                var beforeRun = new XElement(W.r,
                    run.Elements(W.rPr).Select(e => new XElement(e)),
                    new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), beforeText));

                var bookmarkEnd = new XElement(W.bookmarkEnd,
                    new XAttribute(W.id, bookmarkId));

                var afterRun = new XElement(W.r,
                    run.Elements(W.rPr).Select(e => new XElement(e)),
                    new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), afterText));

                run.ReplaceWith(beforeRun, bookmarkEnd, afterRun);
            }
        }

        private static void CreateBookmarkFromIndices(
            WordprocessingDocument wordDoc,
            string bookmarkName,
            AnnotationRange range)
        {
            var mainDoc = wordDoc.MainDocumentPart.GetXDocument();
            var body = mainDoc.Root?.Element(W.body);
            if (body == null)
                throw new InvalidOperationException("Document body not found.");

            var paragraphs = body.Elements(W.p).ToList();

            if (!range.StartParagraphIndex.HasValue || !range.EndParagraphIndex.HasValue)
                throw new ArgumentException("Paragraph indices are required.");

            int startParaIdx = range.StartParagraphIndex.Value;
            int endParaIdx = range.EndParagraphIndex.Value;

            if (startParaIdx < 0 || startParaIdx >= paragraphs.Count)
                throw new ArgumentOutOfRangeException(nameof(range), "Start paragraph index out of range.");
            if (endParaIdx < 0 || endParaIdx >= paragraphs.Count)
                throw new ArgumentOutOfRangeException(nameof(range), "End paragraph index out of range.");
            if (startParaIdx > endParaIdx)
                throw new ArgumentException("Start paragraph index must not be greater than end paragraph index.");

            var startPara = paragraphs[startParaIdx];
            var endPara = paragraphs[endParaIdx];

            // Generate unique bookmark ID
            var existingIds = mainDoc.Descendants(W.bookmarkStart)
                .Select(b => (int?)b.Attribute(W.id))
                .Where(id => id.HasValue)
                .Select(id => id.Value)
                .ToHashSet();

            int newId = 1;
            while (existingIds.Contains(newId)) newId++;

            // Determine where to insert bookmark start
            XElement startInsertPoint;
            if (range.StartRunIndex.HasValue)
            {
                var runs = startPara.Elements(W.r).ToList();
                if (range.StartRunIndex.Value < 0 || range.StartRunIndex.Value >= runs.Count)
                    throw new ArgumentOutOfRangeException(nameof(range), "Start run index out of range.");
                startInsertPoint = runs[range.StartRunIndex.Value];
            }
            else
            {
                startInsertPoint = startPara.Elements().FirstOrDefault(e => e.Name == W.r || e.Name == W.bookmarkStart);
            }

            // Insert bookmark start
            var bookmarkStart = new XElement(W.bookmarkStart,
                new XAttribute(W.id, newId),
                new XAttribute(W.name, bookmarkName));

            if (startInsertPoint != null)
                startInsertPoint.AddBeforeSelf(bookmarkStart);
            else
                startPara.AddFirst(bookmarkStart);

            // Determine where to insert bookmark end
            XElement endInsertPoint;
            if (range.EndRunIndex.HasValue)
            {
                var runs = endPara.Elements(W.r).ToList();
                if (range.EndRunIndex.Value < 0 || range.EndRunIndex.Value >= runs.Count)
                    throw new ArgumentOutOfRangeException(nameof(range), "End run index out of range.");
                endInsertPoint = runs[range.EndRunIndex.Value];
            }
            else
            {
                endInsertPoint = endPara.Elements(W.r).LastOrDefault();
            }

            // Insert bookmark end
            var bookmarkEnd = new XElement(W.bookmarkEnd,
                new XAttribute(W.id, newId));

            if (endInsertPoint != null)
                endInsertPoint.AddAfterSelf(bookmarkEnd);
            else
                endPara.Add(bookmarkEnd);

            wordDoc.MainDocumentPart.PutXDocument();
        }

        private static void RemoveBookmark(WordprocessingDocument wordDoc, string bookmarkName)
        {
            var mainDoc = wordDoc.MainDocumentPart.GetXDocument();

            var bookmarkStart = mainDoc.Descendants(W.bookmarkStart)
                .FirstOrDefault(b => (string)b.Attribute(W.name) == bookmarkName);

            if (bookmarkStart == null) return;

            var bookmarkId = (string)bookmarkStart.Attribute(W.id);

            var bookmarkEnd = mainDoc.Descendants(W.bookmarkEnd)
                .FirstOrDefault(b => (string)b.Attribute(W.id) == bookmarkId);

            bookmarkStart.Remove();
            bookmarkEnd?.Remove();

            wordDoc.MainDocumentPart.PutXDocument();
        }

        private static string GetTextInBookmark(WordprocessingDocument wordDoc, string bookmarkName)
        {
            if (string.IsNullOrEmpty(bookmarkName)) return null;

            var mainDoc = wordDoc.MainDocumentPart.GetXDocument();

            var bookmarkStart = mainDoc.Descendants(W.bookmarkStart)
                .FirstOrDefault(b => (string)b.Attribute(W.name) == bookmarkName);

            if (bookmarkStart == null) return null;

            var bookmarkId = (string)bookmarkStart.Attribute(W.id);

            // Collect all text between bookmark start and end
            var inBookmark = false;
            var textBuilder = new System.Text.StringBuilder();

            foreach (var element in mainDoc.Descendants())
            {
                if (element.Name == W.bookmarkStart && (string)element.Attribute(W.name) == bookmarkName)
                {
                    inBookmark = true;
                    continue;
                }

                if (element.Name == W.bookmarkEnd && (string)element.Attribute(W.id) == bookmarkId)
                {
                    break;
                }

                if (inBookmark && element.Name == W.t)
                {
                    textBuilder.Append(element.Value);
                }
            }

            return textBuilder.ToString();
        }

        #endregion
    }
}
