// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;

namespace Docxodus
{
    /// <summary>
    /// Represents a custom annotation on a document range.
    /// Annotations are stored in a Custom XML Part and linked to document content via bookmarks.
    /// </summary>
    public class DocumentAnnotation
    {
        /// <summary>
        /// Unique annotation identifier.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Label category/type identifier (e.g., "CLAUSE_TYPE_A", "DATE_REF").
        /// Used for categorizing and filtering annotations.
        /// </summary>
        public string LabelId { get; set; }

        /// <summary>
        /// Human-readable label text displayed in the UI.
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// Highlight color in hex format (e.g., "#FFEB3B").
        /// </summary>
        public string Color { get; set; }

        /// <summary>
        /// Author who created the annotation.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Creation timestamp.
        /// </summary>
        public DateTime? Created { get; set; }

        /// <summary>
        /// Internal bookmark name linking to document range.
        /// Format: _Docxodus_Ann_{Id}
        /// </summary>
        public string BookmarkName { get; set; }

        /// <summary>
        /// Cached start page number (may be stale if document changed).
        /// </summary>
        public int? StartPage { get; set; }

        /// <summary>
        /// Cached end page number (may be stale if document changed).
        /// </summary>
        public int? EndPage { get; set; }

        /// <summary>
        /// Whether cached page info needs recalculation.
        /// </summary>
        public bool PageInfoStale { get; set; } = true;

        /// <summary>
        /// Timestamp when page info was last computed.
        /// </summary>
        public DateTime? PageInfoComputedAt { get; set; }

        /// <summary>
        /// Extensible key-value metadata.
        /// </summary>
        public Dictionary<string, string> Metadata { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// The annotated text content (populated when reading from document).
        /// </summary>
        public string AnnotatedText { get; set; }

        /// <summary>
        /// Creates a new DocumentAnnotation with default values.
        /// </summary>
        public DocumentAnnotation()
        {
        }

        /// <summary>
        /// Creates a new DocumentAnnotation with required fields.
        /// </summary>
        public DocumentAnnotation(string id, string labelId, string label, string color)
        {
            Id = id ?? throw new ArgumentNullException(nameof(id));
            LabelId = labelId ?? throw new ArgumentNullException(nameof(labelId));
            Label = label ?? throw new ArgumentNullException(nameof(label));
            Color = color ?? throw new ArgumentNullException(nameof(color));
            BookmarkName = AnnotationManager.BookmarkPrefix + id;
            Created = DateTime.UtcNow;
        }
    }

    /// <summary>
    /// Specifies how to identify the text range for annotation.
    /// </summary>
    public class AnnotationRange
    {
        /// <summary>
        /// Search for text and annotate the Nth occurrence.
        /// </summary>
        public string SearchText { get; set; }

        /// <summary>
        /// Which occurrence to annotate (1-based). Default: 1
        /// </summary>
        public int Occurrence { get; set; } = 1;

        /// <summary>
        /// Use an existing bookmark by name instead of creating a new one.
        /// </summary>
        public string ExistingBookmarkName { get; set; }

        /// <summary>
        /// Start paragraph index (0-based).
        /// </summary>
        public int? StartParagraphIndex { get; set; }

        /// <summary>
        /// End paragraph index (0-based, inclusive).
        /// </summary>
        public int? EndParagraphIndex { get; set; }

        /// <summary>
        /// Start run index within start paragraph (0-based).
        /// If null, starts at beginning of paragraph.
        /// </summary>
        public int? StartRunIndex { get; set; }

        /// <summary>
        /// End run index within end paragraph (0-based, inclusive).
        /// If null, ends at end of paragraph.
        /// </summary>
        public int? EndRunIndex { get; set; }

        /// <summary>
        /// Creates a range specification for searching text.
        /// </summary>
        public static AnnotationRange FromSearch(string searchText, int occurrence = 1)
        {
            return new AnnotationRange
            {
                SearchText = searchText,
                Occurrence = occurrence
            };
        }

        /// <summary>
        /// Creates a range specification for an existing bookmark.
        /// </summary>
        public static AnnotationRange FromBookmark(string bookmarkName)
        {
            return new AnnotationRange
            {
                ExistingBookmarkName = bookmarkName
            };
        }

        /// <summary>
        /// Creates a range specification for paragraph indices.
        /// </summary>
        public static AnnotationRange FromParagraphs(int startIndex, int endIndex)
        {
            return new AnnotationRange
            {
                StartParagraphIndex = startIndex,
                EndParagraphIndex = endIndex
            };
        }

        /// <summary>
        /// Creates a range specification for specific runs within paragraphs.
        /// </summary>
        public static AnnotationRange FromRuns(
            int startParagraphIndex, int startRunIndex,
            int endParagraphIndex, int endRunIndex)
        {
            return new AnnotationRange
            {
                StartParagraphIndex = startParagraphIndex,
                StartRunIndex = startRunIndex,
                EndParagraphIndex = endParagraphIndex,
                EndRunIndex = endRunIndex
            };
        }
    }

    /// <summary>
    /// Specifies how annotation labels are displayed in HTML output.
    /// </summary>
    public enum AnnotationLabelMode
    {
        /// <summary>
        /// Floating label positioned above the highlight.
        /// </summary>
        Above = 0,

        /// <summary>
        /// Label displayed inline at start of highlight.
        /// </summary>
        Inline = 1,

        /// <summary>
        /// Label shown only on hover (tooltip).
        /// </summary>
        Tooltip = 2,

        /// <summary>
        /// No labels displayed, only highlights.
        /// </summary>
        None = 3
    }
}
