// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Docxodus;

/// <summary>
/// Specifies the target location for an annotation.
/// Supports multiple targeting modes: element ID, indices, or text search.
/// </summary>
public class AnnotationTarget
{
    /// <summary>
    /// Target element by its unique ID (e.g., "doc/tbl-0/tr-1/tc-2").
    /// </summary>
    public string? ElementId { get; set; }

    /// <summary>
    /// Type of element being targeted (used with index-based targeting).
    /// </summary>
    public DocumentElementType? ElementType { get; set; }

    /// <summary>
    /// Paragraph index (zero-based) within the document or parent element.
    /// </summary>
    public int? ParagraphIndex { get; set; }

    /// <summary>
    /// Run index (zero-based) within a paragraph.
    /// </summary>
    public int? RunIndex { get; set; }

    /// <summary>
    /// Table index (zero-based) within the document.
    /// </summary>
    public int? TableIndex { get; set; }

    /// <summary>
    /// Row index (zero-based) within a table.
    /// </summary>
    public int? RowIndex { get; set; }

    /// <summary>
    /// Cell index (zero-based) within a row.
    /// </summary>
    public int? CellIndex { get; set; }

    /// <summary>
    /// Column index (zero-based) for table column annotations.
    /// </summary>
    public int? ColumnIndex { get; set; }

    /// <summary>
    /// Text to search for within the target element.
    /// </summary>
    public string? SearchText { get; set; }

    /// <summary>
    /// Which occurrence of SearchText to target (1-based). Default is 1.
    /// </summary>
    public int Occurrence { get; set; } = 1;

    /// <summary>
    /// End of a range (for spanning multiple elements).
    /// </summary>
    public AnnotationTarget? RangeEnd { get; set; }

    #region Factory Methods

    /// <summary>
    /// Target a specific element by its ID.
    /// </summary>
    public static AnnotationTarget Element(string elementId) =>
        new() { ElementId = elementId };

    /// <summary>
    /// Target a paragraph by index.
    /// </summary>
    public static AnnotationTarget Paragraph(int index) =>
        new() { ElementType = DocumentElementType.Paragraph, ParagraphIndex = index };

    /// <summary>
    /// Target a range of paragraphs.
    /// </summary>
    public static AnnotationTarget ParagraphRange(int startIndex, int endIndex) =>
        new()
        {
            ElementType = DocumentElementType.Paragraph,
            ParagraphIndex = startIndex,
            RangeEnd = new AnnotationTarget { ParagraphIndex = endIndex }
        };

    /// <summary>
    /// Target a specific run within a paragraph.
    /// </summary>
    public static AnnotationTarget Run(int paragraphIndex, int runIndex) =>
        new()
        {
            ElementType = DocumentElementType.Run,
            ParagraphIndex = paragraphIndex,
            RunIndex = runIndex
        };

    /// <summary>
    /// Target a table by index.
    /// </summary>
    public static AnnotationTarget Table(int tableIndex) =>
        new() { ElementType = DocumentElementType.Table, TableIndex = tableIndex };

    /// <summary>
    /// Target a table row.
    /// </summary>
    public static AnnotationTarget TableRow(int tableIndex, int rowIndex) =>
        new()
        {
            ElementType = DocumentElementType.TableRow,
            TableIndex = tableIndex,
            RowIndex = rowIndex
        };

    /// <summary>
    /// Target a table cell.
    /// </summary>
    public static AnnotationTarget TableCell(int tableIndex, int rowIndex, int cellIndex) =>
        new()
        {
            ElementType = DocumentElementType.TableCell,
            TableIndex = tableIndex,
            RowIndex = rowIndex,
            CellIndex = cellIndex
        };

    /// <summary>
    /// Target a table column (all cells in that column).
    /// </summary>
    public static AnnotationTarget TableColumn(int tableIndex, int columnIndex) =>
        new()
        {
            ElementType = DocumentElementType.TableColumn,
            TableIndex = tableIndex,
            ColumnIndex = columnIndex
        };

    /// <summary>
    /// Target by text search.
    /// </summary>
    public static AnnotationTarget Search(string text, int occurrence = 1) =>
        new() { SearchText = text, Occurrence = occurrence };

    /// <summary>
    /// Target text within a specific element.
    /// </summary>
    public static AnnotationTarget SearchInElement(string elementId, string text, int occurrence = 1) =>
        new() { ElementId = elementId, SearchText = text, Occurrence = occurrence };

    #endregion

    /// <summary>
    /// Determines the effective targeting mode based on which properties are set.
    /// </summary>
    public AnnotationTargetMode GetTargetMode()
    {
        if (!string.IsNullOrEmpty(ElementId))
        {
            if (!string.IsNullOrEmpty(SearchText))
                return AnnotationTargetMode.SearchInElement;
            return AnnotationTargetMode.ElementId;
        }

        if (!string.IsNullOrEmpty(SearchText))
            return AnnotationTargetMode.TextSearch;

        if (ElementType == DocumentElementType.TableColumn)
            return AnnotationTargetMode.TableColumn;

        if (ElementType.HasValue)
            return AnnotationTargetMode.IndexBased;

        return AnnotationTargetMode.Unknown;
    }
}

/// <summary>
/// The mode used for targeting an annotation.
/// </summary>
public enum AnnotationTargetMode
{
    /// <summary>Target mode could not be determined</summary>
    Unknown,
    /// <summary>Target by element ID</summary>
    ElementId,
    /// <summary>Target by text search</summary>
    TextSearch,
    /// <summary>Target by type and indices</summary>
    IndexBased,
    /// <summary>Target a table column (special handling)</summary>
    TableColumn,
    /// <summary>Search for text within a specific element</summary>
    SearchInElement,
}
