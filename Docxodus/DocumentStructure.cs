// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

/// <summary>
/// Types of document elements that can be annotated.
/// </summary>
public enum DocumentElementType
{
    /// <summary>Root document element</summary>
    Document,
    /// <summary>A paragraph (w:p)</summary>
    Paragraph,
    /// <summary>A run within a paragraph (w:r)</summary>
    Run,
    /// <summary>A table (w:tbl)</summary>
    Table,
    /// <summary>A table row (w:tr)</summary>
    TableRow,
    /// <summary>A table cell (w:tc)</summary>
    TableCell,
    /// <summary>A virtual table column (not a real OOXML element)</summary>
    TableColumn,
    /// <summary>A hyperlink (w:hyperlink)</summary>
    Hyperlink,
    /// <summary>An image/drawing (w:drawing)</summary>
    Image,
}

/// <summary>
/// Represents a document element in the structure tree.
/// </summary>
public class DocumentElement
{
    /// <summary>
    /// Unique identifier for this element (path-based, e.g., "tbl-0/tr-1/tc-2").
    /// </summary>
    public string Id { get; init; } = "";

    /// <summary>
    /// Type of this element.
    /// </summary>
    public DocumentElementType Type { get; init; }

    /// <summary>
    /// Preview of text content (first ~100 characters).
    /// </summary>
    public string? TextPreview { get; init; }

    /// <summary>
    /// Position index within parent element.
    /// </summary>
    public int Index { get; init; }

    /// <summary>
    /// Child elements.
    /// </summary>
    public List<DocumentElement> Children { get; init; } = new();

    /// <summary>
    /// For table rows: the row index within the table.
    /// </summary>
    public int? RowIndex { get; init; }

    /// <summary>
    /// For table cells: the column index (accounting for grid span).
    /// </summary>
    public int? ColumnIndex { get; init; }

    /// <summary>
    /// For table cells: number of rows this cell spans.
    /// </summary>
    public int? RowSpan { get; init; }

    /// <summary>
    /// For table cells: number of columns this cell spans.
    /// </summary>
    public int? ColumnSpan { get; init; }

    /// <summary>
    /// Reference to the underlying XElement (not serialized).
    /// </summary>
    internal XElement? XmlElement { get; init; }
}

/// <summary>
/// Information about a virtual table column.
/// </summary>
public class TableColumnInfo
{
    /// <summary>
    /// ID of the table this column belongs to.
    /// </summary>
    public string TableId { get; init; } = "";

    /// <summary>
    /// Zero-based column index.
    /// </summary>
    public int ColumnIndex { get; init; }

    /// <summary>
    /// IDs of all cells in this column.
    /// </summary>
    public List<string> CellIds { get; init; } = new();

    /// <summary>
    /// Total number of rows in this column.
    /// </summary>
    public int RowCount => CellIds.Count;
}

/// <summary>
/// The complete document structure analysis result.
/// </summary>
public class DocumentStructure
{
    /// <summary>
    /// Root document element containing all children.
    /// </summary>
    public DocumentElement Root { get; init; } = new();

    /// <summary>
    /// Lookup dictionary for finding elements by ID.
    /// </summary>
    public Dictionary<string, DocumentElement> ElementsById { get; init; } = new();

    /// <summary>
    /// Information about table columns (keyed by "tableId/col-N").
    /// </summary>
    public Dictionary<string, TableColumnInfo> TableColumns { get; init; } = new();

    /// <summary>
    /// Find an element by its ID.
    /// </summary>
    public DocumentElement? FindById(string id)
    {
        return ElementsById.TryGetValue(id, out var element) ? element : null;
    }

    /// <summary>
    /// Find all elements of a specific type.
    /// </summary>
    public IEnumerable<DocumentElement> FindByType(DocumentElementType type)
    {
        return ElementsById.Values.Where(e => e.Type == type);
    }

    /// <summary>
    /// Search for elements containing specific text.
    /// </summary>
    public IEnumerable<DocumentElement> Search(string text)
    {
        if (string.IsNullOrEmpty(text)) return Enumerable.Empty<DocumentElement>();

        return ElementsById.Values.Where(e =>
            e.TextPreview != null &&
            e.TextPreview.Contains(text, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Get column information for a specific table.
    /// </summary>
    public IEnumerable<TableColumnInfo> GetTableColumns(string tableId)
    {
        return TableColumns.Values.Where(c => c.TableId == tableId).OrderBy(c => c.ColumnIndex);
    }
}

/// <summary>
/// Analyzes document structure to produce a navigable element tree.
/// </summary>
public static class DocumentStructureAnalyzer
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// Analyze a WmlDocument and return its structure.
    /// </summary>
    public static DocumentStructure Analyze(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wordDoc = WordprocessingDocument.Open(ms, false);

        var mainPart = wordDoc.MainDocumentPart;
        if (mainPart?.Document?.Body == null)
        {
            return new DocumentStructure
            {
                Root = new DocumentElement
                {
                    Id = "doc",
                    Type = DocumentElementType.Document,
                    TextPreview = "(empty document)"
                }
            };
        }

        var body = XElement.Parse(mainPart.Document.Body.OuterXml);
        var elementsById = new Dictionary<string, DocumentElement>();
        var tableColumns = new Dictionary<string, TableColumnInfo>();

        var rootChildren = AnalyzeChildren(body, "doc", elementsById, tableColumns);

        var root = new DocumentElement
        {
            Id = "doc",
            Type = DocumentElementType.Document,
            TextPreview = GetTextPreview(body),
            Index = 0,
            Children = rootChildren
        };

        elementsById["doc"] = root;

        return new DocumentStructure
        {
            Root = root,
            ElementsById = elementsById,
            TableColumns = tableColumns
        };
    }

    private static List<DocumentElement> AnalyzeChildren(
        XElement parent,
        string parentId,
        Dictionary<string, DocumentElement> elementsById,
        Dictionary<string, TableColumnInfo> tableColumns)
    {
        var children = new List<DocumentElement>();

        int paragraphIndex = 0;
        int tableIndex = 0;

        foreach (var child in parent.Elements())
        {
            if (child.Name == W + "p")
            {
                var element = AnalyzeParagraph(child, parentId, paragraphIndex, elementsById);
                children.Add(element);
                paragraphIndex++;
            }
            else if (child.Name == W + "tbl")
            {
                var element = AnalyzeTable(child, parentId, tableIndex, elementsById, tableColumns);
                children.Add(element);
                tableIndex++;
            }
        }

        return children;
    }

    private static DocumentElement AnalyzeParagraph(
        XElement para,
        string parentId,
        int index,
        Dictionary<string, DocumentElement> elementsById)
    {
        var id = $"{parentId}/p-{index}";
        var runs = new List<DocumentElement>();

        int runIndex = 0;
        int hyperlinkIndex = 0;

        foreach (var child in para.Elements())
        {
            if (child.Name == W + "r")
            {
                var runElement = AnalyzeRun(child, id, runIndex, elementsById);
                runs.Add(runElement);
                runIndex++;
            }
            else if (child.Name == W + "hyperlink")
            {
                var hyperlinkElement = AnalyzeHyperlink(child, id, hyperlinkIndex, runIndex, elementsById);
                runs.Add(hyperlinkElement);
                hyperlinkIndex++;
                // Count runs inside hyperlink for proper indexing
                runIndex += child.Elements(W + "r").Count();
            }
        }

        var element = new DocumentElement
        {
            Id = id,
            Type = DocumentElementType.Paragraph,
            TextPreview = GetTextPreview(para),
            Index = index,
            Children = runs,
            XmlElement = para
        };

        elementsById[id] = element;
        return element;
    }

    private static DocumentElement AnalyzeRun(
        XElement run,
        string parentId,
        int index,
        Dictionary<string, DocumentElement> elementsById)
    {
        var id = $"{parentId}/r-{index}";

        // Check if this run contains an image
        var drawing = run.Descendants(W + "drawing").FirstOrDefault();
        var hasImage = drawing != null;

        var element = new DocumentElement
        {
            Id = id,
            Type = hasImage ? DocumentElementType.Image : DocumentElementType.Run,
            TextPreview = hasImage ? "[Image]" : GetTextPreview(run),
            Index = index,
            XmlElement = run
        };

        elementsById[id] = element;
        return element;
    }

    private static DocumentElement AnalyzeHyperlink(
        XElement hyperlink,
        string parentId,
        int hyperlinkIndex,
        int runStartIndex,
        Dictionary<string, DocumentElement> elementsById)
    {
        var id = $"{parentId}/hl-{hyperlinkIndex}";

        var children = new List<DocumentElement>();
        int runIndex = runStartIndex;

        foreach (var run in hyperlink.Elements(W + "r"))
        {
            var runElement = AnalyzeRun(run, id, runIndex - runStartIndex, elementsById);
            children.Add(runElement);
            runIndex++;
        }

        var element = new DocumentElement
        {
            Id = id,
            Type = DocumentElementType.Hyperlink,
            TextPreview = GetTextPreview(hyperlink),
            Index = hyperlinkIndex,
            Children = children,
            XmlElement = hyperlink
        };

        elementsById[id] = element;
        return element;
    }

    private static DocumentElement AnalyzeTable(
        XElement table,
        string parentId,
        int index,
        Dictionary<string, DocumentElement> elementsById,
        Dictionary<string, TableColumnInfo> tableColumns)
    {
        var id = $"{parentId}/tbl-{index}";
        var rows = new List<DocumentElement>();

        // Track column info
        var columnCells = new Dictionary<int, List<string>>();

        int rowIndex = 0;
        foreach (var tr in table.Elements(W + "tr"))
        {
            var rowElement = AnalyzeTableRow(tr, id, rowIndex, elementsById, tableColumns, columnCells);
            rows.Add(rowElement);
            rowIndex++;
        }

        // Build TableColumnInfo entries
        foreach (var (colIdx, cellIds) in columnCells)
        {
            var colId = $"{id}/col-{colIdx}";
            tableColumns[colId] = new TableColumnInfo
            {
                TableId = id,
                ColumnIndex = colIdx,
                CellIds = cellIds
            };
        }

        var element = new DocumentElement
        {
            Id = id,
            Type = DocumentElementType.Table,
            TextPreview = $"[Table: {rowIndex} rows]",
            Index = index,
            Children = rows,
            XmlElement = table
        };

        elementsById[id] = element;
        return element;
    }

    private static DocumentElement AnalyzeTableRow(
        XElement tr,
        string tableId,
        int rowIndex,
        Dictionary<string, DocumentElement> elementsById,
        Dictionary<string, TableColumnInfo> tableColumns,
        Dictionary<int, List<string>> columnCells)
    {
        var id = $"{tableId}/tr-{rowIndex}";
        var cells = new List<DocumentElement>();

        int columnIndex = 0;
        int cellIndex = 0;

        foreach (var tc in tr.Elements(W + "tc"))
        {
            var cellElement = AnalyzeTableCell(tc, id, cellIndex, columnIndex, rowIndex, elementsById, tableColumns);
            cells.Add(cellElement);

            // Track cell for column info
            if (!columnCells.ContainsKey(columnIndex))
            {
                columnCells[columnIndex] = new List<string>();
            }
            columnCells[columnIndex].Add(cellElement.Id);

            // Account for column span
            var gridSpan = tc.Element(W + "tcPr")?.Element(W + "gridSpan")?.Attribute(W + "val")?.Value;
            var span = gridSpan != null ? int.Parse(gridSpan) : 1;
            columnIndex += span;
            cellIndex++;
        }

        var element = new DocumentElement
        {
            Id = id,
            Type = DocumentElementType.TableRow,
            TextPreview = $"[Row {rowIndex + 1}: {cellIndex} cells]",
            Index = cellIndex,
            RowIndex = rowIndex,
            Children = cells,
            XmlElement = tr
        };

        elementsById[id] = element;
        return element;
    }

    private static DocumentElement AnalyzeTableCell(
        XElement tc,
        string rowId,
        int cellIndex,
        int columnIndex,
        int rowIndex,
        Dictionary<string, DocumentElement> elementsById,
        Dictionary<string, TableColumnInfo> tableColumns)
    {
        var id = $"{rowId}/tc-{cellIndex}";

        // Get span info
        var tcPr = tc.Element(W + "tcPr");
        var gridSpanAttr = tcPr?.Element(W + "gridSpan")?.Attribute(W + "val")?.Value;
        var columnSpan = gridSpanAttr != null ? int.Parse(gridSpanAttr) : 1;

        var vMerge = tcPr?.Element(W + "vMerge");
        int? rowSpan = null;
        if (vMerge != null)
        {
            var val = vMerge.Attribute(W + "val")?.Value;
            rowSpan = val == "restart" ? 1 : 0; // 0 means continuation
        }

        // Analyze cell content (paragraphs and nested tables)
        var children = AnalyzeChildren(tc, id, elementsById, tableColumns);

        var element = new DocumentElement
        {
            Id = id,
            Type = DocumentElementType.TableCell,
            TextPreview = GetTextPreview(tc),
            Index = cellIndex,
            ColumnIndex = columnIndex,
            RowIndex = rowIndex,
            ColumnSpan = columnSpan > 1 ? columnSpan : null,
            RowSpan = rowSpan,
            Children = children,
            XmlElement = tc
        };

        elementsById[id] = element;
        return element;
    }

    private static string? GetTextPreview(XElement element, int maxLength = 100)
    {
        var sb = new StringBuilder();

        foreach (var text in element.Descendants(W + "t"))
        {
            sb.Append(text.Value);
            if (sb.Length >= maxLength)
            {
                break;
            }
        }

        var result = sb.ToString();
        if (string.IsNullOrWhiteSpace(result))
        {
            return null;
        }

        return result.Length > maxLength ? result.Substring(0, maxLength) + "..." : result;
    }
}
