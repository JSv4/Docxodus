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
    // ─── Tier D: table cell content ──────────────────────────────────────

    public EditResult ReplaceCellContent(string cellAnchorId, string markdownPayload)
    {
        if (ResolveCell(cellAnchorId, out _, out var cell, out _, out _, out var target) is { } resolveError)
            return resolveError;

        var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
        if (!parsed.Success)
            return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, cellAnchorId);
        if (ValidatePendingHyperlinks(parsed.Blocks.SelectMany(b => b.RunElements), cellAnchorId) is { } linkError)
            return linkError;

        // Replacing cell content removes every block in the cell, including nested tables and
        // structured wrappers; validate the whole cell subtree rather than only direct paragraphs.
        if (ValidateBookmarkRemoval(new[] { cell! }, cellAnchorId) is { } bookmarkError)
            return bookmarkError;
        var hyperlinkOwner = Internal.OwnedPartRelationships.FindOwner(_doc!, cell!);
        var oldHyperlinkIds = cell!.Descendants(W.hyperlink)
            .Select(h => (string?)h.Attribute(R.id)).Where(id => !string.IsNullOrEmpty(id)).Cast<string>().ToList();

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            foreach (var p in cell!.Elements(W.p).ToList()) p.Remove();

            var lists = new PayloadListState(parsed.Blocks);
            foreach (var block in parsed.Blocks)
            {
                var p = BuildParagraphFromParsedBlock(block);
                AssignPayloadListNumbering(p, block, lists);
                UnidHelper.AssignToSelfAndDescendants(p);
                cell.Add(p);
                PromoteHyperlinkRelationships(p);
            }
            if (hyperlinkOwner is { } owner)
            {
                foreach (var relationshipId in oldHyperlinkIds)
                    Internal.OwnedPartRelationships.DeleteReferenceRelationshipIfOrphaned(owner.Part, relationshipId, R.id);
                Internal.OwnedPartRelationships.SweepOrphanedImages(owner.Part);
            }
            // A table cell must contain at least one paragraph per OOXML schema.
            if (!cell.Elements(W.p).Any())
                cell.Add(new XElement(W.p));

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target!.Anchor },
                Patch = PatchFor(target!),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, cellAnchorId);
        }
    }

    // ─── Table editing (row / column CRUD), addressed by a canonical tc anchor ─────────────
    //
    // The grid model (issue #340): a row's cells tile w:tblGrid columns left→right, each
    // covering w:gridSpan columns (default 1) from an origin shifted by w:trPr/w:gridBefore.
    // A vertical merge is a column-aligned run of rows whose lead cell carries
    // w:vMerge w:val="restart" and whose followers carry a bare w:vMerge. Row/column CRUD is
    // grid-aware, not cell-index-aware: inserting across a span extends it, deleting through
    // one narrows it, and a merge a structural edit cannot preserve is rejected — never
    // silently torn.

    /// <summary>Resolve the canonical <c>tc</c> anchor to its cell/row/table. For the documented
    /// compatibility window, a legacy paragraph/heading/list-item anchor is translated to its
    /// nearest ancestor cell; every returned target is canonicalized to that cell. Structural
    /// anchors are never climbed through, which prevents a nested <c>tc</c>/<c>tr</c>/<c>tbl</c>
    /// from silently retargeting its enclosing outer cell.</summary>
    private EditResult? ResolveCell(string cellAnchorId, out XElement? p, out XElement? tc,
        out XElement? tr, out XElement? tbl, out AnchorTarget? target)
    {
        p = tc = tr = tbl = null; target = null;
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        target = FindAnchor(cellAnchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {cellAnchorId}", cellAnchorId);
        p = target.Resolve(_doc!);
        if (p is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", cellAnchorId);
        if (target.Anchor.Kind == "tc" && p.Name == W.tc)
            tc = p;
        else if (target.Anchor.Kind is "p" or "h" or "li")
            tc = p.Ancestors(W.tc).FirstOrDefault();
        if (tc is null)
            return EditResult.Fail(EditErrorCode.TableAnchorMigrationRequired,
                "table cell operations require a canonical tc anchor; legacy p/h/li anchors are translated only when physically inside the intended cell. Use GetTableMetadata or ResolveTableCellCoordinate to obtain the tc anchor.",
                cellAnchorId);
        tr = tc.Ancestors(W.tr).FirstOrDefault();
        tbl = tr?.Ancestors(W.tbl).FirstOrDefault();
        if (tr is null || tbl is null)
            return EditResult.Fail(EditErrorCode.InternalError, "malformed table (cell has no row/table)", cellAnchorId);
        var canonical = AnchorForElement(tc);
        if (canonical is null || FindAnchor(canonical.Value.Id) is not { } canonicalTarget)
            return EditResult.Fail(EditErrorCode.InternalError, "cell has no canonical tc anchor", cellAnchorId);
        target = canonicalTarget;
        return null;
    }

    /// <summary>The row's cells with their shared-model grid geometry, left→right.</summary>
    private static List<GridCell> RowGrid(XElement tr) => Internal.TableGridModel.RowGrid(tr);

    /// <summary>The cell covering <paramref name="gridCol"/>, or null when the row has none.</summary>
    private static GridCell? CellCovering(IEnumerable<GridCell> grid, int gridCol) =>
        Internal.TableGridModel.CellCovering(grid, gridCol);

    /// <summary>The cell of <paramref name="tr"/> occupying exactly <paramref name="shape"/>'s grid
    /// columns — how a vertical-merge run is followed from row to row.</summary>
    private static XElement? AlignedCell(XElement tr, GridCell shape) =>
        Internal.TableGridModel.AlignedCell(tr, shape);

    /// <summary>The cell's vertical-merge role: null = none, true = <c>w:vMerge w:val="restart"</c>
    /// (a merge's lead cell), false = a continuation (bare <c>w:vMerge</c>, or val="continue").</summary>
    private static bool? VMergeRestart(XElement tc) =>
        Internal.TableGridModel.VerticalMergeRole(tc) switch
        {
            TableVerticalMergeRole.Restart => true,
            TableVerticalMergeRole.Continue => false,
            _ => null,
        };

    private static void SetVMerge(XElement tc, bool? restart)
    {
        if (restart is null) { tc.Element(W.tcPr)?.Elements(W.vMerge).Remove(); return; }
        SetChildInOrder(GetOrCreateTcPr(tc),
            restart.Value ? new XElement(W.vMerge, new XAttribute(W.val, "restart")) : new XElement(W.vMerge),
            TcPrChildOrder);
    }

    /// <summary>Write (or, for a span of 1, remove) the cell's <c>w:gridSpan</c>.</summary>
    private static void SetGridSpan(XElement tc, int span)
    {
        if (span <= 1) { tc.Element(W.tcPr)?.Elements(W.gridSpan).Remove(); return; }
        SetChildInOrder(GetOrCreateTcPr(tc),
            new XElement(W.gridSpan, new XAttribute(W.val, span)), TcPrChildOrder);
    }

    private static void SetCellWidth(XElement tc, int twips) =>
        SetChildInOrder(GetOrCreateTcPr(tc),
            new XElement(W.tcW, new XAttribute(W._w, twips), new XAttribute(W.type, "dxa")),
            TcPrChildOrder);

    /// <summary>Grow/shrink an existing dxa cell width by <paramref name="delta"/> twips; a cell
    /// sized any other way (pct/auto) is left alone.</summary>
    private static void BumpCellWidth(XElement tc, int delta)
    {
        var tcW = tc.Element(W.tcPr)?.Element(W.tcW);
        if (tcW is null || (string?)tcW.Attribute(W.type) != "dxa") return;
        tcW.SetAttributeValue(W._w, Math.Max(0, ((int?)tcW.Attribute(W._w) ?? 0) + delta));
    }

    private static void SetRowGridOmission(XElement row, XName name, int value)
    {
        var property = row.Element(W.trPr)?.Element(name);
        if (value <= 0)
        {
            property?.Remove();
            return;
        }
        if (property is null)
            throw new InvalidOperationException($"row has no existing {name.LocalName} omission to adjust");
        property.SetAttributeValue(W.val, value);
    }

    private static List<int> GridColWidths(XElement tbl) =>
        tbl.Element(W.tblGrid)?.Elements(W.gridCol).Select(g => (int?)g.Attribute(W._w) ?? 0).ToList()
        ?? new List<int>();

    /// <summary>Materialize real gridCol elements only inside a structural transaction. Metadata
    /// inspection remains read-only and reports virtual columns; the before/after mapping then
    /// invalidates those virtual identities and adds these real anchors explicitly.</summary>
    private static void EnsureGridColumnsForMutation(XElement table)
    {
        int count = GridColumnCount(table);
        var grid = table.Element(W.tblGrid);
        if (grid is null)
        {
            grid = new XElement(W.tblGrid);
            var properties = table.Element(W.tblPr);
            if (properties is null) table.AddFirst(grid);
            else properties.AddAfterSelf(grid);
        }
        int missing = count - grid.Elements(W.gridCol).Count();
        for (int index = 0; index < missing; index++)
        {
            var column = new XElement(W.gridCol);
            UnidHelper.AssignToSelfAndDescendants(column);
            grid.Add(column);
        }
    }

    private static int SumGridWidths(List<int> widths, int from, int toExclusive)
    {
        int sum = 0;
        for (int i = from; i < toExclusive && i < widths.Count; i++) sum += widths[i];
        return sum;
    }

    /// <summary>The table's grid width — w:tblGrid's column count, falling back to the widest row.</summary>
    private static int GridColumnCount(XElement tbl) => Internal.TableGridModel.GridColumnCount(tbl);

    /// <summary>After a structural edit, resolve freshly-projected anchors for live elements.</summary>
    private List<Anchor> ResolveAnchorsForElements(IEnumerable<XElement> elements)
    {
        _ = AnchorIndex();
        var result = new List<Anchor>();
        foreach (var element in elements)
        {
            var unid = (string?)element.Attribute(PtOpenXml.Unid);
            if (unid is not null && AnchorForUnid(unid, PartUriOf(element)) is { } a)
                result.Add(a);
        }
        return result;
    }

    private TableMetadata CaptureTableMetadata(XElement table) =>
        Internal.TableGridModel.BuildMetadata(table, AnchorForElement);

    private TableAnchorMapping CompleteTableMapping(TableMetadata before, XElement table) =>
        Internal.TableGridModel.Map(before,
            table.Parent is null ? null : Internal.TableGridModel.BuildMetadata(table, AnchorForElement));

    private static IReadOnlyList<Anchor> InvalidatedCellAnchors(TableAnchorMapping mapping) =>
        mapping.Invalidated
            .Where(location => location.EntityKind == TableAnchorEntityKind.Cell)
            .Select(location => location.Anchor)
            .ToList();

    private EditResult? RefuseUnresolvedTableStructure(XElement table, string anchorId)
    {
        var pending = BuildRevisionRegistry().Entries.FirstOrDefault(group =>
            group.Units.Any(unit => ReferenceEquals(unit.Table, table))
            && (group.Family == RevisionFamily.CellInsert
                || group.Family == RevisionFamily.CellDelete
                || group.Family == RevisionFamily.CellMerge));
        return pending is null ? null : EditResult.Fail(
            EditErrorCode.UnresolvedStructuralRevision,
            $"table has unresolved {pending.Family} revision {pending.Id}; resolve it before another structural mutation",
            anchorId);
    }

    private static EditResult TrackedStructureUnsupported(string operation, string anchorId) =>
        EditResult.Fail(EditErrorCode.TrackedOperationUnsupported,
            $"{operation} has no reversible native tracked-change encoding on this document shape; no changes were made",
            anchorId);

    private EditResult? RefuseNestedTrackedPropertyChange(
        IEnumerable<(XElement? Properties, XName ChangeName)> properties, string anchorId)
    {
        if (_trackedChanges != TrackedChangeMode.RenderInline) return null;
        var pending = properties.FirstOrDefault(pair => pair.Properties?.Element(pair.ChangeName) is not null);
        return pending.Properties is null ? null : EditResult.Fail(
            EditErrorCode.UnresolvedStructuralRevision,
            $"the target already has an unresolved {pending.ChangeName.LocalName}; resolve it before another tracked property mutation",
            anchorId);
    }

    private static XElement PropertySnapshot(XElement? properties, XName propertyName, XName changeName,
        params XName[] excluded)
    {
        var exclude = excluded.Append(changeName).ToHashSet();
        return new XElement(propertyName,
            properties?.Attributes().Where(a => !a.IsNamespaceDeclaration && a.Name != PtOpenXml.Unid),
            properties?.Elements().Where(e => !exclude.Contains(e.Name)).Select(e => new XElement(e)));
    }

    private static bool PropertySnapshotEquals(XElement snapshot, XElement? current, XName changeName,
        params XName[] excluded) =>
        XNode.DeepEquals(snapshot, PropertySnapshot(current, snapshot.Name, changeName, excluded));

    /// <summary>Append a native *PrChange only when the live base properties differ from the
    /// captured old value. Callers that change several properties as one operation pass the same
    /// author/date stamp so the registry exposes one atomic table-format revision.</summary>
    private void TrackPropertyMutation(XElement current, XElement oldProperties, XName changeName,
        string author, string date, params XName[] excluded)
    {
        var oldBase = PropertySnapshot(oldProperties, current.Name, changeName, excluded);
        if (PropertySnapshotEquals(oldBase, current, changeName, excluded)) return;
        current.Add(CreateRevisionEnvelope(changeName, author, date, oldBase));
    }

    private void MarkRowAsTrackedRevision(XElement row, bool inserted, string author, string date,
        bool markContent = true)
    {
        var wrapperName = inserted ? W.ins : W.del;
        var trPr = row.Element(W.trPr);
        if (trPr is null)
        {
            trPr = new XElement(W.trPr);
            if (row.Element(W.tblPrEx) is { } exceptions) exceptions.AddAfterSelf(trPr);
            else row.AddFirst(trPr);
        }
        // Schema position, not append: CT_TrPr orders base properties → ins → del →
        // trPrChange, so a row that already carries a trPrChange from a tracked
        // SetTableRowOptions would otherwise get an out-of-order w:del after it.
        if (trPr.Element(wrapperName) is null)
            SetChildInOrder(trPr, CreateRevisionEnvelope(wrapperName, author, date), TrPrChildOrder);
        if (markContent)
            foreach (var paragraph in row.Descendants(W.p).ToList())
                MarkParagraphContentAndMark(paragraph, wrapperName, author, date);
    }

    /// <summary>An empty clone of <paramref name="referenceCell"/>'s shell (width, borders, shading,
    /// valign). Merge markup is always dropped — a clone is a fresh cell, never half of somebody
    /// else's merge — except <c>w:gridSpan</c> when <paramref name="keepSpan"/> is set, which a new
    /// row needs so its cells still line up with w:tblGrid.</summary>
    private static XElement NewEmptyCellLike(XElement referenceCell, bool keepSpan = false)
    {
        var tcPr = referenceCell.Element(W.tcPr);
        var tc = new XElement(W.tc);
        if (tcPr is not null)
        {
            var clone = new XElement(tcPr);
            clone.Elements(W.vMerge).Remove();
            clone.Elements(W.hMerge).Remove();
            if (!keepSpan) clone.Elements(W.gridSpan).Remove();
            tc.Add(clone);
        }
        tc.Add(new XElement(W.p));
        return tc;
    }

    /// <summary>Insert a row before/after the row containing <paramref name="cellAnchorId"/>. The new
    /// row mirrors the reference row's grid shape (cell widths and <c>w:gridSpan</c>s) and starts
    /// empty; where a vertical merge crosses the insertion boundary the new row joins it as a
    /// continuation rather than punching a hole through it. Returns the new canonical cell anchors.</summary>
    public EditResult InsertTableRow(string cellAnchorId, Position pos)
    {
        if (ResolveCell(cellAnchorId, out _, out _, out var tr, out var tbl, out var target) is { } err)
            return err;
        if (RefuseUnresolvedTableStructure(tbl!, cellAnchorId) is { } pending) return pending;

        var before = CaptureTableMetadata(tbl!);
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // A vertical merge crosses the insertion boundary exactly when the row on the far
            // side of it carries a continuation at that grid position.
            var acrossRow = pos == Position.Before ? tr : tr!.ElementsAfterSelf(W.tr).FirstOrDefault();
            var acrossGrid = acrossRow is null ? null : RowGrid(acrossRow);

            var newTr = new XElement(W.tr);
            // Grid-shape row properties (columns skipped at either end) must come along, or the
            // new row's cells would not line up with w:tblGrid.
            var shape = tr!.Element(W.trPr)?.Elements()
                .Where(e => e.Name == W.gridBefore || e.Name == W.gridAfter
                         || e.Name == W.wBefore || e.Name == W.wAfter)
                .Select(e => new XElement(e)).ToList();
            if (shape is { Count: > 0 }) newTr.Add(new XElement(W.trPr, shape));

            var newCells = new List<XElement>();
            foreach (var g in RowGrid(tr))
            {
                var newTc = NewEmptyCellLike(g.Tc, keepSpan: true);
                if (acrossGrid is not null && CellCovering(acrossGrid, g.Start) is { } across
                    && VMergeRestart(across.Tc) == false)
                    SetVMerge(newTc, restart: false);
                newCells.Add(newTc);
                newTr.Add(newTc);
            }
            UnidHelper.AssignToSelfAndDescendants(newTr);
            if (pos == Position.Before) tr.AddBeforeSelf(newTr);
            else tr.AddAfterSelf(newTr);

            if (_trackedChanges == TrackedChangeMode.RenderInline)
            {
                MarkRowAsTrackedRevision(newTr, inserted: true,
                    _revisionAuthor ?? "docxodus", NextTrackedFormatRevisionDate());
            }

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Created = ResolveAnchorsForElements(newCells),
                TableAnchors = CompleteTableMapping(before, tbl!),
                Patch = PatchFor(target!),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, cellAnchorId);
        }
    }

    /// <summary>Insert a grid column before/after the one holding <paramref name="cellAnchorId"/>:
    /// a new empty cell in every row (cloning that column's width) plus a matching w:gridCol. A row
    /// whose cell straddles the new boundary — a horizontal merge spanning it — widens by one
    /// column instead of gaining a cell, so the grid stays consistent. Returns the new
    /// canonical cell anchors (top→bottom); rows that only widened contribute none.</summary>
    public EditResult InsertTableColumn(string cellAnchorId, Position pos)
    {
        if (ResolveCell(cellAnchorId, out _, out var tc, out var tr, out var tbl, out var target) is { } err)
            return err;
        if (RefuseUnresolvedTableStructure(tbl!, cellAnchorId) is { } pending) return pending;

        var anchorCell = RowGrid(tr!).First(g => g.Tc == tc);
        int boundary = pos == Position.Before ? anchorCell.Start : anchorCell.End;

        var before = CaptureTableMetadata(tbl!);
        var tracked = _trackedChanges == TrackedChangeMode.RenderInline;
        var oldGrid = tracked ? new XElement(tbl!.Element(W.tblGrid) ?? new XElement(W.tblGrid)) : null;
        var oldRows = tracked ? tbl!.Elements(W.tr).ToDictionary(row => row,
            row => PropertySnapshot(row.Element(W.trPr), W.trPr, W.trPrChange, W.ins, W.del)) : null;
        var oldCells = tracked ? tbl!.Descendants(W.tc).ToDictionary(cell => cell,
            cell => PropertySnapshot(cell.Element(W.tcPr), W.tcPr, W.tcPrChange,
                W.cellIns, W.cellDel, W.cellMerge)) : null;
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            EnsureGridColumnsForMutation(tbl!);
            // Mirror the structural change in w:tblGrid first, so the new column's width is
            // known before the cells that must carry it are written.
            var widths = GridColWidths(tbl!);
            int srcCol = Math.Clamp(pos == Position.Before ? anchorCell.Start : anchorCell.End - 1,
                0, Math.Max(0, widths.Count - 1));
            int newWidth = widths.Count > 0 ? widths[srcCol] : 0;
            if (tbl!.Element(W.tblGrid) is { } grid && grid.Elements(W.gridCol).ToList() is { Count: > 0 } cols)
            {
                var clone = new XElement(W.gridCol,
                    cols[srcCol].Attributes().Where(attribute => attribute.Name != PtOpenXml.Unid));
                UnidHelper.AssignToSelfAndDescendants(clone);
                if (boundary >= cols.Count) cols[^1].AddAfterSelf(clone);
                else cols[boundary].AddBeforeSelf(clone);
            }

            var newCells = new List<XElement>();
            foreach (var row in tbl.Elements(W.tr))
            {
                var rowGrid = RowGrid(row);
                int gridBefore = Internal.TableGridModel.GridBefore(row);
                int rowEnd = rowGrid.Count == 0 ? gridBefore : rowGrid[^1].End;
                int gridAfter = Internal.TableGridModel.GridAfter(row);
                if (boundary < gridBefore)
                {
                    SetRowGridOmission(row, W.gridBefore, gridBefore + 1);
                    continue;
                }
                if (boundary > rowEnd && boundary <= rowEnd + gridAfter)
                {
                    SetRowGridOmission(row, W.gridAfter, gridAfter + 1);
                    continue;
                }
                // A cell straddling the boundary extends rather than splits: inserting "inside"
                // a horizontal merge widens it.
                if (rowGrid.FirstOrDefault(c => c.Start < boundary && c.End > boundary) is { Tc: not null } straddle)
                {
                    SetGridSpan(straddle.Tc, straddle.Span + 1);
                    BumpCellWidth(straddle.Tc, newWidth);
                    continue;
                }
                var left = rowGrid.LastOrDefault(c => c.End <= boundary);
                var right = rowGrid.FirstOrDefault(c => c.Start >= boundary);
                var refTc = left.Tc ?? right.Tc;
                if (refTc is null) continue; // an empty row has nothing to clone a shell from
                var newTc = NewEmptyCellLike(refTc);
                if (newWidth > 0) SetCellWidth(newTc, newWidth);
                UnidHelper.AssignToSelfAndDescendants(newTc);
                newCells.Add(newTc);
                if (left.Tc is not null) left.Tc.AddAfterSelf(newTc);
                else right.Tc!.AddBeforeSelf(newTc);
            }

            if (tracked)
            {
                var author = _revisionAuthor ?? "docxodus";
                var date = NextTrackedFormatRevisionDate();
                var trackedGrid = tbl.Element(W.tblGrid)
                    ?? throw new InvalidOperationException("tracked column insertion has no table grid");
                trackedGrid.Add(CreateRevisionEnvelope(W.tblGridChange, author, date, oldGrid!));

                foreach (var pair in oldRows!)
                {
                    if (PropertySnapshotEquals(pair.Value, pair.Key.Element(W.trPr), W.trPrChange,
                        W.ins, W.del)) continue;
                    var trPr = pair.Key.Element(W.trPr) ?? new XElement(W.trPr);
                    if (trPr.Parent is null) pair.Key.AddFirst(trPr);
                    trPr.Add(CreateRevisionEnvelope(W.trPrChange, author, date, pair.Value));
                }
                foreach (var pair in oldCells!)
                {
                    if (PropertySnapshotEquals(pair.Value, pair.Key.Element(W.tcPr), W.tcPrChange,
                        W.cellIns, W.cellDel, W.cellMerge)) continue;
                    var tcPr = GetOrCreateTcPr(pair.Key);
                    tcPr.Add(CreateRevisionEnvelope(W.tcPrChange, author, date, pair.Value));
                }
                foreach (var cell in newCells)
                {
                    GetOrCreateTcPr(cell).Add(CreateRevisionEnvelope(W.cellIns, author, date));
                    foreach (var paragraph in cell.Elements(W.p))
                        MarkParagraphContentAndMark(paragraph, W.ins, author, date);
                }
            }

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Created = ResolveAnchorsForElements(newCells),
                TableAnchors = CompleteTableMapping(before, tbl),
                Patch = PatchFor(target!),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, cellAnchorId);
        }
    }

    /// <summary>Delete the row containing <paramref name="cellAnchorId"/>. A vertical merge whose
    /// lead row this is survives: the next row's continuation is promoted to the merge's new
    /// restart, so the run is never left headless. Deleting the last row removes the whole table.</summary>
    public EditResult DeleteTableRow(string cellAnchorId)
    {
        if (ResolveCell(cellAnchorId, out _, out _, out var tr, out var tbl, out var target) is { } err)
            return err;
        var hyperlinkOwner = Internal.OwnedPartRelationships.FindOwner(_doc!, tbl!);
        var removalRoot = WordprocessingMLUtil.TableRows(tbl!).Count() <= 1 ? tbl! : tr!;
        // Tracked mode marks the row deleted instead of removing it, so nothing that lives
        // inside it is actually going away — validating a removal that does not happen is a
        // false refusal.
        if (_trackedChanges != TrackedChangeMode.RenderInline
            && ValidateBookmarkRemoval(new[] { removalRoot }, cellAnchorId) is { } bookmarkError)
            return bookmarkError;
        if (RefuseUnresolvedTableStructure(tbl!, cellAnchorId) is { } pending) return pending;
        if (RefuseNestedTrackedPropertyChange(
                new[] { (tr!.Element(W.trPr), W.trPrChange) }, cellAnchorId) is { } pendingRowFormat)
            return pendingRowFormat;
        if (_trackedChanges == TrackedChangeMode.RenderInline
            && tr!.ElementsAfterSelf(W.tr).FirstOrDefault() is { } trackedNext
            && RowGrid(tr).Any(g => VMergeRestart(g.Tc) == true
                && AlignedCell(trackedNext, g) is { } heir && VMergeRestart(heir) == false))
            return TrackedStructureUnsupported(
                "DeleteTableRow across a vertical-merge restart", cellAnchorId);

        var before = CaptureTableMetadata(tbl!);
        var referencedNotesBefore = ReferencedNoteIds();
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (_trackedChanges == TrackedChangeMode.RenderInline)
            {
                MarkRowAsTrackedRevision(tr!, inserted: false,
                    _revisionAuthor ?? "docxodus", NextTrackedFormatRevisionDate());
            }
            else if (WordprocessingMLUtil.TableRows(tbl!).Count() <= 1) tbl!.Remove();
            else
            {
                if (tr!.ElementsAfterSelf(W.tr).FirstOrDefault() is { } next)
                    foreach (var g in RowGrid(tr).Where(g => VMergeRestart(g.Tc) == true))
                        if (AlignedCell(next, g) is { } heir && VMergeRestart(heir) == false)
                            SetVMerge(heir, restart: true);
                tr.Remove();
            }

            var prunedNoteAnchors = new List<Anchor>();
            AppendPrunedNoteAnchors(
                PruneOrphanedNotes(referencedNotesBefore),
                prunedNoteAnchors,
                new HashSet<string>(StringComparer.Ordinal));
            if (hyperlinkOwner is { } owner)
                SweepOrphanedStoryRelationships(owner.Part);

            InvalidateProjectionCache();
            var mapping = CompleteTableMapping(before, tbl!);
            return new EditResult
            {
                Success = true,
                Removed = InvalidatedCellAnchors(mapping).Concat(prunedNoteAnchors).ToList(),
                TableAnchors = mapping,
                Patch = PatchFor(target!),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, cellAnchorId);
        }
    }

    /// <summary>Delete the grid column holding <paramref name="cellAnchorId"/> from every row (and
    /// its w:gridCol). A cell spanning the doomed column narrows by one instead of disappearing,
    /// keeping its remaining columns and the rest of the grid intact. Deleting the last column
    /// removes the whole table.</summary>
    public EditResult DeleteTableColumn(string cellAnchorId)
    {
        if (ResolveCell(cellAnchorId, out _, out var tc, out var tr, out var tbl, out var target) is { } err)
            return err;
        var hyperlinkOwner = Internal.OwnedPartRelationships.FindOwner(_doc!, tbl!);
        if (RefuseUnresolvedTableStructure(tbl!, cellAnchorId) is { } pending) return pending;
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported("DeleteTableColumn", cellAnchorId);

        int doomed = RowGrid(tr!).First(g => g.Tc == tc).Start;
        int existingColumns = GridColumnCount(tbl!);
        var removalRoots = existingColumns <= 1
            ? new List<XElement> { tbl! }
            : WordprocessingMLUtil.TableRows(tbl!).Select(row => CellCovering(RowGrid(row), doomed))
                .Where(cell => cell.HasValue && cell.Value.Span == 1)
                .Select(cell => cell!.Value.Tc).ToList();
        if (ValidateBookmarkRemoval(removalRoots, cellAnchorId) is { } bookmarkError)
            return bookmarkError;

        var before = CaptureTableMetadata(tbl!);
        var referencedNotesBefore = ReferencedNoteIds();
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            EnsureGridColumnsForMutation(tbl!);
            var grid = tbl!.Element(W.tblGrid);
            int colCount = existingColumns;
            int lostWidth = GridColWidths(tbl) is { } widths && doomed < widths.Count ? widths[doomed] : 0;

            if (colCount <= 1) tbl.Remove();
            else
            {
                foreach (var row in WordprocessingMLUtil.TableRows(tbl).ToList())
                {
                    var rowGrid = RowGrid(row);
                    int gridBefore = Internal.TableGridModel.GridBefore(row);
                    int rowEnd = rowGrid.Count == 0 ? gridBefore : rowGrid[^1].End;
                    int gridAfter = Internal.TableGridModel.GridAfter(row);
                    if (doomed < gridBefore)
                    {
                        SetRowGridOmission(row, W.gridBefore, gridBefore - 1);
                        continue;
                    }
                    if (doomed >= rowEnd && doomed < rowEnd + gridAfter)
                    {
                        SetRowGridOmission(row, W.gridAfter, gridAfter - 1);
                        continue;
                    }
                    if (CellCovering(rowGrid, doomed) is not { } cell) continue;
                    if (cell.Span > 1)
                    {
                        SetGridSpan(cell.Tc, cell.Span - 1);
                        BumpCellWidth(cell.Tc, -lostWidth);
                        continue;
                    }
                    cell.Tc.Remove();
                }
                var cols = grid?.Elements(W.gridCol).ToList();
                if (cols is not null && doomed < cols.Count) cols[doomed].Remove();
            }

            var prunedNoteAnchors = new List<Anchor>();
            AppendPrunedNoteAnchors(
                PruneOrphanedNotes(referencedNotesBefore),
                prunedNoteAnchors,
                new HashSet<string>(StringComparer.Ordinal));
            if (hyperlinkOwner is { } owner)
                SweepOrphanedStoryRelationships(owner.Part);

            InvalidateProjectionCache();
            var mapping = CompleteTableMapping(before, tbl);
            return new EditResult
            {
                Success = true,
                Removed = InvalidatedCellAnchors(mapping).Concat(prunedNoteAnchors).ToList(),
                TableAnchors = mapping,
                Patch = PatchFor(target!),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, cellAnchorId);
        }
    }

    // ─── Cell merge / unmerge (issue #340 Stage B), addressed by a canonical tc anchor ─────
    //
    // Anchor semantics: the surviving tc retains its canonical anchor; absorbed tc anchors are
    // invalidated and reported in Removed/TableAnchors.Invalidated. Append moves absorbed content
    // into the survivor, preserving those paragraph anchors. A vertical-merge continuation cell
    // remains a canonical tc and keeps exactly one empty w:p because CT_Tc requires a block child.
    // It stays addressable even though Word renders nothing for it; unmerge before writing visible
    // content. The markdown projection still treats any table carrying a merge as opaque.

    private static EditResult MergeFail(string message, string anchorId) =>
        EditResult.Fail(EditErrorCode.InvalidTableMerge, message, anchorId);

    /// <summary>A block with no text, image or line break — what an absorbed/continuation cell may
    /// be reduced to without losing anything.</summary>
    private static bool IsEmptyBlock(XElement e) =>
        e.Name == W.p
        && !e.Descendants(W.t).Any(t => ((string)t).Length > 0)
        && !e.Descendants().Any(d => d.Name == W.drawing || d.Name == W.pict || d.Name == W.br);

    private static List<XElement> CellBlocks(XElement tc) =>
        tc.Elements().Where(e => e.Name != W.tcPr).ToList();

    /// <summary>Reduce a cell to a single empty paragraph — what Word writes in a vertical-merge
    /// continuation, which renders nothing. Returns the fresh paragraph, or null when the cell
    /// already held exactly one empty paragraph (whose anchor is then left intact).</summary>
    private static XElement? EmptyCellBody(XElement tc)
    {
        var blocks = CellBlocks(tc);
        if (blocks.Count == 1 && IsEmptyBlock(blocks[0])) return null;
        foreach (var b in blocks) b.Remove();
        var p = new XElement(W.p);
        UnidHelper.AssignToSelfAndDescendants(p);
        tc.Add(p);
        return p;
    }

    /// <summary>
    /// Merge the rectangle of cells anchored at <paramref name="cellAnchorId"/> and running
    /// <paramref name="rowSpan"/> rows down × <paramref name="colSpan"/> cells right (Word's
    /// *Merge Cells*). The horizontal extent becomes <c>w:gridSpan</c> on each row's surviving
    /// cell; a vertical extent becomes <c>w:vMerge w:val="restart"</c> on the lead cell and a bare
    /// <c>w:vMerge</c> on the rows beneath, whose bodies are emptied the way Word writes them.
    /// <para>
    /// The rectangle must tile the same whole grid columns in every row it covers and must not
    /// clip a vertical merge entering from above or continuing below; a partial overlap is
    /// rejected (<see cref="EditErrorCode.InvalidTableMerge"/>) rather than silently tearing the
    /// grid. Absorbed cells' content is appended to the surviving cell by default — see
    /// <see cref="TableMergeOptions.Content"/>.
    /// </para>
    /// </summary>
    public EditResult MergeCells(string cellAnchorId, int rowSpan, int colSpan,
        TableMergeOptions? options = null)
    {
        if (ResolveCell(cellAnchorId, out _, out var tc, out var tr, out var tbl, out var target) is { } err)
            return err;
        if (RefuseUnresolvedTableStructure(tbl!, cellAnchorId) is { } pending) return pending;
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported("MergeCells", cellAnchorId);

        var opts = options ?? new TableMergeOptions();
        if (rowSpan < 1 || colSpan < 1 || (long)rowSpan * colSpan < 2)
            return MergeFail($"a merge must cover at least two cells; got {rowSpan}×{colSpan}", cellAnchorId);

        var rows = tbl!.Elements(W.tr).ToList();
        int r0 = rows.IndexOf(tr!);
        if (r0 + rowSpan > rows.Count)
            return MergeFail(
                $"rowSpan {rowSpan} runs past the table's last row (only {rows.Count - r0} rows at and below the anchor)",
                cellAnchorId);

        var leadRow = RowGrid(tr!);
        int lead = leadRow.FindIndex(g => g.Tc == tc);
        if (lead < 0 || lead + colSpan > leadRow.Count)
            return MergeFail(
                $"colSpan {colSpan} runs past the row's last cell (only {leadRow.Count - lead} cells at and right of the anchor)",
                cellAnchorId);
        int c0 = leadRow[lead].Start, c1 = leadRow[lead + colSpan - 1].End;

        // Every covered row must tile exactly the same grid columns, or an existing span
        // straddles the rectangle's edge and merging would leave the grid ragged.
        var rect = new List<List<GridCell>>();
        for (int r = r0; r < r0 + rowSpan; r++)
        {
            var band = RowGrid(rows[r]).Where(g => g.End > c0 && g.Start < c1).ToList();
            if (band.Count == 0 || band[0].Start != c0 || band[^1].End != c1)
                return MergeFail(
                    $"row {r + 1}'s cells do not tile grid columns {c0}–{c1 - 1}: an existing merge overlaps the rectangle's edge",
                    cellAnchorId);
            rect.Add(band);
        }
        if (rect[0].Any(g => VMergeRestart(g.Tc) == false))
            return MergeFail("the rectangle's first row continues a vertical merge started above it", cellAnchorId);
        if (r0 + rowSpan < rows.Count && RowGrid(rows[r0 + rowSpan])
                .Any(g => g.End > c0 && g.Start < c1 && VMergeRestart(g.Tc) == false))
            return MergeFail("a vertical merge continues past the rectangle's last row", cellAnchorId);

        var absorbed = rect.SelectMany(b => b).Select(g => g.Tc).Where(x => x != tc).ToList();
        if (opts.Content == TableMergeContent.Reject && absorbed.Any(x => CellBlocks(x).Any(b => !IsEmptyBlock(b))))
            return MergeFail(
                "absorbed cells are not empty (use Content = Append to keep their content, or Discard to drop it)",
                cellAnchorId);

        var before = CaptureTableMetadata(tbl);
        var discardedBlocks = opts.Content == TableMergeContent.Append
            ? absorbed.SelectMany(CellBlocks).Where(IsEmptyBlock).ToList()
            : absorbed.SelectMany(CellBlocks).ToList();
        if (ValidateBookmarkRemoval(discardedBlocks, cellAnchorId) is { } bookmarkError)
            return bookmarkError;

        var hyperlinkOwner = Internal.OwnedPartRelationships.FindOwner(_doc!, tbl);
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // Content first: everything the merge absorbs MOVES into the surviving cell. Detach
            // before re-adding — XContainer.Add clones a still-parented node, which would leave
            // the original behind (and duplicate its Unid).
            if (opts.Content == TableMergeContent.Append)
                foreach (var block in absorbed.SelectMany(CellBlocks).Where(b => !IsEmptyBlock(b)).ToList())
                {
                    block.Remove();
                    tc!.Add(block);
                }

            // Then structure: one cell per row, spanning the rectangle's grid columns.
            int width = SumGridWidths(GridColWidths(tbl), c0, c1);
            for (int i = 0; i < rect.Count; i++)
            {
                var keep = rect[i][0].Tc;
                foreach (var g in rect[i].Skip(1)) g.Tc.Remove();
                SetGridSpan(keep, c1 - c0);
                if (width > 0) SetCellWidth(keep, width);
                if (rowSpan == 1) continue;
                SetVMerge(keep, restart: i == 0);
                if (i > 0) _ = EmptyCellBody(keep);
            }

            if (hyperlinkOwner is { } owner)
                SweepOrphanedStoryRelationships(owner.Part);

            InvalidateProjectionCache();
            var mapping = CompleteTableMapping(before, tbl);
            return new EditResult
            {
                Success = true,
                Removed = InvalidatedCellAnchors(mapping),
                Modified = new[] { AnchorForUnid(target!.Unid, target.PartUri) ?? target.Anchor },
                TableAnchors = mapping,
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, cellAnchorId);
        }
    }

    /// <summary>
    /// Split the merged cell at <paramref name="cellAnchorId"/> back into unit cells (Word's
    /// *Split Cells* undoing a merge): drop its <c>w:gridSpan</c> and restore one cell per grid
    /// column, and — for a vertical merge — do the same for every row of the run while dropping
    /// the <c>w:vMerge</c> markup. Addressing a continuation cell unmerges the whole run, not just
    /// that row. Restored cells clone the merged cell's shell (borders, shading, valign) without
    /// its merge markup, start empty and take their column's <c>w:tblGrid</c> width; the merged
    /// cell keeps its content. A cell carrying no merge markup is rejected
    /// (<see cref="EditErrorCode.InvalidTableMerge"/>).
    /// </summary>
    public EditResult UnmergeCells(string cellAnchorId)
    {
        if (ResolveCell(cellAnchorId, out _, out var tc, out var tr, out var tbl, out var target) is { } err)
            return err;
        if (RefuseUnresolvedTableStructure(tbl!, cellAnchorId) is { } pending) return pending;
        if (_trackedChanges == TrackedChangeMode.RenderInline)
            return TrackedStructureUnsupported("UnmergeCells", cellAnchorId);

        var shape = RowGrid(tr!).First(g => g.Tc == tc);
        bool? vMerge = VMergeRestart(tc!);
        if (shape.Span <= 1 && vMerge is null)
            return MergeFail("cell is not merged (no w:gridSpan and no w:vMerge)", cellAnchorId);

        // A continuation cell stands for the whole run: walk up to the restart, then down through
        // every column-aligned continuation.
        var rows = tbl!.Elements(W.tr).ToList();
        int r0 = rows.IndexOf(tr!), r1 = r0;
        if (vMerge is not null)
        {
            while (r0 > 0 && VMergeRestart(AlignedCell(rows[r0], shape)!) == false
                   && AlignedCell(rows[r0 - 1], shape) is { } up && VMergeRestart(up) is not null)
                r0--;
            while (r1 + 1 < rows.Count
                   && AlignedCell(rows[r1 + 1], shape) is { } down && VMergeRestart(down) == false)
                r1++;
        }

        var before = CaptureTableMetadata(tbl);
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var widths = GridColWidths(tbl);
            int Width(int col) => col < widths.Count ? widths[col] : 0;

            var created = new List<XElement>();
            for (int r = r0; r <= r1; r++)
            {
                if (AlignedCell(rows[r], shape) is not { } cell) continue;
                SetVMerge(cell, null);
                SetGridSpan(cell, 1);
                if (Width(shape.Start) > 0) SetCellWidth(cell, Width(shape.Start));

                var tail = cell;
                for (int col = shape.Start + 1; col < shape.End; col++)
                {
                    var unit = NewEmptyCellLike(cell);
                    if (Width(col) > 0) SetCellWidth(unit, Width(col));
                    UnidHelper.AssignToSelfAndDescendants(unit);
                    tail.AddAfterSelf(unit);
                    tail = unit;
                    created.Add(unit);
                }
            }

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Created = ResolveAnchorsForElements(created),
                Modified = new[] { AnchorForUnid(target!.Unid, target.PartUri) ?? target.Anchor },
                TableAnchors = CompleteTableMapping(before, tbl),
                Patch = PatchFor(target),
            };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, cellAnchorId);
        }
    }

    /// <summary>Insert (replacing any existing) a child at its correct schema position per
    /// <paramref name="order"/> — the generalized <see cref="SetPPrChildInOrder"/>.</summary>
    private static void SetChildInOrder(XElement parent, XElement child, string[] order)
    {
        parent.Elements(child.Name).Remove();
        int idx = Array.IndexOf(order, child.Name.LocalName);
        XElement? after = null;
        foreach (var e in parent.Elements())
        {
            int ei = Array.IndexOf(order, e.Name.LocalName);
            if (ei >= 0 && ei < idx) after = e;
            else if (ei >= idx) break;
        }
        if (after is null) parent.AddFirst(child);
        else after.AddAfterSelf(child);
    }

    /// <summary>w:tblPr must be the table's first child.</summary>
    private static XElement GetOrCreateTblPr(XElement tbl)
    {
        var tblPr = tbl.Element(W.tblPr);
        if (tblPr is null) { tblPr = new XElement(W.tblPr); tbl.AddFirst(tblPr); }
        return tblPr;
    }

    /// <summary>w:tcPr must be the cell's first child.</summary>
    private static XElement GetOrCreateTcPr(XElement tc)
    {
        var tcPr = tc.Element(W.tcPr);
        if (tcPr is null) { tcPr = new XElement(W.tcPr); tc.AddFirst(tcPr); }
        return tcPr;
    }

    /// <summary>The shared "styling applied" result: the target anchor in Modified + a patch.</summary>
    private EditResult TableStyleResult(AnchorTarget target, TableAnchorMapping? tableAnchors = null)
    {
        InvalidateProjectionCache();
        var updated = AnchorForUnid(target.Unid, target.PartUri) ?? target.Anchor;
        return new EditResult
        {
            Success = true,
            Modified = new[] { updated },
            Patch = PatchFor(target),
            TableAnchors = tableAnchors,
        };
    }

    /// <summary>
    /// Retune the column widths of the table containing <paramref name="cellAnchorId"/> —
    /// the post-insert counterpart of <see cref="TableInsertOptions.ColumnWidths"/>. Rewrites
    /// <c>w:tblGrid</c> and every row's <c>w:tcW</c>, sizes the table to the widths' sum
    /// (dxa) and pins <c>w:tblLayout</c> fixed, exactly as inserting with explicit widths
    /// would. One positive twip value per column is required.
    /// </summary>
    public EditResult SetColumnWidths(string cellAnchorId, IReadOnlyList<int> widthsTwips)
    {
        if (ResolveCell(cellAnchorId, out _, out _, out _, out var tbl, out var target) is { } err)
            return err;

        var before = Internal.TableGridModel.BuildMetadata(tbl!, AnchorForElement);
        var grid = tbl!.Element(W.tblGrid);
        int colCount = GridColumnCount(tbl);
        if (widthsTwips is null || widthsTwips.Count != colCount || widthsTwips.Any(w => w <= 0))
            return EditResult.Fail(EditErrorCode.InvalidTableStyling,
                $"widths must list one positive twip value per column ({colCount}); got {widthsTwips?.Count ?? 0}",
                cellAnchorId);

        var tracked = _trackedChanges == TrackedChangeMode.RenderInline;
        var tableProperties = tbl.Element(W.tblPr);
        var cellProperties = tbl.Descendants(W.tc).Select(cell => cell.Element(W.tcPr)).ToList();
        if (RefuseNestedTrackedPropertyChange(
                new[] { (grid, W.tblGridChange), (tableProperties, W.tblPrChange) }
                    .Concat(cellProperties.Select(properties => (properties, W.tcPrChange))),
                cellAnchorId) is { } pending)
            return pending;
        var oldGrid = tracked ? new XElement(grid ?? new XElement(W.tblGrid)) : null;
        var oldTableProperties = tracked ? new XElement(tableProperties ?? new XElement(W.tblPr)) : null;
        var oldCellProperties = tracked ? tbl.Descendants(W.tc).ToDictionary(cell => cell,
            cell => new XElement(cell.Element(W.tcPr) ?? new XElement(W.tcPr))) : null;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (grid is null)
            {
                grid = new XElement(W.tblGrid);
                var pr = tbl.Element(W.tblPr);
                if (pr is not null) pr.AddAfterSelf(grid);
                else tbl.AddFirst(grid);
            }
            var gridColumns = grid.Elements(W.gridCol).ToList();
            for (int index = 0; index < widthsTwips.Count; index++)
            {
                if (index < gridColumns.Count)
                {
                    gridColumns[index].SetAttributeValue(W._w, widthsTwips[index]);
                    continue;
                }
                var column = new XElement(W.gridCol, new XAttribute(W._w, widthsTwips[index]));
                UnidHelper.AssignToSelfAndDescendants(column);
                grid.Add(column);
            }

            // A merged cell is as wide as the grid columns it spans, so widths are summed over
            // each cell's grid range rather than read off its position in the row.
            var widths = widthsTwips.ToList();
            foreach (var row in tbl.Elements(W.tr))
                foreach (var cell in RowGrid(row))
                    if (SumGridWidths(widths, cell.Start, cell.End) is > 0 and var w)
                        SetCellWidth(cell.Tc, w);

            var tblPr = GetOrCreateTblPr(tbl);
            SetChildInOrder(tblPr,
                new XElement(W.tblW, new XAttribute(W._w, widthsTwips.Sum()), new XAttribute(W.type, "dxa")),
                TblPrChildOrder);
            SetChildInOrder(tblPr,
                new XElement(W.tblLayout, new XAttribute(W.type, "fixed")),
                TblPrChildOrder);

            if (tracked)
            {
                var author = _revisionAuthor ?? "docxodus";
                var date = NextTrackedFormatRevisionDate();
                TrackPropertyMutation(grid, oldGrid!, W.tblGridChange, author, date);
                TrackPropertyMutation(tblPr, oldTableProperties!, W.tblPrChange, author, date);
                foreach (var pair in oldCellProperties!)
                    TrackPropertyMutation(GetOrCreateTcPr(pair.Key), pair.Value,
                        W.tcPrChange, author, date, W.cellIns, W.cellDel, W.cellMerge);
            }

            return TableStyleResult(target!, CompleteTableMapping(before, tbl));
        }
        catch (Exception ex)
        {
            return FailInternal(ex, cellAnchorId);
        }
    }

    /// <summary>
    /// Set the table-level borders (<c>w:tblPr/w:tblBorders</c>) of the table containing
    /// <paramref name="cellAnchorId"/>. Only the edges named by <see cref="TableBorderSpec.Scope"/>
    /// are written (as explicit edges, so style-inherited borders are overridden); the rest are
    /// left untouched. Style "none" removes the targeted edges the way
    /// <see cref="TableInsertOptions.Borderless"/> does. Cell-level <c>w:tcBorders</c>, where a
    /// document has them, still win over these — v1 does not touch per-cell borders.
    /// </summary>
    public EditResult SetTableBorders(string cellAnchorId, TableBorderSpec? spec = null)
    {
        if (ResolveCell(cellAnchorId, out _, out _, out _, out var tbl, out var target) is { } err)
            return err;

        var s = spec ?? new TableBorderSpec();
        if (s.Size is < 0)
            return EditResult.Fail(EditErrorCode.InvalidTableStyling,
                "border size (eighths of a point) must be >= 0", cellAnchorId);

        var tracked = _trackedChanges == TrackedChangeMode.RenderInline;
        var existingTblPr = tbl!.Element(W.tblPr);
        if (RefuseNestedTrackedPropertyChange(
                new[] { (existingTblPr, W.tblPrChange) }, cellAnchorId) is { } pending)
            return pending;
        var oldTblPr = tracked ? new XElement(existingTblPr ?? new XElement(W.tblPr)) : null;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var tblPr = GetOrCreateTblPr(tbl!);
            var borders = tblPr.Element(W.tblBorders);
            if (borders is null)
            {
                borders = new XElement(W.tblBorders);
                SetChildInOrder(tblPr, borders, TblPrChildOrder);
            }

            var edges = s.Scope switch
            {
                TableBorderScope.Outside => new[] { W.top, W.left, W.bottom, W.right },
                TableBorderScope.Inside => new[] { W.insideH, W.insideV },
                _ => new[] { W.top, W.left, W.bottom, W.right, W.insideH, W.insideV },
            };

            bool none = string.Equals(s.Style, "none", StringComparison.OrdinalIgnoreCase);
            foreach (var edgeName in edges)
            {
                var edge = none
                    ? new XElement(edgeName, new XAttribute(W.val, "none"), new XAttribute(W.sz, 0),
                        new XAttribute(W.space, 0), new XAttribute(W.color, "auto"))
                    : new XElement(edgeName,
                        new XAttribute(W.val, string.IsNullOrEmpty(s.Style) ? "single" : s.Style),
                        new XAttribute(W.sz, s.Size ?? 4),
                        new XAttribute(W.space, 0),
                        new XAttribute(W.color, string.IsNullOrEmpty(s.Color) ? "auto" : s.Color));
                SetChildInOrder(borders, edge, TblBordersEdgeOrder);
            }

            if (tracked)
                TrackPropertyMutation(tblPr, oldTblPr!, W.tblPrChange,
                    _revisionAuthor ?? "docxodus", NextTrackedFormatRevisionDate());

            return TableStyleResult(target!);
        }
        catch (Exception ex)
        {
            return FailInternal(ex, cellAnchorId);
        }
    }

    /// <summary>
    /// Shade the cell containing <paramref name="cellAnchorId"/> — or, with
    /// <see cref="TableShadingScope.Row"/>, every cell of its row (header-row banding).
    /// <paramref name="fillColor"/> is a hex RRGGBB triplet (a leading '#' is tolerated) or
    /// "auto"; null/empty removes the shading. Writes <c>w:tcPr/w:shd</c> with
    /// <c>w:val="clear"</c>, Word's plain-fill idiom.
    /// </summary>
    public EditResult SetCellShading(string cellAnchorId, string? fillColor,
        TableShadingScope scope = TableShadingScope.Cell)
    {
        if (ResolveCell(cellAnchorId, out _, out var tc, out var tr, out _, out var target) is { } err)
            return err;

        bool clear = string.IsNullOrEmpty(fillColor);
        string fill = "auto";
        if (!clear)
        {
            fill = fillColor!.TrimStart('#');
            if (!string.Equals(fill, "auto", StringComparison.OrdinalIgnoreCase))
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(fill, "^[0-9A-Fa-f]{6}$"))
                    return EditResult.Fail(EditErrorCode.InvalidTableStyling,
                        $"fill must be a hex RRGGBB triplet or \"auto\"; got '{fillColor}'", cellAnchorId);
                fill = fill.ToUpperInvariant();
            }
            else fill = "auto";
        }

        var cells = scope == TableShadingScope.Row ? tr!.Elements(W.tc).ToList() : new List<XElement> { tc! };
        if (RefuseNestedTrackedPropertyChange(
                cells.Select(cell => (cell.Element(W.tcPr), W.tcPrChange)),
                cellAnchorId) is { } pending)
            return pending;
        var tracked = _trackedChanges == TrackedChangeMode.RenderInline;
        var oldCellProperties = tracked ? cells.ToDictionary(cell => cell,
            cell => new XElement(cell.Element(W.tcPr) ?? new XElement(W.tcPr))) : null;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            foreach (var cell in cells)
            {
                if (clear)
                {
                    cell.Element(W.tcPr)?.Elements(W.shd).Remove();
                    continue;
                }
                SetChildInOrder(GetOrCreateTcPr(cell),
                    new XElement(W.shd, new XAttribute(W.val, "clear"), new XAttribute(W.color, "auto"),
                        new XAttribute(W.fill, fill)),
                    TcPrChildOrder);
            }

            if (tracked)
            {
                var author = _revisionAuthor ?? "docxodus";
                var date = NextTrackedFormatRevisionDate();
                foreach (var pair in oldCellProperties!)
                    if (!PropertySnapshotEquals(pair.Value, pair.Key.Element(W.tcPr),
                            W.tcPrChange, W.cellIns, W.cellDel, W.cellMerge))
                        TrackPropertyMutation(GetOrCreateTcPr(pair.Key), pair.Value,
                            W.tcPrChange, author, date, W.cellIns, W.cellDel, W.cellMerge);
            }

            return TableStyleResult(target!);
        }
        catch (Exception ex)
        {
            return FailInternal(ex, cellAnchorId);
        }
    }

    /// <summary>
    /// Mark (or unmark) the row containing <paramref name="cellAnchorId"/> as a repeating
    /// header row (<c>w:trPr/w:tblHeader</c>), so a multi-page table re-shows it on every page.
    /// Word only honors the flag on a run of rows starting at the table's first row — setting
    /// it elsewhere is legal but ignored by renderers.
    /// </summary>
    public EditResult SetRepeatHeaderRow(string cellAnchorId, bool repeat)
        => SetTableRowOptions(cellAnchorId, new TableRowOptions { RepeatHeader = repeat });

    /// <summary>
    /// Set row-level layout properties for the row containing <paramref name="cellAnchorId"/>:
    /// repeat-header (<c>w:tblHeader</c>), page-split policy (<c>w:cantSplit</c>), and explicit
    /// height (<c>w:trHeight</c>). Null properties are left unchanged; height zero clears it.
    /// </summary>
    public EditResult SetTableRowOptions(string cellAnchorId, TableRowOptions? options = null)
    {
        if (ResolveCell(cellAnchorId, out _, out _, out var tr, out _, out var target) is { } err)
            return err;

        var opts = options ?? new TableRowOptions();
        if (opts.HeightTwips is < 0)
            return EditResult.Fail(EditErrorCode.InvalidTableStyling,
                "row height in twips must be >= 0", cellAnchorId);

        var tracked = _trackedChanges == TrackedChangeMode.RenderInline;
        var existingTrPr = tr!.Element(W.trPr);
        if (RefuseNestedTrackedPropertyChange(
                new[] { (existingTrPr, W.trPrChange) }, cellAnchorId) is { } pending)
            return pending;
        var oldTrPr = tracked ? new XElement(existingTrPr ?? new XElement(W.trPr)) : null;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var trPr = tr!.Element(W.trPr);

            if (opts.RepeatHeader is { } repeat)
            {
                if (repeat)
                {
                    if (trPr is null) { trPr = new XElement(W.trPr); tr.AddFirst(trPr); }
                    SetChildInOrder(trPr, new XElement(W.tblHeader), TrPrChildOrder);
                }
                else
                {
                    trPr?.Elements(W.tblHeader).Remove();
                }
            }

            if (opts.AllowBreakAcrossPages is { } allowBreak)
            {
                if (allowBreak)
                    trPr?.Elements(W.cantSplit).Remove();
                else
                {
                    if (trPr is null) { trPr = new XElement(W.trPr); tr.AddFirst(trPr); }
                    SetChildInOrder(trPr, new XElement(W.cantSplit), TrPrChildOrder);
                }
            }

            if (opts.HeightTwips is { } height)
            {
                if (height == 0)
                    trPr?.Elements(W.trHeight).Remove();
                else
                {
                    if (trPr is null) { trPr = new XElement(W.trPr); tr.AddFirst(trPr); }
                    var rule = opts.HeightRule switch
                    {
                        TableRowHeightRule.Auto => "auto",
                        TableRowHeightRule.Exact => "exact",
                        _ => "atLeast",
                    };
                    SetChildInOrder(trPr,
                        new XElement(W.trHeight,
                            new XAttribute(W.val, height),
                            new XAttribute(W.hRule, rule)),
                        TrPrChildOrder);
                }
            }

            // An emptied trPr is dropped entirely. Only element children matter: CT_TrPr has
            // no schema attributes, and the in-memory tree may carry pt bookkeeping attributes
            // (Unid) that Save() strips anyway.
            if (trPr is not null && !trPr.HasElements) trPr.Remove();

            if (tracked)
            {
                trPr = tr.Element(W.trPr);
                if (!PropertySnapshotEquals(oldTrPr!, trPr, W.trPrChange, W.ins, W.del))
                {
                    if (trPr is null) { trPr = new XElement(W.trPr); tr.AddFirst(trPr); }
                    TrackPropertyMutation(trPr, oldTrPr!, W.trPrChange,
                        _revisionAuthor ?? "docxodus", NextTrackedFormatRevisionDate(), W.ins, W.del);
                }
            }

            return TableStyleResult(target!);
        }
        catch (Exception ex)
        {
            return FailInternal(ex, cellAnchorId);
        }
    }
}
