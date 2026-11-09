"""Canonical table addressing and complete Python table-op ripple (#450)."""

from __future__ import annotations

from docx_scalpel import (
    Position,
    TableAnchorEntityKind,
    TableBorderSpec,
    TableInsertOptions,
    TableRowHeightRule,
    TableRowOptions,
    open_session,
)


def test_table_metadata_resolution_and_all_mutations(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        body = next(
            anchor for anchor in session.project().anchor_index.values()
            if anchor.scope == "body" and anchor.kind in ("p", "h", "li")
        )
        inserted = session.insert_table(
            body.id,
            Position.AFTER,
            2,
            2,
            TableInsertOptions(
                cell_contents=("A", "B", "C", "D"),
                column_widths=(1800, 2200),
            ),
        )
        assert inserted.success
        assert inserted.created and all(anchor.kind == "tc" for anchor in inserted.created)
        assert inserted.table_anchors is not None
        table_id = next(
            location.anchor.id for location in inserted.table_anchors.added
            if location.entity_kind is TableAnchorEntityKind.TABLE
        )

        metadata_result = session.get_table_metadata(table_id)
        assert metadata_result.success and metadata_result.metadata is not None
        metadata = metadata_result.metadata
        assert metadata.anchor.kind == "tbl"
        assert [column.anchor.kind for column in metadata.columns] == ["col", "col"]
        assert all(not column.is_virtual for column in metadata.columns)
        assert [row.anchor.kind for row in metadata.rows] == ["tr", "tr"]
        cells = [cell for row in metadata.rows for cell in row.cells]
        assert len(cells) == 4 and all(cell.anchor.kind == "tc" for cell in cells)

        first = cells[0]
        by_anchor = session.resolve_table_cell_anchor(first.anchor.id)
        by_coordinate = session.resolve_table_cell_coordinate(table_id, 0, 0)
        assert by_anchor.success and by_anchor.cell == first
        assert by_coordinate.success and by_coordinate.cell == first

        assert session.replace_cell_content(first.anchor.id, "replaced").success
        assert session.set_column_widths(first.anchor.id, (2000, 2400)).success
        assert session.set_table_borders(
            first.anchor.id, TableBorderSpec(scope="outside", style="single", size=8)
        ).success
        assert session.set_cell_shading(first.anchor.id, "D9EAF7").success
        assert session.set_repeat_header_row(first.anchor.id, True).success
        assert session.set_table_row_options(
            first.anchor.id,
            TableRowOptions(
                repeat_header=True,
                allow_break_across_pages=False,
                height_twips=480,
                height_rule=TableRowHeightRule.AT_LEAST,
            ),
        ).success

        inserted_row = session.insert_table_row(first.anchor.id, Position.AFTER)
        assert inserted_row.success and inserted_row.table_anchors is not None
        inserted_column = session.insert_table_column(first.anchor.id, Position.AFTER)
        assert inserted_column.success and inserted_column.table_anchors is not None

        current = session.get_table_metadata(table_id).metadata
        assert current is not None
        merge_anchor = current.rows[0].cells[0].anchor.id
        merged = session.merge_cells(merge_anchor, 1, 2)
        assert merged.success and merged.table_anchors is not None
        assert all(anchor.kind == "tc" for anchor in merged.removed)
        unmerged = session.unmerge_cells(merge_anchor)
        assert unmerged.success and unmerged.table_anchors is not None
        assert all(anchor.kind == "tc" for anchor in unmerged.created)

        current = session.get_table_metadata(table_id).metadata
        assert current is not None
        deleted_row = session.delete_table_row(current.rows[-1].cells[0].anchor.id)
        assert deleted_row.success and deleted_row.table_anchors is not None
        current = session.get_table_metadata(table_id).metadata
        assert current is not None
        deleted_column = session.delete_table_column(current.rows[0].cells[-1].anchor.id)
        assert deleted_column.success and deleted_column.table_anchors is not None

        retained_column_ids = tuple(
            column.anchor.id for column in session.get_table_metadata(table_id).metadata.columns
        )
        saved = session.save(persist_anchor_ids=True)

    with open_session(saved) as reopened:
        reopened_metadata = reopened.get_table_metadata(table_id)
        assert reopened_metadata.success and reopened_metadata.metadata is not None
        assert tuple(column.anchor.id for column in reopened_metadata.metadata.columns) == retained_column_ids


def test_table_style_options_tracking_and_stale_preconditions(tour_plan_bytes: bytes) -> None:
    import io
    import zipfile
    import xml.etree.ElementTree as ET
    from docx_scalpel import EditErrorCode, MutationPreconditions, TrackedChangeMode

    # Independent style IDs with a real table border, added to the existing test document.
    ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(tour_plan_bytes)) as source, zipfile.ZipFile(output, "w") as target:
        for entry in source.infolist():
            data = source.read(entry)
            if entry.filename == "word/styles.xml":
                styles = ET.fromstring(data)
                for style_id in ("PythonTableA", "PythonTableB"):
                    style = ET.SubElement(styles, ns + "style", {ns + "type": "table", ns + "styleId": style_id})
                    ET.SubElement(style, ns + "name", {ns + "val": style_id})
                    borders = ET.SubElement(ET.SubElement(style, ns + "tblPr"), ns + "tblBorders")
                    ET.SubElement(borders, ns + "top", {ns + "val": "single", ns + "sz": "8"})
                data = ET.tostring(styles)
            target.writestr(entry, data)
    with open_session(output.getvalue()) as session:
        paragraph = next(a for a in session.project().anchor_index.values() if a.scope == "body" and a.kind == "p")
        inserted = session.insert_table(paragraph.id, Position.AFTER, 1, 1, TableInsertOptions(style_id="PythonTableA"))
        assert inserted.success
        table = next(a.anchor.id for a in inserted.table_anchors.added if a.entity_kind is TableAnchorEntityKind.TABLE)
        before = session.raw.get_xml(table)
        assert 'PythonTableA' in before
        version = session.get_version()
        with session.preconditioned(MutationPreconditions(expected_version=version - 1)):
            refused = session.set_table_style(table, "PythonTableB")
        assert not refused.success
        assert refused.error.code is EditErrorCode.PRECONDITION_FAILED
        assert session.get_version() == version
        assert session.raw.get_xml(table) == before
        session.set_tracked_changes(TrackedChangeMode.RENDER_INLINE)
        assert session.set_table_style(table, "PythonTableB").success
        revisions = session.list_revisions()
        assert len(revisions) == 1 and revisions[0].family == "properties_change"
        assert session.reject_revision(revisions[0].id).success
        # Property resolution renews internal property identities; content and formatting must restore.
        def without_internal_ids(xml: str) -> bytes:
            root = ET.fromstring(xml)
            for element in root.iter():
                for key in list(element.attrib):
                    if key == "{http://powertools.codeplex.com/2011}Unid":
                        del element.attrib[key]
            return ET.tostring(root)
        assert without_internal_ids(session.raw.get_xml(table)) == without_internal_ids(before)
