"""Word's Table Properties: ``set_table_properties``, ``set_cell_properties`` and
``get_table_properties``."""

from __future__ import annotations

from docx_scalpel import (
    CellPropertiesOp,
    Position,
    TableCellMargins,
    TableInsertOptions,
    TablePropertiesOp,
    TableRowOptions,
    open_session,
)


def _table(session):
    body = next(
        anchor for anchor in session.project().anchor_index.values()
        if anchor.scope == "body" and anchor.kind in ("p", "h", "li")
    )
    inserted = session.insert_table(
        body.id, Position.AFTER, 2, 3, TableInsertOptions(cell_contents=("A", "B", "C", "D", "E", "F")),
    )
    assert inserted.success
    return [anchor.id for anchor in inserted.created]


def test_set_table_properties_writes_and_reads_back_and_undoes(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        cells = _table(session)
        before = session.get_table_properties(cells[4])
        assert before is not None and before.alignment is None and len(before.column_widths_twips) == 3

        result = session.set_table_properties(cells[0], TablePropertiesOp(
            width_type="percent", width=80, alignment="center", indent_twips=0,
            cell_margins=TableCellMargins(left_twips=144, right_twips=144),
            alt_title="Letters", alt_description="Six letters",
        ))

        assert result.success, result.error
        info = session.get_table_properties(cells[4])
        assert info is not None
        assert (info.width_type, info.width, info.alignment) == ("percent", 80.0, "center")
        assert info.cell_margins is not None and info.cell_margins.left_twips == 144
        assert info.cell_margins.top_twips is None
        assert (info.alt_title, info.alt_description) == ("Letters", "Six letters")
        assert (info.row_index, info.column_index, info.row_count) == (1, 1, 2)
        assert session.undo()
        assert session.get_table_properties(cells[4]).alignment is None


def test_set_cell_properties_by_scope_and_row_options_read_back(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        cells = _table(session)
        assert session.set_cell_properties(cells[1], CellPropertiesOp(vertical_alignment="bottom"), "column").success
        assert session.get_table_properties(cells[4]).vertical_alignment == "bottom"
        assert session.get_table_properties(cells[3]).vertical_alignment is None
        assert session.set_cell_properties(cells[3], CellPropertiesOp(vertical_alignment="center")).success
        assert session.get_table_properties(cells[3]).vertical_alignment == "center"

        assert session.set_table_row_options(cells[3], TableRowOptions(height_twips=720)).success
        info = session.get_table_properties(cells[5])
        assert (info.row_height_twips, info.row_height_rule) == (720, "atLeast")


def test_bad_values_are_refused_and_outside_a_table_reads_none(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        cells = _table(session)
        refused = session.set_table_properties(cells[0], TablePropertiesOp(alignment="middle"))
        assert not refused.success
        assert refused.error.code.value == "invalid_table_styling"
        refused = session.set_cell_properties(cells[0], CellPropertiesOp(vertical_alignment="middle"))
        assert not refused.success
        outside = next(
            anchor.id for anchor in session.project().anchor_index.values()
            if anchor.scope == "body" and anchor.kind in ("p", "h", "li")
        )
        assert session.get_table_properties(outside) is None
