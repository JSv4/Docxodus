"""Portable PageMap registration and citation transport coverage."""

from __future__ import annotations

from docx_scalpel import (
    PageCitationRequest,
    PageMap,
    PageMapFragment,
    PageMapPage,
    PageMapRect,
    open_session,
)


def test_register_and_consume_page_map(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        target = next(
            anchor
            for anchor in session.project().anchor_index.values()
            if anchor.scope == "body" and anchor.kind == "tbl"
        )
        request = PageCitationRequest(
            document_version=session.get_version(),
            renderer_fingerprint="python-page-map-v1",
        )
        page_map = PageMap(
            document_version=request.document_version,
            renderer_fingerprint=request.renderer_fingerprint,
            pages=(
                PageMapPage(
                    page_number=1,
                    page_in_section=1,
                    width=612,
                    height=792,
                    page_name="docxodus-section-0",
                    section_index=0,
                ),
            ),
            fragments=(
                PageMapFragment(
                    fragment_id=f"p1-f0-{target.id}",
                    anchor_id=target.id,
                    fragment_index=0,
                    page_number=1,
                    geometry=PageMapRect(72, 90, 468, 120),
                    story="body",
                ),
            ),
        )

        assert session.register_page_map(page_map).success
        assert session.get_page_map_status(request).availability == "available"

        citation = session.get_page_citation(target.id, request)
        assert citation.availability == "available"
        assert citation.fragments[0].page_number == 1

        structural = session.find_by_kind("tbl", "body", request)
        cited = next(anchor for anchor in structural if anchor.id == target.id)
        assert cited.citation is not None
        assert cited.citation.fragments[0].fragment_id == f"p1-f0-{target.id}"

        mismatch = session.get_page_map_status(
            PageCitationRequest(request.document_version, "different-renderer")
        )
        assert mismatch.availability == "unavailable"
        assert mismatch.unavailable_reason == "renderer_fingerprint_mismatch"
