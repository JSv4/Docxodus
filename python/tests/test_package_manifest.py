"""End-to-end coverage for the stateless and live-session manifest surfaces."""

from __future__ import annotations

import hashlib

from docx_scalpel import (
    PackageContentTypeSource,
    PackageKind,
    PackageManifestInspectionLimits,
    generate_package_manifest,
    open_session,
)


def test_stateless_package_manifest_is_typed_deterministic_and_exact(
    tour_plan_bytes: bytes,
) -> None:
    first = generate_package_manifest(tour_plan_bytes)
    second = generate_package_manifest(tour_plan_bytes)

    assert first == second
    assert first.schema == "https://docxodus.dev/schemas/verification/package-manifest/v1"
    assert first.schema_version == 1
    assert first.package_kind is PackageKind.OPC
    assert first.is_valid
    assert first.raw_package_bytes_digest.value == hashlib.sha256(tour_plan_bytes).hexdigest()
    assert first.entries
    assert all(
        isinstance(entry.content_type_source, PackageContentTypeSource)
        for entry in first.entries
    )
    # ZIP64 lengths use decimal strings on the JSON wire but remain exact Python ints.
    assert all(isinstance(entry.size, int) and entry.size >= 0 for entry in first.entries)
    assert all(
        isinstance(entry.compressed_size, int) and entry.compressed_size >= 0
        for entry in first.entries
    )


def test_live_session_manifest_overlays_edits_without_touching_history(
    tour_plan_bytes: bytes,
) -> None:
    with open_session(tour_plan_bytes) as session:
        projection = session.project()
        anchor = next(
            target.id
            for target in projection.anchor_index.values()
            if target.kind in ("p", "h", "li") and target.scope == "body"
        )

        before = session.get_package_manifest()
        version_before_read = session.get_version()
        assert session.get_package_manifest() == before
        assert session.get_version() == version_before_read

        edit = session.replace_text(anchor, "package manifest Python ripple")
        assert edit.success
        after = session.get_package_manifest()
        version_after_edit = session.get_version()
        assert after.normalized_semantic_digest != before.normalized_semantic_digest
        assert session.get_version() == version_after_edit

        assert session.undo()
        assert (
            session.get_package_manifest().normalized_semantic_digest
            == before.normalized_semantic_digest
        )
        assert session.redo()
        assert (
            session.get_package_manifest().normalized_semantic_digest
            == after.normalized_semantic_digest
        )


def test_stateless_manifest_honours_lowered_inspection_limits(
    tour_plan_bytes: bytes,
) -> None:
    """A caller's lowered ceiling must constrain the inspection itself.

    The engine reports the breach as a structured finding on an otherwise well-formed
    package rather than raising, so the same call under default ceilings stays valid.
    """
    default = generate_package_manifest(tour_plan_bytes)
    assert default.is_valid

    constrained = generate_package_manifest(
        tour_plan_bytes,
        PackageManifestInspectionLimits(opc_entries=1),
    )

    assert constrained.findings
    assert constrained != default
