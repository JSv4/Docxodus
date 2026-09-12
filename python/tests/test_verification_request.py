"""The full deliverable-verification request through the stdio host (issue #747)."""

from __future__ import annotations

import hashlib

from docx_scalpel import (
    DeliverableArtifactRole,
    DeliverableCompanionArtifact,
    DeliverableRenderDiagnostic,
    DeliverableVerificationDecision,
    DeliverableVerificationOptions,
    DeliverableVerificationRequest,
    VerificationDigest,
    docx_diff_get_semantic_changes,
    open_session,
    verify_deliverable,
)


def _sha256(data: bytes) -> VerificationDigest:
    return VerificationDigest(algorithm="SHA-256", value=hashlib.sha256(data).hexdigest())


def _deliverable(baseline: bytes) -> bytes:
    with open_session(baseline) as session:
        first = next(
            a.id for a in session.project().anchor_index.values()
            if a.scope == "body" and a.kind == "p"
        )
        assert session.replace_text(first, "Changed for delivery.").success
        return session.save()


def test_full_request_reports_unexpected_change_and_stale_artifact(tour_plan_bytes: bytes) -> None:
    deliverable = _deliverable(tour_plan_bytes)
    request = DeliverableVerificationRequest(
        options=DeliverableVerificationOptions(fail_on_unexpected_changes=True),
        expected_package_changes=(),
        companion_artifacts=(
            DeliverableCompanionArtifact(
                artifact_id="pdf-1",
                role=DeliverableArtifactRole.PDF,
                media_type="application/pdf",
                data=b"%PDF-1.4\n%stale\n",
                page_count=1,
                renderer_fingerprint="renderer/1.0",
                source_package_digest=_sha256(tour_plan_bytes),
                render_diagnostics=(
                    DeliverableRenderDiagnostic(kind="missingFont", message="Aptos substituted"),
                ),
            ),
        ),
    )

    report = verify_deliverable(deliverable, tour_plan_bytes, request=request)

    codes = {finding.code for finding in report.findings}
    assert "delta.package_change_unexpected" in codes
    assert "artifact.source_digest_mismatch" in codes
    assert report.decision is DeliverableVerificationDecision.FAILED
    assert [a.artifact_id for a in report.companion_artifacts] == ["pdf-1"]
    assert report.companion_artifacts[0].render_diagnostic_count == 1


def test_default_request_matches_the_simple_call(tour_plan_bytes: bytes) -> None:
    deliverable = _deliverable(tour_plan_bytes)
    assert verify_deliverable(deliverable, tour_plan_bytes) == verify_deliverable(
        deliverable, tour_plan_bytes, request=DeliverableVerificationRequest()
    )


def test_approved_semantic_changes_round_trip_and_unknown_options_are_rejected(
    tour_plan_bytes: bytes,
) -> None:
    deliverable = _deliverable(tour_plan_bytes)
    # Expectations come from the same byte-level comparison the verifier runs.
    expected = docx_diff_get_semantic_changes(tour_plan_bytes, deliverable)
    assert expected.change_count > 0

    approved = verify_deliverable(deliverable, tour_plan_bytes, request=DeliverableVerificationRequest(
        options=DeliverableVerificationOptions(fail_on_unexpected_changes=True),
        expected_semantic_changes=expected,
    ))
    codes = {finding.code for finding in approved.findings}
    assert "delta.semantic_change_unexpected" not in codes
    assert "delta.semantic_change_missing" not in codes

    with open_session(tour_plan_bytes) as session:
        first = next(
            a.id for a in session.project().anchor_index.values()
            if a.scope == "body" and a.kind == "p"
        )
        assert session.replace_text(first, "Changed for delivery.").success
        live = session.verify_deliverable(DeliverableVerificationRequest(
            options=DeliverableVerificationOptions(fail_on_unexpected_changes=True),
            expected_semantic_changes=expected,
        ))
        live_codes = {finding.code for finding in live.findings}
        assert "delta.semantic_change_unexpected" not in live_codes
        assert "delta.semantic_change_missing" not in live_codes

        try:
            session.verify_deliverable(DeliverableVerificationRequest(
                options=DeliverableVerificationOptions(package_manifest={"maxEntries": 1}),
            ))
        except Exception as error:  # the host names the unknown property
            assert "maxEntries" in str(error)
        else:
            raise AssertionError("an unknown option must be rejected")
