"""Typed stateless and live-session deliverable-verification transport coverage."""

from __future__ import annotations

import base64
import hashlib

import pytest

from docx_scalpel import (
    DeliverableCheckStatus,
    DeliverableVerificationDecision,
    DeliverableVerificationMode,
    DocxodusTransportError,
    open_session,
    verify_deliverable,
)
from docx_scalpel._transport import call as transport_call


def test_stateless_verification_is_typed_deterministic_and_byte_bound(
    tour_plan_bytes: bytes,
) -> None:
    first = verify_deliverable(tour_plan_bytes)
    second = verify_deliverable(tour_plan_bytes)

    assert first == second
    assert first.schema == (
        "https://docxodus.dev/schemas/verification/deliverable-verification/v1"
    )
    assert first.schema_version == 1
    assert first.mode is DeliverableVerificationMode.STANDARD
    assert first.mode.value == "standard"
    assert isinstance(first.decision, DeliverableVerificationDecision)
    assert first.decision.value[0].islower()
    assert not first.baseline_compared
    assert first.baseline_package is None
    assert first.deliverable_package.raw_package_bytes_digest.value == hashlib.sha256(
        tour_plan_bytes
    ).hexdigest()
    assert first.checks
    assert all(isinstance(check.status, DeliverableCheckStatus) for check in first.checks)

    compared = verify_deliverable(tour_plan_bytes, baseline=tour_plan_bytes)
    assert compared.baseline_compared
    assert compared.baseline_package is not None
    expected_digest = hashlib.sha256(tour_plan_bytes).hexdigest()
    assert compared.baseline_package.raw_package_bytes_digest.value == expected_digest
    assert compared.deliverable_package.raw_package_bytes_digest.value == expected_digest
    assert compared.semantic_delta is not None
    assert compared.semantic_delta.change_count == 0


def test_session_verification_uses_opening_baseline_and_does_not_mutate_version(
    tour_plan_bytes: bytes,
) -> None:
    with open_session(tour_plan_bytes) as session:
        projection = session.project()
        anchor = next(
            target.id
            for target in projection.anchor_index.values()
            if target.kind in ("p", "h", "li") and target.scope == "body"
        )
        edit = session.replace_text(anchor, "deliverable verification transport edit")
        assert edit.success

        checkpoint_bytes = session.save()
        version_before = session.get_version()
        report = session.verify_deliverable()

        assert session.get_version() == version_before
        assert report.baseline_compared
        assert report.baseline_package is not None
        assert report.baseline_package.raw_package_bytes_digest.value == hashlib.sha256(
            tour_plan_bytes
        ).hexdigest()
        assert report.deliverable_package.raw_package_bytes_digest.value == hashlib.sha256(
            checkpoint_bytes
        ).hexdigest()
        assert report.semantic_delta is not None
        assert report.semantic_delta.change_count == len(report.semantic_delta.changes)
        assert report.package_changes


@pytest.mark.parametrize(
    ("args", "message"),
    [
        ({"docxB64": None}, '"docxB64" must be a string'),
        ({"docxB64": 42}, '"docxB64" must be a string'),
        ({"docxB64": "", "baselineB64": None}, '"baselineB64" must be a string'),
        ({"baselineB64": base64.b64encode(b"baseline").decode("ascii")},
         '"baselineB64" requires string "docxB64"'),
    ],
)
def test_stateless_verification_rejects_malformed_operation_shapes(
    args: dict[str, object],
    message: str,
) -> None:
    with pytest.raises(DocxodusTransportError, match=message) as exc_info:
        transport_call("verify_deliverable", args)
    assert exc_info.value.code == "malformed_request"
