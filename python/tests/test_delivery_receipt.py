"""Portable delivery-receipt verification transport coverage (issue #520).

Uses the vendored cross-language fixture ``TestFiles/Delivery/DR001-*`` — the same
files the C# ``DCR055`` pin verifies — so a canonical-format drift is caught on
both sides of the wire.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_scalpel import (
    DeliveryArtifactVerificationStatus,
    DeliveryReceiptVerificationResult,
    verify_delivery_receipt,
)


@pytest.fixture(scope="session")
def receipt_fixture(test_files_dir: Path) -> tuple[str, dict[str, bytes]]:
    delivery = test_files_dir / "Delivery"
    receipt_json = (delivery / "DR001-Receipt.json").read_text(encoding="utf-8")
    artifacts = {
        "clean-docx": (test_files_dir / "HC001-5DayTourPlanTemplate.docx").read_bytes(),
        "semantic-source-to-delivered": (delivery / "DR001-Semantic.json").read_bytes(),
    }
    return receipt_json, artifacts


def test_receipt_verifies_with_exact_artifacts(
    receipt_fixture: tuple[str, dict[str, bytes]],
) -> None:
    receipt_json, artifacts = receipt_fixture

    result = verify_delivery_receipt(receipt_json, artifacts)

    assert isinstance(result, DeliveryReceiptVerificationResult)
    assert result.is_valid, result.findings
    assert result.receipt_digest_valid
    assert result.contract_valid
    assert result.citation_bindings_valid
    assert not result.findings
    statuses = {artifact.artifact_id: artifact.status for artifact in result.artifacts}
    assert statuses == {
        "clean-docx": DeliveryArtifactVerificationStatus.VERIFIED,
        "semantic-source-to-delivered": DeliveryArtifactVerificationStatus.VERIFIED,
    }
    clean = next(a for a in result.artifacts if a.artifact_id == "clean-docx")
    assert clean.expected_digest is not None
    assert clean.expected_digest.algorithm == "SHA-256"
    assert clean.expected_length == len(artifacts["clean-docx"])


def test_tampered_artifact_and_absence_are_detected(
    receipt_fixture: tuple[str, dict[str, bytes]],
) -> None:
    receipt_json, artifacts = receipt_fixture

    tampered = dict(artifacts)
    tampered["clean-docx"] = artifacts["clean-docx"][:-1] + bytes(
        [artifacts["clean-docx"][-1] ^ 0xFF]
    )
    rejected = verify_delivery_receipt(receipt_json, tampered)
    assert not rejected.is_valid
    assert rejected.receipt_digest_valid  # the envelope itself is untouched
    statuses = {a.artifact_id: a.status for a in rejected.artifacts}
    assert statuses["clean-docx"] is DeliveryArtifactVerificationStatus.DIGEST_MISMATCH

    bare = verify_delivery_receipt(receipt_json)
    assert not bare.is_valid
    assert bare.receipt_digest_valid
    assert all(
        a.status is DeliveryArtifactVerificationStatus.MISSING for a in bare.artifacts
    )


def test_malformed_receipt_is_a_structured_verdict() -> None:
    result = verify_delivery_receipt('{"nope": true}')

    assert not result.is_valid
    assert not result.receipt_digest_valid
    assert result.findings
