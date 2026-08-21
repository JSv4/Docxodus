"""Typed transport coverage for the redline reversibility proof.

The proof engine is covered by the .NET suite. These tests assert the stdio host and
the typed client hand Python callers the same schema-v1 document, correctly decoded.
"""

from __future__ import annotations

import hashlib
from pathlib import Path

import pytest

from docx_scalpel import (
    DocxodusTransportError,
    RedlineProofDirection,
    RedlineRevisionDisposition,
    docx_diff_compare,
    prove_redline_reversibility,
)

PAIR = ("WC/WC001-Digits.docx", "WC/WC001-Digits-Mod.docx")


@pytest.fixture(scope="module")
def triple(test_files_dir: Path) -> tuple[bytes, bytes, bytes]:
    """A baseline, an intended final, and a redline genuinely generated between them."""
    baseline_path, final_path = (test_files_dir / rel for rel in PAIR)
    if not baseline_path.exists() or not final_path.exists():
        pytest.skip(f"fixture absent: {PAIR[0]} / {PAIR[1]}")
    baseline, intended_final = baseline_path.read_bytes(), final_path.read_bytes()
    return baseline, intended_final, docx_diff_compare(baseline, intended_final)


def test_proof_is_typed_deterministic_and_bound_to_its_three_inputs(
    triple: tuple[bytes, bytes, bytes],
) -> None:
    baseline, intended_final, redline = triple

    proof = prove_redline_reversibility(baseline, intended_final, redline)

    assert proof == prove_redline_reversibility(baseline, intended_final, redline)
    assert proof.schema == (
        "https://docxodus.dev/schemas/verification/redline-reversibility-proof/v1"
    )
    assert proof.schema_version == 1
    assert isinstance(proof.success, bool)

    # Each recorded package identity is the digest of the exact bytes that were passed.
    for package, data in (
        (proof.baseline_package, baseline),
        (proof.intended_final_package, intended_final),
        (proof.redline_package, redline),
    ):
        assert package.raw_package_bytes_digest.algorithm == "SHA-256"
        assert package.raw_package_bytes_digest.value == hashlib.sha256(data).hexdigest()


def test_generated_revisions_are_classified_and_both_paths_are_typed(
    triple: tuple[bytes, bytes, bytes],
) -> None:
    baseline, intended_final, redline = triple

    proof = prove_redline_reversibility(baseline, intended_final, redline)

    # A redline the engine just generated owns its revisions; none can be conflicted,
    # because there is no competing pre-existing review state to conflict with.
    assert proof.revision_classifications
    assert any(
        item.disposition is RedlineRevisionDisposition.GENERATED
        for item in proof.revision_classifications
    )
    assert all(
        item.disposition is not RedlineRevisionDisposition.CONFLICTED
        for item in proof.revision_classifications
    )
    assert all(item.reason for item in proof.revision_classifications)

    accept, reject = proof.accept_to_final, proof.reject_to_baseline
    assert accept is not None and reject is not None
    assert accept.direction is RedlineProofDirection.ACCEPT_TO_FINAL
    assert reject.direction is RedlineProofDirection.REJECT_TO_BASELINE
    # Each path's expected package is the document that path must reproduce.
    assert accept.expected_package.raw_package_bytes_digest.value == hashlib.sha256(
        intended_final
    ).hexdigest()
    assert reject.expected_package.raw_package_bytes_digest.value == hashlib.sha256(
        baseline
    ).hexdigest()
    # Enums decode as enums, not bare strings, on every nested divergence.
    for divergence in (*accept.divergences, *reject.divergences):
        assert divergence.kind.value in {"added", "removed", "modified"}
        assert divergence.part_uri.startswith("/")


def test_malformed_package_is_a_typed_finding_not_a_raised_error(
    triple: tuple[bytes, bytes, bytes],
) -> None:
    baseline, intended_final, _ = triple

    proof = prove_redline_reversibility(baseline, intended_final, b"not-a-package")

    assert not proof.success
    assert proof.findings
    assert any(finding.severity.value == "error" for finding in proof.findings)
    # Fail-closed: neither path is attempted, so no partial result can be misread as evidence.
    assert proof.accept_to_final is None
    assert proof.reject_to_baseline is None


def test_host_refuses_a_missing_package_rather_than_proving_against_nothing() -> None:
    from docx_scalpel._transport import call as transport_call

    with pytest.raises(DocxodusTransportError):
        transport_call("prove_redline_reversibility", {"baselineB64": ""})
