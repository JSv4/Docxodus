"""Host-captured delivery evidence and receipt-bearing delivery through the stdio host (issue #748)."""

from __future__ import annotations

from docx_scalpel import (
    DeliveryReceiptPrivacyProfile,
    DocxSession,
    DocxSessionSettings,
    MutationBatchStep,
    open_session,
    verify_delivery_receipt,
)


def _first_paragraph(session: DocxSession) -> str:
    projection = session.project()
    return next(
        anchor.id
        for anchor in projection.anchor_index.values()
        if anchor.scope == "body" and anchor.kind in ("p", "h", "li")
    )


def test_captured_edits_mint_a_verifiable_receipt(tour_plan_bytes: bytes) -> None:
    settings = DocxSessionSettings(capture_delivery_evidence=True)
    with open_session(tour_plan_bytes, settings) as session:
        anchor = _first_paragraph(session)

        # A direct call, a transactional batch (retried), a failed call, undo and redo.
        assert session.replace_text(anchor, "Directly edited.").success
        steps = [MutationBatchStep("insert_paragraph", {
            "anchorId": anchor, "position": "after", "markdown": "Batched.",
        })]
        first = session.execute_batch(steps, transaction_id="tx-deliver")
        assert session.execute_batch(steps, transaction_id="tx-deliver") == first
        assert not session.replace_text("p:body:missing", "Never.").success
        assert session.undo() and session.redo()

        status = session.get_delivery_evidence_status()
        assert status.enabled and status.unavailable_reason is None
        assert status.transaction_count == 3
        assert status.lineage_event_count == 2
        assert status.unlabeled_transaction_count == 0

        bundle = session.build_delivery_receipt(
            privacy_profile=DeliveryReceiptPrivacyProfile.FULL_EVIDENCE,
        )
        assert bundle.status == "complete" and bundle.verified
        assert bundle.evidence is not None and bundle.evidence.unavailable_reason is None
        receipt = bundle.artifact("change-receipt")
        assert receipt is not None and receipt.availability == "available"
        assert receipt.bytes is not None

        artifacts = {
            a.artifact_id: a.bytes
            for a in bundle.artifacts
            if a.bytes is not None and a.artifact_id != "change-receipt"
        }
        verification = verify_delivery_receipt(receipt.bytes.decode("utf-8"), artifacts)
        assert verification.is_valid, verification.findings

        final = bundle.artifact("final-docx")
        assert final is not None and final.bytes is not None
        with open_session(final.bytes) as delivered:
            markdown = delivered.project().markdown
            assert "Directly edited." in markdown and "Batched." in markdown


def test_without_capture_the_receipt_is_explicitly_unavailable(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = _first_paragraph(session)
        assert session.replace_text(anchor, "Unrecorded.").success

        status = session.get_delivery_evidence_status()
        assert not status.enabled
        assert status.unavailable_reason is not None

        bundle = session.build_delivery_receipt()
        assert bundle.status == "incomplete"
        receipt = bundle.artifact("change-receipt")
        assert receipt is not None and receipt.availability == "unavailable"
        assert receipt.bytes is None and receipt.unavailable_reason
        assert bundle.evidence is not None and not bundle.evidence.enabled
        final = bundle.artifact("final-docx")
        assert final is not None and final.bytes is not None
        with open_session(final.bytes) as delivered:
            assert "Unrecorded." in delivered.project().markdown
