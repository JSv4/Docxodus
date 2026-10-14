"""Explicit revision repairs through the stdio host (issues #754–#758)."""

from __future__ import annotations

import io
import zipfile

from docx_scalpel import EditErrorCode, RevisionRepairRequest, open_session


def _with_idless_insertion(docx: bytes) -> bytes:
    """Wrap the first run of the body in a w:ins that carries no w:id."""
    source = zipfile.ZipFile(io.BytesIO(docx))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "word/document.xml":
                text = data.decode("utf-8")
                start = text.index("<w:r>") if "<w:r>" in text else text.index("<w:r ")
                end = text.index("</w:r>", start) + len("</w:r>")
                run = text[start:end]
                text = (text[:start]
                        + '<w:ins w:author="Anonymous" w:date="2026-01-01T00:00:00Z">' + run + "</w:ins>"
                        + text[end:])
                data = text.encode("utf-8")
            target.writestr(item, data)
    return out.getvalue()


def test_missing_id_is_proposed_repaired_and_then_resolvable(tour_plan_bytes: bytes) -> None:
    with open_session(_with_idless_insertion(tour_plan_bytes)) as session:
        listed = session.list_revisions()
        broken = next(r for r in listed if r.diagnostic and r.diagnostic.code == "missing_revision_id")
        proposals = session.list_revision_repairs()
        proposal = next(p for p in proposals if p.revision_id == broken.id)
        assert proposal.kind == "assign_identity" and proposal.repairable
        assert any(c.startswith("w:ins@") for c in proposal.carriers)

        result = session.repair_revisions([RevisionRepairRequest(broken.id, "assign_identity")])

        assert result.success, result.error
        identity = result.repairs[0].identities[0]
        assert identity.old_id is None and identity.new_id.isdigit()
        repaired = next(r for r in session.list_revisions() if identity.new_id in r.constituent_ids)
        assert repaired.resolution_status == "supported"
        assert session.accept_revision(repaired.id).success
        assert session.undo() and session.undo()
        assert any(r.diagnostic and r.diagnostic.code == "missing_revision_id"
                   for r in session.list_revisions())


def test_unknown_or_unoffered_repairs_are_refused_without_mutation(tour_plan_bytes: bytes) -> None:
    with open_session(_with_idless_insertion(tour_plan_bytes)) as session:
        version = session.get_version()
        unknown = session.repair_revisions([RevisionRepairRequest("rev2-nope", "assign_identity")])
        assert not unknown.success and unknown.error is not None
        assert unknown.error.code is EditErrorCode.REVISION_NOT_FOUND

        broken = next(r for r in session.list_revisions()
                      if r.diagnostic and r.diagnostic.code == "missing_revision_id")
        wrong_kind = session.repair_revisions([RevisionRepairRequest(broken.id, "restore_orphan_text")])
        assert not wrong_kind.success and wrong_kind.error is not None
        assert wrong_kind.error.code is EditErrorCode.REVISION_REPAIR_REJECTED
        assert session.get_version() == version
