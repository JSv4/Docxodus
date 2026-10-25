"""Account ownership through the public Python client and the real stdio host."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import quoteattr

import pytest

from docx_scalpel import (
    DocxSessionSettings,
    RevisionAuthorNameStatus,
    RevisionIdentity,
    TrackedChangeMode,
    open_session,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.microsoft.com/office/word/2012/wordml"
OUR_ACCOUNT = "S::pat@ourfirm.example::22222222-2222-2222-2222-222222222222"
OTHER_ACCOUNT = "S::pat@counterparty.example::11111111-1111-1111-1111-111111111111"
PROPOSAL = "Counterparty proposal."


def identity_docx(author: str = "Pat", account: str | None = OTHER_ACCOUNT) -> bytes:
    """Original three-paragraph input with a pending insertion and optional people mapping."""
    types = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    for part, kind in (("document", "document.main"), ("styles", "styles"), ("settings", "settings")):
        types += f'<Override PartName="/word/{part}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>'
    if account is not None:
        types += '<Override PartName="/word/people.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"/>'
    relationships = f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyles" Type="{R}/styles" Target="styles.xml"/><Relationship Id="rSettings" Type="{R}/settings" Target="settings.xml"/>'
    if account is not None:
        relationships += '<Relationship Id="rPeople" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"/>'
    stamp = f'w:author={quoteattr(author)} w:date="2026-05-01T09:00:00Z"'
    parts = {
        "[Content_Types].xml": types + "</Types>",
        "_rels/.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/_rels/document.xml.rels": relationships + "</Relationships>",
        "word/document.xml": f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>Before.</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:ins w:id="1" {stamp}/></w:rPr></w:pPr><w:ins w:id="2" {stamp}><w:r><w:t>{PROPOSAL}</w:t></w:r></w:ins></w:p><w:p><w:r><w:t>After.</w:t></w:r></w:p></w:body></w:document>',
        "word/styles.xml": f'<w:styles xmlns:w="{W}"/>',
        "word/settings.xml": f'<w:settings xmlns:w="{W}"/>',
    }
    if account is not None:
        parts["word/people.xml"] = f'<w15:people xmlns:w15="{P}"><w15:person w15:author={quoteattr(author)}><w15:presenceInfo w15:providerId="AD" w15:userId={quoteattr(account)}/></w15:person></w15:people>'
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w") as archive:
        for name, content in parts.items():
            archive.writestr(zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0)), content.encode())
    return stream.getvalue()


def part(docx: bytes, name: str) -> ET.Element | None:
    with zipfile.ZipFile(io.BytesIO(docx)) as archive:
        return ET.fromstring(archive.read(name)) if name in archive.namelist() else None


@pytest.mark.parametrize("author,account,configured,own", [
    ("Pat", OTHER_ACCOUNT, True, False),
    ("Pat", None, True, False),
    ("Word Pat", "S::PAT@OURFIRM.EXAMPLE::33333333-3333-3333-3333-333333333333", True, True),
    ("Word Pat", "S::former@ourfirm.example::22222222-2222-2222-2222-222222222222", True, True),
    ("Pat", OTHER_ACCOUNT, False, True),
])
def test_account_ownership_and_legacy_control(author: str, account: str | None, configured: bool, own: bool, tmp_path: Path) -> None:
    source = identity_docx(author, account)
    (tmp_path / "source.docx").write_bytes(source)
    settings = DocxSessionSettings(
        tracked_changes=TrackedChangeMode.RENDER_INLINE,
        revision_author="Pat",
        revision_identities=[RevisionIdentity("AD", OUR_ACCOUNT)] if configured else None,
    )
    with open_session(source, settings) as session:
        info = session.get_revision_author_name_status(author)
        assert info.author == author
        assert info.status is (RevisionAuthorNameStatus.OWN if own else RevisionAuthorNameStatus.TAKEN)
        found = session.find_by_text(PROPOSAL)
        assert found is not None
        assert session.delete_block(found.id).success
        saved = session.save()
    (tmp_path / "result.docx").write_bytes(saved)
    document = part(saved, "word/document.xml")
    assert document is not None
    assert len(document.findall(f"{{{W}}}body/{{{W}}}p")) == (2 if own else 3)
    assert bool(document.findall(f".//{{{W}}}del")) is not own


def test_setting_identities_mid_session_and_clearing_them() -> None:
    with open_session(identity_docx(), DocxSessionSettings(revision_author="Pat")) as session:
        assert session.get_revision_author_name_status().status is RevisionAuthorNameStatus.OWN
        session.set_revision_identities([RevisionIdentity("AD", OUR_ACCOUNT)])
        assert session.get_revision_author_name_status().status is RevisionAuthorNameStatus.TAKEN
        assert session.get_revision_author_name_status("Available Reviewer").status is RevisionAuthorNameStatus.FREE
        session.set_revision_identities(None)
        assert session.get_revision_author_name_status().status is RevisionAuthorNameStatus.OWN


def test_a_new_claim_is_saved_and_undo_redo_restores_its_part(tmp_path: Path) -> None:
    source = identity_docx("Earlier", None)
    (tmp_path / "source.docx").write_bytes(source)
    with open_session(source, DocxSessionSettings(
        tracked_changes=TrackedChangeMode.RENDER_INLINE, revision_author="Pat",
        revision_identities=[RevisionIdentity("AD", OUR_ACCOUNT)],
    )) as session:
        assert session.get_revision_author_name_status().status is RevisionAuthorNameStatus.FREE
        found = session.find_by_text("Before.")
        assert found is not None
        assert session.replace_text(found.id, "Edited.").success
        saved = session.save()
        people = part(saved, "word/people.xml")
        assert people is not None
        person = people.find(f"{{{P}}}person")
        assert person is not None and person.get(f"{{{P}}}author") == "Pat"
        presence = person.find(f"{{{P}}}presenceInfo")
        assert presence is not None and presence.get(f"{{{P}}}userId") == OUR_ACCOUNT
        assert session.undo()
        assert part(session.save(), "word/people.xml") is None
        assert session.redo()
        assert session.get_revision_author_name_status().status is RevisionAuthorNameStatus.OWN
    (tmp_path / "result.docx").write_bytes(saved)
