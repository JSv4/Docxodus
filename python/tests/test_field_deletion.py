"""A field is one native, reversible deletion, including through atomic batches."""
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from docx_scalpel import DocxSessionSettings, MutationBatchStep, TrackedChangeMode, open_session

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _source(package: bytes) -> bytes:
    stream = BytesIO()
    with ZipFile(BytesIO(package)) as original, ZipFile(stream, "w", ZIP_DEFLATED) as output:
        for entry in original.infolist():
            content = original.read(entry.filename)
            if entry.filename == "word/document.xml":
                root = ET.fromstring(content)
                body = root.find(W + "body")
                for child in list(body):
                    if child.tag != W + "sectPr":
                        body.remove(child)
                paragraph = ET.fromstring('''<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                    <w:r><w:t xml:space="preserve">Before </w:t></w:r>
                    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
                    <w:r><w:instrText> REF heading </w:instrText></w:r>
                    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
                    <w:r><w:t>Section 4</w:t></w:r>
                    <w:r><w:fldChar w:fldCharType="end"/></w:r>
                    <w:r><w:t xml:space="preserve"> after.</w:t></w:r>
                </w:p>''')
                body.insert(0, paragraph)
                content = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            output.writestr(entry, content)
    return stream.getvalue()


def test_delete_field_and_replace_text_in_one_atomic_batch(tour_plan_bytes: bytes):
    settings = DocxSessionSettings(tracked_changes=TrackedChangeMode.RENDER_INLINE)
    with open_session(_source(tour_plan_bytes), settings) as session:
        anchor = session.grep("Section 4")[0].enclosing_anchor.id
        result = session.execute_batch([
            MutationBatchStep("delete_field_at_span", {"anchorId": anchor, "spanStart": 7, "spanLength": 9}),
            MutationBatchStep("replace_text_at_span", {"anchorId": anchor, "spanStart": 0, "spanLength": 6, "replace": "Prior"}),
        ])
        assert result.success
        content = session.save()
    with open_session(content) as accepted:
        assert accepted.accept_all_revisions().success
        assert accepted.grep("Prior  after")
    with open_session(content) as rejected:
        assert rejected.reject_all_revisions().success
        assert rejected.grep("Before Section 4 after")


def test_failed_field_batch_rolls_back_prior_steps(tour_plan_bytes: bytes):
    with open_session(_source(tour_plan_bytes)) as session:
        anchor = session.grep("Section 4")[0].enclosing_anchor.id
        result = session.execute_batch([
            MutationBatchStep("replace_text_at_span", {"anchorId": anchor, "spanStart": 0, "spanLength": 6, "replace": "Prior"}),
            MutationBatchStep("delete_field_at_span", {"anchorId": anchor, "spanStart": 6, "spanLength": 7}),
        ])
        assert not result.success and result.rolled_back
        assert session.grep("Before Section 4 after")


def test_replace_field_with_plain_text_and_reject_it(tour_plan_bytes: bytes):
    settings = DocxSessionSettings(tracked_changes=TrackedChangeMode.RENDER_INLINE)
    with open_session(_source(tour_plan_bytes), settings) as session:
        anchor = session.grep("Section 4")[0].enclosing_anchor.id
        result = session.execute_batch([
            MutationBatchStep("delete_field_at_span", {"anchorId": anchor, "spanStart": 7, "spanLength": 9}),
            MutationBatchStep("replace_text_at_span", {"anchorId": anchor, "spanStart": 7, "spanLength": 0, "replace": "Section 5"}),
        ])
        assert result.success
        content = session.save()
    with open_session(content) as accepted:
        assert accepted.accept_all_revisions().success
        assert accepted.grep("Before Section 5 after")
    with open_session(content) as rejected:
        assert rejected.reject_all_revisions().success
        assert rejected.grep("Before Section 4 after")
