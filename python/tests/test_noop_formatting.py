"""A no-op formatting call does not split an inherited tracked revision."""
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from docx_scalpel import DocxSessionSettings, FormatOp, TrackedChangeMode, open_session

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
                    <w:r><w:rPr><w:u w:val="single"/><w:rPrChange w:id="9" w:author="Counsel" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>Earlier formatting</w:t></w:r>
                </w:p>''')
                body.insert(0, paragraph)
                content = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            output.writestr(entry, content)
    return stream.getvalue()


def test_noop_inside_existing_format_revision_preserves_version(tour_plan_bytes: bytes):
    settings = DocxSessionSettings(tracked_changes=TrackedChangeMode.RENDER_INLINE)
    with open_session(_source(tour_plan_bytes), settings) as session:
        match = session.grep("Earlier")[0]
        before = session.list_revisions()
        version = session.get_version()
        from docx_scalpel import CharSpan
        result = session.apply_format(match.enclosing_anchor.id, CharSpan(2, 5), FormatOp(underline=True))
        assert result.success
        assert session.list_revisions() == before
        assert session.get_version() == version
        assert not session.undo()
