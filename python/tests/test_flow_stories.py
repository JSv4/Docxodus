"""Effective default stories must cross both public HTML conversion entry points."""
from io import BytesIO
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import pytest

from docx_scalpel import HtmlOptions, convert_docx_to_html, open_session


@pytest.mark.parametrize("empty_final", [False, True])
def test_flow_inherits_the_nearest_default_story(empty_final: bool) -> None:
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    types = "http://schemas.openxmlformats.org/package/2006/content-types"
    rels = "http://schemas.openxmlformats.org/package/2006/relationships"
    parts = {"word/styles.xml": f'<w:styles xmlns:w="{w}"/>',
             "word/settings.xml": f'<w:settings xmlns:w="{w}"/>'}
    overrides = {"document": "document.main", "styles": "styles", "settings": "settings"}
    relationships = '<Relationship Id="styles" Type="' + r + '/styles" Target="styles.xml"/>'
    relationships += f'<Relationship Id="settings" Type="{r}/settings" Target="settings.xml"/>'
    paragraphs = ""
    for name in ("Old", "Current"):
        references = ""
        for kind in ("header", "footer"):
            part = kind + name
            tag = "hdr" if kind == "header" else "ftr"
            parts[f"word/{part}.xml"] = f'<w:{tag} xmlns:w="{w}"><w:p><w:r><w:t>{name} {kind}</w:t></w:r></w:p></w:{tag}>'
            overrides[part] = kind
            relationships += f'<Relationship Id="{part}" Type="{r}/{kind}" Target="{part}.xml"/>'
            references += f'<w:{kind}Reference w:type="default" r:id="{part}"/>'
        paragraphs += f'<w:p><w:pPr><w:sectPr>{references}</w:sectPr></w:pPr><w:r><w:t>{name} section</w:t></w:r></w:p>'
    final = ('<w:sectPr><w:headerReference w:type="default" r:id="headerOld"/>'
             '<w:footerReference w:type="default" r:id="footerOld"/></w:sectPr>' if empty_final else
             '<w:p><w:r><w:t>Final section</w:t></w:r></w:p><w:sectPr/>')
    parts["word/document.xml"] = f'<w:document xmlns:w="{w}" xmlns:r="{r}"><w:body>{paragraphs}{final}</w:body></w:document>'
    parts["word/_rels/document.xml.rels"] = f'<Relationships xmlns="{rels}">{relationships}</Relationships>'
    parts["_rels/.rels"] = f'<Relationships xmlns="{rels}"><Relationship Id="document" Type="{r}/officeDocument" Target="word/document.xml"/></Relationships>'
    parts["[Content_Types].xml"] = f'<Types xmlns="{types}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' + "".join(
        f'<Override PartName="/word/{name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>'
        for name, kind in overrides.items()) + '</Types>'
    stream = BytesIO()
    with ZipFile(stream, "w") as package:
        for name, value in parts.items():
            package.writestr(name, value)
    data = stream.getvalue()
    with open_session(data) as session:
        for convert in (session.to_html, lambda options: convert_docx_to_html(data, options)):
            html = ET.fromstring(convert(HtmlOptions(render_headers_and_footers=True)))
            for kind in ("header", "footer"):
                stories = html.findall(f".//{{http://www.w3.org/1999/xhtml}}{kind}")
                assert len(stories) == 1
                assert "".join(stories[0].itertext()).strip() == ("Old " if empty_final else "Current ") + kind
