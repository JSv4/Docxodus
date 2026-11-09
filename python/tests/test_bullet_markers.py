"""Original familiar bullets and a checkbox through both public HTML APIs."""
from io import BytesIO
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import pytest

from docx_scalpel import HtmlOptions, convert_docx_to_html, open_session


@pytest.mark.parametrize("family,glyph,css_type", [("Symbol", "\uf0b7", "disc"), ("Courier New", "o", "circle")])
def test_normalized_bullets_preserve_checkbox_content(family, glyph, css_type):
    w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rels = "http://schemas.openxmlformats.org/package/2006/relationships"
    definitions, paragraphs = "", ""
    for num, font, marker, text in [(1, family, glyph, "Harbor"), (2, "Liberation Serif", "☐", "Confirm")]:
        definitions += (f'<w:abstractNum w:abstractNumId="{num}"><w:lvl w:ilvl="0"><w:start w:val="1"/>'
                        f'<w:numFmt w:val="bullet"/><w:lvlText w:val="{marker}"/><w:pPr>'
                        '<w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr>'
                        f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}"/></w:rPr></w:lvl></w:abstractNum>')
        paragraphs += (f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num}"/></w:numPr>'
                       f'</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>')
    definitions += ''.join(f'<w:num w:numId="{n}"><w:abstractNumId w:val="{n}"/></w:num>' for n in (1, 2))
    parts = {
        "document": f'<w:document xmlns:w="{w}"><w:body>{paragraphs}<w:sectPr/></w:body></w:document>',
        "numbering": f'<w:numbering xmlns:w="{w}">{definitions}</w:numbering>',
        "settings": f'<w:settings xmlns:w="{w}"/>',
        "styles": f'<w:styles xmlns:w="{w}"><w:docDefaults><w:rPrDefault><w:rPr>'
                  '<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/>'
                  '</w:rPr></w:rPrDefault></w:docDefaults></w:styles>',
    }
    stream = BytesIO()
    with ZipFile(stream, "w") as package:
        for name, content in parts.items():
            package.writestr(f'word/{name}.xml', content)
        package.writestr('_rels/.rels', f'<Relationships xmlns="{rels}"><Relationship Id="document" Type="{r}/officeDocument" Target="word/document.xml"/></Relationships>')
        package.writestr('word/_rels/document.xml.rels', f'<Relationships xmlns="{rels}">' + ''.join(
            f'<Relationship Id="{name}" Type="{r}/{name}" Target="{name}.xml"/>' for name in parts if name != 'document') + '</Relationships>')
        package.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                         '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' + ''.join(
                             f'<Override PartName="/word/{name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{"document.main" if name == "document" else name}+xml"/>' for name in parts) + '</Types>')
    data = stream.getvalue()
    with open_session(data) as session:
        for convert in (session.to_html, lambda options: convert_docx_to_html(data, options)):
            for normalize in (False, True):
                root = ET.fromstring(convert(HtmlOptions(semantic_lists=True, normalize_bullet_markers=normalize,
                                                        fabricate_css_classes=False)))
                lists = root.findall('.//{http://www.w3.org/1999/xhtml}ul')
                assert len(lists) == 2
                assert f'list-style-type: {css_type if normalize else "none"};' in lists[0].attrib['style']
                assert 'list-style-type: none;' in lists[1].attrib['style']
                assert bool(lists[0][0].findall('./*[@data-list-marker]')) is not normalize
                assert lists[1][0].findall('./*[@data-list-marker]')
                assert '☐' in ''.join(lists[1].itertext())
                if normalize:
                    assert ''.join(lists[0][0].itertext()).strip() == 'Harbor'
