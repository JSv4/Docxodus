"""Writes line-baselines.docx, the Word-reference fixture for exported baseline placement (issue #908).

Seven cases, one per page, each a paragraph several lines long that starts at the top margin, so the
first baseline sits a measurable distance below the margin and later lines give the pitch:

  CASE0..2  11 pt Calibri, w:line 240 / 276 / 480 (single, 1.15, double), w:lineRule auto
  CASE3..5  12 pt Arial,   the same three spacings
  CASE6     11 pt Calibri at 1.15 with a raised run (w:position 6 = 3 pt up) on its first line

The zip is written with fixed timestamps so the file's SHA-256 is stable. Re-run only to change the
fixture, and re-capture the Word measurements when you do.
"""
import zipfile
from pathlib import Path

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
SENTENCE = "The quick brown fox jumps over the lazy dog while the editor counts every line. "
CASES = [
    ("Calibri", 22, 240), ("Calibri", 22, 276), ("Calibri", 22, 480),
    ("Arial", 24, 240), ("Arial", 24, 276), ("Arial", 24, 480),
]


def run(font, half_points, text, position=None):
    pos = f'<w:position w:val="{position}"/>' if position is not None else ""
    return (f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>{pos}'
            f'<w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/></w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t></w:r>')


def paragraph(index, font, half_points, line, raised=False):
    # The paragraph mark carries the runs' font too, as Word writes it, so the paragraph's own line
    # height comes from the same font as its text.
    ppr = (('<w:pageBreakBefore/>' if index > 0 else '')
           + f'<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="auto"/>'
           + f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
           + f'<w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/></w:rPr>')
    body = run(font, half_points, f"CASE{index} ")
    if raised:
        body += run(font, half_points, "raised", position=6) + run(font, half_points, " " + SENTENCE * 4)
    else:
        body += run(font, half_points, SENTENCE * 4)
    return f"<w:p><w:pPr>{ppr}</w:pPr>{body}</w:p>"


paragraphs = [paragraph(i, *case) for i, case in enumerate(CASES)]
paragraphs.append(paragraph(6, "Calibri", 22, 276, raised=True))

parts = {
    "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>""",
    "_rels/.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>""",
    "word/_rels/document.xml.rels": """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>""",
    "word/styles.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>""",
    "word/settings.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="{W}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>""",
    "word/document.xml": f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}"><w:body>{''.join(paragraphs)}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>""",
}

out = Path(__file__).with_name("line-baselines.docx")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    for name, data in parts.items():
        info = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        z.writestr(info, data.encode("utf-8"))
print(out)
