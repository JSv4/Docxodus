"""Writes line-baselines.docx, the Word-reference fixture for exported baseline placement (issue #908).

Ten cases, one per page, each a paragraph several lines long that starts at the top margin, so the
first baseline sits a measurable distance below the margin and later lines give the pitch:

  CASE0..2  11 pt Calibri, w:line 240 / 276 / 480 (single, 1.15, double), w:lineRule auto
  CASE3..5  12 pt Arial,   the same three spacings
  CASE6     11 pt Calibri at 1.15 with a raised run (w:position 6 = 3 pt up) on its first line
  CASE7     11 pt Calibri at 1.15 with a lowered run (w:position -6 = 3 pt down) on its first line
  CASE8..9  10 pt Calibri runs under a 20 pt Calibri paragraph mark, single and 1.15 (issue #949),
            each followed on its page by a plain 10 pt paragraph (AFTER8, AFTER9)

It also writes line-baselines-plain-marks.docx: the same cases, CASE8 and CASE9 aside, with paragraph marks that carry no run
properties, so each mark keeps the default 11 pt Calibri under its runs' font (issue #940). Word's baselines do
not depend on the mark here: Word for the web gave identical baselines for a plain-mark version of
line-baselines.docx and for the current one.

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


def paragraph(index, font, half_points, line, position=None, plain_mark=False, mark_half_points=None):
    # The paragraph mark carries the runs' font too, as Word writes it, so the paragraph's own line
    # height comes from the same font as its text. A plain mark has no run properties at all.
    mark = ('' if plain_mark else
            f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
            + f'<w:sz w:val="{mark_half_points or half_points}"/><w:szCs w:val="{mark_half_points or half_points}"/></w:rPr>')
    ppr = (('<w:pageBreakBefore/>' if index > 0 else '')
           + f'<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="auto"/>'
           + mark)
    body = run(font, half_points, f"CASE{index} ")
    if position is not None:
        word = "raised" if position > 0 else "lowered"
        body += run(font, half_points, word, position=position) + run(font, half_points, " " + SENTENCE * 4)
    else:
        body += run(font, half_points, SENTENCE * 4)
    return f"<w:p><w:pPr>{ppr}</w:pPr>{body}</w:p>"


def follower(label, font, half_points, line):
    """A plain paragraph on the same page as the case before it, its mark the size of its runs."""
    mark = (f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
            + f'<w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/></w:rPr>')
    ppr = f'<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="auto"/>' + mark
    return f"<w:p><w:pPr>{ppr}</w:pPr>{run(font, half_points, label + SENTENCE)}</w:p>"


def parts(plain_marks):
    paragraphs = [paragraph(i, *case, plain_mark=plain_marks) for i, case in enumerate(CASES)]
    paragraphs.append(paragraph(6, "Calibri", 22, 276, position=6, plain_mark=plain_marks))
    paragraphs.append(paragraph(7, "Calibri", 22, 276, position=-6, plain_mark=plain_marks))
    # The mark is the point of CASE8 and CASE9, so the plain-mark variant leaves them out.
    # A plain 10 pt paragraph follows each, so the page shows how tall the tall-mark paragraph's last line is.
    if not plain_marks:
        for index, line in ((8, 240), (9, 276)):
            paragraphs.append(paragraph(index, "Calibri", 20, line, mark_half_points=40))
            paragraphs.append(follower(f"AFTER{index} ", "Calibri", 20, line))
    return {
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

for name, plain_marks in (("line-baselines.docx", False), ("line-baselines-plain-marks.docx", True)):
    out = Path(__file__).with_name(name)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for part, data in parts(plain_marks).items():
            info = zipfile.ZipInfo(part, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data.encode("utf-8"))
    print(out)
