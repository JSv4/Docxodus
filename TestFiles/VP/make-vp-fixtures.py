#!/usr/bin/env python3
"""Generate the VP visual-parity corpus fixtures (issue #400).

The visual-parity corpus (npm/tests/visual-parity/corpus.ts) only admits Git-tracked
fixtures whose worktree blob matches HEAD, so these documents are committed as bytes.
This script is the reproducible source for the authored ones — rerunning it must
produce byte-identical output (fixed zip timestamps, fixed member order, no RNG):

  VP001-Chart-Stacked-Column.docx  HC043-Chart.docx with the clustered column chart
                                   regrouped as stacked (grouping + overlap 100).
  VP002-Image-Wrap-Tight.docx      HC042-Image-Png.docx with the inline picture
                                   re-anchored as a floating, tight-wrapped image
                                   surrounded by enough text to wrap.
  VP003-Two-Column-Section.docx    Authored: single-column title section followed by
                                   a continuous two-column (`w:cols w:num="2"`) section.
  VP004-Legal-Contract.docx        Authored: realistic services agreement — cached TOC
                                   with hyperlink entries and PAGEREF fields, multilevel
                                   heading numbering (1. / 1.1), (a)/(i) sub-lists,
                                   bookmarks with cached REF cross-references, and a
                                   borderless signature table.

Run from the repository root:  python3 TestFiles/VP/make-vp-fixtures.py
"""

import io
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
VP = ROOT / 'TestFiles' / 'VP'

# Fixed timestamp so regenerating cannot churn the committed bytes.
ZIP_DATE = (1980, 1, 1, 0, 0, 0)


def write_docx(path: Path, members: dict[str, bytes]) -> None:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name in members:  # dict order is the member order — keep it stable
            info = zipfile.ZipInfo(name, date_time=ZIP_DATE)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            archive.writestr(info, members[name])
    path.write_bytes(buffer.getvalue())


def read_docx(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


# ---------------------------------------------------------------------------
# VP001 — stacked column chart, derived from the tracked clustered fixture
# ---------------------------------------------------------------------------

def make_vp001() -> None:
    members = read_docx(ROOT / 'TestFiles' / 'HC043-Chart.docx')
    chart = members['word/charts/chart1.xml'].decode('utf-8')
    if '<c:grouping val="clustered"/>' not in chart or '<c:overlap val="-27"/>' not in chart:
        raise SystemExit('HC043-Chart.docx chart1.xml no longer matches the expected clustered shape')
    chart = chart.replace('<c:grouping val="clustered"/>', '<c:grouping val="stacked"/>')
    chart = chart.replace('<c:overlap val="-27"/>', '<c:overlap val="100"/>')
    members['word/charts/chart1.xml'] = chart.encode('utf-8')
    write_docx(VP / 'VP001-Chart-Stacked-Column.docx', members)


# ---------------------------------------------------------------------------
# VP002 — floating, tight-wrapped picture, derived from the tracked inline one
# ---------------------------------------------------------------------------

FILLER = (
    'Video provides a powerful way to help you prove your point. When you click Online '
    'Video, you can paste in the embed code for the video you want to add. You can also '
    'type a keyword to search online for the video that best fits your document. To make '
    'your document look professionally produced, Word provides header, footer, cover page, '
    'and text box designs that complement each other. For example, you can add a matching '
    'cover page, header, and sidebar. Click Insert and then choose the elements you want '
    'from the different galleries.'
)


def make_vp002() -> None:
    members = read_docx(ROOT / 'TestFiles' / 'HC042-Image-Png.docx')
    document = members['word/document.xml'].decode('utf-8')

    # Three times the inline extent so the wrap is visually load-bearing.
    cx, cy = 419158 * 3, 314369 * 3
    anchor = (
        '<w:drawing>'
        '<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        'relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:wrapTight wrapText="bothSides">'
        '<wp:wrapPolygon edited="0">'
        '<wp:start x="0" y="0"/>'
        '<wp:lineTo x="0" y="21600"/>'
        '<wp:lineTo x="21600" y="21600"/>'
        '<wp:lineTo x="21600" y="0"/>'
        '<wp:lineTo x="0" y="0"/>'
        '</wp:wrapPolygon>'
        '</wp:wrapTight>'
        '<wp:docPr id="1" name="Picture 1"/>'
        '<wp:cNvGraphicFramePr>'
        '<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:nvPicPr><pic:cNvPr id="1" name="Capture.PNG"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rId4"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        '<pic:spPr><a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        '</pic:pic></a:graphicData></a:graphic>'
        '</wp:anchor>'
        '</w:drawing>'
    )

    body = (
        '<w:body>'
        f'<w:p><w:r><w:rPr><w:noProof/></w:rPr>{anchor}</w:r>'
        f'<w:r><w:t xml:space="preserve">{FILLER} </w:t></w:r>'
        f'<w:r><w:t xml:space="preserve">{FILLER}</w:t></w:r></w:p>'
        f'<w:p><w:r><w:t xml:space="preserve">{FILLER}</w:t></w:r></w:p>'
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>'
        '</w:body>'
    )
    document = re.sub(r'<w:body>.*</w:body>', body, document, count=1, flags=re.S)
    members['word/document.xml'] = document.encode('utf-8')
    write_docx(VP / 'VP002-Image-Wrap-Tight.docx', members)


# ---------------------------------------------------------------------------
# Shared scaffolding for the authored (from-scratch) documents
# ---------------------------------------------------------------------------

W_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)

CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>'
    '{numbering_override}'
    '</Types>'
)

ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="word/document.xml"/>'
    '</Relationships>'
)

DOCUMENT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>'
    '{numbering_rel}'
    '</Relationships>'
)

SETTINGS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<w:settings {W_NS}>'
    '<w:zoom w:percent="100"/>'
    '<w:defaultTabStop w:val="720"/>'
    '<w:characterSpacingControl w:val="doNotCompress"/>'
    '</w:settings>'
)


def font_table(fonts: list[str]) -> str:
    entries = ''.join(f'<w:font w:name="{name}"/>' for name in fonts)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:fonts {W_NS}>{entries}</w:fonts>'
    )


def styles_xml(default_font: str, default_size: int, extra_styles: str) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles {W_NS}>'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        f'<w:rFonts w:ascii="{default_font}" w:hAnsi="{default_font}" w:cs="{default_font}"/>'
        f'<w:sz w:val="{default_size}"/><w:szCs w:val="{default_size}"/>'
        '<w:lang w:val="en-US"/>'
        '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/><w:qFormat/></w:style>'
        '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
        '<w:name w:val="Default Paragraph Font"/></w:style>'
        f'{extra_styles}'
        '</w:styles>'
    )


def scratch_docx(path: Path, document: str, styles: str, fonts: list[str],
                 numbering: str | None = None) -> None:
    members: dict[str, bytes] = {
        '[Content_Types].xml': CONTENT_TYPES.format(
            numbering_override=(
                '<Override PartName="/word/numbering.xml" ContentType='
                '"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
                if numbering else ''
            )
        ).encode('utf-8'),
        '_rels/.rels': ROOT_RELS.encode('utf-8'),
        'word/_rels/document.xml.rels': DOCUMENT_RELS.format(
            numbering_rel=(
                '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/'
                'officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'
                if numbering else ''
            )
        ).encode('utf-8'),
        'word/document.xml': document.encode('utf-8'),
        'word/styles.xml': styles.encode('utf-8'),
        'word/settings.xml': SETTINGS.encode('utf-8'),
        'word/fontTable.xml': font_table(fonts).encode('utf-8'),
    }
    if numbering:
        members['word/numbering.xml'] = numbering.encode('utf-8')
    write_docx(path, members)


def para(text: str, ppr: str = '', rpr: str = '') -> str:
    ppr_xml = f'<w:pPr>{ppr}</w:pPr>' if ppr else ''
    rpr_xml = f'<w:rPr>{rpr}</w:rPr>' if rpr else ''
    return f'<w:p>{ppr_xml}<w:r>{rpr_xml}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


# ---------------------------------------------------------------------------
# VP003 — single-column title, then a continuous two-column section
# ---------------------------------------------------------------------------

COLUMN_TEXT = [
    'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Maecenas porttitor congue '
    'massa. Fusce posuere, magna sed pulvinar ultricies, purus lectus malesuada libero, '
    'sit amet commodo magna eros quis urna.',
    'Nunc viverra imperdiet enim. Fusce est. Vivamus a tellus. Pellentesque habitant morbi '
    'tristique senectus et netus et malesuada fames ac turpis egestas. Proin pharetra '
    'nonummy pede. Mauris et orci.',
    'Aenean nec lorem. In porttitor. Donec laoreet nonummy augue. Suspendisse dui purus, '
    'scelerisque at, vulputate vitae, pretium mattis, nunc. Mauris eget neque at sem '
    'venenatis eleifend. Ut nonummy.',
    'Fusce aliquet pede non pede. Suspendisse dapibus lorem pellentesque magna. Integer '
    'nulla. Donec blandit feugiat ligula. Donec hendrerit, felis et imperdiet euismod, '
    'purus ipsum pretium metus, in lacinia nulla nisl eget sapien.',
    'Donec ut est in lectus consequat consequat. Etiam eget dui. Aliquam erat volutpat. '
    'Sed at lorem in nunc porta tristique. Proin nec augue. Quisque aliquam tempor magna. '
    'Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis '
    'egestas.',
    'Nunc ac magna. Maecenas odio dolor, vulputate vel, auctor ac, accumsan id, felis. '
    'Pellentesque cursus sagittis felis. Pellentesque porttitor, velit lacinia egestas '
    'auctor, diam eros tempus arcu, nec vulputate augue magna vel risus.',
    'Cras non magna vel ante adipiscing rhoncus. Vivamus a mi. Morbi neque. Aliquam erat '
    'volutpat. Integer ultrices lobortis eros. Pellentesque habitant morbi tristique '
    'senectus et netus et malesuada fames ac turpis egestas.',
    'Proin semper, ante vitae sollicitudin posuere, metus quam iaculis nibh, vitae '
    'scelerisque nunc massa eget pede. Sed velit urna, interdum vel, ultricies vel, '
    'faucibus at, quam. Donec elit est, consectetuer eget, consequat quis, tempus quis, '
    'wisi.',
]


def make_vp003() -> None:
    heading_style = (
        '<w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>'
        '<w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
        '</w:style>'
    )
    title_section = (
        '<w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/>'
        # This sectPr ends the single-column title section.
        '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>'
        '</w:pPr>'
        '<w:r><w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>'
        '<w:t>The Docxodus Gazette</w:t></w:r>'
    )
    body_paragraphs = ''.join(
        para(text, ppr='<w:spacing w:after="120"/><w:jc w:val="both"/>')
        for text in COLUMN_TEXT
    )
    column_heading = para('A Two-Column Layout Exercise', ppr='<w:pStyle w:val="Heading1"/>')
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document {W_NS}><w:body>'
        f'<w:p>{title_section}</w:p>'
        f'{column_heading}'
        f'{body_paragraphs}'
        # The trailing sectPr is the two-column section; continuous so it shares the page.
        '<w:sectPr><w:type w:val="continuous"/>'
        '<w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:num="2" w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>'
        '</w:body></w:document>'
    )
    scratch_docx(VP / 'VP003-Two-Column-Section.docx', document,
                 styles_xml('Calibri', 22, heading_style), ['Calibri'])


# ---------------------------------------------------------------------------
# VP004 — realistic legal contract: numbering + cached TOC + cross-references
# ---------------------------------------------------------------------------

def contract_numbering() -> str:
    # abstractNum 0: multilevel heading numbering (1. / 1.1) bound to Heading1/Heading2.
    # abstractNum 1: (a) / (i) sub-clause lists used in the body.
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:numbering {W_NS}>'
        '<w:abstractNum w:abstractNumId="0">'
        '<w:multiLevelType w:val="multilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:pStyle w:val="Heading1"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="432" w:hanging="432"/></w:pPr></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/>'
        '<w:pStyle w:val="Heading2"/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="576" w:hanging="576"/></w:pPr></w:lvl>'
        '</w:abstractNum>'
        '<w:abstractNum w:abstractNumId="1">'
        '<w:multiLevelType w:val="multilevel"/>'
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/>'
        '<w:lvlText w:val="(%1)"/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="1080" w:hanging="360"/></w:pPr></w:lvl>'
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/>'
        '<w:lvlText w:val="(%2)"/><w:lvlJc w:val="left"/>'
        '<w:pPr><w:ind w:left="1800" w:hanging="360"/></w:pPr></w:lvl>'
        '</w:abstractNum>'
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
        '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
        # Fresh instance per (a)-list so each restarts at (a).
        '<w:num w:numId="3"><w:abstractNumId w:val="1"/>'
        '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>'
        '<w:num w:numId="4"><w:abstractNumId w:val="1"/>'
        '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>'
        '</w:numbering>'
    )


def contract_styles() -> str:
    return (
        '<w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/>'
        '<w:pPr><w:keepNext/><w:numPr><w:numId w:val="1"/></w:numPr>'
        '<w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>'
        '<w:rPr><w:b/><w:caps/></w:rPr>'
        '</w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading2">'
        '<w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/>'
        '<w:pPr><w:keepNext/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>'
        '<w:spacing w:before="120" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr>'
        '<w:rPr><w:b/></w:rPr>'
        '</w:style>'
        '<w:style w:type="paragraph" w:styleId="TOC1">'
        '<w:name w:val="toc 1"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:after="100"/></w:pPr>'
        '</w:style>'
        '<w:style w:type="paragraph" w:styleId="TOC2">'
        '<w:name w:val="toc 2"/><w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:after="100"/><w:ind w:left="220"/></w:pPr>'
        '</w:style>'
        '<w:style w:type="character" w:styleId="Hyperlink">'
        '<w:name w:val="Hyperlink"/>'
        '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>'
        '</w:style>'
    )


SECTIONS: list[tuple[int, str, str]] = [
    # (level, bookmark, heading text) — page numbers in the cached TOC below.
    (1, '_Toc400000001', 'Definitions'),
    (1, '_Toc400000002', 'Services'),
    (2, '_Toc400000003', 'Statements of Work'),
    (2, '_Toc400000004', 'Change Orders'),
    (1, '_Toc400000005', 'Fees and Payment'),
    (2, '_Toc400000006', 'Fees'),
    (2, '_Toc400000007', 'Invoicing; Late Payment'),
    (1, '_Toc400000008', 'Term and Termination'),
    (1, '_Toc400000009', 'Confidentiality'),
    (1, '_Toc400000010', 'Limitation of Liability'),
    (1, '_Toc400000011', 'General Provisions'),
]

TOC_PAGES = ['1', '2', '2', '2', '2', '2', '2', '3', '3', '3', '3']

TOC_NUMBERS = ['1.', '2.', '2.1', '2.2', '3.', '3.1', '3.2', '4.', '5.', '6.', '7.']


def toc_entry(level: int, anchor: str, number: str, text: str, page: str) -> str:
    return (
        f'<w:p><w:pPr><w:pStyle w:val="TOC{level}"/>'
        '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>'
        '<w:rPr><w:noProof/></w:rPr></w:pPr>'
        f'<w:hyperlink w:anchor="{anchor}" w:history="1">'
        '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr>'
        f'<w:t xml:space="preserve">{number} {text}</w:t></w:r>'
        '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r>'
        '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr>'
        f'<w:instrText xml:space="preserve"> PAGEREF {anchor} \\h </w:instrText></w:r>'
        '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:t>{page}</w:t></w:r>'
        '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
        '</w:hyperlink></w:p>'
    )


def heading(level: int, anchor: str, text: str, bookmark_id: int,
            named_bookmark: str | None = None) -> str:
    named_start = named_end = ''
    if named_bookmark:
        named_start = f'<w:bookmarkStart w:id="{bookmark_id + 100}" w:name="{named_bookmark}"/>'
        named_end = f'<w:bookmarkEnd w:id="{bookmark_id + 100}"/>'
    return (
        f'<w:p><w:pPr><w:pStyle w:val="Heading{level}"/></w:pPr>'
        f'<w:bookmarkStart w:id="{bookmark_id}" w:name="{anchor}"/>{named_start}'
        f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
        f'{named_end}<w:bookmarkEnd w:id="{bookmark_id}"/></w:p>'
    )


def cross_ref(bookmark: str, cached: str) -> str:
    """A cached REF field, as Word writes Insert > Cross-reference (number only)."""
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve"> REF {bookmark} \\r \\h </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r><w:t>{cached}</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


def body_para(runs: str, extra_ppr: str = '') -> str:
    return (
        f'<w:p><w:pPr><w:spacing w:after="120"/><w:jc w:val="both"/>{extra_ppr}</w:pPr>'
        f'{runs}</w:p>'
    )


def text_run(text: str, bold: bool = False) -> str:
    rpr = '<w:rPr><w:b/></w:rPr>' if bold else ''
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def list_para(text: str, num_id: int, ilvl: int = 0) -> str:
    return (
        f'<w:p><w:pPr><w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr>'
        '<w:spacing w:after="60"/><w:jc w:val="both"/></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )


def make_vp004() -> None:
    toc_field_open = (
        '<w:p><w:pPr><w:pStyle w:val="TOC1"/>'
        '<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>'
        '<w:rPr><w:noProof/></w:rPr></w:pPr>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-2" \\h \\z \\u </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    )
    # First entry lives in the paragraph that opens the TOC field, as Word writes it.
    first = SECTIONS[0]
    toc_first_entry = (
        f'<w:hyperlink w:anchor="{first[1]}" w:history="1">'
        '<w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr>'
        f'<w:t xml:space="preserve">{TOC_NUMBERS[0]} {first[2]}</w:t></w:r>'
        '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r>'
        '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr>'
        f'<w:instrText xml:space="preserve"> PAGEREF {first[1]} \\h </w:instrText></w:r>'
        '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:t>{TOC_PAGES[0]}</w:t></w:r>'
        '<w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>'
        '</w:hyperlink></w:p>'
    )
    toc_rest = ''.join(
        toc_entry(level, anchor, number, text, page)
        for (level, anchor, text), number, page in zip(SECTIONS[1:], TOC_NUMBERS[1:], TOC_PAGES[1:])
    )
    toc_close = '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'

    h = {anchor: (level, text) for level, anchor, text in SECTIONS}
    body_parts: list[str] = []
    body_parts.append(
        '<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr>'
        '<w:r><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
        '<w:t>MASTER SERVICES AGREEMENT</w:t></w:r></w:p>'
    )
    body_parts.append(body_para(
        text_run('This Master Services Agreement (this ')
        + text_run('"Agreement"', bold=True)
        + text_run(') is entered into as of March 1, 2026 (the ')
        + text_run('"Effective Date"', bold=True)
        + text_run(') by and between Meridian Consulting Group LLC, a Delaware limited '
                   'liability company (')
        + text_run('"Provider"', bold=True)
        + text_run('), and Atlas Manufacturing Corporation, a Michigan corporation (')
        + text_run('"Client"', bold=True)
        + text_run('). Provider and Client are each a ')
        + text_run('"Party"', bold=True)
        + text_run(' and together the ')
        + text_run('"Parties"', bold=True)
        + text_run('.')
    ))
    body_parts.append(para('TABLE OF CONTENTS',
                           ppr='<w:jc w:val="center"/><w:spacing w:before="240" w:after="120"/>',
                           rpr='<w:b/>'))
    body_parts.append(toc_field_open + toc_first_entry + toc_rest + toc_close)

    body_parts.append(heading(1, '_Toc400000001', 'Definitions', 1))
    body_parts.append(body_para(text_run(
        'As used in this Agreement, the following terms have the meanings set forth below. '
        'Other capitalized terms are defined where they first appear.')))
    body_parts.append(list_para(
        '"Confidential Information" means all non-public information disclosed by a Party '
        'in connection with this Agreement, whether oral or written, that is designated as '
        'confidential or that reasonably should be understood to be confidential.', 3))
    body_parts.append(list_para(
        '"Deliverables" means the work product identified in a Statement of Work to be '
        'delivered by Provider to Client.', 3))
    body_parts.append(list_para(
        '"Services" means the professional services described in a Statement of Work, '
        'including:', 3))
    body_parts.append(list_para('advisory and analysis services;', 3, ilvl=1))
    body_parts.append(list_para('implementation and configuration services; and', 3, ilvl=1))
    body_parts.append(list_para('training and knowledge-transfer services.', 3, ilvl=1))
    body_parts.append(list_para(
        '"Statement of Work" or "SOW" means a written statement of work executed by both '
        'Parties that references this Agreement.', 3))

    body_parts.append(heading(1, '_Toc400000002', 'Services', 2))
    body_parts.append(heading(2, '_Toc400000003', 'Statements of Work', 3, 'Ref_SOW'))
    body_parts.append(body_para(
        text_run('Provider shall perform the Services described in each SOW in a '
                 'professional and workmanlike manner, in accordance with generally '
                 'accepted industry standards. Each SOW is incorporated into this '
                 'Agreement by reference. In the event of a conflict between this '
                 'Agreement and a SOW, this Agreement controls unless the SOW expressly '
                 'states otherwise.')))
    body_parts.append(heading(2, '_Toc400000004', 'Change Orders', 4))
    body_parts.append(body_para(
        text_run('Either Party may propose changes to a SOW. No change is binding until '
                 'set forth in a written change order signed by both Parties. Provider '
                 'may equitably adjust the fees payable under Section ')
        + cross_ref('Ref_Fees', '3.1')
        + text_run(' to reflect any agreed change in scope, in accordance with the '
                   'procedures in Section ')
        + cross_ref('Ref_SOW', '2.1')
        + text_run('.')))

    body_parts.append(heading(1, '_Toc400000005', 'Fees and Payment', 5))
    body_parts.append(heading(2, '_Toc400000006', 'Fees', 6, 'Ref_Fees'))
    body_parts.append(body_para(
        text_run('Client shall pay Provider the fees set forth in the applicable SOW. '
                 'Except as expressly stated in a SOW, fees are exclusive of taxes, and '
                 'Client is responsible for all sales, use, and excise taxes other than '
                 'taxes on Provider’s income.')))
    body_parts.append(heading(2, '_Toc400000007', 'Invoicing; Late Payment', 7, 'Ref_Invoicing'))
    body_parts.append(body_para(
        text_run('Provider shall invoice Client monthly in arrears. Undisputed amounts '
                 'are due within thirty (30) days after the invoice date. Late payments '
                 'bear interest at the lesser of one percent (1%) per month or the '
                 'maximum rate permitted by law.')))

    body_parts.append(heading(1, '_Toc400000008', 'Term and Termination', 8, 'Ref_Term'))
    body_parts.append(body_para(
        text_run('This Agreement begins on the Effective Date and continues for an '
                 'initial term of two (2) years, renewing automatically for successive '
                 'one (1) year terms unless either Party gives notice of non-renewal at '
                 'least sixty (60) days before the end of the then-current term. Either '
                 'Party may terminate this Agreement:')))
    body_parts.append(list_para(
        'for material breach, if the breach is not cured within thirty (30) days after '
        'written notice describing the breach in reasonable detail;', 4))
    body_parts.append(list_para(
        'immediately, if the other Party becomes insolvent, makes an assignment for the '
        'benefit of creditors, or becomes subject to bankruptcy proceedings not dismissed '
        'within sixty (60) days; or', 4))
    body_parts.append(list_para(
        'for convenience, on ninety (90) days’ written notice, subject to payment of '
        'fees for Services performed through the effective date of termination as '
        'provided in Section ', 4).replace(
            '</w:t></w:r></w:p>',
            '</w:t></w:r>' + cross_ref('Ref_Invoicing', '3.2')
            + '<w:r><w:t xml:space="preserve">.</w:t></w:r></w:p>'))

    body_parts.append(heading(1, '_Toc400000009', 'Confidentiality', 9, 'Ref_Confidentiality'))
    body_parts.append(body_para(
        text_run('Each Party shall protect the other Party’s Confidential Information '
                 'with at least the degree of care it uses for its own confidential '
                 'information, and no less than reasonable care. Confidential Information '
                 'may be used only to perform or receive the Services. The obligations in '
                 'this Section ')
        + cross_ref('Ref_Confidentiality', '5')
        + text_run(' survive termination of this Agreement for five (5) years, except for '
                   'trade secrets, which are protected for as long as they remain trade '
                   'secrets under applicable law.')))

    body_parts.append(heading(1, '_Toc400000010', 'Limitation of Liability', 10, 'Ref_Liability'))
    body_parts.append(body_para(
        text_run('EXCEPT FOR BREACHES OF SECTION ')
        + cross_ref('Ref_Confidentiality', '5')
        + text_run(' OR A PARTY’S INDEMNIFICATION OBLIGATIONS, NEITHER PARTY IS '
                   'LIABLE FOR ANY INDIRECT, INCIDENTAL, SPECIAL, CONSEQUENTIAL, OR '
                   'PUNITIVE DAMAGES, AND EACH PARTY’S TOTAL AGGREGATE LIABILITY '
                   'UNDER THIS AGREEMENT IS LIMITED TO THE FEES PAID OR PAYABLE BY CLIENT '
                   'IN THE TWELVE (12) MONTHS PRECEDING THE EVENT GIVING RISE TO THE '
                   'CLAIM.')))

    body_parts.append(heading(1, '_Toc400000011', 'General Provisions', 11))
    body_parts.append(body_para(
        text_run('Notices. All notices under this Agreement must be in writing and are '
                 'deemed given when delivered personally, one (1) business day after '
                 'deposit with a nationally recognized overnight courier, or three (3) '
                 'business days after mailing by certified mail, return receipt '
                 'requested, to the addresses set forth in the applicable SOW.')))
    body_parts.append(body_para(
        text_run('Governing Law. This Agreement is governed by the laws of the State of '
                 'Delaware, without regard to its conflict of laws principles. The '
                 'Parties consent to the exclusive jurisdiction of the state and federal '
                 'courts located in Wilmington, Delaware. Termination rights are as set '
                 'out in Section ')
        + cross_ref('Ref_Term', '4')
        + text_run('; limitations of liability are as set out in Section ')
        + cross_ref('Ref_Liability', '6')
        + text_run('.')))
    body_parts.append(body_para(
        text_run('IN WITNESS WHEREOF, the Parties have executed this Agreement as of the '
                 'Effective Date.'),
        extra_ppr='<w:spacing w:before="240"/>'))

    signature_cell = (
        '<w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr>'
        '{content}</w:tc>'
    )
    def sig_lines(party: str, name: str, title: str) -> str:
        return (
            para(party, rpr='<w:b/>')
            + para('')
            + para('By: _______________________________')
            + para(f'Name: {name}')
            + para(f'Title: {title}')
        )
    signature_table = (
        '<w:tbl><w:tblPr><w:tblW w:w="9350" w:type="dxa"/>'
        '<w:tblLayout w:type="fixed"/><w:tblCellMar>'
        '<w:left w:w="0" w:type="dxa"/><w:right w:w="115" w:type="dxa"/>'
        '</w:tblCellMar></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="4675"/><w:gridCol w:w="4675"/></w:tblGrid>'
        '<w:tr>'
        + signature_cell.format(content=sig_lines(
            'MERIDIAN CONSULTING GROUP LLC', 'Dana Whitfield', 'Managing Member'))
        + signature_cell.format(content=sig_lines(
            'ATLAS MANUFACTURING CORPORATION', 'Jordan Alvarez', 'Chief Operating Officer'))
        + '</w:tr></w:tbl>'
        '<w:p/>'
    )
    body_parts.append(signature_table)

    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document {W_NS}><w:body>'
        + ''.join(body_parts)
        + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="720" w:footer="720" w:gutter="0"/>'
        '<w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr>'
        '</w:body></w:document>'
    )
    scratch_docx(VP / 'VP004-Legal-Contract.docx', document,
                 styles_xml('Times New Roman', 24, contract_styles()),
                 ['Times New Roman'], numbering=contract_numbering())


def main() -> None:
    VP.mkdir(parents=True, exist_ok=True)
    make_vp001()
    make_vp002()
    make_vp003()
    make_vp004()
    for name in sorted(VP.glob('VP0*.docx')):
        print(f'{name.relative_to(ROOT)}  {name.stat().st_size} bytes')


if __name__ == '__main__':
    main()
