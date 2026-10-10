"""Generate independent, deterministic DOCX fixtures from literal XML only."""
import argparse
import json
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def run(text):
    return '<w:r><w:t xml:space="preserve">' + escape(text) + '</w:t></w:r>'


def para(text, style=None, properties=""):
    properties = (f'<w:pStyle w:val="{style}"/>' if style else "") + properties
    return '<w:p>' + (f'<w:pPr>{properties}</w:pPr>' if properties else "") + run(text) + '</w:p>'


def font(name, size):
    return f'<w:rFonts w:ascii="{name}" w:hAnsi="{name}" w:eastAsia="{name}" w:cs="{name}"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'


def spacing(after, line=240):
    return f'<w:spacing w:before="0" w:after="{after}" w:line="{line}" w:lineRule="auto"/>'


def style(ident="Normal", name="Normal", run_props="", para_props="", default=True, extra=""):
    return (f'<w:style w:type="paragraph" w:styleId="{ident}"' + (' w:default="1"' if default else '') + '>'
            f'<w:name w:val="{name}"/>{extra}<w:pPr>{para_props}</w:pPr><w:rPr>{run_props}</w:rPr></w:style>')


def styles(items=None, default_run=None, default_para=None):
    if items is None:
        items = style(run_props=font("DejaVu Sans", 22), para_props=spacing(0))
    defaults = ""
    if default_run is not None or default_para is not None:
        defaults = ('<w:docDefaults><w:rPrDefault><w:rPr>' + (default_run or '') + '</w:rPr></w:rPrDefault>'
                    '<w:pPrDefault><w:pPr>' + (default_para or '') + '</w:pPr></w:pPrDefault></w:docDefaults>')
    return f'<w:styles xmlns:w="{W}">{defaults}{items}</w:styles>'


def cell_margins(top, bottom=0):
    return (f'<w:tblCellMar><w:top w:w="{top}" w:type="dxa"/>'
            f'<w:left w:w="120" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/>'
            '<w:right w:w="120" w:type="dxa"/></w:tblCellMar>')


def table(style_id=None, properties=""):
    props = (f'<w:tblStyle w:val="{style_id}"/>' if style_id else '') + '<w:tblW w:w="4800" w:type="dxa"/>' + properties
    return ('<w:tbl><w:tblPr>' + props + '</w:tblPr><w:tblGrid><w:gridCol w:w="4800"/></w:tblGrid>'
            '<w:tr><w:tc><w:tcPr><w:tcW w:w="4800" w:type="dxa"/></w:tcPr>'
            + para('The cobalt capsule contains five ceramic lenses.') + '</w:tc></w:tr></w:tbl>')


def package(folder, name, body, style_xml=None, top=1440):
    types = [('document', 'document.main'), ('styles', 'styles'), ('settings', 'settings')]
    parts = {
        '[Content_Types].xml': ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>' + ''.join(
            f'<Override PartName="/word/{name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>'
            for name, kind in types) + '</Types>'),
        '_rels/.rels': f'<Relationships xmlns="{REL}"><Relationship Id="rDocument" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>',
        'word/_rels/document.xml.rels': f'<Relationships xmlns="{REL}"><Relationship Id="rStyles" Type="{R}/styles" Target="styles.xml"/><Relationship Id="rSettings" Type="{R}/settings" Target="settings.xml"/></Relationships>',
        'word/document.xml': (f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body}<w:sectPr>'
            '<w:pgSz w:w="12240" w:h="15840"/>'
            f'<w:pgMar w:top="{top}" w:right="1440" w:bottom="1440" w:left="1440" w:header="360" w:footer="360" w:gutter="0"/>'
            '</w:sectPr></w:body></w:document>'),
        'word/styles.xml': style_xml or styles(),
        'word/settings.xml': f'<w:settings xmlns:w="{W}"/>',
    }
    path = folder / f'{name}.docx'
    with ZipFile(path, 'w', ZIP_DEFLATED) as archive:
        for entry, xml in sorted(parts.items()):
            info = ZipInfo(entry, (2025, 1, 1, 0, 0, 0))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + xml).encode())
    return path.name


def generate(folder):
    folder.mkdir(parents=True, exist_ok=True)
    cases = []
    def pair(name, group, left_body, right_body, left_style=None, right_style=None, check='effective-format', left_top=1440, right_top=1440):
        left = package(folder, name + '-original', left_body, left_style, left_top)
        right = package(folder, name + '-revised', right_body, right_style, right_top)
        cases.append(dict(name=name, group=group, left=left, right=right, check=check))

    invented = ['Cobalt station logs the evening signal.', 'Five ceramic lenses rest inside the capsule.',
                'Quartz markers identify the northern relay.', 'The violet dial records a steady pulse.']
    plain = ''.join(para(t) for t in invented)
    old_style = styles(style('BaseText', run_props=font('DejaVu Sans', 20), para_props=spacing(0)))
    new_style = styles(style('FreshText', run_props=font('DejaVu Serif', 28), para_props=spacing(360, 360)))
    pair('default-style-id', 'style-correspondence', plain, plain, old_style, new_style)
    pair('named-style-id', 'style-correspondence', ''.join(para(t, 'BaseText') for t in invented),
         ''.join(para(t, 'FreshText') for t in invented), old_style, new_style)
    same_id = styles(style('BaseText', run_props=font('DejaVu Serif', 28), para_props=spacing(360, 360)))
    pair('same-style-id-control', 'style-correspondence', plain, plain, old_style, same_id)

    shared = style()
    old_defaults = styles(shared, font('DejaVu Sans', 20), spacing(0))
    new_defaults = styles(shared, font('DejaVu Serif', 28), spacing(360, 360))
    pair('defaults-plain-control', 'effective-defaults', plain, plain, old_defaults, new_defaults)
    pair('defaults-with-table', 'effective-defaults', plain + table(), plain + table(), old_defaults, new_defaults)

    head = para('Cobalt log begins here.')
    tail = para('Cobalt log ends here.')
    left_a = 'Quartz relay signal chamber enabled.'
    left_b = 'Amber sensor pressure chamber active.'
    right_a = 'Quartz relay signal chamber disabled.'
    right_b = 'Amber sensor pressure chamber idle.'
    heading = 'Quartz relay signal chamber maintenance schedule for the observatory team.'
    alignment_styles = styles(style() + style('PanelHeading', 'Panel Heading', font('DejaVu Sans', 28),
        '<w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="1"/>', False))
    pair('inserted-heading-gap', 'paragraph-pairing', head + para(left_a) + para(left_b) + tail,
         head + para(heading, 'PanelHeading') + para(right_a) + para(right_b) + tail,
         alignment_styles, alignment_styles, check='heading-pairing')
    pair('body-edits-control', 'paragraph-pairing', head + para(left_a) + para(left_b) + tail,
         head + para(right_a) + para(right_b) + tail, alignment_styles, alignment_styles, check='heading-pairing')

    def table_style(top, bottom):
        return styles(style() + '<w:style w:type="table" w:styleId="CapsuleGrid"><w:name w:val="Capsule Grid"/>'
                      '<w:tblPr>' + cell_margins(top, bottom) + '</w:tblPr></w:style>')
    pair('table-style-cell-margins', 'layout-properties', table('CapsuleGrid'), table('CapsuleGrid'),
         table_style(0, 0), table_style(300, 180), check='table-margins')
    pair('table-direct-margins-control', 'layout-properties', table(properties=cell_margins(0, 0)),
         table(properties=cell_margins(300, 180)), check='table-margins')
    pair('section-margin-control', 'layout-properties', plain, plain, check='section-margin', right_top=2160)
    pair('paragraph-spacing-control', 'layout-properties', para(invented[0], properties=spacing(0)),
         para(invented[0], properties=spacing(360, 360)))
    (folder / 'cases.json').write_text(json.dumps(cases, indent=2) + '\n')
    return cases


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    result = generate(args.output.resolve())
    print(f'Generated {len(result)} synthetic pairs in {args.output.resolve()}')
