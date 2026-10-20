"""OOXML fixtures for conversion, pagination, and revision resolution."""
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZIP_DEFLATED

OUT = Path(__file__).parent / "inputs"
OUT.mkdir(exist_ok=True)
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS = f'xmlns:w="{W}" xmlns:r="{R}" xmlns:a="{A}" xmlns:wp="{WP}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"'

def run(text, props=""):
    return f'<w:r>{("<w:rPr>"+props+"</w:rPr>") if props else ""}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'

def para(content, props=""):
    return f'<w:p>{("<w:pPr>"+props+"</w:pPr>") if props else ""}{content}</w:p>'

def package(name, body, interval=720, parts=None, rels=None, footer=None):
    parts = dict(parts or {})
    relationships = [
        ("styles", "styles", "styles.xml"), ("settings", "settings", "settings.xml")
    ] + list(rels or [])
    section = '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
    if footer:
        parts["footer1.xml"] = (f'<w:ftr {NS}>{footer}</w:ftr>', "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
        relationships.append(("footer", "footer", "footer1.xml"))
        section = '<w:footerReference w:type="default" r:id="footer"/>' + section
    parts["document.xml"] = (f'<w:document {NS} mc:Ignorable="wps wpg"><w:body>{body}<w:sectPr>{section}</w:sectPr></w:body></w:document>', "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
    parts["styles.xml"] = (f'<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>', "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")
    parts["settings.xml"] = (f'<w:settings xmlns:w="{W}"><w:defaultTabStop w:val="{interval}"/></w:settings>', "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml")
    types = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    types += "".join(f'<Override PartName="/word/{p}" ContentType="{t}"/>' for p, (_, t) in parts.items()) + '</Types>'
    with ZipFile(OUT / f"{name}.docx", "w", ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", types)
        z.writestr("_rels/.rels", f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/_rels/document.xml.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "".join(f'<Relationship Id="{i}" Type="{R}/{t}" Target="{p}"/>' for i,t,p in relationships) + '</Relationships>')
        for path, (xml, _) in parts.items():
            z.writestr("word/"+path, xml)

for interval in (0, 12, 720):
    tabs = "".join(f'<w:r><w:rPr><w:color w:val="{color}"/></w:rPr><w:tab/></w:r>' + run(label) for color,label in (("112233","North"),("223344","East"),("334455","South")))
    package(f"tabs-{interval}", para(run("Observatory inventory") + tabs + run("Copper lens")), interval)


stamp = 'w:id="31" w:author="Review Example" w:date="2026-09-12T10:00:00Z"'
revision_text = f'<w:del {stamp}><w:r><w:delText>Amber prism</w:delText></w:r></w:del><w:ins w:id="32" w:author="Review Example" w:date="2026-09-12T10:01:00Z">{run("Indigo prism")}</w:ins>'
package("revisions-transitional", para(revision_text))
with ZipFile(OUT/"revisions-transitional.docx") as src:
    payload = {name:src.read(name) for name in src.namelist()}
with ZipFile(OUT/"revisions-strict.docx", "w", ZIP_DEFLATED) as dst:
    for name, content in payload.items():
        value = content.decode().replace(W, "http://purl.oclc.org/ooxml/wordprocessingml/main").replace(R, "http://purl.oclc.org/ooxml/officeDocument/relationships").replace(A, "http://purl.oclc.org/ooxml/drawingml/main").replace(WP, "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing")
        dst.writestr(name, value)
note_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"

for family, size in [("Arial",20)]:
    name = family.lower().replace(" ","-")+str(size)
    props = f'<w:rFonts w:ascii="{family}" w:hAnsi="{family}"/><w:sz w:val="{size}"/><w:rPrChange {stamp}><w:rPr><w:i/></w:rPr></w:rPrChange>'
    pprops = f'<w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:rFonts w:ascii="{family}" w:hAnsi="{family}"/><w:sz w:val="{size}"/></w:rPr>'
    note = f'<w:footnotes {NS}><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1">{para(run("A spare quartz plate remains in the cabinet.",props),pprops)}</w:footnote></w:footnotes>'
    package("note-"+name, para(run("Receiver maintenance")+'<w:r><w:footnoteReference w:id="1"/></w:r>'), parts={"footnotes.xml":(note,note_type)},rels=[("notes","footnotes","footnotes.xml")])
    if name == "arial20":
        control = note.replace(f'<w:rPrChange {stamp}><w:rPr><w:i/></w:rPr></w:rPrChange>', "")
        package("note-arial20-control", para(run("Receiver maintenance")+'<w:r><w:footnoteReference w:id="1"/></w:r>'), parts={"footnotes.xml":(control,note_type)},rels=[("notes","footnotes","footnotes.xml")])

def field(code, cached, deleted=False):
    instruction = "delInstrText" if deleted else "instrText"
    text = "delText" if deleted else "t"
    return f'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:{instruction} xml:space="preserve"> {code} </w:{instruction}></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:{text}>{cached}</w:{text}></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>'
body = para(run("Lunar receiver inspection")) + para('<w:r><w:br w:type="page"/></w:r>'+run("Signal record"))
for deleted in (False, True):
    content = run("Folio ") + field("PAGE","1",deleted) + run(" / ") + field("NUMPAGES","2",deleted)
    if deleted: content = f'<w:del {stamp}>{content.replace("<w:t ", "<w:delText ").replace("</w:t>", "</w:delText>")}</w:del>'
    package("footer-deleted" if deleted else "footer-control", body, footer=para(content))

def shape():
    return '<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="19647E"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp>'

for grouped in (False, True):
    data = shape()
    uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    if grouped:
        uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
        data = '<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="1828800" cy="914400"/></a:xfrm></wpg:grpSpPr>' + data + '</wpg:wgp>'
    drawing = f'<w:r><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="7" name="Teal inspection panel"/><a:graphic><a:graphicData uri="{uri}">{data}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
    for deleted in (False, True):
        payload = f'<w:del {stamp}>{drawing}</w:del>' if deleted else drawing
        package(("group" if grouped else "shape") + ("-deleted" if deleted else "-plain"), para(run("Optical panel")) + para(payload) + para(run("End of inspection")))

print(f"Created {len(list(OUT.glob('*.docx')))} independent inputs in {OUT}")
