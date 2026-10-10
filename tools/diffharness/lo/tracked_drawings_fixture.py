"""Generate tracked DrawingML controls and a 48-case renderer isolation matrix.

Usage: tracked_drawings_fixture.py REPO OUTPUT_DIR
Fresh literal OOXML, Python standard library; .NET runs comparison and schema checks.
"""
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZIP_DEFLATED, ZipInfo
import json
import sys
import subprocess

OUT = Path(sys.argv[2]).resolve()
OUT.mkdir(parents=True,exist_ok=True)
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
DECL = f'xmlns:w="{W}" xmlns:r="{R}" xmlns:a="{A}" xmlns:wp="{WP}" xmlns:wps="{WPS}" xmlns:wpg="{WPG}" xmlns:v="urn:schemas-microsoft-com:vml"'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

def run(text):
    return '<w:r><w:t xml:space="preserve">' + escape(text) + '</w:t></w:r>'

def para(text, style=None):
    props = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ''
    return '<w:p>' + props + run(text) + '</w:p>'

def styles(default_run='', default_para='', normal_run='', normal_para='', extra=''):
    return (f'<w:styles xmlns:w="{W}"><w:docDefaults>'
            f'<w:rPrDefault><w:rPr>{default_run}</w:rPr></w:rPrDefault>'
            f'<w:pPrDefault><w:pPr>{default_para}</w:pPr></w:pPrDefault></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
            f'<w:pPr>{normal_para}</w:pPr><w:rPr>{normal_run}</w:rPr></w:style>{extra}</w:styles>')

STANDARD_STYLES = styles('<w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="22"/>')

def package(name, body, style_xml=STANDARD_STYLES, header=None, theme=None):
    parts = {}
    types = [('document','application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'),
             ('styles','application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'),
             ('settings','application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml')]
    rels = [('rStyles','styles','styles.xml'),('rSettings','settings','settings.xml')]
    refs = ''
    if header is not None:
        parts['word/header1.xml'] = f'<w:hdr {DECL}>{header}</w:hdr>'
        types.append(('header1','application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'))
        rels.append(('rHeader','header','header1.xml'))
        refs = '<w:headerReference w:type="default" r:id="rHeader"/>'
    if theme is not None:
        parts['word/theme/theme1.xml'] = theme
        types.append(('theme/theme1','application/vnd.openxmlformats-officedocument.theme+xml'))
        rels.append(('rTheme','theme','theme/theme1.xml'))
    parts['[Content_Types].xml'] = ('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>' + ''.join(
            f'<Override PartName="/word/{part}.xml" ContentType="{content}"/>' for part,content in types) + '</Types>')
    parts['_rels/.rels'] = f'<Relationships xmlns="{REL}"><Relationship Id="rDocument" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>'
    parts['word/_rels/document.xml.rels'] = f'<Relationships xmlns="{REL}">' + ''.join(
        f'<Relationship Id="{ident}" Type="{R}/{kind}" Target="{target}"/>' for ident,kind,target in rels) + '</Relationships>'
    parts['word/document.xml'] = (f'<w:document {DECL}><w:body>{body}<w:sectPr>{refs}'
        '<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="360" w:footer="360" w:gutter="0"/>'
        '</w:sectPr></w:body></w:document>')
    parts['word/styles.xml'] = style_xml
    parts['word/settings.xml'] = f'<w:settings xmlns:w="{W}"/>'
    path = OUT / f'{name}.docx'
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        for entry,data in sorted(parts.items()):
            zi=ZipInfo(entry,(2024,1,1,0,0,0));zi.compress_type=ZIP_DEFLATED
            z.writestr(zi,('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+data).encode())
    return name

PAIRS=[]
def pair(name,left,right):
    PAIRS.append(dict(name=name,left=f'{left}.docx',right=f'{right}.docx'))

# Revision wrappers around a newly drawn header label. No images, linked files, or external content.
vml='<w:r><w:pict><v:rect id="header-box" style="width:120pt;height:24pt" fillcolor="#336699"><v:textbox><w:txbxContent>'+para('Harbor workshop')+'</w:txbxContent></v:textbox></v:rect></w:pict></w:r>'
def shape(ident,label,x):
    return (f'<wps:wsp><wps:cNvPr id="{ident}" name="Box {ident}"/><wps:cNvSpPr txBox="1"/><wps:spPr>'
        f'<a:xfrm><a:off x="{x}" y="0"/><a:ext cx="1200000" cy="400000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="D6EAF8"/></a:solidFill></wps:spPr>'
        f'<wps:txbx><w:txbxContent>{para(label)}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>')
group=('<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
       '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>'
       '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
       '<wp:extent cx="2500000" cy="400000"/><wp:wrapNone/><wp:docPr id="1" name="Two boxes"/><wp:cNvGraphicFramePr/>'
       '<a:graphic><a:graphicData uri="'+WPG+'"><wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm>'
       '<a:off x="0" y="0"/><a:ext cx="2500000" cy="400000"/><a:chOff x="0" y="0"/><a:chExt cx="2500000" cy="400000"/>'
       '</a:xfrm></wpg:grpSpPr>'+shape(2,'Harbor workshop',0)+shape(3,'Room seven',1300000)+
       '</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')
header_base=package('header-absent',para('The workshop opens at nine.'))
for kind,graphic in [('vml',vml),('group',group)]:
    clean=package(f'header-{kind}-plain',para('The workshop opens at nine.'),header='<w:p>'+graphic+'</w:p>')
    pair(f'header-insert-{kind}',header_base,clean)
    pair(f'header-delete-{kind}',clean,header_base)
    for revision in ['ins','del']:
        markup=f'<w:{revision} w:id="7" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">{graphic}</w:{revision}>'
        package(f'header-{kind}-{revision}',para('The workshop opens at nine.'),header='<w:p>'+markup+'</w:p>')



# Isolate each dimension while retaining the tracked insertion/deletion controls.
import copy
import itertools
import xml.etree.ElementTree as ET
for ns, uri in [('w', W), ('r', R), ('a', A), ('wp', WP), ('wps', WPS), ('wpg', WPG)]:
    ET.register_namespace(ns, uri)
MATRIX = []
for grouped, textbox, anchored, header in itertools.product([False, True], repeat=4):
    graphic = ET.fromstring('<root ' + DECL + '>' + group + '</root>')[0]
    data = graphic.find('.//{' + A + '}graphicData')
    if not grouped:
        single = copy.deepcopy(data.find('.//{' + WPS + '}wsp'))
        data.clear()
        data.set('uri', WPS)
        data.append(single)
    if not textbox:
        for wsp in graphic.findall('.//{' + WPS + '}wsp'):
            txbx = wsp.find('{' + WPS + '}txbx')
            wsp.remove(txbx)
            wsp.find('{' + WPS + '}cNvSpPr').set('txBox', '0')
    if not anchored:
        anchor = graphic.find('.//{' + WP + '}anchor')
        anchor.tag = '{' + WP + '}inline'
        anchor.attrib.clear()
        for child in list(anchor):
            if child.tag.rsplit('}', 1)[-1] in ['simplePos', 'positionH', 'positionV', 'wrapNone']:
                anchor.remove(child)
    xml = ET.tostring(graphic, encoding='unicode')
    stem = 'matrix-' + '-'.join(['group' if grouped else 'single', 'textbox' if textbox else 'no-textbox',
                               'anchor' if anchored else 'inline', 'header' if header else 'body'])
    for revision in ['plain', 'ins', 'del']:
        tracked = xml if revision == 'plain' else ('<w:' + revision + ' w:id="7" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">' + xml + '</w:' + revision + '>')
        paragraph = '<w:p>' + tracked + '</w:p>'
        name = package(stem + '-' + revision, para('The workshop opens at nine.') + ('' if header else paragraph),
                       header=paragraph if header else None)
        MATRIX.append(dict(name=name + '.docx', grouped=grouped, textbox=textbox, anchored=anchored, header=header, revision=revision))
(OUT/'matrix.json').write_text(json.dumps(MATRIX, indent=2))

# Paragraph-mark-only tracking is a diagnostic: it does not track the drawing run.
for revision in ['ins', 'del']:
    marker='<w:pPr><w:rPr><w:'+revision+' w:id="7" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"/></w:rPr></w:pPr>'
    package('header-group-paragraph-mark-'+revision, para('The workshop opens at nine.'), header='<w:p>'+marker+group+'</w:p>')

RUNNER = r'''
using Docxodus;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;

var folder = Path.GetFullPath(args[0]);
var output = Path.Combine(folder, "generated");
Directory.CreateDirectory(output);
XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

object Inspect(string file)
{
    using var document = WordprocessingDocument.Open(file, false);
    var errors = new OpenXmlValidator(FileFormatVersions.Office2019) { MaxNumberOfErrors = 0 }
        .Validate(document).Select(e => new { e.Id, e.Description, Part=e.Part?.Uri.ToString(), Path=e.Path?.XPath }).ToArray();
    using var zip = ZipFile.OpenRead(file);
    XDocument Xml(string path) { using var stream=zip.GetEntry(path)!.Open(); return XDocument.Load(stream); }
    var contents = zip.Entries.Where(e => e.FullName.EndsWith(".xml")).Select(e => Xml(e.FullName)).ToArray();
    var drawingCount = contents.Sum(xml => xml.Descendants(w+"drawing").Count());
    var pictureCount = contents.Sum(xml => xml.Descendants(w+"pict").Count());
    return new { File=Path.GetFileName(file), Sha256=Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(file))).ToLowerInvariant(), Errors=errors, Drawings=drawingCount, Pictures=pictureCount };
}

var reports = new List<object>();
using var pairs = JsonDocument.Parse(File.ReadAllText(Path.Combine(folder,"pairs.json")));
foreach (var pair in pairs.RootElement.EnumerateArray())
{
    var name=pair.GetProperty("name").GetString()!;
    var leftPath=Path.Combine(folder,pair.GetProperty("left").GetString()!);
    var rightPath=Path.Combine(folder,pair.GetProperty("right").GetString()!);
    try
    {
        var compared=DocxCompare.Compare(new WmlDocument(leftPath),new WmlDocument(rightPath));
        var comparedPath=Path.Combine(output,name+".docx");
        compared.SaveAs(comparedPath);
        var acceptedPath=Path.Combine(output,name+"-accepted.docx");
        var rejectedPath=Path.Combine(output,name+"-rejected.docx");
        RevisionProcessor.AcceptRevisions(compared).SaveAs(acceptedPath);
        RevisionProcessor.RejectRevisions(compared).SaveAs(rejectedPath);
        reports.Add(new { Name=name, Left=Inspect(leftPath), Right=Inspect(rightPath), Compared=Inspect(comparedPath), Accepted=Inspect(acceptedPath), Rejected=Inspect(rejectedPath) });
        Console.WriteLine(name+": generated");
    }
    catch (Exception ex) { reports.Add(new { Name=name, Error=ex.ToString() }); Console.WriteLine(name+": "+ex.Message); }
}
foreach (var file in Directory.GetFiles(folder, "header-group-paragraph-mark-*.docx"))
{
    var source = new WmlDocument(file);
    var name = Path.GetFileNameWithoutExtension(file);
    var accepted = Path.Combine(output, name + "-accepted.docx");
    var rejected = Path.Combine(output, name + "-rejected.docx");
    RevisionProcessor.AcceptRevisions(source).SaveAs(accepted);
    RevisionProcessor.RejectRevisions(source).SaveAs(rejected);
    reports.Add(new { Name=name, Mode="paragraph-mark-diagnostic", Source=Inspect(file), Accepted=Inspect(accepted), Rejected=Inspect(rejected) });
}
File.WriteAllText(Path.Combine(folder,"results.json"),JsonSerializer.Serialize(reports,new JsonSerializerOptions{WriteIndented=true}));
var standalone=Directory.GetFiles(folder,"*.docx").Select(Inspect).ToArray();
File.WriteAllText(Path.Combine(folder,"input-validation.json"),JsonSerializer.Serialize(standalone,new JsonSerializerOptions{WriteIndented=true}));
'''


(OUT/'pairs.json').write_text(json.dumps(PAIRS,indent=2))
repo=Path(sys.argv[1]).resolve()
project=repo/'Docxodus/Docxodus.csproj'
if not project.is_file():
    raise SystemExit('First argument must be a Docxodus repository checkout')
(OUT/'Repro.csproj').write_text('<Project Sdk="Microsoft.NET.Sdk"><PropertyGroup><OutputType>Exe</OutputType><TargetFramework>net10.0</TargetFramework><ImplicitUsings>enable</ImplicitUsings><Nullable>enable</Nullable></PropertyGroup><ItemGroup><ProjectReference Include="'+escape(str(project), {'"':'&quot;'})+'" /></ItemGroup></Project>')
(OUT/'Program.cs').write_text(RUNNER)
subprocess.run(['dotnet','run','--project',str(OUT/'Repro.csproj'),'-c','Release','-p:NuGetAudit=false','--',str(OUT)],check=True)
print('Generated DOCX files and results.json:',OUT)
