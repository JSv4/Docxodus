import io, os, shutil, struct, sys, zipfile

OUT = sys.argv[1]
FIXTURES = sys.argv[2]
os.makedirs(OUT, exist_ok=True)

CT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
<Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
<Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
<Relationship Id="rIdImg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
<Relationship Id="rIdExt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/x" TargetMode="External"/>
</Relationships>"""

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DOC = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}" xmlns:r="{R}">
<w:body>
<w:p><w:r><w:t xml:space="preserve">Hello  world</w:t></w:r>
<w:r><w:rPr><w:ins w:id="9" w:author="a" w:date="2024-01-01T00:00:00Z"/></w:rPr></w:r></w:p>
<w:p><w:ins w:id="1" w:author="a" w:date="2024-01-01T00:00:00Z"><w:r><w:t>ins</w:t></w:r></w:ins>
<w:del w:id="2" w:author="a" w:date="2024-01-01T00:00:00Z"><w:r><w:delText>del</w:delText></w:r></w:del>
<w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rIdImg"/></w:drawing></w:r>
<w:r><w:commentReference w:id="0"/></w:r><w:r><w:footnoteReference w:id="2"/></w:r></w:p>
<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>
<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr>
</w:body></w:document>"""

STYLES = f'<?xml version="1.0"?><w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>'
HEADER = f'<?xml version="1.0"?><w:hdr xmlns:w="{W}"><w:p><w:r><w:t>H</w:t></w:r></w:p></w:hdr>'
FOOTNOTES = f'<?xml version="1.0"?><w:footnotes xmlns:w="{W}"><w:footnote w:id="2"><w:p/></w:footnote></w:footnotes>'
COMMENTS = f'<?xml version="1.0"?><w:comments xmlns:w="{W}"><w:comment w:id="0" w:author="a"><w:p/></w:comment></w:comments>'
CORE = '<?xml version="1.0"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>'
PNG = bytes.fromhex('89504e470d0a1a0a0000000d494844520000000100000001080600000037') + b'\x00'*20

def rich(zip64=False, dup=False, dirs=False, stored=False):
    buf = io.BytesIO()
    comp = zipfile.ZIP_STORED if stored else zipfile.ZIP_DEFLATED
    with zipfile.ZipFile(buf, 'w', comp) as z:
        def w(name, data):
            if isinstance(data, str): data = data.encode()
            if zip64:
                with z.open(zipfile.ZipInfo(name), 'w', force_zip64=True) as f: f.write(data)
            else:
                z.writestr(name, data)
        if dirs: w('word/', b'')
        w('[Content_Types].xml', CT); w('_rels/.rels', ROOT_RELS)
        w('word/document.xml', DOC); w('word/_rels/document.xml.rels', DOC_RELS)
        w('word/styles.xml', STYLES); w('word/header1.xml', HEADER)
        w('word/footnotes.xml', FOOTNOTES); w('word/comments.xml', COMMENTS)
        w('docProps/core.xml', CORE); w('word/media/image1.png', PNG)
        w('customXml/item1.xml', '<data xmlns="urn:opaque"><a>1</a> <b>2</b></data>')
        if dup: w('word/styles.xml', STYLES.replace('Normal', 'Other'))
    return buf.getvalue()

def save(name, data):
    with open(os.path.join(OUT, name), 'wb') as f: f.write(data)

base = rich()
save('rich.zip', base)
save('rich-zip64.zip', rich(zip64=True))
save('rich-dup.zip', rich(dup=True))
save('rich-dirs.zip', rich(dirs=True))
save('rich-stored.zip', rich(stored=True))
save('rich-prepended.zip', b'JUNKJUNK' + base)
save('rich-appended.zip', base + b'TRAILER!')
save('rich-truncated.zip', base[:len(base)//2])

# set the encryption bit (general-purpose flag bit 0) in local + central headers
enc = bytearray(rich(stored=True))
i = 0
while True:
    j = enc.find(b'PK\x03\x04', i)
    if j < 0: break
    enc[j+6] |= 1; i = j+4
i = 0
while True:
    j = enc.find(b'PK\x01\x02', i)
    if j < 0: break
    enc[j+8] |= 1; i = j+4
save('rich-encflag.zip', bytes(enc))

# weird names
buf = io.BytesIO()
with zipfile.ZipFile(buf, 'w') as z:
    z.writestr('[Content_Types].xml', CT)
    z.writestr('word/a%40b.xml', '<x/>')
    z.writestr('word/caf%C3%A9.xml', '<x/>')
    z.writestr('word/café.xml', '<x/>')          # UTF-8 flagged literal
    z.writestr('word/document.xml/extra.xml', '<x/>')  # interleaved
    z.writestr('../escape.xml', '<x/>')
    z.writestr('word/..%2Fup.xml', '<x/>')
save('weird-names.zip', buf.getvalue())

# malformed content types + DTD part
buf = io.BytesIO()
with zipfile.ZipFile(buf, 'w') as z:
    z.writestr('[Content_Types].xml', CT[: len(CT)//2])
    z.writestr('word/document.xml', '<!DOCTYPE x [<!ENTITY e "boom">]><x>&e;</x>')
save('bad-ct-dtd.zip', buf.getvalue())

# ---- valid CFB with EncryptedPackage/EncryptionInfo (MS-CFB v3, 512B sectors) ----
FREE, ENDCHAIN, FATSECT = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD
def dirent(name, etype, color, left, right, child, start, size):
    n = name.encode('utf-16-le') + b'\x00\x00'
    e = n + b'\x00' * (64 - len(n))
    e += struct.pack('<HBBIII', len(n), etype, color, left, right, child)
    e += b'\x00'*16 + struct.pack('<I', 0) + b'\x00'*16   # clsid, state, times
    e += struct.pack('<IQ', start, size)
    assert len(e) == 128, len(e)
    return e

hdr = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' + b'\x00'*16
hdr += struct.pack('<HHHHH', 0x003E, 0x0003, 0xFFFE, 0x0009, 0x0006)
hdr += b'\x00'*6
hdr += struct.pack('<IIIIIIIII', 0, 1, 1, 0, 0x1000, ENDCHAIN, 0, ENDCHAIN, 0)
hdr += struct.pack('<I', 0) + b'\xff'*4*108
assert len(hdr) == 512

fat = [FATSECT, ENDCHAIN]                       # sector0 FAT, sector1 directory
fat += list(range(3, 10)) + [ENDCHAIN]          # sectors 2..9: EncryptionInfo (4096B)
fat += list(range(11, 18)) + [ENDCHAIN]         # sectors 10..17: EncryptedPackage
fat += [FREE] * (128 - len(fat))
fatsec = struct.pack('<128I', *fat)

d = dirent('Root Entry', 5, 1, FREE, FREE, 1, ENDCHAIN, 0)
d += dirent('EncryptionInfo', 2, 1, FREE, 2, FREE, 2, 4096)
d += dirent('EncryptedPackage', 2, 1, FREE, FREE, FREE, 10, 4096)
d += b'\x00' * 128  # unused directory entry
cfb = hdr + fatsec + d + b'\xEE'*4096 + b'\xDD'*4096
save('ole-encrypted.cfb', cfb)
save('ole-magic-only.bin', b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' + b'\x00'*504)

# raw non-zip inputs
save('plain.xml', b'<?xml version="1.0"?><root a="1"> <child/> </root>')
save('empty.bin', b'')
save('tiny.bin', b'PK')
save('random.bin', os.urandom(1024))

# small real fixtures
import glob
picked = 0
for p in sorted(glob.glob(os.path.join(FIXTURES, '**', '*.docx'), recursive=True), key=os.path.getsize):
    if os.path.getsize(p) < 40000:
        shutil.copy(p, os.path.join(OUT, 'fx%d.docx' % picked)); picked += 1
    if picked >= 6: break

print('seeds:', len(os.listdir(OUT)))
