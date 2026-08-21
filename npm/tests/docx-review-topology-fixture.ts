import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

/**
 * Fixtures for the revision and comment families the export profiles cannot draw.
 *
 * Both are built from readable XML rather than committed binaries because the whole point of
 * each is one specific element the manifest counts — a `w:pPrChange`, a `w15:commentEx` with a
 * `paraIdParent` — and a binary fixture would hide exactly that.
 */

const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>`;

const SECT_PR = '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
  + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

/**
 * A document carrying one revision from each family the markup profile cannot draw:
 * a paragraph-property change, a tracked cell insertion, and — for contrast — an ordinary
 * insertion that markup does draw.
 */
export function generateUnrenderableRevisionDocx(): Uint8Array {
  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CT_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`),
    },
    {
      name: '_rels/.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG_REL_NS}">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: 'word/_rels/document.xml.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG_REL_NS}">
  <Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/>
</Relationships>`),
    },
    { name: 'word/styles.xml', data: xml(STYLES_XML) },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"><w:body>
  <w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      <!-- Word records the previous paragraph properties; markup cannot draw this. -->
      <w:pPrChange w:id="10" w:author="Reviewer" w:date="2026-08-16T00:00:00Z">
        <w:pPr><w:jc w:val="left"/></w:pPr>
      </w:pPrChange>
    </w:pPr>
    <w:r><w:t xml:space="preserve">Centered now, left before. </w:t></w:r>
    <w:ins w:id="11" w:author="Reviewer" w:date="2026-08-16T00:00:00Z">
      <w:r><w:t>This insertion is drawn.</w:t></w:r>
    </w:ins>
  </w:p>
  <w:tbl>
    <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>
    <w:tr>
      <w:tc>
        <w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>
        <w:p><w:r><w:t>Existing cell</w:t></w:r></w:p>
      </w:tc>
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="4680" w:type="dxa"/>
          <!-- A tracked cell insertion; markup cannot draw this either. -->
          <w:cellIns w:id="12" w:author="Reviewer" w:date="2026-08-16T00:00:00Z"/>
        </w:tcPr>
        <w:p><w:r><w:t>Inserted cell</w:t></w:r></w:p>
      </w:tc>
    </w:tr>
  </w:tbl>
  <w:p/>
  ${SECT_PR}
</w:body></w:document>`),
    },
  ]);
}

/**
 * A document with two comments where the second is a reply to the first, and the first is
 * marked resolved. The threading and resolved state live only in commentsExtended, which the
 * converter does not read.
 */
export function generateCommentTopologyDocx(): Uint8Array {
  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CT_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>
</Types>`),
    },
    {
      name: '_rels/.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG_REL_NS}">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: 'word/_rels/document.xml.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PKG_REL_NS}">
  <Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${R_NS}/comments" Target="comments.xml"/>
  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>
</Relationships>`),
    },
    { name: 'word/styles.xml', data: xml(STYLES_XML) },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
  <w:p>
    <w:commentRangeStart w:id="1"/>
    <w:r><w:t>The clause under review.</w:t></w:r>
    <w:commentRangeEnd w:id="1"/>
    <w:r><w:commentReference w:id="1"/></w:r>
    <w:r><w:commentReference w:id="2"/></w:r>
  </w:p>
  ${SECT_PR}
</w:body></w:document>`),
    },
    {
      name: 'word/comments.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:comment w:id="1" w:author="First Reviewer" w:initials="FR" w:date="2026-08-16T00:00:00Z">
    <w:p w14:paraId="11111111"><w:r><w:t>Is this clause still needed?</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="2" w:author="Second Reviewer" w:initials="SR" w:date="2026-08-16T01:00:00Z">
    <w:p w14:paraId="22222222"><w:r><w:t>No, it was superseded.</w:t></w:r></w:p>
  </w:comment>
</w:comments>`),
    },
    {
      name: 'word/commentsExtended.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx xmlns:w15="${W15_NS}">
  <w15:commentEx w15:paraId="11111111" w15:done="1"/>
  <w15:commentEx w15:paraId="22222222" w15:paraIdParent="11111111" w15:done="0"/>
</w15:commentsEx>`),
    },
  ]);
}
