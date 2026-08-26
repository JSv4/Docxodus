import { R_NS, storedZip, W_NS, xml } from './docx-zip.js';

/**
 * Fixtures for the revision and comment families the export profiles cannot draw.
 *
 * These are built from readable XML rather than committed binaries because the whole point of
 * each is one specific element the manifest counts — a `w:pPrChange` against a `w:rPrChange`, a
 * `w:customXmlInsRangeStart`, a `w15:commentEx` with a `paraIdParent` — and a binary fixture
 * would hide exactly that.
 */

const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const AUTHORED = 'w:author="Reviewer" w:date="2026-08-16T00:00:00Z"';

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

interface PackageParts {
  /** Body markup, excluding the trailing `w:sectPr`. */
  body: string;
  /** Root element attributes for `w:document` beyond the `w` namespace. */
  documentAttributes?: string;
  /** Extra `[Content_Types].xml` overrides. */
  overrides?: string;
  /** Extra `word/_rels/document.xml.rels` relationships. */
  relationships?: string;
  /** Extra parts keyed by ZIP entry name. */
  parts?: Record<string, string>;
}

/** The smallest package the export accepts, so each fixture contributes only its own subject. */
function minimalPackage({
  body,
  documentAttributes = '',
  overrides = '',
  relationships = '',
  parts = {},
}: PackageParts): Uint8Array {
  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CT_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  ${overrides}
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
  ${relationships}
</Relationships>`),
    },
    { name: 'word/styles.xml', data: xml(STYLES_XML) },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"${documentAttributes}><w:body>
  ${body}
  ${SECT_PR}
</w:body></w:document>`),
    },
    ...Object.entries(parts).map(([name, data]) => ({ name, data: xml(data) })),
  ]);
}

/**
 * A document carrying both sides of the property-revision split plus a custom XML range.
 *
 * Two `w:pPrChange` are unrepresentable; the `w:rPrChange`, the `w:ins`, and (since issue
 * #538) the `w:customXmlInsRangeStart`/`End` pair beside them are drawn, and exist so the
 * warnings can be shown to count only what is actually missing. Manifest counts:
 * `propertyChanges` 3, `runPropertyChanges` 1, `otherChanges` 1, `insertions` 1.
 */
export function generateUnrenderableRevisionDocx(): Uint8Array {
  return minimalPackage({
    body: `<w:p>
    <w:pPr>
      <w:jc w:val="center"/>
      <!-- Word records the previous paragraph properties; markup cannot draw this. -->
      <w:pPrChange w:id="10" ${AUTHORED}>
        <w:pPr><w:jc w:val="left"/></w:pPr>
      </w:pPrChange>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:b/>
        <!-- A run-level format change; markup DOES draw this one. -->
        <w:rPrChange w:id="11" ${AUTHORED}><w:rPr/></w:rPrChange>
      </w:rPr>
      <w:t xml:space="preserve">Bolded by a reviewer. </w:t>
    </w:r>
    <w:ins w:id="12" ${AUTHORED}>
      <w:r><w:t>This insertion is drawn.</w:t></w:r>
    </w:ins>
  </w:p>
  <w:p>
    <w:pPr>
      <w:ind w:left="720"/>
      <w:pPrChange w:id="13" ${AUTHORED}><w:pPr/></w:pPrChange>
    </w:pPr>
    <!-- A custom XML revision range; the converter has no handling for these at all. -->
    <w:customXmlInsRangeStart w:id="14" ${AUTHORED}/>
    <w:r><w:t>Indented, and wrapped in custom XML by a reviewer.</w:t></w:r>
    <w:customXmlInsRangeEnd w:id="14"/>
  </w:p>`,
  });
}

/**
 * A document whose every revision is one the markup profile draws: a run-level format change and
 * an insertion. Nothing about it is unrepresentable, so it must raise no revision warning and
 * must survive `unsupportedContent: "strict"`.
 */
export function generateRenderedRevisionOnlyDocx(): Uint8Array {
  return minimalPackage({
    body: `<w:p>
    <w:r>
      <w:rPr>
        <w:i/>
        <w:rPrChange w:id="20" ${AUTHORED}><w:rPr/></w:rPrChange>
      </w:rPr>
      <w:t xml:space="preserve">Italicised by a reviewer. </w:t>
    </w:r>
    <w:ins w:id="21" ${AUTHORED}>
      <w:r><w:t>And an insertion beside it.</w:t></w:r>
    </w:ins>
  </w:p>`,
  });
}

/**
 * A table whose second cell is a tracked insertion and whose third is a tracked deletion.
 *
 * These are drawn — the converter stamps `rev-cell-ins`/`rev-cell-del` on the cell and emits CSS
 * that tints and strikes it — so they must raise no warning. The fixture exists to hold that line:
 * reporting them would fail a strict export whose content is fully visible.
 */
export function generateCellRevisionDocx(): Uint8Array {
  return minimalPackage({
    body: `<w:tbl>
    <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/></w:tblGrid>
    <w:tr>
      <w:tc>
        <w:tcPr><w:tcW w:w="3120" w:type="dxa"/></w:tcPr>
        <w:p><w:r><w:t>Existing cell</w:t></w:r></w:p>
      </w:tc>
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="3120" w:type="dxa"/>
          <w:cellIns w:id="30" ${AUTHORED}/>
        </w:tcPr>
        <w:p><w:r><w:t>Inserted cell</w:t></w:r></w:p>
      </w:tc>
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="3120" w:type="dxa"/>
          <w:cellDel w:id="31" ${AUTHORED}/>
        </w:tcPr>
        <w:p><w:r><w:t>Deleted cell</w:t></w:r></w:p>
      </w:tc>
    </w:tr>
  </w:tbl>
  <w:p/>`,
  });
}

/**
 * A document with two comments where the second is a reply to the first, and the first is
 * marked resolved. The threading and resolved state live only in commentsExtended, which the
 * converter reads since issue #540: the reply nests beneath its thread root and the resolved
 * root is badged and muted.
 */
export function generateCommentTopologyDocx(): Uint8Array {
  return minimalPackage({
    documentAttributes: ` xmlns:r="${R_NS}"`,
    overrides: '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>\n'
      + '  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>',
    relationships: `<Relationship Id="rId2" Type="${R_NS}/comments" Target="comments.xml"/>\n`
      + '  <Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>',
    body: `<w:p>
    <w:commentRangeStart w:id="1"/>
    <w:r><w:t>The clause under review.</w:t></w:r>
    <w:commentRangeEnd w:id="1"/>
    <w:r><w:commentReference w:id="1"/></w:r>
    <w:r><w:commentReference w:id="2"/></w:r>
  </w:p>`,
    parts: {
      'word/comments.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}">
  <w:comment w:id="1" w:author="First Reviewer" w:initials="FR" w:date="2026-08-16T00:00:00Z">
    <w:p w14:paraId="11111111"><w:r><w:t>Is this clause still needed?</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="2" w:author="Second Reviewer" w:initials="SR" w:date="2026-08-16T01:00:00Z">
    <w:p w14:paraId="22222222"><w:r><w:t>No, it was superseded.</w:t></w:r></w:p>
  </w:comment>
</w:comments>`,
      'word/commentsExtended.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx xmlns:w15="${W15_NS}">
  <w15:commentEx w15:paraId="11111111" w15:done="1"/>
  <w15:commentEx w15:paraId="22222222" w15:paraIdParent="11111111" w15:done="0"/>
</w15:commentsEx>`,
    },
  });
}
