import { storedZip, xml } from './docx-zip.js';

/**
 * A paragraph followed by a one-cell bordered table, for the table's horizontal placement (issue
 * #1219). The compatibility mode, the table's own properties and the document's default table style
 * are the variables; everything else is fixed.
 */

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export interface TableIndentOptions {
  /** `w:compatSetting compatibilityMode`; `null` states none. */
  mode: string | null;
  /** Extra `w:tblPr` children, such as `w:tblInd` or `w:tblCellMar`. */
  tableProps?: string;
  /** Give the document Word's "Normal Table" default style: zero indent, 108-twip side margins. */
  defaultTableStyle?: boolean;
}

export const cellMargins108 =
  '<w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>';

export function tableIndentDocx({ mode, tableProps = '', defaultTableStyle = false }: TableIndentOptions): Uint8Array {
  const borders = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right']
    .map(side => `<w:${side} w:val="single" w:sz="8" w:color="000000"/>`).join('') + '</w:tblBorders>';
  // SDK order places cell margins after borders and layout.
  const margins = tableProps.match(/<w:tblCellMar>[\s\S]*?<\/w:tblCellMar>/)?.[0] ?? '';
  const properties = margins ? tableProps.replace(margins, '') : tableProps;
  const body = '<w:p><w:r><w:t>IndentBefore</w:t></w:r></w:p>'
    + `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/>${properties}${borders}<w:tblLayout w:type="fixed"/>${margins}</w:tblPr>`
    + '<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>'
    + '<w:p><w:r><w:t>IndentCell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t>IndentAfter</w:t></w:r></w:p>';
  const tableStyle = defaultTableStyle
    ? '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr>'
      + `<w:tblInd w:w="0" w:type="dxa"/>${cellMargins108}</w:tblPr></w:style>`
    : '';
  const compat = mode === null ? '' : '<w:compat><w:compatSetting w:name="compatibilityMode" '
    + `w:uri="http://schemas.microsoft.com/office/word" w:val="${mode}"/></w:compat>`;
  const font = 'Liberation Serif';
  return storedZip([
    { name: '[Content_Types].xml', data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>`) },
    { name: '_rels/.rels', data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`) },
    { name: 'word/_rels/document.xml.rels', data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`) },
    { name: 'word/styles.xml', data: xml(`<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr>`
      + `<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>`
      + '<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="320" w:lineRule="exact"/></w:pPr></w:pPrDefault>'
      + `</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>${tableStyle}</w:styles>`) },
    { name: 'word/settings.xml', data: xml(`<w:settings xmlns:w="${W}">${compat}</w:settings>`) },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>`
      + '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
      + '</w:sectPr></w:body></w:document>') },
  ]);
}
