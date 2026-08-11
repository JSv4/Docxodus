import { storedZip, xml } from './docx-zip.js';

/**
 * A generated table whose colours come only from a table STYLE (issue #399): conditional
 * formatting for the header row and the banded rows, plus a table-level border colour. No cell in
 * the document declares a colour of its own, which is exactly the shape of the tracked `HC029`
 * benchmark case.
 *
 * The point of generating it is to make the theme-vs-cache question DECIDABLE. In a real Word file
 * the cached `w:fill`/`w:color` literal and the theme resolution agree, because Word rewrites the
 * cache whenever it applies the theme — so the tracked fixture cannot tell which one a renderer
 * used. Here the two can be made to DISAGREE on purpose, and then the rendered pixel says which
 * one won.
 */

const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** The theme's accent5, and the tints the style asks for. */
export const THEME_ACCENT5 = '4472C4';
/** tint 0x99 of accent5, per Word's tint formula: v*t + 255*(1-t), floored. */
export const ACCENT5_TINT_99 = '8EAADB';
/** tint 0x33 of accent5. */
export const ACCENT5_TINT_33 = 'D9E2F3';

/**
 * Deliberately WRONG cached literals, one colour per property so a failure names which one
 * regressed. A renderer painting these used the cache; one painting the theme-derived value
 * resolved the reference. Word treats the literal as a cache of the last resolution, so the
 * theme is what a conforming consumer must use.
 */
export const STALE_HEADER_FILL = 'FF0000';
export const STALE_BAND_FILL = '00FF00';
export const STALE_BORDER_COLOR = '0000FF';

export interface TableColorFixtureOptions {
  /**
   * When true, the style's cached `w:fill`/`w:color` literals are replaced with a stale value
   * that disagrees with the theme, isolating which source the renderer honours.
   */
  staleCache?: boolean;
}

export function generateTableColorDocx(options: TableColorFixtureOptions = {}): Uint8Array {
  const headerFill = options.staleCache ? STALE_HEADER_FILL : THEME_ACCENT5;
  const bandFill = options.staleCache ? STALE_BAND_FILL : ACCENT5_TINT_33;
  const borderColor = options.staleCache ? STALE_BORDER_COLOR : ACCENT5_TINT_99;

  // Word stamps each row and cell with the conditional formats that apply to it rather than
  // leaving a consumer to derive band membership from w:tblLook and the row index. Docxodus reads
  // those hints, so a generated table must carry them to exercise the same path a real file does.
  const cnf = (kind: 'firstRow' | 'oddHBand' | 'none') => {
    if (kind === 'none') return '';
    const val = kind === 'firstRow' ? '100000000000' : '000000100000';
    const first = kind === 'firstRow' ? '1' : '0';
    const odd = kind === 'oddHBand' ? '1' : '0';
    return `<w:cnfStyle w:val="${val}" w:firstRow="${first}" w:lastRow="0" w:firstColumn="0" ` +
      `w:lastColumn="0" w:oddVBand="0" w:evenVBand="0" w:oddHBand="${odd}" w:evenHBand="0" ` +
      `w:firstRowFirstColumn="0" w:firstRowLastColumn="0" w:lastRowFirstColumn="0" ` +
      `w:lastRowLastColumn="0"/>`;
  };

  const cell = (text: string, kind: 'firstRow' | 'oddHBand' | 'none') =>
    `<w:tc><w:tcPr>${cnf(kind)}<w:tcW w:w="3000" w:type="dxa"/></w:tcPr>` +
    `<w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const row = (cells: string[], kind: 'firstRow' | 'oddHBand' | 'none' = 'none') =>
    `<w:tr><w:trPr>${cnf(kind)}</w:trPr>${cells.map(c => cell(c, kind)).join('')}</w:tr>`;

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${w}">
  <w:body>
    <w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="GridTable4-Accent5"/>
        <w:tblW w:w="0" w:type="auto"/>
        <w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1"
          w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
      ${row(['Header A', 'Header B', 'Header C'], 'firstRow')}
      ${row(['Band one A', 'Band one B', 'Band one C'], 'oddHBand')}
      ${row(['Plain two A', 'Plain two B', 'Plain two C'])}
      ${row(['Band three A', 'Band three B', 'Band three C'], 'oddHBand')}
    </w:tbl>
    <w:p/>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
        w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${w}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
    <w:name w:val="Default Paragraph Font"/>
  </w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>
  <w:style w:type="table" w:styleId="GridTable4-Accent5">
    <w:name w:val="Grid Table 4 Accent 5"/><w:basedOn w:val="TableNormal"/>
    <w:tblPr>
      <w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="${borderColor}" w:themeColor="accent5" w:themeTint="99"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="${borderColor}" w:themeColor="accent5" w:themeTint="99"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="${borderColor}" w:themeColor="accent5" w:themeTint="99"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="${borderColor}" w:themeColor="accent5" w:themeTint="99"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="${borderColor}" w:themeColor="accent5" w:themeTint="99"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="${borderColor}" w:themeColor="accent5" w:themeTint="99"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tblStylePr w:type="firstRow">
      <w:rPr><w:b/><w:color w:val="FFFFFF" w:themeColor="background1"/></w:rPr>
      <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${headerFill}" w:themeFill="accent5"/></w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="band1Horz">
      <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${bandFill}" w:themeFill="accent5" w:themeFillTint="33"/></w:tcPr>
    </w:tblStylePr>
  </w:style>
</w:styles>`;

  const themeXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="${a}" name="Generated">
  <a:themeElements>
    <a:clrScheme name="Generated">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="${THEME_ACCENT5}"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Generated">
      <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Generated">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;

  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`),
    },
    {
      name: '_rels/.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: 'word/_rels/document.xml.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`),
    },
    // Style resolution needs the settings part; without it a style's CSS class is emitted with an
    // empty rule and every declared property is silently lost.
    { name: 'word/settings.xml', data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="${w}"/>`) },
    { name: 'word/theme/theme1.xml', data: xml(themeXml) },
    { name: 'word/styles.xml', data: xml(stylesXml) },
    { name: 'word/document.xml', data: xml(documentXml) },
  ]);
}
