import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Original decoration inventory, with single-line and disabled controls.
export function decorationFixture(runs?: [string, string][]): Uint8Array {
  return storedZip(Object.entries({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="rSettings" Type="${R_NS}/settings" Target="settings.xml"/></Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${(runs ?? [
      ['DoubleUnderline', '<w:u w:val="double"/>'],
      ['DoubleStrike', '<w:dstrike/>'],
      ['CombinedRules', '<w:strike/><w:u w:val="single"/>'],
      ['BlueDotted', '<w:u w:val="dotted" w:color="0000FF"/>'],
      ['SingleUnderline', '<w:u w:val="single"/>'],
      ['SingleStrike', '<w:strike/>'],
      ['IndependentRules', '<w:dstrike/><w:u w:val="dotted" w:color="0000FF"/>'],
      ['DisabledRules', '<w:strike w:val="0"/><w:dstrike w:val="false"/><w:u w:val="none"/>'],
    ]).map(([text, props]) => `<w:p><w:r><w:rPr>${props}</w:rPr><w:t>${text}</w:t></w:r></w:p>`).join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"/>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}

