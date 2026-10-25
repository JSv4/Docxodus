import {storedZip, xml, W_NS, R_NS} from './docx-zip.js';

export const OUR_ACCOUNT = 'S::pat@ourfirm.example::22222222-2222-2222-2222-222222222222';
export const OTHER_ACCOUNT = 'S::pat@counterparty.example::11111111-1111-1111-1111-111111111111';
export const IDENTITY_PROPOSAL = 'Counterparty proposal.';
const P = 'http://schemas.microsoft.com/office/word/2012/wordml';
const attr = (value: string) => value.replace(/[&<>"']/g, c => ({'&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&apos;'}[c]!));

/** Original pending paragraph, with the account recorded separately from its display name. */
export function revisionIdentityDocx(author = 'Pat', account: string | null = OTHER_ACCOUNT): Uint8Array {
  const types = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + [['document','document.main'],['styles','styles'],['settings','settings']].map(([name,type]) =>
      `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`).join('')
    + (account === null ? '' : '<Override PartName="/word/people.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"/>') + '</Types>';
  const stamp = `w:author="${attr(author)}" w:date="2026-05-01T09:00:00Z"`;
  const parts: Record<string,string> = {
    '[Content_Types].xml':types,
    '_rels/.rels':`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels':`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="rSettings" Type="${R_NS}/settings" Target="settings.xml"/>`
      + (account === null ? '' : '<Relationship Id="rPeople" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"/>') + '</Relationships>',
    'word/document.xml':`<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>Before.</w:t></w:r></w:p><w:p><w:pPr><w:rPr><w:ins w:id="1" ${stamp}/></w:rPr></w:pPr><w:ins w:id="2" ${stamp}><w:r><w:t>${IDENTITY_PROPOSAL}</w:t></w:r></w:ins></w:p><w:p><w:r><w:t>After.</w:t></w:r></w:p></w:body></w:document>`,
    'word/styles.xml':`<w:styles xmlns:w="${W_NS}"/>`,
    'word/settings.xml':`<w:settings xmlns:w="${W_NS}"/>`,
  };
  if(account !== null) parts['word/people.xml'] = `<w15:people xmlns:w15="${P}"><w15:person w15:author="${attr(author)}"><w15:presenceInfo w15:providerId="AD" w15:userId="${attr(account)}"/></w15:person></w15:people>`;
  return storedZip(Object.entries(parts).map(([name,value]) => ({name,data:xml(value)})));
}
