import { fittingImageRun, pageFittingDocx } from './docx-page-fitting-fixture.js';

export function linkedTextBoxDocx(mode: 'flow' | 'fitting' | 'unlinked' | 'image' | 'cross-page' | 'tab'): Uint8Array {
  const words = Array.from({ length: 70 }, (_, i) => `Alder${String(i).padStart(3, '0')}`).join(' ');
  const runs = mode === 'image'
    ? `<w:r><w:t xml:space="preserve">${words.split(' ').slice(0, 35).join(' ')} </w:t></w:r>${fittingImageRun(3, 20, 10)}<w:r><w:t xml:space="preserve"> ${words.split(' ').slice(35).join(' ')}</w:t></w:r>`
    : `<w:r><w:t>${words}</w:t></w:r>${mode === 'tab' ? '<w:r><w:tab/><w:t>CedarClosing</w:t></w:r>' : ''}`;
  const paragraphs = `<w:p><w:pPr><w:keepLines w:val="0"/><w:widowControl w:val="0"/><w:spacing w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>${runs}</w:p>`;
  const box = (continuation: boolean) => {
    const width = continuation ? 216 : 144, height = continuation ? 216 : mode === 'fitting' || mode === 'unlinked' ? 360 : 60;
    return `<w:drawing><wp:anchor simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${(continuation ? 240 : 72) * 12700}</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="page"><wp:posOffset>1524000</wp:posOffset></wp:positionV><wp:extent cx="${width * 12700}" cy="${height * 12700}"/>
      <wp:wrapNone/><wp:docPr id="${continuation ? 2 : 1}" name="Alder flow frame"/>
      <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        ${continuation ? '<wps:linkedTxbx id="37" seq="1"/>' : `<wps:txbx id="37"><w:txbxContent>${paragraphs}</w:txbxContent></wps:txbx>`}
        <wps:bodyPr lIns="38100" rIns="38100" tIns="38100" bIns="38100"><a:noAutofit/></wps:bodyPr>
      </wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
  };
  return pageFittingDocx(`<w:p><w:r><w:t>Alder linked frame inventory.</w:t></w:r><w:r>${box(false)}${mode === 'unlinked' || mode === 'cross-page' ? '' : box(true)}</w:r></w:p>${mode === 'cross-page' ? `<w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Alder continuation carrier.</w:t></w:r><w:r>${box(true)}</w:r></w:p>` : ''}`);
}
