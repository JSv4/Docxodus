import { tabDocument } from './docx-multiple-tab-fixture.js';

export const LONG_TAB_TEXT = Array.from({ length: 18 }, (_, i) =>
  `Wren${String(i + 1).padStart(2, '0')} The eastern store keeps every copper token beside the window.`).join(' ') + ' TAIL COMPLETE';

export function rightTabTailDocx({ alignment = 'right', longPrefix = false, longTail = true,
  mixedRun = false, tailHalfPoints = 24, tailSpacing = 0, tailText = '', box = '', prefixText = '' } = {}): Uint8Array {
  const prefix = prefixText || (longPrefix ? LONG_TAB_TEXT : 'Item');
  const tail = tailText || (longTail ? LONG_TAB_TEXT : 'FINISH');
  let body = '<w:p><w:r><w:t>Wren inventory</w:t></w:r></w:p><w:p><w:pPr><w:tabs>' +
    `<w:tab w:val="${alignment}" w:pos="8640" w:leader="dot"/></w:tabs></w:pPr>` +
    `<w:r><w:t>${prefix}</w:t></w:r><w:r><w:tab/>` +
    (mixedRun ? '' : `</w:r><w:r><w:rPr><w:spacing w:val="${tailSpacing}"/><w:sz w:val="${tailHalfPoints}"/></w:rPr>`) +
    `<w:t>${tail}</w:t></w:r></w:p>`;
  if (box) body = `<w:p><w:r><w:pict><v:${box} xmlns:v="urn:schemas-microsoft-com:vml" id="WrenBox" ` +
    'stroked="f" style="position:absolute;z-index:1;mso-position-horizontal-relative:page;' +
    'mso-position-vertical-relative:page;margin-left:72pt;margin-top:100pt;width:468pt;height:500pt">' +
    `<v:textbox inset="0,0,0,0"><w:txbxContent>${body}</w:txbxContent></v:textbox></v:${box}></w:pict></w:r></w:p>`;
  return tabDocument(body);
}
