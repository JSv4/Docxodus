import { test, expect, type Page } from '@playwright/test';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Leader-and-page-number lines: a label, a left tab, a title, a right dot-leader tab at the
// column's right edge (9360 twips = 624px), and a page number. Each variant adds one inert item
// that used to push the right tab back to a fixed server-measured width (issue #1275), and the
// right indent ends the line 864 twips before the stop, which used to break every entry.
const entries = [
  ['2.1.', 'Quarry Access Roads', '3'],
  ['2.2.', 'Spillway', '12'],
  ['2.3.', 'Culvert Inspection and Annual Load Ratings', '27'],
  ['2.4.', 'Riprap', '140'],
] as const;
const variants = ['plain', 'bookmark', 'hyphen', 'underline', 'strike'] as const;

function leaderLineFixture(variant: typeof variants[number], rightIndent: boolean, stop = 9360): Uint8Array {
  const tabRPr = variant === 'underline' ? '<w:rPr><w:u w:val="single"/></w:rPr>'
    : variant === 'strike' ? '<w:rPr><w:strike/></w:rPr>' : '';
  const body = entries.map(([label, title, number], i) =>
    `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1296"/><w:tab w:val="right" w:leader="dot" w:pos="${stop}"/></w:tabs>` +
    `<w:ind w:left="1296"${rightIndent ? ' w:right="864"' : ''} w:hanging="648"/></w:pPr>` +
    (variant === 'bookmark' ? `<w:bookmarkStart w:id="${i}" w:name="_Toc${i}"/><w:bookmarkEnd w:id="${i}"/>` : '') +
    `<w:r><w:t>${label}</w:t></w:r><w:r>${tabRPr}<w:tab/></w:r><w:r><w:t>${title}</w:t></w:r>` +
    (variant === 'hyphen' ? '<w:r><w:noBreakHyphen/><w:t>West</w:t></w:r>' : '') +
    `<w:r>${tabRPr}<w:tab/><w:t>${number}</w:t></w:r></w:p>`).join('');
  return storedZip(Object.entries({
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${['document', 'styles', 'settings'].map(p => `<Override PartName="/word/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${p === 'document' ? 'document.main' : p}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'settings'].map(p => `<Relationship Id="r${p}" Type="${R_NS}/${p}" Target="${p}.xml"/>`).join('')}</Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}

async function layOut(page: Page, variant: typeof variants[number], rightIndent: boolean, stop: number) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  const results = await page.evaluate(({ bytes, entries }) => {
    const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Leader lines', 'docx-', false, '', -1, 'comment-',
      0, 1, 'page-', false, 0, 'annot-', false, false, false, true, true);
    if (html.startsWith('{')) throw new Error(html);
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    document.head.innerHTML = parsed.head.innerHTML;
    document.body.style.cssText = 'margin:0;padding:0';
    document.body.innerHTML = `<div id="host" style="width:6.5in;margin:0;padding:0">${parsed.body.innerHTML}</div>`;
    const host = document.querySelector<HTMLElement>('#host')!;
    const origin = host.getBoundingClientRect().left;
    return entries.map(([label, , number]) => {
      const paragraph = Array.from(host.querySelectorAll('p')).find(p => p.textContent?.includes(label))!;
      const leaf = (text: string) => Array.from(paragraph.querySelectorAll<HTMLElement>('span'))
        .find(e => e.childElementCount === 0 && e.textContent === text)!.getBoundingClientRect();
      const labelRect = leaf(label), numberRect = leaf(number);
      return { labelTop: labelRect.top, numberTop: numberRect.top, numberRight: numberRect.right - origin };
    });
  }, { bytes: Array.from(leaderLineFixture(variant, rightIndent, stop)), entries });
  for (const entry of results) {
    expect(entry.numberTop).toBeCloseTo(entry.labelTop, 0);
    expect(entry.numberRight).toBeCloseTo(stop / 15, 0);
  }
}

for (const variant of variants) for (const rightIndent of [true, false]) {
  test(`leader lines keep one line and end the page number on the right stop (${variant}, right indent ${rightIndent})`, async ({ page }) => {
    await layOut(page, variant, rightIndent, 9360);
  });
}

// A stop past the column edge (10000 twips = 666.67px in a 624px column): LibreOffice 25.8 keeps
// each entry on one line and ends the number on the stop, in the page margin.
for (const rightIndent of [true, false]) {
  test(`a right stop past the column edge keeps one line with the number on the stop (right indent ${rightIndent})`, async ({ page }) => {
    await layOut(page, 'plain', rightIndent, 10000);
  });
}
