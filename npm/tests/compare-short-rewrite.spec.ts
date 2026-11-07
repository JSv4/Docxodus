import { test, expect } from '@playwright/test';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';
import { readZipEntry } from './zip.js';

function document(paragraphs: string[]): Uint8Array {
  const relations = (content: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${content}</Relationships>`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>
      ${['document', 'styles', 'settings'].map(name => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name}${name === 'document' ? '.main' : ''}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': relations(`<Relationship Id="main" Type="${R_NS}/officeDocument" Target="word/document.xml"/>`),
    'word/_rels/document.xml.rels': relations(['styles', 'settings'].map(name =>
      `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('')),
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${paragraphs.map(text =>
      `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`).join('')}<w:sectPr/></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`,
  };
  return storedZip(Object.entries(parts).map(([name, content]) => ({ name, data: xml(content) })));
}

for (const kind of ['short', 'numbered', 'long', 'interior']) for (const reverse of [false, true]) {
  test(`${kind} complete rewrite, reverse=${reverse}: coherent review and exact resolved paragraphs`, async ({ page }) => {
    let before = ['Garden club minutes', 'Roses need pruning.', 'Members voted on the budget.', 'Dues rise in March.'];
    let after = ['Library opening hours', 'Volunteers shelve donated novels every Saturday morning.',
      'Children meet in the reading room.', 'Parking stays free.'];
    if (kind === 'numbered') {
      before = before.map((text, i) => `${i + 1}. ${text}`);
      after = after.map((text, i) => `${i + 1}. ${text}`);
    }
    if (kind === 'long') {
      before = before.map(text => Array(5).fill(text).join(' '));
      after = after.map(text => Array(5).fill(text).join(' '));
    }
    if (kind === 'interior') {
      before = ['Unchanged opening.', ...before, 'Unchanged closing.'];
      after = ['Unchanged opening.', ...after, 'Unchanged closing.'];
    }
    if (reverse) [before, after] = [after, before];
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    const outputs = await page.evaluate(([left, right]) => {
      const T = (window as any).DocxodusTests;
      const compared = T.compareDocuments(new Uint8Array(left), new Uint8Array(right), 'Fixture');
      if (compared.error) throw new Error(JSON.stringify(compared.error));
      return [compared, T.docxDiffAcceptRevisions(compared.docxBytes), T.docxDiffRejectRevisions(compared.docxBytes)]
        .map(result => {
          if (result.error) throw new Error(JSON.stringify(result.error));
          return Array.from(result.docxBytes) as number[];
        });
    }, [Array.from(document(before)), Array.from(document(after))]);
    const views = await page.evaluate(({ parts, ns }) => parts.map(part => {
      const doc = new DOMParser().parseFromString(part, 'application/xml');
      return Array.from(doc.getElementsByTagNameNS(ns, 'p')).map(p => ({
        text: Array.from(p.getElementsByTagNameNS(ns, 't')).map(t => t.textContent).join(''),
        deleted: Array.from(p.getElementsByTagNameNS(ns, 'delText')).map(t => t.textContent).join(''),
        inserted: Array.from(p.getElementsByTagNameNS(ns, 'ins')).flatMap(ins =>
          Array.from(ins.getElementsByTagNameNS(ns, 't'))).map(t => t.textContent).join(''),
      }));
    }), { parts: outputs.map(bytes => readZipEntry(Buffer.from(bytes), 'word/document.xml')), ns: W_NS });
    expect(views[1].map(p => p.text)).toEqual(after);
    expect(views[2].map(p => p.text)).toEqual(before);
    if (kind === 'short' || kind === 'numbered') {
      expect(views[0].filter(p => p.deleted).map(p => p.deleted)).toEqual(before);
      expect(views[0].filter(p => p.inserted).map(p => p.inserted)).toEqual(after);
      let previousEnd = -1;
      for (let i = 0; i < before.length; i++) {
        const oldIndex = views[0].findIndex(p => p.deleted === before[i]);
        const newIndex = views[0].findIndex(p => p.inserted === after[i]);
        expect(Math.abs(oldIndex - newIndex)).toBeLessThanOrEqual(1);
        expect(Math.min(oldIndex, newIndex)).toBeGreaterThan(previousEnd);
        previousEnd = Math.max(oldIndex, newIndex);
      }
    } else {
      const sides = views[0].flatMap(p => [...(p.inserted ? ['insert'] : []), ...(p.deleted ? ['delete'] : [])]);
      expect(sides.slice(1).filter((side, i) => side !== sides[i])).toHaveLength(1);
    }
  });
}
