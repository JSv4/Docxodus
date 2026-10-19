import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { readZipEntry } from './zip.js';

const cases = [
  ...['body', 'header', 'footer'].flatMap(story => ['edited', 'unchanged-control'].map(kind =>
    [`removed-comment-${story}-${kind}`, 'compare', story === 'body' ? 'word/document.xml' : `word/${story}1.xml`, '1', '0'])),
  ['comment-add-customxml', 'compare'],
  ['comment-identity-reused-id', 'compare'],
  ['comment-identity-replaced-id', 'compare'],
  ['comment-identity-distinct-control', 'compare'],
  ['comment-identity-reviewer-collision', 'consolidate'],
  ['comment-identity-reviewer-identical', 'consolidate'],
  ['comment-identity-reviewer-control', 'consolidate'],
  ['comment-remove-sdt', 'compare'],
  ['comment-wrapper-control', 'compare'],
  ['reviewer-added-comment', 'consolidate'],
  ['reviewer-comment-control', 'consolidate'],
  ['duplicate-comment-sdt', 'compare'],
  ['duplicate-comment-sdt-control', 'compare'],
  ['duplicate-comment-customxml', 'compare'],
  ['duplicate-comment-customxml-control', 'compare'],
];
const fixture = (name: string) => readFileSync(new URL(`../../TestFiles/CommentViews/${name}.docx`, import.meta.url));

for (const [name, operation, part = 'word/document.xml', final = 'final', original = 'original'] of cases) {
  test(`${name} preserves comment anchors through browser accept/reject`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
    const inputs = [0, 1, ...(operation === 'consolidate' ? [2] : [])].map(i => Array.from(fixture(`${name}-${i}`)));
    const result = await page.evaluate(({ inputs, operation }) => {
      const T = (window as any).DocxodusTests;
      const bytes = inputs.map(b => new Uint8Array(b));
      const redline = operation === 'compare' ? T.docxDiffCompare(bytes[0], bytes[1]) :
        T.docxDiffConsolidate(bytes[0], bytes.slice(1).map((document, i) => ({ document, author: `Reviewer${i + 1}` })));
      if (redline.error) throw new Error(JSON.stringify(redline.error));
      return [redline, T.docxDiffAcceptRevisions(redline.docxBytes), T.docxDiffRejectRevisions(redline.docxBytes)].map(r => {
        if (r.error) throw new Error(JSON.stringify(r.error));
        return Array.from(r.docxBytes) as number[];
      });
    }, { inputs, operation });
    const packages = [...result.map(bytes => Buffer.from(bytes)), fixture(`${name}-${final}`), fixture(`${name}-${original}`)];
    const xml = packages.map(bytes => {
      const body = readZipEntry(bytes, part);
      return { body, comments: body.includes('commentReference') ? readZipEntry(bytes, 'word/comments.xml') : '<comments/>' };
    });
    const views = await page.evaluate(xml => xml.map(({ body, comments }) => {
      const parse = (s: string) => new DOMParser().parseFromString(s, 'application/xml');
      const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
      const document = parse(body);
      const definitions = new Map(Array.from(parse(comments).getElementsByTagNameNS(ns, 'comment')).map(c =>
        [c.getAttributeNS(ns, 'id'), c.getAttributeNS(ns, 'author') + '|' + c.getAttributeNS(ns, 'date') + '|' +
          Array.from(c.getElementsByTagNameNS(ns, 't')).map(t => t.textContent).join('')]));
      const ranges = Array.from(document.getElementsByTagNameNS(ns, 'commentRangeStart')).map(e => e.getAttributeNS(ns, 'id'));
      if (new Set(ranges).size !== ranges.length) throw new Error('Duplicate comment range ID');
      function walk(e: Element): unknown[] {
        if (e.localName === 't') return [e.textContent];
        if (['commentRangeStart', 'commentRangeEnd', 'commentReference'].includes(e.localName)) {
          const definition = definitions.get(e.getAttributeNS(ns, 'id'));
          if (!definition) throw new Error('Unresolved comment marker');
          return [[e.localName, definition]];
        }
        const children: unknown[] = [];
        for (const child of Array.from(e.children).flatMap(walk)) {
          const last = children[children.length - 1];
          if (typeof last === 'string' && typeof child === 'string') children[children.length - 1] = last + child;
          else children.push(child);
        }
        if (e.localName === 'sdt' || e.localName === 'customXml') {
          const tag = e.localName === 'sdt'
            ? e.getElementsByTagNameNS(ns, 'tag')[0]?.getAttributeNS(ns, 'val')
            : e.getAttributeNS(ns, 'element');
          return [[e.localName, tag, children]];
        }
        return ['p', 'tbl', 'tr', 'tc'].includes(e.localName) ? [[e.localName, children]] : children;
      }
      const authors = Array.from(document.getElementsByTagNameNS(ns, 'ins'))
        .filter(e => e.getElementsByTagNameNS(ns, 'commentReference').length > 0)
        .map(e => e.getAttributeNS(ns, 'author'));
      return { view: walk(document.documentElement), authors };
    }), xml);
    expect(views[1].view).toEqual(views[3].view);
    expect(views[2].view).toEqual(views[4].view);
    expect(views[3].view).not.toEqual(views[4].view);
    if (name === 'reviewer-added-comment') expect(views[0].authors).toEqual(['Reviewer1']);
  });
}
