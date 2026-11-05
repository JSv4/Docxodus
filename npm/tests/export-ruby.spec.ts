import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { rubyDocx, type RubyFixture } from './docx-ruby-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const kind of ['simple', 'zero-raise', 'high-raise', 'multiple', 'long', 'plain', 'long-plain', 'vertical', 'tab'] as RubyFixture[]) {
  test(`ruby ${kind} retains both text layers in the delivered PDF`, async ({ browser }) => {
    test.setTimeout(120_000);
    const unsupported = kind === 'vertical' || kind === 'tab';
    const result = await renderDocxArtifacts(rubyDocx(kind), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
      unsupportedContent: unsupported ? 'warn' : 'strict',
    });
    expect(result.html!.includes('data-content-type="RubyAnnotation"')).toBe(unsupported);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const items = (await Promise.all(Array.from({ length: pdf.numPages }, async (_, i) =>
        (await (await pdf.getPage(i + 1)).getTextContent()).items.filter(item => 'str' in item)))).flat();
      const text = items.map(item => item.str).join('');
      expect(text.match(/CEDAR/g)).toHaveLength(1);
      expect(text.match(/reading/g)).toHaveLength(1);
      if (kind.startsWith('long')) {
        expect(pdf.numPages).toBeGreaterThan(1);
        expect(text.match(/Pine\d{3}/g)).toEqual(Array.from({ length: 100 }, (_, i) => `Pine${String(i).padStart(3, '0')}`));
      }
      if (kind === 'simple' || kind === 'zero-raise' || kind === 'high-raise') {
        const base = items.find(item => item.str.includes('CEDAR'))!, rt = items.find(item => item.str === 'reading')!;
        const raise = kind === 'zero-raise' ? 0 : kind === 'high-raise' ? 22 : 11;
        expect(rt.transform[0]).toBeCloseTo(6, 1);
        // Chromium 143 rounds relative-positioned glyphs to CSS pixels when printing,
        // even though DOM baseline geometry keeps 1/64px precision. The 0/22pt cases
        // still reject an omitted raise correction, unlike an 11pt-only assertion.
        expect(Math.abs(rt.transform[5] - base.transform[5] - raise)).toBeLessThanOrEqual(0.75);
        const preview = await browser.newPage();
        try {
          await preview.setContent(result.html!);
          await preview.evaluate(() => document.fonts.ready);
          const gap = await preview.evaluate(() => {
            const baseline = (element: Element) => {
              const marker = document.createElement('span');
              marker.style.cssText = 'display:inline-block;width:0;height:0;vertical-align:baseline;margin:0;padding:0;border:0';
              element.append(marker); const y = marker.getBoundingClientRect().top; marker.remove(); return y;
            };
            const ruby = document.querySelector('ruby')!;
            return (baseline(ruby.querySelector('[data-docx-ruby-base]')!) - baseline(ruby.querySelector('rt')!)) * 0.75;
          });
          expect(gap).toBeCloseTo(raise, 1);
        } finally { await preview.close(); }
        // Paint order can differ from reading order. Tagged export must keep the annotation
        // associated with its base before the following prose in the semantic structure.
        const page = await pdf.getPage(1), byId = new Map<string, string>(), stack: Array<string | null> = [];
        for (const item of (await page.getTextContent({ includeMarkedContent: true })).items) {
          if ('str' in item) {
            const id = stack[stack.length - 1];
            if (id) byId.set(id, (byId.get(id) ?? '') + item.str);
          } else if (item.type === 'endMarkedContent') stack.pop();
          else stack.push('id' in item && item.id ? item.id : stack[stack.length - 1] ?? null);
        }
        type Tree = { role?: string; type?: string; id?: string; children?: Tree[] };
        const tree = await page.getStructTree() as Tree | null;
        expect(tree).not.toBeNull();
        const textOf = (node: Tree): string => node.type === 'content' ? byId.get(node.id!) ?? ''
          : (node.children ?? []).map(textOf).join('');
        const nodes = (node: Tree): Tree[] => [node, ...(node.children ?? []).flatMap(nodes)];
        // Chromium 143 uses NonStruct for ruby; newer Chromium emits Ruby/RT.
        // Either representation must associate the pair before following prose.
        const pair = nodes(tree!).filter(node => {
          const text = textOf(node); return text.includes('CEDAR') && text.includes('reading') && !text.includes('CedarAfter');
        });
        expect(pair.length).toBeGreaterThan(0);
        expect(textOf(pair[pair.length - 1]).replace(/\s/g, '')).toBe('CEDARreading');
        const ordered = textOf(tree!);
        expect(ordered.indexOf('CEDAR')).toBeLessThan(ordered.indexOf('reading'));
        expect(ordered.indexOf('reading')).toBeLessThan(ordered.indexOf('CedarAfter'));
        expect(ordered.indexOf('CedarAfter')).toBeLessThan(ordered.indexOf('PineTail'));
      }
    } finally { await task.destroy(); }
    if (unsupported) await expect(renderDocxArtifacts(rubyDocx(kind), {
      browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict', timeoutMs: 120_000,
    })).rejects.toMatchObject({ code: 'resource_policy_failure', phase: 'docx_conversion',
      message: 'RubyAnnotation is represented by a visible placeholder.' });
  });
}
