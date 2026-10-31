import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { rubyDocx, type RubyFixture } from './docx-ruby-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const kind of ['simple', 'zero-raise', 'multiple', 'long', 'plain', 'long-plain', 'vertical', 'tab'] as RubyFixture[]) {
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
      if (kind === 'simple' || kind === 'zero-raise') {
        const base = items.find(item => item.str.includes('CEDAR'))!, rt = items.find(item => item.str === 'reading')!;
        expect(rt.transform[0]).toBeCloseTo(6, 1);
        expect(rt.transform[5] - base.transform[5]).toBeCloseTo(kind === 'simple' ? 11 : 0, 1);
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
        const ruby = nodes(tree!).filter(node => node.role === 'Ruby');
        expect(ruby).toHaveLength(1);
        expect(nodes(ruby[0]).filter(node => node.role === 'RT').map(textOf)).toEqual(['reading']);
        const ordered = textOf(tree!);
        expect(ordered.indexOf('CEDAR')).toBeLessThan(ordered.indexOf('reading'));
        expect(ordered.indexOf('reading')).toBeLessThan(ordered.indexOf('CedarAfter'));
        expect(ordered.indexOf('CedarAfter')).toBeLessThan(ordered.indexOf('PineTail'));
      }
    } finally { await task.destroy(); }
    if (unsupported) await expect(renderDocxArtifacts(rubyDocx(kind), {
      browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict', timeoutMs: 120_000,
    })).rejects.toThrow(/unsupported/i);
  });
}
