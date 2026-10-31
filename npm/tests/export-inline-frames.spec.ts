import { test, expect } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { inlineFrameFixture } from './docx-inline-frame-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true } });

for (const story of ['body', 'header', 'footer', 'table'] as const) {
  for (const mode of ['notBeside', 'around', 'none', 'normal'] as const) {
    test(`inline ${story} frame honors ${mode} text flow`, async ({ browser, page }, info) => {
      const normal = mode === 'normal', wrap = normal ? 'notBeside' : mode;
      const result = await renderDocxArtifacts(inlineFrameFixture({ story, wrap, normal }), {
        browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
      });
      writeFileSync(info.outputPath('frame.pdf'), result.pdf!);
      expect(result.pageCount).toBe(1);
      const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
      try {
        const pdf = await task.promise;
        const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item);
        const word = (text: string) => {
          const matches = items.filter(item => item.str === text);
          expect(matches, `${text} survives once`).toHaveLength(1);
          return matches[0];
        };
        const before = word('SageBefore'), first = word('SageFrameFirst'), last = word('SageFrameLast'), after = word('SageAfter');
        expect(before.transform[5] - first.transform[5]).toBeCloseTo(12, 0);
        expect(first.transform[5] - last.transform[5]).toBeCloseTo(12, 0);
        expect(first.transform[4]).toBeCloseTo(story === 'table' || normal ? 72 : 84, 0);
        const below = story === 'table' || normal || wrap === 'notBeside';
        expect(first.transform[5] - after.transform[5]).toBeCloseTo(below ? 24 : 0, 0);
        expect(after.transform[4]).toBeCloseTo(story !== 'table' && wrap === 'around' ? 264 : 72, 0);
        if (story === 'footer') {
          expect(word('SageDocument').transform[5] - before.transform[5]).toBeCloseTo(below ? 636 : 660, 0);
        }
      } finally { await task.destroy(); }
      await page.setContent(result.html!);
      const anchors = await page.locator('.page-box p').filter({ hasText: /^SageFrame(?:First|Last)$/ })
        .evaluateAll(paragraphs => paragraphs.map(p => (p as HTMLElement).dataset.sourceAnchorId));
      expect(new Set(anchors).size).toBe(2);
      for (const anchor of anchors) {
        expect(anchor).toBeTruthy();
        const fragments = result.pageMap.fragments.filter(fragment => fragment.anchorId === anchor);
        expect(fragments).toHaveLength(1);
        expect(fragments[0]).toMatchObject({ pageNumber: 1, story: story === 'table' ? 'body' : story, inTableCell: story === 'table' });
        expect(fragments[0].geometry.height).toBeGreaterThan(0);
      }
    });
  }
}

for (const wrap of ['notBeside', 'around'] as const) {
  test(`inline ${wrap} frame at the page bottom preserves the following paragraph`, async ({ browser }) => {
    const result = await renderDocxArtifacts(inlineFrameFixture({ wrap, spill: true }), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    });
    expect(result.pageCount).toBe(wrap === 'notBeside' ? 2 : 1);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const words: { text: string; page: number; x: number; y: number }[] = [];
      for (let page = 1; page <= pdf.numPages; page++) {
        const items = (await (await pdf.getPage(page)).getTextContent()).items;
        words.push(...items.filter(item => 'str' in item).filter(item => item.str.trim())
          .map(item => ({ text: item.str, page, x: item.transform[4], y: item.transform[5] })));
      }
      for (const label of Array.from({ length: 52 }, (_, i) => `SageFiller${String(i).padStart(3, '0')}`)) {
        expect(words.filter(word => word.text === label)).toEqual([expect.objectContaining({ text: label, page: 1 })]);
      }
      for (const text of ['SageFrameFirst', 'SageFrameLast']) expect(words.filter(word => word.text === text)).toEqual([expect.objectContaining({ text, page: 1 })]);
      expect(words.filter(word => word.text === 'SageAfter')).toEqual([expect.objectContaining({ text: 'SageAfter', page: wrap === 'notBeside' ? 2 : 1 })]);
      if (wrap === 'around') {
        const after = words.find(word => word.text === 'SageAfter')!, first = words.find(word => word.text === 'SageFrameFirst')!;
        expect(after.x).toBeCloseTo(264, 0);
        expect(after.y).toBeCloseTo(first.y, 1);
      }
    } finally { await task.destroy(); }
  });
}

for (const wrap of ['notBeside', 'around'] as const) for (const fillerCount of [53, 54]) {
  test(`inline ${wrap} frame moves whole after ${fillerCount} body lines`, async ({ browser }) => {
    const result = await renderDocxArtifacts(inlineFrameFixture({ wrap, fillerCount }), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    });
    expect(result.pageCount).toBe(2);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      for (let page = 1; page <= 2; page++) {
        const items = (await (await pdf.getPage(page)).getTextContent()).items.filter(item => 'str' in item);
        const labels = items.map(item => item.str).filter(text => ['SageFrameFirst', 'SageFrameLast', 'SageAfter'].includes(text));
        expect(labels.sort()).toEqual(page === 1 ? [] : ['SageAfter', 'SageFrameFirst', 'SageFrameLast']);
      }
    } finally { await task.destroy(); }
  });
}

for (const story of ['header', 'footer'] as const) for (const wrap of ['notBeside', 'around'] as const) {
  test(`inline ${story} ${wrap} frame preserves paragraph spacing and body clearance`, async ({ browser }) => {
    const result = await renderDocxArtifacts(inlineFrameFixture({ story, wrap, spaceBefore: 6, spaceAfter: 4 }), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    });
    expect(result.pageCount).toBe(1);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item);
      const word = (text: string) => {
        const matches = items.filter(item => item.str === text);
        expect(matches).toHaveLength(1);
        return matches[0];
      };
      const before = word('SageBefore'), first = word('SageFrameFirst'), last = word('SageFrameLast'), after = word('SageAfter'), body = word('SageDocument');
      // The converter's existing inch quantization and Word's substitute font differ by <0.8pt.
      const near = (actual: number, expected: number) => expect(Math.abs(actual - expected)).toBeLessThan(0.8);
      near(before.transform[5] - first.transform[5], 22);
      near(first.transform[5] - last.transform[5], 18);
      near(before.transform[5] - after.transform[5], wrap === 'around' ? 18 : 58);
      if (story === 'footer') near(body.transform[5] - before.transform[5], wrap === 'around' ? 644 : 604);
      else {
        near(before.transform[5] - body.transform[5], wrap === 'around' ? 40 : 80);
        near(body.transform[4], wrap === 'around' ? 264 : 72);
      }
    } finally { await task.destroy(); }
  });
}

for (const wrap of ['around', 'notBeside'] as const) {
  test(`inline ${wrap} horizontal text distance preserves frame paint`, async ({ browser }) => {
    const result = await renderDocxArtifacts(inlineFrameFixture({ wrap, horizontalTextDistance: 6 }), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      expect(pdf.numPages).toBe(1);
      const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item);
      const [before, first, last, after] = ['SageBefore', 'SageFrameFirst', 'SageFrameLast', 'SageAfter'].map(text => {
        const matches = items.filter(item => item.str === text);
        expect(matches).toHaveLength(1);
        return matches[0];
      });
      expect(first.transform[4]).toBeCloseTo(84, 0);
      expect(first.transform[5] - last.transform[5]).toBeCloseTo(12, 0);
      expect(before.transform[4]).toBeCloseTo(72, 0);
      expect(after.transform[4]).toBeCloseTo(wrap === 'around' ? 270 : 72, 0);
      expect(first.transform[5] - before.transform[5]).toBeCloseTo(-12, 0);
      expect(first.transform[5] - after.transform[5]).toBeCloseTo(wrap === 'around' ? 0 : 24, 0);
      expect(result.renderReport.unsupportedContent).toEqual(wrap === 'around'
        ? [expect.objectContaining({ contentType: 'Other', action: 'placeholder' })] : []);
    } finally { await task.destroy(); }
  });
}

for (const [name, frameOptions, left] of [
  ['centered wrapping', { wrap: 'around', horizontalAlign: 'center' }, 216],
  ['right wrapping', { wrap: 'around', horizontalAlign: 'right' }, 360],
  ['page-relative wrapping', { wrap: 'around', horizontalAnchor: 'page' }, 12],
  ['around vertical distance', { wrap: 'around', verticalTextDistance: 4 }, 84],
  ['notBeside vertical distance', { wrap: 'notBeside', verticalTextDistance: 4 }, 84],
] as const) {
  test(`inline ${name} reports complete-flow approximation and honors strict policy`, async ({ browser, page }) => {
    const input = inlineFrameFixture(frameOptions);
    const options = { browser, outputs: ['html', 'pdf'] as const, reviewProfile: 'final' as const,
      commentProfile: 'hidden' as const, timeoutMs: 120_000 };
    const result = await renderDocxArtifacts(input, options);
    expect(result.warnings.map(warning => warning.code)).toContain('unsupported_content_placeholder');
    expect(result.renderReport.unsupportedContent).toEqual([
      expect.objectContaining({ contentType: 'Other', action: 'placeholder' }),
    ]);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      expect(pdf.numPages).toBe(1);
      const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item);
      const [before, first, last, after] = ['SageBefore', 'SageFrameFirst', 'SageFrameLast', 'SageAfter'].map(text => {
        const matches = items.filter(item => item.str === text);
        expect(matches).toHaveLength(1);
        return matches[0];
      });
      expect(first.transform[4]).toBeCloseTo(left, 0);
      expect(before.transform[5] - first.transform[5]).toBeCloseTo(12, 0);
      expect(first.transform[5] - last.transform[5]).toBeCloseTo(12, 0);
      expect(last.transform[5] - after.transform[5]).toBeCloseTo(12, 0);
      expect(after.transform[4]).toBeCloseTo(72, 0);
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    for (const text of ['SageBefore', 'SageFrameFirst', 'SageFrameLast', 'SageAfter']) {
      const id = await page.locator('.page-box p').filter({ hasText: new RegExp(`^${text}$`) }).getAttribute('data-source-anchor-id');
      expect(id).toBeTruthy();
      expect(result.pageMap.fragments.filter(fragment => fragment.anchorId === id)).toEqual([
        expect.objectContaining({ pageNumber: 1, story: 'body' }),
      ]);
    }
    await expect(renderDocxArtifacts(input, { ...options, unsupportedContent: 'strict' }))
      .rejects.toMatchObject({ code: 'resource_policy_failure' });
  });
}

test('inline notBeside right alignment keeps exact flow without an approximation warning', async ({ browser }) => {
  const result = await renderDocxArtifacts(inlineFrameFixture({ horizontalAlign: 'right' }), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict', timeoutMs: 120_000,
  });
  expect(result.renderReport.unsupportedContent).toEqual([]);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item);
    const first = items.find(item => item.str === 'SageFrameFirst')!, after = items.find(item => item.str === 'SageAfter')!;
    expect(first.transform[4]).toBeCloseTo(360, 0);
    expect(first.transform[5] - after.transform[5]).toBeCloseTo(24, 0);
    expect(after.transform[4]).toBeCloseTo(72, 0);
  } finally { await task.destroy(); }
});

test('inline left-edge horizontal text distance remains exact under strict policy', async ({ browser }) => {
  const result = await renderDocxArtifacts(inlineFrameFixture({ wrap: 'around', horizontalOffset: '0', horizontalTextDistance: 6 }), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict', timeoutMs: 120_000,
  });
  expect(result.renderReport.unsupportedContent).toEqual([]);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item);
    expect(items.find(item => item.str === 'SageFrameFirst')!.transform[4]).toBeCloseTo(72, 0);
    expect(items.find(item => item.str === 'SageAfter')!.transform[4]).toBeCloseTo(258, 0);
  } finally { await task.destroy(); }
});

test('inline explicit interior offset diagnoses conservative one-sided wrapping', async ({ browser }) => {
  const input = inlineFrameFixture({ wrap: 'around', horizontalOffset: '2880', tailWords: 16 });
  const options = { browser, outputs: ['pdf'] as const, reviewProfile: 'final' as const, commentProfile: 'hidden' as const, timeoutMs: 120_000 };
  const result = await renderDocxArtifacts(input, options);
  expect(result.renderReport.unsupportedContent).toEqual([expect.objectContaining({ contentType: 'Other' })]);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(1);
    const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item);
    const tokens = items.flatMap(item => item.str.trim().split(/\s+/)).filter(Boolean);
    for (const text of ['SageBefore', 'SageFrameFirst', 'SageFrameLast',
      ...Array.from({ length: 16 }, (_, i) => `SageTail${String(i).padStart(3, '0')}`)]) {
      expect(tokens.filter(token => token === text)).toHaveLength(1);
    }
    const frame = items.find(item => item.str === 'SageFrameFirst')!, tail = items.find(item => item.str.startsWith('SageTail000'))!;
    expect(frame.transform[4]).toBeCloseTo(216, 0);
    expect(tail.transform[4]).toBeCloseTo(396, 0);
    expect(tail.transform[5]).toBeCloseTo(frame.transform[5], 0);
  } finally { await task.destroy(); }
  await expect(renderDocxArtifacts(input, { ...options, unsupportedContent: 'strict' }))
    .rejects.toMatchObject({ code: 'resource_policy_failure' });
});

for (const mode of ['notBeside', 'around', 'floating'] as const) {
  test(`${mode} frame accepts equivalent universal-measure offsets`, async ({ browser }) => {
    const positions = [];
    for (const horizontalOffset of ['240', '12pt']) {
      const result = await renderDocxArtifacts(inlineFrameFixture({ horizontalOffset,
        wrap: mode === 'floating' ? 'none' : mode, floating: mode === 'floating' }), {
        browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
      });
      const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
      try {
        const pdf = await task.promise;
        expect(pdf.numPages).toBe(1);
        positions.push((await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item)
          .filter(item => item.str.trim()).map(item => ({ text: item.str, x: item.transform[4], y: item.transform[5] })));
      } finally { await task.destroy(); }
    }
    expect(positions[1]).toEqual(positions[0]);
    expect(positions[0].find(item => item.text === 'SageFrameFirst')!.x).toBeCloseTo(84, 0);
  });
}

test('adjacent distinct frames exclude the union of their rectangles', async ({ browser }) => {
  const result = await renderDocxArtifacts(inlineFrameFixture({ wrap: 'around', horizontalOffset: '480', secondOffset: '240' }), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
  });
  expect(result.pageCount).toBe(1);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item);
    const word = (text: string) => {
      const matches = items.filter(item => item.str === text);
      expect(matches).toHaveLength(1);
      return matches[0];
    };
    const first = word('SageFrameFirst'), last = word('SageFrameLast'), after = word('SageAfter');
    expect(first.transform[4]).toBeCloseTo(96, 0);
    expect(last.transform[4]).toBeCloseTo(84, 0);
    expect(after.transform[4]).toBeCloseTo(276, 0);
    expect(first.transform[5]).toBeCloseTo(last.transform[5], 1);
    expect(first.transform[5]).toBeCloseTo(after.transform[5], 1);
  } finally { await task.destroy(); }
});

for (const story of ['body', 'header', 'footer'] as const) for (const normal of [false, true]) {
  test(`separated ${story} ${normal ? 'ordinary control' : 'frames'} retain source order and text positions`, async ({ browser }) => {
    const result = await renderDocxArtifacts(inlineFrameFixture({ story, normal, wrap: 'around', separate: true, heightRule: 'exact' }), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      expect(pdf.numPages).toBe(1);
      const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item);
      const words = ['SageBefore', 'SageFrameFirst', 'SageBetween', 'SageFrameLast', 'SageAfter'].map(text => {
        const matches = items.filter(item => item.str === text);
        expect(matches, `${text} survives once`).toHaveLength(1);
        return matches[0];
      });
      for (const [i, word] of words.entries()) {
        expect(words[0].transform[5] - word.transform[5]).toBeCloseTo((normal ? i : Math.ceil(i / 2)) * 12, 0);
        expect(word.transform[4]).toBeCloseTo(normal || i === 0 ? 72 : i % 2 ? 84 : 264, 0);
      }
    } finally { await task.destroy(); }
  });
}

for (const wrap of ['around', 'notBeside'] as const) for (const heightRule of ['exact', 'atLeast', 'auto'] as const) {
  test(`inline ${wrap} frame honors ${heightRule} height`, async ({ browser }) => {
    const result = await renderDocxArtifacts(inlineFrameFixture({ wrap, heightRule }), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      expect(pdf.numPages).toBe(1);
      const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item);
      const word = (text: string) => {
        const matches = items.filter(item => item.str === text);
        expect(matches).toHaveLength(1);
        return matches[0];
      };
      const first = word('SageFrameFirst'), last = word('SageFrameLast'), after = word('SageAfter');
      expect(first.transform[5] - last.transform[5]).toBeCloseTo(12, 0);
      expect(first.transform[5] - after.transform[5]).toBeCloseTo(wrap === 'around' ? 0 : heightRule === 'auto' ? 24 : 40, 0);
      expect(after.transform[4]).toBeCloseTo(wrap === 'around' ? 264 : 72, 0);
    } finally { await task.destroy(); }
  });
}

test('a long paragraph beside an inline frame continues with every word in order', async ({ browser }) => {
  const result = await renderDocxArtifacts(inlineFrameFixture({ wrap: 'around', fillerCount: 50, tailWords: 90 }), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
  });
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(2);
    const words: string[] = [];
    for (let page = 1; page <= pdf.numPages; page++) {
      const items = (await (await pdf.getPage(page)).getTextContent()).items.filter(item => 'str' in item);
      words.push(...items.flatMap(item => item.str.match(/SageTail\d+/g) ?? []));
      const first = items.find(item => /SageTail\d+/.test(item.str))!;
      expect(first.transform[4]).toBeCloseTo(page === 1 ? 264 : 72, 0);
      const frames = items.filter(item => /^SageFrame/.test(item.str)).map(item => item.str);
      expect(frames.sort()).toEqual(page === 1 ? ['SageFrameFirst', 'SageFrameLast'] : []);
    }
    expect(words).toEqual(Array.from({ length: 90 }, (_, i) => `SageTail${String(i).padStart(3, '0')}`));
  } finally { await task.destroy(); }
});
