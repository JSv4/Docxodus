import { createHash } from 'node:crypto';
import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { balloonDocx, commented, formatted } from './docx-markup-balloons-fixture.js';

const author = 'w:author="Spruce Reviewer" w:date="2026-01-02T09:00:00Z"';
const digest = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex');
const fixture = (format = true, comments = true, text = false) => balloonDocx({
  body: '<w:p>' + (format ? formatted(1, 'SpruceBold ') : '<w:r><w:t>SprucePlain </w:t></w:r>')
    + (text ? `<w:del w:id="2" ${author}><w:r><w:delText>SpruceOld </w:delText></w:r></w:del>`
      + `<w:ins w:id="3" ${author}><w:r><w:t>SpruceNew </w:t></w:r></w:ins>` : '')
    + (comments ? commented(0, 'SpruceAnchor') : '<w:r><w:t>SpruceEnd</w:t></w:r>') + '</w:p>',
  comments: comments ? [{ id: 0, author: 'Spruce Reviewer', initials: 'SR', text: 'SpruceNote' }] : [],
});

async function textOf(pdf: Uint8Array): Promise<string> {
  const task = getDocument({ data: new Uint8Array(pdf), verbosity: 0 });
  try {
    const document = await task.promise;
    const texts = await Promise.all(Array.from({ length: document.numPages }, async (_, i) =>
      (await (await document.getPage(i + 1)).getTextContent()).items.filter(item => 'str' in item).map(item => item.str)));
    return texts.flat().join(' ').replace(/\s+/g, ' ');
  } finally { await task.destroy(); }
}

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const formatting of [true, false]) for (const insertionsAndDeletions of [true, false]) for (const comments of [true, false]) {
  test(`independent visibility formatting=${formatting}, text=${insertionsAndDeletions}, comments=${comments}`, async ({ browser, page }) => {
    const source = fixture(true, true, true), before = digest(source);
    const visibility = { formatting, insertionsAndDeletions, comments };
    const options: Parameters<typeof renderDocxArtifacts>[1] = {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', markupPresentation: 'word',
      commentProfile: 'margin', commentColumnPlacement: 'balloons', markupVisibility: visibility,
    };
    const result = await renderDocxArtifacts(source, options);
    expect(digest(source)).toBe(before);
    expect(result.renderReport.source.rawPackageBytesDigest).toBe(before);
    expect(result.renderReport.derivedProfileSource).toBeUndefined();
    const deviations = Object.fromEntries(Object.entries(visibility).filter(([, shown]) => !shown));
    expect(result.renderReport.options.markupVisibility)
      .toEqual(Object.keys(deviations).length ? deviations : undefined);
    const text = await textOf(result.pdf!);
    expect(text).toContain('SpruceBold');
    expect(text).toContain('SpruceNew');
    expect(text.includes('SpruceOld')).toBe(insertionsAndDeletions);
    expect(text.includes('Formatted:')).toBe(formatting);
    expect(text.includes('SpruceNote')).toBe(comments);
    await page.setContent(result.html!);
    expect(await page.locator('.page-markup-area').count()).toBe(formatting || comments ? 1 : 0);
    expect(await page.locator('ins').count()).toBe(insertionsAndDeletions ? 1 : 0);
    expect(await page.locator('del').count()).toBe(insertionsAndDeletions ? 1 : 0);
    expect(await page.locator('.rev-changed-line').count()).toBe(formatting || insertionsAndDeletions ? 1 : 0);
    expect(await page.locator('span').evaluateAll(spans => spans.some(span =>
      span.textContent?.trim() === 'SpruceBold' && Number(getComputedStyle(span).fontWeight) >= 700))).toBe(true);
  });
}

for (const comments of [false, true]) test(`format-only toggle preserves ${comments ? 'comment-only geometry' : 'full page width'}`, async ({ browser, page }) => {
  const source = fixture(!comments, comments);
  const render = async (formatting?: boolean) => {
    const options: Parameters<typeof renderDocxArtifacts>[1] = {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', markupPresentation: 'word',
      commentProfile: 'margin', commentColumnPlacement: 'balloons',
      ...(formatting === undefined ? {} : { markupVisibility: { formatting } }),
    };
    return renderDocxArtifacts(source, options);
  };
  const before = await render(), hidden = await render(false);
  expect(before.pageMap.pages).toEqual(hidden.pageMap.pages);
  await page.setContent(hidden.html!);
  expect(await page.locator('.page-markup-area').count()).toBe(comments ? 1 : 0);
  if (comments) {
    expect(await textOf(before.pdf!)).toBe(await textOf(hidden.pdf!));
    expect(before.pageMap.fragments.map(f => f.geometry)).toEqual(hidden.pageMap.fragments.map(f => f.geometry));
  } else {
    expect(await page.locator('.page-markup-layer').count()).toBe(0);
    expect(await textOf(hidden.pdf!)).not.toContain('Formatted:');
  }
});

test('hiding the only deleted table produces an explicit blank page without fabricated source anchors',async ({browser,page})=>{
  const source=balloonDocx({body:'<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="2400"/></w:tblPr>'
    + '<w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>'
    + `<w:tr><w:trPr><w:del w:id="1" ${author}/></w:trPr><w:tc><w:p><w:r><w:t>SpruceDeletedTable</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`});
  const result=await renderDocxArtifacts(source,{browser,outputs:['html','pdf'],reviewProfile:'markup',commentProfile:'hidden',
    markupVisibility:{insertionsAndDeletions:false}});
  expect(result.pageCount).toBe(1);
  expect(result.pageMap.fragments).toEqual([]);
  expect((await textOf(result.pdf!)).trim()).toBe('');
  expect(result.renderReport.source.rawPackageBytesDigest).toBe(digest(source));
  expect(result.renderReport.derivedProfileSource).toBeUndefined();
  await page.setContent(result.html!);
  expect(await page.locator('[data-docx-empty-review-display="true"]').count()).toBe(1);
  expect(await page.locator('[data-source-anchor-id]').count()).toBe(0);
});

test('explicit defaults preserve layout identity and a nondefault policy changes it', async ({ browser }) => {
  const source = fixture(false, false);
  const options: Parameters<typeof renderDocxArtifacts>[1] = {
    browser, outputs: ['html'], reviewProfile: 'markup', commentProfile: 'hidden',
  };
  const baseline = await renderDocxArtifacts(source, options);
  for (const markupVisibility of [{}, { formatting: true, insertionsAndDeletions: true, comments: true }]) {
    const result = await renderDocxArtifacts(source, { ...options, markupVisibility });
    expect(result.renderReport.options.layoutDigest).toBe(baseline.renderReport.options.layoutDigest);
    expect(result.renderReport.options.markupVisibility).toBeUndefined();
  }
  const hidden = await renderDocxArtifacts(source, { ...options, markupVisibility: { formatting: false } });
  expect(hidden.renderReport.options.layoutDigest).not.toBe(baseline.renderReport.options.layoutDigest);
  expect(hidden.renderReport.source).toEqual(baseline.renderReport.source);
});

test('browser validation refuses malformed visibility policies', async ({ page }) => {
  await page.goto('/standalone-export-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
  for (const markupVisibility of [null, [], { formatting: 'false' }, { unknown: false }]) {
    const result = await page.evaluate(async ({ bytes, markupVisibility }) => (window as any).DocxodusStandalone
      .convertFailure(bytes, { reviewProfile: 'markup', commentProfile: 'hidden', markupVisibility }),
    { bytes: Array.from(fixture(false, false)), markupVisibility });
    expect(result.code).toBe('invalid_argument');
    expect(result.message).toContain('markupVisibility');
  }
});

for (const alignment of ['left', 'right', 'center', 'decimal']) {
  test(`a ${alignment} tab across a hidden paragraph mark matches a literal joined paragraph`, async ({ browser, page }) => {
    const properties = `<w:tabs><w:tab w:val="${alignment}" w:pos="2880"/></w:tabs>`;
    const prefix = '<w:r><w:t>SpruceFirst</w:t></w:r>';
    const suffix = '<w:r><w:tab/></w:r><w:r><w:t>12.34</w:t></w:r>';
    const bodies = [
      `<w:p><w:pPr><w:rPr><w:del w:id="1" ${author}/></w:rPr></w:pPr>${prefix}</w:p>`
        + `<w:p><w:pPr>${properties}</w:pPr>${suffix}</w:p>`,
      `<w:p><w:pPr>${properties}</w:pPr>${prefix}${suffix}</w:p>`,
    ];
    const positions: Array<{ x: number; y: number }> = [];
    for (const body of bodies) {
      const result = await renderDocxArtifacts(balloonDocx({ body }), {
        browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', commentProfile: 'hidden',
        markupVisibility: { insertionsAndDeletions: false },
      });
      expect(await textOf(result.pdf!)).toContain('SpruceFirst');
      expect(await textOf(result.pdf!)).toContain('12.34');
      await page.setContent(result.html!);
      positions.push(await page.evaluate(async () => {
        await document.fonts.ready;
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
        let node: Node | null;
        while ((node = walker.nextNode())) if (node.textContent === '12.34') {
          const range = document.createRange(); range.selectNodeContents(node);
          const rect = range.getBoundingClientRect();
          return { x: rect.x, y: rect.y };
        }
        throw new Error('Missing tab suffix');
      }));
      if (positions.length === 1)
        expect(new Set(result.pageMap.fragments.map(fragment => fragment.anchorId)).size).toBe(2);
    }
    expect(positions[0].x).toBeCloseTo(positions[1].x, 1);
    expect(positions[0].y).toBeCloseTo(positions[1].y, 1);
  });
}

for (const field of ['PAGE', 'NUMPAGES']) test(`joined header ${field} substitution retains both source paragraphs`, async ({ browser, page }) => {
  const source = balloonDocx({
    body: '<w:p><w:r><w:t>SpruceFirstPage</w:t><w:br w:type="page"/><w:t>SpruceSecondPage</w:t></w:r></w:p>',
    header: `<w:p><w:pPr><w:rPr><w:del w:id="1" ${author}/></w:rPr></w:pPr>`
      + `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> ${field} </w:instrText></w:r></w:p>`
      + '<w:p><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>9</w:t></w:r>'
      + '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
  });
  const result = await renderDocxArtifacts(source, { browser, outputs: ['html', 'pdf'], reviewProfile: 'markup',
    commentProfile: 'hidden', markupVisibility: { insertionsAndDeletions: false } });
  expect(result.pageCount).toBe(2);
  const headers = result.pageMap.fragments.filter(fragment => fragment.anchorId.startsWith('p:hdr'));
  expect(new Set(headers.map(fragment => fragment.anchorId)).size).toBe(2);
  expect(headers).toHaveLength(4);
  await page.setContent(result.html!);
  expect(await page.locator(`[data-field="${field}"]`).allTextContents()).toEqual(field === 'PAGE' ? ['1', '2'] : ['2', '2']);
});

test('worker transports independent revision visibility and preserves defaults', async ({ page }) => {
  await page.goto('/worker-test-harness.html');
  const result = await page.evaluate(async bytes => {
    await (window as any).createDocxodusWorker();
    const worker = (window as any).DocxodusWorker;
    const states = [];
    try {
      for (const policy of [{}, { showFormattingChanges: false }, { showInsertionsAndDeletions: false }]) {
        const html = await worker.convertDocxToHtml(new Uint8Array(bytes), {
          renderTrackedChanges: true, revisionPresentation: 1, ...policy,
        });
        const document = new DOMParser().parseFromString(html, 'text/html');
        states.push({ text: document.body.textContent, ins: document.querySelectorAll('ins').length,
          formatting: document.querySelectorAll('[data-format-description]').length });
      }
    } finally { worker.terminate(); }
    return states;
  }, Array.from(fixture(true, false, true)));
  expect(result.map(state => state.ins)).toEqual([1, 1, 0]);
  expect(result.map(state => state.formatting)).toEqual([1, 0, 1]);
  expect(result.map(state => state.text?.includes('SpruceOld'))).toEqual([true, true, false]);
});
