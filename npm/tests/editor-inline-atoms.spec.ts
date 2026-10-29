import { test, expect, Page } from '@playwright/test';

const atoms = [
  ['br', '<w:br/>', '\n'],
  ['cr', '<w:cr/>', '\n'],
  ['page break', '<w:br w:type="page"/>', '\n'],
  ['column break', '<w:br w:type="column"/>', '\n'],
  ['ptab', '<w:ptab w:alignment="right" w:relativeTo="margin" w:leader="none"/>', '\t'],
  ['symbol', '<w:sym w:font="Symbol" w:char="03A9"/>', 'Ω'],
  ['soft hyphen', '<w:softHyphen/>', '\u00ad'],
  ['nonbreaking hyphen', '<w:noBreakHyphen/>', '\u2011'],
] as const;

async function open(page: Page, atom: string, edges = false) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  await page.evaluate(({ atom, edges }) => {
    const w = window as any;
    const D = w.Docxodus;
    const s = D.openTypedSession(D.DocxSessionBridge.CreateBlankDocx());
    const anchor = Object.keys(s.project().anchorIndex).find((id) => id.startsWith('p:'))!;
    const original = s.raw.getXml(anchor) as string;
    const start = original.match(/^<w:p\b[^>]*>/)![0].replace(/\/>$/, '>');
    const content = edges ? `${atom}<w:t>BBB</w:t>${atom}` : `<w:t>AAA</w:t>${atom}<w:t>BBB</w:t>`;
    const edit = s.raw.replaceXml(anchor, `${start}<w:r>${content}</w:r></w:p>`);
    if (!edit.success) throw new Error(JSON.stringify(edit));
    const bytes = s.save();
    s.close();
    const container = document.createElement('div');
    document.body.appendChild(container);
    w.failures = [];
    w.editor = D.DocxEditor.open(container, bytes, D, {
      onEditFailed: (e: unknown) => w.failures.push(e),
    });
    w.block = () => container.querySelector('p[data-anchor][contenteditable="true"]');
  }, { atom, edges });
}

async function selectB(page: Page, from: number, length = 0) {
  await page.evaluate(({ from, length }) => {
    const block: HTMLElement = (window as any).block();
    block.focus();
    const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const at = (n.textContent ?? '').indexOf('BBB');
      if (at < 0) continue;
      const r = document.createRange();
      r.setStart(n, at + from);
      r.setEnd(n, at + from + length);
      window.getSelection()!.removeAllRanges();
      window.getSelection()!.addRange(r);
      return;
    }
    throw new Error('no BBB text node');
  }, { from, length });
}

async function saved(page: Page) {
  await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur());
  return page.evaluate(() => {
    const w = window as any;
    const s = w.Docxodus.openTypedSession(w.editor.save());
    try {
      const projection = s.project();
      const ids = Object.keys(projection.anchorIndex).filter((id) => id.startsWith('p:'));
      return {
        text: ids.map((id) => s.getAnchorInfo(id).visibleText),
        xml: ids.map((id) => s.raw.getXml(id)),
        failures: w.failures,
      };
    } finally { s.close(); }
  });
}

for (const [name, atom, text] of atoms) {
  test(`${name}: loaded baseline and typing use the native text coordinates`, async ({ page }) => {
    await open(page, atom);
    expect(await page.evaluate(() => (window as any).block().dataset.committedText)).toBe(`AAA${text}BBB`);
    await selectB(page, 3);
    await page.keyboard.type('CCC');
    const out = await saved(page);
    expect(out.text).toEqual([`AAA${text}BBBCCC`]);
    expect(out.failures).toEqual([]);
    expect(out.xml[0]).toContain(atom.match(/^<w:(\w+)/)![1]);
  });

  test(`${name}: formatting and Enter after the atom address the displayed text`, async ({ page }) => {
    await open(page, atom);
    await selectB(page, 0, 3);
    await page.evaluate(() => (window as any).editor.format('bold', true));
    const formatted = await saved(page);
    expect(formatted.xml[0]).toMatch(/<w:b\b[^>]*\/>[\s\S]*?<w:t[^>]*>BBB<\/w:t>/);
    await selectB(page, 3);
    await page.keyboard.press('Enter');
    const out = await saved(page);
    expect(out.text).toEqual([`AAA${text}BBB`, '']);
    expect(out.failures).toEqual([]);
  });
}

for (const kind of ['br', 'cr']) {
  test(`${kind}: deleting only a loaded break commits its removal`, async ({ page }) => {
    await open(page, `<w:${kind}/>`);
    await selectB(page, 0);
    await page.keyboard.press('Backspace');
    const out = await saved(page);
    expect(out.text).toEqual(['AAABBB']);
    expect(out.xml[0]).not.toMatch(/<w:(?:br|cr)\b/);
    expect(out.failures).toEqual([]);
  });

  test(`${kind}: edge breaks survive typing and splitting`, async ({ page }) => {
    await open(page, `<w:${kind}/>`, true);
    expect(await page.evaluate(() => (window as any).block().dataset.committedText)).toBe('\nBBB\n');
    await selectB(page, 1);
    await page.keyboard.type('Q');
    await page.keyboard.press('Enter');
    const out = await saved(page);
    expect(out.text).toEqual(['\nBQ', 'BB\n']);
    expect(out.failures).toEqual([]);
  });
}

for (const count of [1, 2]) {
  test(`${count} Shift+Enter at the end keeps exactly that many breaks after blur and reopen`, async ({ page }) => {
    await open(page, '');
    await selectB(page, 3);
    for (let i = 0; i < count; i++) await page.keyboard.press('Shift+Enter');
    const out = await saved(page);
    expect(out.text).toEqual(['AAABBB' + '\n'.repeat(count)]);
    expect(out.xml[0].match(/<w:br\b/g)).toHaveLength(count);
    expect(out.failures).toEqual([]);
  });
}

test('Shift+Enter replaces a selected suffix and typing continues after the break', async ({ page }) => {
  await open(page, '');
  await selectB(page, 1, 2);
  await page.keyboard.press('Shift+Enter');
  await page.keyboard.type('CCC');
  const out = await saved(page);
  expect(out.text).toEqual(['AAAB\nCCC']);
  expect(out.xml[0].match(/<w:br\b/g)).toHaveLength(1);
  expect(out.failures).toEqual([]);
});
