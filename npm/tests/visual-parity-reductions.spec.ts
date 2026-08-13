import { test, expect, type Page } from '@playwright/test';
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { storedZip, xml, R_NS, W_NS } from './docx-zip.js';
import { assertLibreOfficeContract } from './visual-parity/environment-contract.js';
import { FONT_CONTRACT_FILE, assertFontContract } from './visual-parity/font-contract.js';
import { VISUAL_THRESHOLDS, background } from './visual-parity/metrics.js';
import { decodePng, encodePng, type RgbaImage } from './visual-parity/png.js';

/**
 * Reduced same-font layout cases for the environment-attributed corpus residuals (issue #404).
 *
 * `environment` means "the two engines lay out the SAME fonts differently" — a legitimate
 * resting state for a Chromium-vs-LibreOffice comparison, but as a whole-fixture impression it
 * is unfalsifiable. Each probe here reduces one attributed case's dominant residual to a
 * minimal GENERATED document and measures the same observable in both engines with the same
 * ink model, so the disposition can cite numbers from an isolated mechanism instead:
 *
 *  - `landscape-section` → paragraph pitch (line box + `w:spacing w:after`) in a landscape
 *    section: uniform in both engines, differing only by the per-line font metric.
 *  - `inline-image` → an inline picture's rendered extent and the position where following
 *    text resumes.
 *  - `tracked-deletion` → the heading line box (Calibri Light → Carlito) over a body line.
 *
 * Assertions pin what DOCXODUS owes the OOXML (declared extents honored, uniform pitch with no
 * accumulation, structure identical across engines). Engine-vs-engine deltas are MEASURED and
 * logged — they are the environment residual being isolated, not a defect either engine is
 * accused of. Word evidence for whether Docxodus's choice is also Word-correct is captured via
 * the issue-#402 procedure (see WORD_REFERENCE.md).
 */

test.skip(process.env.DOCXODUS_VISUAL_PARITY !== '1',
  'set DOCXODUS_VISUAL_PARITY=1 on a host with libreoffice and pdftoppm');

const __dirname = dirname(fileURLToPath(import.meta.url));

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const ROOT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`;

function docx(bodyXml: string, extras: { name: string; data: Buffer }[] = [],
  documentRels = ''): Uint8Array {
  return storedZip([
    { name: '[Content_Types].xml', data: xml(CONTENT_TYPES_XML) },
    { name: '_rels/.rels', data: xml(ROOT_RELS_XML) },
    ...(documentRels ? [{
      name: 'word/_rels/document.xml.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${documentRels}</Relationships>`),
    }] : []),
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>${bodyXml}</w:body></w:document>`),
    },
    ...extras,
  ]);
}

const run = (family: string, halfPoints: number, text: string): string =>
  `<w:r><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/>` +
  `<w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/></w:rPr>` +
  `<w:t xml:space="preserve">${text}</w:t></w:r>`;

/** REDUCTION 1 (landscape-section): four identical single-line paragraphs, landscape letter. */
function landscapeSpacingDocx(): Uint8Array {
  const paragraph =
    `<w:p><w:pPr><w:spacing w:before="0" w:after="160" w:line="259" w:lineRule="auto"/></w:pPr>` +
    run('Calibri', 22, 'Measured paragraph pitch probe.') + `</w:p>`;
  return docx(`
  ${paragraph.repeat(4)}
  <w:sectPr>
    <w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
      w:header="720" w:footer="720" w:gutter="0"/>
    <w:cols w:space="720"/>
  </w:sectPr>`);
}

/** REDUCTION 2 (inline-image): text, an inline 150x75 px picture, text. */
function inlineImageDocx(): Uint8Array {
  const pngWidth = 150;
  const pngHeight = 75;
  const data = new Uint8Array(pngWidth * pngHeight * 4);
  for (let i = 0; i < data.length; i += 4) {
    data[i] = 32; data[i + 1] = 64; data[i + 2] = 128; data[i + 3] = 255;
  }
  const png = Buffer.from(encodePng({ width: pngWidth, height: pngHeight, data }));
  const emu = (px: number) => px * 9525;
  const drawing =
    `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" ` +
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">` +
    `<wp:extent cx="${emu(pngWidth)}" cy="${emu(pngHeight)}"/>` +
    `<wp:docPr id="1" name="reduction-image"/>` +
    `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">` +
    `<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:nvPicPr><pic:cNvPr id="1" name="reduction-image"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rIdImg1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(pngWidth)}" cy="${emu(pngHeight)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const spacing =
    `<w:pPr><w:spacing w:before="0" w:after="160" w:line="259" w:lineRule="auto"/></w:pPr>`;
  return docx(`
  <w:p>${spacing}${run('Calibri', 22, 'Text before the inline image.')}</w:p>
  <w:p>${spacing}${drawing}</w:p>
  <w:p>${spacing}${run('Calibri', 22, 'Text after the inline image.')}</w:p>
  <w:sectPr>
    <w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
      w:header="720" w:footer="720" w:gutter="0"/>
    <w:cols w:space="720"/>
  </w:sectPr>`,
  [{ name: 'word/media/image1.png', data: png }],
  `<Relationship Id="rIdImg1" Type="${R_NS}/image" Target="media/image1.png"/>`);
}

/** REDUCTION 3 (tracked-deletion): a Calibri Light heading line over a Calibri body line. */
function headingMetricsDocx(): Uint8Array {
  const single = `<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>`;
  return docx(`
  <w:p>${single}${run('Calibri Light', 32, 'Heading line metrics probe')}</w:p>
  <w:p>${single}${run('Calibri', 22, 'Body line under the heading.')}</w:p>
  <w:sectPr>
    <w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
      w:header="720" w:footer="720" w:gutter="0"/>
    <w:cols w:space="720"/>
  </w:sectPr>`);
}

// ---------------------------------------------------------------------------
// Rendering and measurement (the corpus runner's contract, self-contained).
// ---------------------------------------------------------------------------

function libreofficeEnv(work: string): NodeJS.ProcessEnv {
  const runtimeDir = join(work, 'runtime');
  const homeDir = join(work, 'home');
  for (const directory of [runtimeDir, homeDir]) mkdirSync(directory, { recursive: true, mode: 0o700 });
  return {
    ...process.env,
    HOME: homeDir,
    XDG_RUNTIME_DIR: runtimeDir,
    LANG: 'C.UTF-8',
    LC_ALL: 'C.UTF-8',
    TZ: 'UTC',
    FONTCONFIG_FILE: FONT_CONTRACT_FILE,
  };
}

function renderLibreOfficePages(docxPath: string, work: string): RgbaImage[] {
  const pdfDir = join(work, 'pdf');
  const profileDir = join(work, 'profile');
  for (const directory of [pdfDir, profileDir]) mkdirSync(directory, { recursive: true, mode: 0o700 });
  execFileSync('libreoffice', [
    `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
    '--headless', '--nologo', '--nodefault', '--nofirststartwizard', '--norestore',
    '--convert-to', 'pdf', '--outdir', pdfDir, docxPath,
  ], { env: libreofficeEnv(work), stdio: 'pipe', timeout: 120000 });
  const pdfPath = join(pdfDir, `${docxPath.split('/').pop()!.replace(/\.docx$/i, '')}.pdf`);
  if (!existsSync(pdfPath)) throw new Error(`LibreOffice did not produce ${pdfPath}`);
  execFileSync('pdftoppm', ['-r', '96', '-png', pdfPath, join(work, 'lo-page')],
    { env: libreofficeEnv(work), stdio: 'pipe', timeout: 120000 });
  const pattern = /^lo-page-(\d+)\.png$/;
  return readdirSync(work)
    .map(name => ({ name, match: name.match(pattern) }))
    .filter((item): item is { name: string; match: RegExpMatchArray } => item.match !== null)
    .sort((a, b) => Number(a.match[1]) - Number(b.match[1]))
    .map(item => decodePng(readFileSync(join(work, item.name))));
}

async function renderDocxodusPages(page: Page, bytes: Uint8Array, work: string): Promise<RgbaImage[]> {
  const result = await page.evaluate((input) => {
    try {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(input),
        'Document', 'docx-', true, '', -1, 'comment-',
        1, 1, 'page-',
        false, 0, 'annot-',
        true, true,
        false, false, false,
      );
      if (html.startsWith('{') && html.includes('"Error"')) return { error: html };
      return { html };
    } catch (error) {
      return { error: String(error) };
    }
  }, Array.from(bytes));
  if (!result.html) throw new Error(`Docxodus conversion failed: ${result.error}`);

  await page.evaluate((html) => {
    document.documentElement.style.background = 'white';
    document.body.innerHTML = '<main id="reduction-root"></main>';
    document.body.style.cssText = 'margin:0;padding:0;background:white;';
    (window as any).DocxodusPagination.paginateHtml(html, document.getElementById('reduction-root'), {
      scale: 1, showPageNumbers: false, pageGap: 0, fragmentParagraphs: false,
    });
    const style = document.createElement('style');
    style.textContent = `*, *::before, *::after { animation: none !important;
      caret-color: transparent !important; transition: none !important; }
      #pagination-container, .page-container { margin: 0 !important; padding: 0 !important; }
      .page-box { box-shadow: none !important; margin: 0 !important; }`;
    document.head.appendChild(style);
  }, result.html);
  await page.evaluate(async () => {
    await document.fonts.ready;
    await Promise.all(Array.from(document.images).map(image => image.complete
      ? Promise.resolve()
      : new Promise<void>(done => {
          image.addEventListener('load', () => done(), { once: true });
          image.addEventListener('error', () => done(), { once: true });
        })));
    await new Promise<void>(done => requestAnimationFrame(() => requestAnimationFrame(() => done())));
  });

  const boxes = page.locator('#reduction-root .page-box');
  const count = await boxes.count();
  if (!count) throw new Error('Docxodus pagination produced no page boxes');
  const pages: RgbaImage[] = [];
  for (let index = 0; index < count; index++) {
    const path = join(work, `docxodus-${index + 1}.png`);
    await boxes.nth(index).screenshot({ path, animations: 'disabled', caret: 'hide' });
    pages.push(decodePng(readFileSync(path)));
  }
  return pages;
}

export interface InkBand { top: number; bottom: number; left: number; right: number }

/** Contiguous horizontal bands of rows containing ink, under the corpus's shared ink model. */
function inkRowBands(image: RgbaImage): InkBand[] {
  const bg = background(image);
  const bands: InkBand[] = [];
  let current: InkBand | null = null;
  for (let y = 0; y < image.height; y++) {
    let rowLeft = -1;
    let rowRight = -1;
    for (let x = 0; x < image.width; x++) {
      const i = (y * image.width + x) * 4;
      const distance = Math.max(
        Math.abs(image.data[i] - bg[0]),
        Math.abs(image.data[i + 1] - bg[1]),
        Math.abs(image.data[i + 2] - bg[2]),
      );
      if (distance > VISUAL_THRESHOLDS.inkBackgroundDistance) {
        if (rowLeft < 0) rowLeft = x;
        rowRight = x;
      }
    }
    if (rowLeft >= 0) {
      if (current && y === current.bottom + 1) {
        current.bottom = y;
        current.left = Math.min(current.left, rowLeft);
        current.right = Math.max(current.right, rowRight);
      } else {
        current = { top: y, bottom: y, left: rowLeft, right: rowRight };
        bands.push(current);
      }
    }
  }
  return bands;
}

const bandTable = (label: string, bands: InkBand[]): string =>
  `${label}: ${bands.map(band =>
    `[y ${band.top}-${band.bottom}, x ${band.left}-${band.right}]`).join(' ')}`;

interface Reduction {
  docxodus: { page: RgbaImage; bands: InkBand[] };
  libreoffice: { page: RgbaImage; bands: InkBand[] };
}

async function reduce(page: Page, name: string, bytes: Uint8Array): Promise<Reduction> {
  const work = mkdtempSync(join(tmpdir(), `docxodus-reduction-${name}-`));
  try {
    const docxPath = join(work, `${name}.docx`);
    writeFileSync(docxPath, Buffer.from(bytes));
    const libreofficePages = renderLibreOfficePages(docxPath, work);
    const docxodusPages = await renderDocxodusPages(page, bytes, work);
    expect(docxodusPages.length, `${name}: both engines must produce one page`).toBe(1);
    expect(libreofficePages.length, `${name}: both engines must produce one page`).toBe(1);
    const result = {
      docxodus: { page: docxodusPages[0], bands: inkRowBands(docxodusPages[0]) },
      libreoffice: { page: libreofficePages[0], bands: inkRowBands(libreofficePages[0]) },
    };
    console.log(`[reduction ${name}] page D ${result.docxodus.page.width}x${result.docxodus.page.height}, ` +
      `L ${result.libreoffice.page.width}x${result.libreoffice.page.height}`);
    console.log(`[reduction ${name}] ${bandTable('D', result.docxodus.bands)}`);
    console.log(`[reduction ${name}] ${bandTable('L', result.libreoffice.bands)}`);
    return result;
  } finally {
    rmSync(work, { recursive: true, force: true });
  }
}

const tops = (bands: InkBand[]): number[] => bands.map(band => band.top);
const pitches = (bands: InkBand[]): number[] =>
  bands.slice(1).map((band, index) => band.top - bands[index].top);

test.describe('issue-#404 reduced environment cases', () => {
  test.beforeAll(() => {
    assertFontContract();
    assertLibreOfficeContract();
  });

  test.beforeEach(async ({ page }) => {
    await page.setViewportSize({ width: 1400, height: 1200 });
    await page.emulateMedia({ reducedMotion: 'reduce' });
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 90000 });
    await page.addScriptTag({ url: '/pagination.bundle.js' });
  });

  test('landscape-section reduces to paragraph pitch of the same font', async ({ page }) => {
    test.setTimeout(5 * 60 * 1000);
    const result = await reduce(page, 'landscape-spacing', landscapeSpacingDocx());

    // Landscape page geometry is not in dispute — both engines must agree exactly.
    expect(result.docxodus.page.width).toBe(result.libreoffice.page.width);
    expect(result.docxodus.page.height).toBe(result.libreoffice.page.height);
    // Four single-line paragraphs make four ink bands in BOTH engines: structure agrees,
    // so whatever differs below is per-line metrics, not lost or rewrapped content.
    expect(result.docxodus.bands).toHaveLength(4);
    expect(result.libreoffice.bands).toHaveLength(4);

    // What Docxodus owes the OOXML: uniform pitch — `w:spacing` must not accumulate error
    // down the page (the issue-#396 defect class). Same invariant holds for LibreOffice.
    for (const engine of [result.docxodus, result.libreoffice]) {
      const engineSteps = pitches(engine.bands);
      for (const step of engineSteps) {
        expect(Math.abs(step - engineSteps[0]),
          'paragraph pitch must be uniform down the page').toBeLessThanOrEqual(1);
      }
    }

    // The isolated residual, measured: per-paragraph pitch delta between the engines laying
    // out the SAME substituted font. Logged for the disposition; bounded loosely so a real
    // regression (a broken spacing model) still fails while metric drift does not.
    const deltaPerParagraph = pitches(result.docxodus.bands)[0] - pitches(result.libreoffice.bands)[0];
    console.log(`[reduction landscape-spacing] pitch D=${pitches(result.docxodus.bands)[0]}px ` +
      `L=${pitches(result.libreoffice.bands)[0]}px delta=${deltaPerParagraph}px/paragraph; ` +
      `first band top D=${result.docxodus.bands[0].top} L=${result.libreoffice.bands[0].top}`);
    expect(Math.abs(deltaPerParagraph),
      'pitch delta beyond a couple of pixels is a spacing-model difference, not line metrics')
      .toBeLessThanOrEqual(2);
  });

  test('inline-image reduces to extent fidelity and where following text resumes', async ({ page }) => {
    test.setTimeout(5 * 60 * 1000);
    const result = await reduce(page, 'inline-image', inlineImageDocx());

    // Text, image, text: three bands in both engines.
    expect(result.docxodus.bands).toHaveLength(3);
    expect(result.libreoffice.bands).toHaveLength(3);

    // What Docxodus owes the OOXML: the declared `wp:extent` (150x75 px at 96 DPI), exactly.
    const docxodusImage = result.docxodus.bands[1];
    expect(docxodusImage.bottom - docxodusImage.top + 1).toBeGreaterThanOrEqual(74);
    expect(docxodusImage.bottom - docxodusImage.top + 1).toBeLessThanOrEqual(76);
    expect(docxodusImage.right - docxodusImage.left + 1).toBeGreaterThanOrEqual(149);
    expect(docxodusImage.right - docxodusImage.left + 1).toBeLessThanOrEqual(151);
    // Both engines start the image at the left margin (96 px at 96 DPI, 1440 twips).
    expect(Math.abs(docxodusImage.left - 96)).toBeLessThanOrEqual(1);
    expect(Math.abs(result.libreoffice.bands[1].left - 96)).toBeLessThanOrEqual(1);

    const libreofficeImage = result.libreoffice.bands[1];
    console.log('[reduction inline-image] image ' +
      `D=(${docxodusImage.left},${docxodusImage.top})-(${docxodusImage.right},${docxodusImage.bottom}) ` +
      `L=(${libreofficeImage.left},${libreofficeImage.top})-(${libreofficeImage.right},${libreofficeImage.bottom}); ` +
      `text-after top D=${result.docxodus.bands[2].top} L=${result.libreoffice.bands[2].top}`);

    // Following text resumes below the image in both engines; the top-position delta is the
    // same-font line-metric residual the corpus case's disposition cites.
    expect(result.docxodus.bands[2].top).toBeGreaterThan(docxodusImage.bottom);
    expect(result.libreoffice.bands[2].top).toBeGreaterThan(libreofficeImage.bottom);
  });

  test('tracked-deletion reduces to the heading (Calibri Light) line box', async ({ page }) => {
    test.setTimeout(5 * 60 * 1000);
    const result = await reduce(page, 'heading-metrics', headingMetricsDocx());

    // A heading line and a body line: two bands in both engines.
    expect(result.docxodus.bands).toHaveLength(2);
    expect(result.libreoffice.bands).toHaveLength(2);

    // Heading-to-body advance: the single-spaced heading line box. The engines measure the
    // SAME substituted face (Carlito for Calibri Light) differently; the delta is the case's
    // residual, measured in isolation. Glyph-ink heights are logged for completeness.
    const advance = (bands: InkBand[]) => bands[1].top - bands[0].top;
    console.log(`[reduction heading-metrics] heading advance D=${advance(result.docxodus.bands)}px ` +
      `L=${advance(result.libreoffice.bands)}px; heading ink height ` +
      `D=${result.docxodus.bands[0].bottom - result.docxodus.bands[0].top + 1} ` +
      `L=${result.libreoffice.bands[0].bottom - result.libreoffice.bands[0].top + 1}`);
    expect(Math.abs(advance(result.docxodus.bands) - advance(result.libreoffice.bands)),
      'heading advance delta beyond a few pixels would be a spacing bug, not font metrics')
      .toBeLessThanOrEqual(3);
  });
});
