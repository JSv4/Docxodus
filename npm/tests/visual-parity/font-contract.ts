import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { lstatSync, readFileSync, realpathSync } from 'node:fs';
import { basename, dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  FONT_SUBSTITUTION_CONTRACT,
  type FontSubstitutionEntry,
} from '../../src/font-contract.js';
import { storedZip, xml, R_NS, W_NS } from '../docx-zip.js';

/**
 * The font-substitution contract (issue #379): the proprietary Office families the corpus uses,
 * each pinned to a license-safe metric-compatible substitute by `fonts.conf`, which both
 * renderers load via FONTCONFIG_FILE. This module owns declaring the contract, verifying the
 * host satisfies it, and recording the resolved fonts in the report.
 */

const __dirname = dirname(fileURLToPath(import.meta.url));

/** Absolute path both renderers must receive as FONTCONFIG_FILE. */
export const FONT_CONTRACT_FILE = resolve(__dirname, 'fonts.conf');

export type FontContractEntry = FontSubstitutionEntry;
export const FONT_CONTRACT = FONT_SUBSTITUTION_CONTRACT;

/**
 * Test-environment install hints. These intentionally stay out of the browser-portable
 * production substitution contract because package names are deployment-specific.
 */
export const FONT_CONTRACT_PACKAGES: Readonly<Record<string, string>> = Object.freeze({
  Calibri: 'fonts-crosextra-carlito',
  'Calibri Light': 'fonts-crosextra-carlito',
  Cambria: 'fonts-crosextra-caladea',
  'Times New Roman': 'fonts-liberation2',
  Arial: 'fonts-liberation2',
  'Courier New': 'fonts-liberation2',
});

export interface ResolvedFont {
  family: string;
  substitute: string;
  resolvedFamily: string;
  file: string;
  fileSha256: string;
  fontVersion: string;
  metricCompatible: boolean;
}

/** What each declared family resolves to under the contract, per fc-match. */
export function resolveContractFonts(fcMatchExecutable = 'fc-match'): ResolvedFont[] {
  return FONT_CONTRACT.map(entry => {
    const out = execFileSync(fcMatchExecutable, ['-f', '%{family}\t%{file}\t%{fontversion}', entry.family], {
      encoding: 'utf8',
      env: { ...process.env, FONTCONFIG_FILE: FONT_CONTRACT_FILE },
      timeout: 15_000,
      maxBuffer: 1024 * 1024,
    });
    const [resolvedFamily = '', file = '', fontVersion = ''] = out.split('\t');
    const resolvedFile = realpathSync(file);
    const metadata = lstatSync(resolvedFile);
    if (!metadata.isFile() || metadata.isSymbolicLink()) {
      throw new Error(`fc-match returned a non-regular font file for ${entry.family}`);
    }
    return {
      family: entry.family,
      substitute: entry.substitute,
      // fc-match may report a family list; the substitute must be one of its names.
      resolvedFamily,
      file: basename(resolvedFile),
      fileSha256: createHash('sha256').update(readFileSync(resolvedFile)).digest('hex'),
      fontVersion,
      metricCompatible: entry.metricCompatible,
    };
  });
}

/**
 * Fails clearly when the host cannot satisfy the contract, naming the packages to install.
 * Returns the resolutions for the report on success.
 */
export function assertFontContract(fcMatchExecutable = 'fc-match'): ResolvedFont[] {
  let resolved: ResolvedFont[];
  try {
    resolved = resolveContractFonts(fcMatchExecutable);
  } catch (error) {
    throw new Error(`fc-match is unavailable; the font contract cannot be verified: ${error}`);
  }
  const broken = resolved.filter(r =>
    !r.resolvedFamily.split(',').some(name => name.trim() === r.substitute));
  if (broken.length) {
    const detail = broken.map(r => {
      const pkg = FONT_CONTRACT_PACKAGES[r.family] ?? '(no test package hint configured)';
      return `  ${r.family} -> ${r.resolvedFamily || '(nothing)'} (contract: ${r.substitute}; install ${pkg})`;
    }).join('\n');
    throw new Error(`Font-substitution contract not satisfied:\n${detail}`);
  }
  return resolved;
}

/** The report block `summary.json` records for the environment. */
export function fontContractReport(repoRoot: string, fcMatchExecutable = 'fc-match') {
  return {
    file: relative(repoRoot, FONT_CONTRACT_FILE),
    sha256: createHash('sha256').update(readFileSync(FONT_CONTRACT_FILE)).digest('hex'),
    families: assertFontContract(fcMatchExecutable),
  };
}

/**
 * The drift probe document: one multi-line paragraph per contract family, in that family, at
 * 11pt in a 6.5in column. Wrapping is the metric-sensitive observable — if either renderer
 * resolves a family differently, advance widths change and the paragraph wraps to a different
 * number of lines. The pangram mixes widths so a substitution cannot hide behind quantization.
 */
export const PROBE_TEXT =
  'Sphinx of black quartz, judge my vow; 0123456789 — pack my box with five dozen liquor jugs. '
    .repeat(4)
    .trim();

/** Marker prefixing each family's paragraph so both renderers' output can be attributed. */
export const probeMarker = (index: number): string => `@@F${index}@@`;

export function generateFontProbeDocx(): Uint8Array {
  const paragraphs = FONT_CONTRACT.map((entry, index) =>
    `<w:p><w:pPr><w:spacing w:before="0" w:after="120" w:line="240" w:lineRule="auto"/></w:pPr>` +
    `<w:r><w:rPr><w:rFonts w:ascii="${entry.family}" w:hAnsi="${entry.family}"/><w:sz w:val="22"/></w:rPr>` +
    `<w:t xml:space="preserve">${probeMarker(index)} ${PROBE_TEXT}</w:t></w:r></w:p>`).join('\n  ');

  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`),
    },
    {
      name: '_rels/.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
  ${paragraphs}
  <w:sectPr>
    <w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"
      w:header="720" w:footer="720" w:gutter="0"/>
    <w:cols w:space="720"/>
  </w:sectPr>
</w:body></w:document>`),
    },
  ]);
}

/** Lines per probe paragraph in a `pdftotext -layout` dump, attributed by marker order. */
export function probeLineCountsFromPdfText(text: string): number[] {
  const lines = text.split('\n').map(line => line.trim()).filter(Boolean);
  const starts = FONT_CONTRACT.map((_, index) =>
    lines.findIndex(line => line.includes(probeMarker(index))));
  if (starts.some(start => start < 0)) {
    throw new Error(`probe markers missing from pdftotext output:\n${text.slice(0, 500)}`);
  }
  return starts.map((start, index) =>
    (index + 1 < starts.length ? starts[index + 1] : lines.length) - start);
}
