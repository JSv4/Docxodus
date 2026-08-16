import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import {
  FONT_SUBSTITUTION_CONTRACT,
  type FontSubstitutionEntry,
} from '../src/font-contract.js';
import {
  FONT_CONTRACT_FILE,
  FONT_CONTRACT_PACKAGES,
} from './visual-parity/font-contract.js';

interface FontConfigSubstitution {
  family: string;
  substitute: string;
}

function substitutionsFromFontConfig(xml: string): FontConfigSubstitution[] {
  const matchBlocks = Array.from(xml.matchAll(/<match\s+target="pattern">([\s\S]*?)<\/match>/gu));
  const substitutions = matchBlocks.map(([, block], index) => {
    const family = block.match(
      /<test\s+name="family"><string>([^<]+)<\/string><\/test>/u,
    )?.[1];
    const substitute = block.match(
      /<edit\s+name="family"\s+mode="assign"\s+binding="same"><string>([^<]+)<\/string><\/edit>/u,
    )?.[1];
    expect(family, `fonts.conf match ${index + 1} must declare one source family`).toBeTruthy();
    expect(substitute, `fonts.conf match ${index + 1} must assign one substitute family`).toBeTruthy();
    return { family: family!, substitute: substitute! };
  });
  expect(substitutions.length, 'fonts.conf must contain at least one substitution').toBeGreaterThan(0);
  return substitutions;
}

test.describe('visual-parity font substitution contract', () => {
  test('fonts.conf aliases exactly match the browser-portable production contract', () => {
    const configured = substitutionsFromFontConfig(readFileSync(FONT_CONTRACT_FILE, 'utf8'));
    const production = FONT_SUBSTITUTION_CONTRACT.map(({ family, substitute }) => ({
      family,
      substitute,
    }));

    expect(configured).toEqual(production);
    expect(new Set(configured.map(({ family }) => family)).size).toBe(configured.length);
  });

  test('deployment package hints remain complete and test-local', () => {
    expect(Object.keys(FONT_CONTRACT_PACKAGES).sort()).toEqual(
      FONT_SUBSTITUTION_CONTRACT.map(({ family }) => family).sort(),
    );
    for (const entry of FONT_SUBSTITUTION_CONTRACT as readonly FontSubstitutionEntry[]) {
      expect(entry).not.toHaveProperty('package');
      expect(FONT_CONTRACT_PACKAGES[entry.family]).toBeTruthy();
    }
  });
});
