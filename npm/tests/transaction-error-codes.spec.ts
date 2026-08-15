import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import type { EditErrorCode } from '../src/types.js';

// The MCP server produces these codes; the browser package has no transaction surface of its
// own (idempotent retries are MCP-only). It does type every EditError it decodes, so the union
// has to name them. Declaring the list as EditErrorCode[] makes `npm run typecheck` fail if a
// member is dropped or misspelled, and the source assertion below catches it at runtime too.
const TRANSACTION_ERROR_CODES: readonly EditErrorCode[] = [
  'invalid_transaction',
  'transaction_conflict',
  'transaction_result_evicted',
  'transaction_incomplete',
];

const typesSource = readFileSync(
  fileURLToPath(new URL('../src/types.ts', import.meta.url)),
  'utf8',
);

test('EditErrorCode names every MCP mutation-transaction wire string', () => {
  for (const code of TRANSACTION_ERROR_CODES) {
    expect(typesSource).toContain(`| "${code}"`);
  }
  expect(new Set(TRANSACTION_ERROR_CODES).size).toBe(TRANSACTION_ERROR_CODES.length);
});
