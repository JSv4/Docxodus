import { buildSync } from 'esbuild';
import { fileURLToPath } from 'node:url';

// Exercise the built production module and its dependencies in isolated setContent pages.
export const wordTextScript = buildSync({
  entryPoints: [fileURLToPath(new URL('../dist/word-text.js', import.meta.url))],
  bundle: true, format: 'iife', globalName: 'wordText', write: false,
}).outputFiles[0].text + '\nObject.assign(window, wordText);';
