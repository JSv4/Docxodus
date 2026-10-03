import { build } from 'esbuild';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * Re-emit dist/editor.js with its third-party drag-and-drop dependencies inlined (issue #854).
 *
 * `@atlaskit/pragmatic-drag-and-drop` has no `exports` map: the editor imports it by directory
 * (`…/element/adapter`), and its own files import each other without file extensions. Bundlers
 * accept both, but Node's ESM resolver rejects them, so `import 'docxodus'` failed in plain Node
 * before calling anything. Inlining those packages here leaves dist/editor.js importing only the
 * package's own modules, by relative path, so it still shares `core.js` — and its single WASM
 * runtime — with every other entry point.
 *
 * Runs as part of `build:ts`, right after `tsc` writes the unbundled file, so no build path can
 * ship the unresolvable imports.
 */
const packageRoot = dirname(dirname(fileURLToPath(import.meta.url)));

await build({
  entryPoints: [join(packageRoot, 'src', 'editor.ts')],
  outfile: join(packageRoot, 'dist', 'editor.js'),
  bundle: true,
  format: 'esm',
  platform: 'browser',
  target: 'es2020',
  sourcemap: true,
  allowOverwrite: true,
  logLevel: 'warning',
  plugins: [{
    name: 'keep-package-modules-external',
    setup(builder) {
      // The editor's own modules stay separate files (tsc already emitted them); only bare
      // package specifiers are inlined.
      builder.onResolve({ filter: /^\.\.?\// }, (args) =>
        args.importer.startsWith(join(packageRoot, 'src')) ? { path: args.path, external: true } : undefined);
    },
  }],
});
