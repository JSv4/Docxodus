import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import {
  copyFileSync,
  readFileSync,
  readdirSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs';
import { dirname, extname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const packageRoot = dirname(dirname(fileURLToPath(import.meta.url)));
const dist = join(packageRoot, 'dist');
const framework = join(dist, 'wasm', '_framework');
const reportSchemaSource = resolve(packageRoot, '..', 'docs', 'schemas', 'render-report-v1.schema.json');
const reportSchemaDestination = join(dist, 'render-report-v1.schema.json');
const packageJson = JSON.parse(readFileSync(join(packageRoot, 'package.json'), 'utf8'));

copyFileSync(reportSchemaSource, reportSchemaDestination);

const requiredRoots = [
  join(dist, 'export-browser.bundle.js'),
  join(dist, 'docxodus.worker.js'),
  join(dist, 'pagination.bundle.js'),
  reportSchemaDestination,
  join(dist, 'export-resource-limits-v1.json'),
];
for (const file of requiredRoots) assert.ok(statSync(file).isFile(), `missing export asset: ${file}`);
assert.ok(statSync(framework).isDirectory(), `missing WASM framework directory: ${framework}`);

for (const bundle of ['export-browser', 'docxodus.worker']) {
  const metafilePath = join(dist, `${bundle}.meta.json`);
  const metafile = JSON.parse(readFileSync(metafilePath, 'utf8'));
  const externalImports = Object.values(metafile.outputs)
    .flatMap((output) => output.imports ?? [])
    .filter((entry) => entry.external);
  assert.deepEqual(externalImports, [], `${bundle} bundle must not retain external module imports`);
  unlinkSync(metafilePath);
}

const frameworkFiles = readdirSync(framework)
  .map((name) => join(framework, name))
  .filter((file) => statSync(file).isFile() && !file.endsWith('.br'));

const mediaTypes = {
  '.css': 'text/css',
  '.dat': 'application/octet-stream',
  '.js': 'text/javascript',
  '.json': 'application/json',
  '.wasm': 'application/wasm',
};

const assets = [...requiredRoots, ...frameworkFiles]
  .map((file) => {
    const bytes = readFileSync(file);
    const path = `./${relative(dist, file).split(sep).join('/')}`;
    return {
      path,
      mediaType: mediaTypes[extname(file)] ?? 'application/octet-stream',
      byteLength: bytes.byteLength,
      sha256: createHash('sha256').update(bytes).digest('hex'),
    };
  })
  .sort((left, right) => left.path.localeCompare(right.path));

const manifest = {
  schema: 'https://docxodus.dev/schemas/export/export-assets/v1',
  schemaVersion: 1,
  packageVersion: packageJson.version,
  assets,
};
writeFileSync(join(dist, 'export-assets.json'), `${JSON.stringify(manifest, null, 2)}\n`);
console.log(`export asset graph: ${assets.length} hashed runtime assets`);
