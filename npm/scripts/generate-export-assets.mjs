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
const reportSchemaSources = [1, 2].map((version) =>
  resolve(packageRoot, '..', 'docs', 'schemas', `render-report-v${version}.schema.json`));
const reportSchemaDestinations = [1, 2].map((version) =>
  join(dist, `render-report-v${version}.schema.json`));
const limitsContractSource = join(packageRoot, 'src', 'export-resource-limits-v1.json');
const limitsContractDestination = join(dist, 'export-resource-limits-v1.json');
const packageJson = JSON.parse(readFileSync(join(packageRoot, 'package.json'), 'utf8'));

reportSchemaSources.forEach((source, index) =>
  copyFileSync(source, reportSchemaDestinations[index]));
copyFileSync(limitsContractSource, limitsContractDestination);

const requiredRoots = [
  join(dist, 'export-browser.bundle.js'),
  join(dist, 'docxodus.worker.js'),
  join(dist, 'pagination.bundle.js'),
  ...reportSchemaDestinations,
  limitsContractDestination,
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
  .sort((left, right) => left.path < right.path ? -1 : left.path > right.path ? 1 : 0);

const manifest = {
  schema: 'https://docxodus.dev/schemas/export/export-assets/v1',
  schemaVersion: 1,
  packageVersion: packageJson.version,
  assets,
};
writeFileSync(join(dist, 'export-assets.json'), `${JSON.stringify(manifest, null, 2)}\n`);
console.log(`export asset graph: ${assets.length} hashed runtime assets`);
