// Preserve the exact runtime and JS used by one arm before building the next.
import { cp, mkdir, readFile, readdir, stat, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const repo = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const [arm, profilePath, root = '/tmp/docxodus-issue-783/bundles'] = process.argv.slice(2);
if (!/^[a-z-]+$/.test(arm ?? '') || !profilePath) throw new Error('Usage: node snapshot.mjs ARM PROFILE_PATH [BUNDLE_ROOT]');
const dest = join(resolve(root), arm);
await mkdir(resolve(root), { recursive: true });
await mkdir(dest); // Refuse to overwrite a measured arm.
await cp(join(repo, 'npm/dist/wasm/_framework'), join(dest, '_framework'), { recursive: true });
await cp(join(repo, 'npm/tests/test-harness.html'), join(dest, 'test-harness.html'));
for (const name of ['editor.bundle.js', 'session.bundle.js']) await cp(join(repo, 'npm/dist', name), join(dest, name));
await cp(resolve(profilePath), join(dest, 'profile.aotprofile'));
const digest = async path => createHash('sha256').update(await readFile(path)).digest('hex');
const framework = join(dest, '_framework');
let rawBytes = 0, brotliBytes = 0;
for (const name of await readdir(framework)) {
  const size = (await stat(join(framework, name))).size;
  if (name.endsWith('.br')) brotliBytes += size;
  else rawBytes += size;
}
const metadata = {
  arm,
  sourceCommit: execFileSync('git', ['rev-parse', 'HEAD'], { cwd: repo, encoding: 'utf8' }).trim(),
  sdk: execFileSync('dotnet', ['--version'], { cwd: repo, encoding: 'utf8' }).trim(),
  profileSha256: await digest(join(dest, 'profile.aotprofile')),
  profileBytes: (await stat(join(dest, 'profile.aotprofile'))).size,
  nativeSha256: await digest(join(framework, 'dotnet.native.wasm')),
  nativeBytes: (await stat(join(framework, 'dotnet.native.wasm'))).size,
  editorSha256: await digest(join(dest, 'editor.bundle.js')),
  sessionSha256: await digest(join(dest, 'session.bundle.js')),
  rawBytes, brotliBytes,
  wireBudgetBytes: Number((await readFile(join(repo, 'scripts/build-wasm.sh'), 'utf8'))
    .match(/WIRE_BUDGET_BYTES=\$\(\((\d+) \* 1024\)\)/)[1]) * 1024,
};
await writeFile(join(dest, 'build.json'), JSON.stringify(metadata, null, 2) + '\n');
console.log(JSON.stringify(metadata, null, 2));
