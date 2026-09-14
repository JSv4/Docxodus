// Adapted from the native-only reproduction linked in issue #788:
// https://github.com/JSv4/react-docxodus-viewer/blob/8e94365/scripts/benchmark-native-transactions.mjs
import { chromium } from '../../npm/node_modules/@playwright/test/index.mjs';
import { createServer } from 'node:http';
import { createReadStream } from 'node:fs';
import { mkdir, readFile, stat, writeFile } from 'node:fs/promises';
import { dirname, resolve, sep, extname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { cpus, platform, arch } from 'node:os';

// Native-only diagnostic: no React, editor canvas, rendering, or registered page map.
// Run browsers serially and keep builds/other CPU-heavy work out of the measurement.

const repo = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const roots = {
  batch: resolve(process.env.DOCXODUS_BATCH_ROOT || resolve(repo, 'npm')),
  typing: resolve(repo, 'npm'),
};
const fixture = await readFile(resolve(repo, 'TestFiles/NVCA-Model-COI.docx'));
const sha256 = createHash('sha256').update(fixture).digest('hex');
if (sha256 !== 'd75600769c12724990de48149d7a2bb161f3522daa54b1783672f93697d87d29') throw new Error('NVCA fixture hash mismatch');
const packages = Object.fromEntries(await Promise.all(Object.entries(roots).map(async ([label, root]) => [label, {
  version: JSON.parse(await readFile(resolve(root, 'package.json'), 'utf8')).version,
  sessionHash: createHash('sha256').update(await readFile(resolve(root, 'dist/session.js'))).digest('hex'),
  libraryHash: createHash('sha256').update(await readFile(resolve(root, 'dist/wasm/_framework/Docxodus.wasm'))).digest('hex'),
  nativeHash: createHash('sha256').update(await readFile(resolve(root, 'dist/wasm/_framework/dotnet.native.wasm'))).digest('hex'),
}])));
const server = createServer(async (req, res) => {
  res.setHeader('Cross-Origin-Opener-Policy', 'same-origin');
  res.setHeader('Cross-Origin-Embedder-Policy', 'require-corp');
  res.setHeader('Cache-Control', 'no-store');
  const url = new URL(req.url, 'http://localhost');
  if (url.pathname === '/') { res.setHeader('Content-Type', 'text/html'); res.end('<!doctype html><title>Native batch benchmark</title>'); return; }
  if (url.pathname === '/fixture.docx') { res.end(fixture); return; }
  const [,version,...parts] = decodeURIComponent(url.pathname).split('/');
  const root = roots[version], file = root && resolve(root, ...parts);
  if (!file || !file.startsWith(root + sep)) { res.writeHead(404).end(); return; }
  try {
    if (!(await stat(file)).isFile()) throw new Error('Not a file');
    res.setHeader('Content-Type', ({'.js':'text/javascript','.wasm':'application/wasm','.json':'application/json'})[extname(file)] || 'application/octet-stream');
    createReadStream(file).pipe(res);
  } catch { res.writeHead(404).end(); }
});
await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
const base = `http://127.0.0.1:${server.address().port}/`;
const browser = await chromium.launch({headless:true});
const runs = [];
const startedAt = new Date().toISOString();
try {
  // Alternate operations in fresh contexts, with no parallel browser workload.
  for (const label of [...Object.keys(roots), ...Object.keys(roots).reverse()]) {
    const context = await browser.newContext();
    const page = await context.newPage();
    const errors = [];
    page.on('pageerror', error => errors.push(error.message));
    await page.goto(base);
    const result = await page.evaluate(async ({label, version}) => {
      const api = await import(`/${label}/dist/core.js`);
      await api.initialize(`/${label}/dist/wasm/`);
      const bytes = new Uint8Array(await (await fetch('/fixture.docx')).arrayBuffer());
      const s = api.openDocxSession(bytes, {emitMarkdownPatch:false, persistAnchorIds:true});
      const anchor = Object.entries(s.project().anchorIndex).find(([,item]) => item.scope==='body' && item.textPreview.startsWith('The Certificate of Incorporation'))[0];
      const formatting = () => s.getFormatting(anchor).runs.map(({text, span, effective}) => ({text, span, effective}));
      const beforeRuns = formatting();
      const before = beforeRuns.map(run=>run.text).join('');
      const at = before.indexOf('document');
      if (at<0) throw new Error('Fixture paragraph did not match');
      const bridge = api.getWasmExports().DocxSessionBridge;
      const calls = [];
      for (const name of ['BeginTransaction','GetPackageContentHash','ReplaceTextAtSpan','ReplaceTextAtSpanWithFormat','ApplyFormat','ListRevisions']) {
        const original=bridge[name];
        if (!original) continue;
        bridge[name]=function(...args) { const start=performance.now(); try {return original.apply(this,args);} finally {calls.push({name,ms:performance.now()-start});} };
      }
      const match = {text:'document', enclosingAnchor:{id:anchor,kind:'p',scope:'body',unid:anchor.split(':').at(-1)},span:{start:at,length:8},fragments:[],contextBefore:'',contextAfter:'',groups:[]};
      const attempts = [];
      // First transaction plus two repeats on the same session after undo/redo verification.
      for (let iteration = 0; iteration < 3; iteration++) {
        const versionBefore = s.getVersion();
        calls.length = 0;
        const start=performance.now();
        const batch=label === 'typing' ? s.replaceMatch(match,'new document',{bold:true}) : s.executeBatch([
          {tool:'benchmark',action:'text',mutation:()=>s.replaceMatch(match,'new document')},
          {tool:'benchmark',action:'format',mutation:()=>s.applyFormat(anchor,{start:at,length:12},{bold:true})},
        ]);
        const ms=performance.now()-start;
        const measuredCalls = [...calls];
        if(!batch.success) throw new Error(JSON.stringify(batch));
        if(label === 'batch' && !batch.packageHash) throw new Error('Batch receipt hash missing');
        if(s.getVersion() !== versionBefore + 1) throw new Error('Batch must advance the native version once');
        const afterRuns = formatting();
        if(afterRuns.map(run=>run.text).join('')!==before.slice(0,at)+'new document'+before.slice(at+8)) throw new Error('Text mismatch');
        const formatted = afterRuns.filter(run=>run.span.start<at+12 && run.span.start+run.span.length>at);
        if(!formatted.length || !formatted.every(run=>run.effective.bold)) throw new Error('Formatting mismatch');
        const restored = () => JSON.stringify(formatting()) === JSON.stringify(beforeRuns);
        if(!s.undo() || !restored()) throw new Error('Native undo must restore text and formatting together');
        if(!s.redo() || JSON.stringify(formatting()) !== JSON.stringify(afterRuns)) throw new Error('Native redo mismatch');
        if(!s.undo() || !restored()) throw new Error('Native undo after redo mismatch');
        attempts.push({iteration, ms, calls:measuredCalls, hashPresent:!!batch.packageHash,
          checks:{text:true, bold:true, oneVersion:true, undo:true, redo:true}});
      }
      s.close();
      return {version,attempts,crossOriginIsolated};
    }, {label, version: packages[label].version});
    runs.push({label,...result,errors}); console.log(JSON.stringify({label,...result,errors}));
    await context.close();
  }
  const output=process.argv[2] || '/tmp/docxodus-atomic-typing.json';
  await mkdir(dirname(output), {recursive:true});
  await writeFile(output,JSON.stringify({startedAt,sha256,packages,browser:browser.version(),
    machine:{cpu:cpus()[0]?.model, logicalCpus:cpus().length, platform:platform(), arch:arch()},
    workload:'Text-plus-Bold atomic batch versus replaceMatch with format; first attempt and two repeats after undo/redo, in fresh ABBA browser contexts.',
    cpuThrottling:false,runs},null,2)+'\n');
} finally {await browser.close();await new Promise(resolve=>server.close(resolve));}
