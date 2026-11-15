// The pages and files that tell search engines and agents what Docxodus is —
// llms.txt, the sitemap, the landing page's structured data, package tags —
// are prose nothing compiles. This keeps them true: every link resolves to a
// file in the repository, and every API they name exists on the surface they
// name it for.
import assert from 'node:assert/strict';
import test from 'node:test';
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

const repoRoot = fileURLToPath(new URL('../../../', import.meta.url));
const read = (relative) => readFileSync(join(repoRoot, relative), 'utf8');

const RAW = 'https://raw.githubusercontent.com/JSv4/Docxodus/main/';
const SITE = 'https://jsv4.github.io/Docxodus/';

/** The repository file a published URL is served from, or null for an external URL. */
function repoFileFor(url) {
  if (url.startsWith(RAW)) return url.slice(RAW.length);
  if (url.startsWith(SITE)) {
    const path = url.slice(SITE.length);
    return join('docs', path === '' || path.endsWith('/') ? `${path}index.html` : path);
  }
  return null;
}

/** All source text under a directory, for "is this name defined anywhere here". */
function sources(dir, extension) {
  let text = '';
  for (const entry of readdirSync(join(repoRoot, dir), { withFileTypes: true })) {
    if (entry.name === 'bin' || entry.name === 'obj' || entry.name === 'node_modules') continue;
    const relative = join(dir, entry.name);
    if (entry.isDirectory()) text += sources(relative, extension);
    else if (entry.name.endsWith(extension)) text += read(relative);
  }
  return text;
}

const llms = read('docs/llms.txt');
const section = (heading) => {
  const start = llms.indexOf(`## ${heading}\n`);
  assert.notEqual(start, -1, `llms.txt has no "## ${heading}" section`);
  const end = llms.indexOf('\n## ', start + 1);
  return llms.slice(start, end === -1 ? undefined : end);
};
const catalog = read('tools/mcp-server/ToolCatalog.cs');
const mcpTools = [...catalog
  .matchAll(/new ToolDefinition\(\s*"(docxodus_[a-z_]+)"/g)].map((m) => m[1]);

test('llms.txt follows the llmstxt.org shape', () => {
  const lines = llms.split('\n');
  assert.match(lines[0], /^# Docxodus$/, 'the first line is the H1 project name');
  assert.match(lines[2], /^> \S/, 'a blockquote summary follows the H1');
  assert.ok(llms.includes('\n## Optional\n'), 'secondary links sit under "## Optional"');
  for (const line of lines.filter((l) => l.includes(']('))) {
    assert.match(line, /^- \[[^\]]+\]\(https:\/\/[^)\s]+\): \S/,
      `link lines are "- [name](url): notes" — got: ${line}`);
  }
});

test('every llms.txt and sitemap link to this project resolves to a checked-in file', () => {
  const urls = [
    ...[...llms.matchAll(/\]\((https:[^)]+)\)/g)].map((m) => m[1]),
    ...[...read('docs/sitemap.xml').matchAll(/<loc>([^<]+)<\/loc>/g)].map((m) => m[1]),
  ];
  assert.ok(urls.length > 20, 'the link scan found the links');
  for (const url of urls) {
    const file = repoFileFor(url);
    if (file === null) continue;
    assert.ok(existsSync(join(repoRoot, file)) && statSync(join(repoRoot, file)).isFile(),
      `${url} points at ${file}, which does not exist`);
  }
});

test('the sitemap lists the site URLs and nothing off-site', () => {
  const locs = [...read('docs/sitemap.xml').matchAll(/<loc>([^<]+)<\/loc>/g)].map((m) => m[1]);
  assert.ok(locs.includes(`${SITE}demo/`) && locs.includes(`${SITE}llms.txt`));
  for (const loc of locs) assert.ok(loc.startsWith(SITE), `${loc} is not on the Pages site`);
});

test('every API llms.txt names exists on the surface it names it for', () => {
  const corpus = {
    '.NET': sources('Docxodus', '.cs'),
    npm: sources('npm/src', '.ts') + sources('npm-export/src', '.ts'),
    Python: sources('python/src/docx_scalpel', '.py'),
  };
  const cliCommands = { redline: 'tools/redline', docx2html: 'tools/docx2html', docx2oc: 'tools/docx2oc' };
  const exportBins = Object.keys(JSON.parse(read('npm-export/package.json')).bin);
  let checked = 0;

  for (const line of section('Task to API').split('\n').filter((l) => l.startsWith('- '))) {
    // "- Task: .NET `A.B`; npm `c`; Python `d`; MCP `e`; CLI `f g`."
    for (const part of line.slice(line.indexOf(': ') + 2).split('; ')) {
      const surface = Object.keys(corpus).concat('MCP', 'CLI').find((s) => part.includes(`${s} `));
      for (const [, code] of part.matchAll(/`([^`]+)`/g)) {
        checked++;
        if (code.startsWith('docxodus_')) {
          assert.ok(mcpTools.includes(code), `llms.txt names MCP tool ${code}, which ToolCatalog does not define`);
        } else if (surface === 'CLI') {
          const command = code.split(' ')[0];
          assert.ok(cliCommands[command] ? existsSync(join(repoRoot, cliCommands[command])) : exportBins.includes(command),
            `llms.txt names the command "${command}", which no tool here provides`);
        } else if (surface === 'MCP') {
          // An argument value, as in "`docxodus_get_content` with format `html`".
          assert.ok(catalog.includes(`"${code}"`), `llms.txt names MCP value "${code}", which ToolCatalog does not accept`);
        } else {
          assert.ok(surface, `cannot tell which surface \`${code}\` belongs to in: ${line}`);
          // `new DocxSession(bytes)` / `session.find_by_text(...)` / `@docxodus/export` → the names.
          const names = code.startsWith('@') ? [] : code.replace(/\([^)]*\)/g, '')
            .split(/[.\s]+/).filter((n) => n && n !== 'new' && n !== 'session');
          for (const name of names) {
            assert.match(corpus[surface], new RegExp(`\\b${name}\\b`),
              `llms.txt names ${surface} \`${name}\`, which is not in that surface's source`);
          }
        }
      }
    }
  }
  assert.ok(checked > 40, `the scan read the API names (${checked})`);
});

test('the advertised MCP tool count matches the catalog', () => {
  for (const file of ['docs/llms.txt', 'README.md', 'tools/mcp-server/README.md']) {
    const counts = [...read(file).matchAll(/\b(\d+) tools\b/g)].map((m) => Number(m[1]));
    assert.ok(counts.length > 0, `${file} states the tool count`);
    for (const count of counts) assert.equal(count, mcpTools.length, `${file} tool count`);
  }
});

test('the landing page carries valid structured data and one canonical URL', () => {
  const html = read('docs/demo/index.html');
  const [, json] = html.match(/<script type="application\/ld\+json">([\s\S]*?)<\/script>/) ?? [];
  assert.ok(json, 'the landing page has a JSON-LD block');
  const data = JSON.parse(json);
  assert.equal(data['@type'], 'SoftwareApplication');
  assert.equal(data.name, 'Docxodus');
  assert.ok(data.featureList.length >= 5 && data.description.length > 80);
  const canonical = html.match(/<link rel="canonical" href="([^"]+)"/)?.[1];
  assert.equal(canonical, html.match(/<meta property="og:url" content="([^"]+)"/)?.[1]);
  assert.equal(canonical, data.url);
  const llmsLink = html.match(/<link rel="alternate" type="text\/plain"[^>]*href="([^"]+)"/)?.[1];
  assert.equal(repoFileFor(llmsLink), 'docs/llms.txt');
});

test('package search metadata names only what the packages do', () => {
  // SpreadsheetML and PresentationML support was removed; tags outlived it once.
  const removed = /\b(xlsx|excel|pptx|powerpoint|spreadsheet|presentation)\b/i;
  const csproj = (file, tag) => read(file).match(new RegExp(`<${tag}>([^<]*)</${tag}>`))[1];
  const npm = JSON.parse(read('npm/package.json'));
  const python = read('python/pyproject.toml');
  const fields = {
    'Docxodus.csproj PackageTags': csproj('Docxodus/Docxodus.csproj', 'PackageTags'),
    'Docxodus.csproj Description': csproj('Docxodus/Docxodus.csproj', 'Description'),
    'npm keywords': npm.keywords.join(';'),
    'npm description': npm.description,
    'pyproject keywords': python.match(/keywords = \[([\s\S]*?)\]/)[1],
  };
  for (const [where, value] of Object.entries(fields)) {
    assert.doesNotMatch(value, removed, `${where} advertises a format Docxodus no longer handles`);
    // The three things people search this project for, in every registry.
    const terms = where.startsWith('npm') ? ['docx', 'redline', 'tracked', 'javascript', 'typescript'] : ['docx', 'redline', 'tracked'];
    for (const term of terms) {
      assert.match(value.toLowerCase(), new RegExp(term), `${where} is missing "${term}"`);
    }
  }
});
