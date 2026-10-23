import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { readdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

import { planLaneRuns, serialLaneSpecs, splitSpecFiles } from './playwright-lanes.mjs';

const npmDir = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const playwrightCli = createRequire(join(npmDir, 'package.json')).resolve('@playwright/test/cli');

test('a bare run plans the pool lane and then the serial lane, both strict', () => {
  const runs = planLaneRuns([]);
  assert.deepEqual(runs.map(run => run.lane), ['pool', 'serial']);
  for (const run of runs) {
    assert.equal(run.env.DOCXODUS_PLAYWRIGHT_LANE, run.lane);
    assert.deepEqual(run.args, ['test']);
  }
});

test('a narrowed run forwards the arguments to both lanes and tolerates an empty lane', () => {
  const runs = planLaneRuns(['worker.spec.ts', '--project=chromium-serial']);
  assert.deepEqual(runs[0].args, ['test', 'worker.spec.ts', '--project=chromium-serial', '--pass-with-no-tests']);
  assert.deepEqual(runs[1].args, runs[0].args);
});

test('a run that already tolerates an empty lane is not told twice', () => {
  const runs = planLaneRuns(['--pass-with-no-tests', 'worker.spec.ts']);
  assert.deepEqual(runs[0].args, ['test', '--pass-with-no-tests', 'worker.spec.ts']);
});

test('a preset lane runs only that lane', () => {
  assert.deepEqual(planLaneRuns([], { lanes: ['serial'] }).map(run => run.lane), ['serial']);
});

test('options alone do not narrow a run, so an empty lane still fails', () => {
  const runs = planLaneRuns(['--reporter=blob', '--workers=2']);
  assert.deepEqual(runs[0].args, ['test', '--reporter=blob', '--workers=2']);
  assert.deepEqual(planLaneRuns(['--grep', 'footnote'])[0].args, ['test', '--grep', 'footnote', '--pass-with-no-tests']);
});

const five = ['a.spec.ts', 'b.spec.ts', 'c.spec.ts', 'd.spec.ts', 'e.spec.ts'];

test('a split hands each spec file to exactly one shard, round-robin, in a stable order', () => {
  const shards = [1, 2, 3].map(index => splitSpecFiles(five, index, 3));
  assert.deepEqual(shards, [['a.spec.ts', 'd.spec.ts'], ['b.spec.ts', 'e.spec.ts'], ['c.spec.ts']]);
  assert.deepEqual(shards.flat().sort(), five);
});

test('--split=i/N becomes that shard\'s file filters and is not passed on to Playwright', () => {
  const runs = planLaneRuns(['--split=2/3', '--reporter=blob'], { lanes: ['pool'], specFiles: five });
  assert.deepEqual(runs, [{
    lane: 'pool',
    env: { DOCXODUS_PLAYWRIGHT_LANE: 'pool' },
    args: ['test', '/b\\.spec\\.ts$', '/e\\.spec\\.ts$', '--reporter=blob', '--pass-with-no-tests'],
  }]);
});

test('a malformed split is refused', () => {
  for (const bad of ['--split=0/3', '--split=4/3', '--split=2', '--split=a/b']) {
    assert.throws(() => planLaneRuns([bad], { specFiles: five }), /--split/, bad);
  }
});

// Ask Playwright itself which spec files each lane runs: the configuration is the only
// place that decides, and this is the shape every lane job in CI relies on.
function listSpecFiles(lane) {
  const env = { ...process.env, DOCXODUS_PLAYWRIGHT_LANE: lane };
  if (lane === undefined) delete env.DOCXODUS_PLAYWRIGHT_LANE;
  const result = spawnSync(process.execPath, [playwrightCli, 'test', '--list', '--reporter=json'], {
    cwd: npmDir, env, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024,
  });
  let report;
  try { report = JSON.parse(result.stdout); } catch { report = undefined; }
  if (result.status !== 0) {
    // Listing loads every spec file, so a spec whose import is missing (an unbuilt companion,
    // say) fails here; the JSON report carries those load errors, stderr usually nothing.
    const errors = (report?.errors ?? []).map(error => error.message ?? error.value ?? '').join('\n');
    assert.fail(`playwright test --list exited ${result.status}\n${errors}\n${result.stderr}`);
  }
  assert.ok(report, 'playwright test --list printed no JSON report');
  const filesByProject = new Map(report.config.projects.map(project => [project.name, new Set()]));
  const visit = (suite, file) => {
    for (const child of suite.suites ?? []) visit(child, file);
    for (const spec of suite.specs ?? []) {
      for (const entry of spec.tests) filesByProject.get(entry.projectName).add(file);
    }
  };
  for (const fileSuite of report.suites) visit(fileSuite, fileSuite.file);
  return filesByProject;
}

const allSpecFiles = new Set(readdirSync(join(npmDir, 'tests')).filter(name => name.endsWith('.spec.ts')));

test('the pool and serial lanes partition the spec files between their Chromium projects', () => {
  const pool = listSpecFiles('pool');
  const serial = listSpecFiles('serial');
  assert.deepEqual([...pool.keys()], ['chromium']);
  assert.deepEqual([...serial.keys()], ['chromium-serial', 'chromium-pixel5', 'firefox-cross-block-selection']);

  // Firefox re-runs one pool spec on purpose; the Chromium projects are the partition.
  const poolFiles = pool.get('chromium');
  const serialFiles = new Set([...serial.get('chromium-serial'), ...serial.get('chromium-pixel5')]);
  const inBoth = [...poolFiles].filter(file => serialFiles.has(file));
  assert.deepEqual(inBoth, [], 'a spec file must belong to one lane');
  const inNeither = [...allSpecFiles].filter(file => !poolFiles.has(file) && !serialFiles.has(file));
  assert.deepEqual(inNeither, [], 'every spec file must be run by one lane');
  // A misspelt entry would match no file and quietly leave that spec in the pool.
  assert.deepEqual([...serial.get('chromium-serial')].sort(), [...serialLaneSpecs].sort(),
    'the serial lane runs exactly the listed specs');
  assert.ok(poolFiles.size > serialFiles.size, 'the serial lane is the exception, not the rule');
});

test('without a lane, an ad-hoc run still reaches every project', () => {
  const projects = [...listSpecFiles(undefined).keys()];
  assert.deepEqual(projects, ['chromium', 'chromium-serial', 'chromium-pixel5', 'firefox-cross-block-selection']);
});

test('an unknown lane is refused instead of silently running nothing', () => {
  const result = spawnSync(process.execPath, [playwrightCli, 'test', '--list'], {
    cwd: npmDir, env: { ...process.env, DOCXODUS_PLAYWRIGHT_LANE: 'fast' }, encoding: 'utf8',
  });
  assert.notEqual(result.status, 0);
  assert.match(result.stderr + result.stdout, /DOCXODUS_PLAYWRIGHT_LANE/);
});
