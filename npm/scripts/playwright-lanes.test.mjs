import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { readdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

import { planLaneRuns, serialLaneSpecs } from './playwright-lanes.mjs';

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
  assert.deepEqual([...pool.keys()], ['chromium', 'chromium-pixel5']);
  assert.deepEqual([...serial.keys()], ['chromium-serial', 'firefox-cross-block-selection']);

  const poolFiles = new Set([...pool.get('chromium'), ...pool.get('chromium-pixel5')]);
  const serialFiles = serial.get('chromium-serial');
  const inBoth = [...poolFiles].filter(file => serialFiles.has(file));
  assert.deepEqual(inBoth, [], 'a spec file must belong to one lane');
  const inNeither = [...allSpecFiles].filter(file => !poolFiles.has(file) && !serialFiles.has(file));
  assert.deepEqual(inNeither, [], 'every spec file must be run by one lane');
  // A misspelt entry would match no file and quietly leave that spec in the pool.
  assert.deepEqual([...serialFiles].sort(), [...serialLaneSpecs].sort(), 'the serial lane runs exactly the listed specs');
  assert.ok(poolFiles.size > serialFiles.size, 'the serial lane is the exception, not the rule');
});

test('without a lane, an ad-hoc run still reaches every project', () => {
  const projects = [...listSpecFiles(undefined).keys()];
  assert.deepEqual(projects, ['chromium', 'chromium-pixel5', 'chromium-serial', 'firefox-cross-block-selection']);
});

test('an unknown lane is refused instead of silently running nothing', () => {
  const result = spawnSync(process.execPath, [playwrightCli, 'test', '--list'], {
    cwd: npmDir, env: { ...process.env, DOCXODUS_PLAYWRIGHT_LANE: 'fast' }, encoding: 'utf8',
  });
  assert.notEqual(result.status, 0);
  assert.match(result.stderr + result.stdout, /DOCXODUS_PLAYWRIGHT_LANE/);
});
