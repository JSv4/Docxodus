#!/usr/bin/env node
// Record how long each spec file's tests took, for `--split` to balance the CI pool shards.
//
// The shards are dealt by duration (see splitSpecFiles in playwright-lanes.mjs), so the
// recorded figures only need to be roughly right and only need refreshing when the suite's
// shape changes: a spec that is missing from the file is counted at the median. One spec's time
// varies by a third or more between runs, so pass several runs' reports; each spec's figure is
// its mean over the reports that ran it.
//
//   gh run download <run-id> --pattern 'playwright-lane-pool-*' --dir lanes
//   mkdir blobs && cp lanes/*/blob-report/*.zip blobs/
//   npx playwright merge-reports --reporter json blobs > report.json
//   node scripts/record-spec-durations.mjs report.json [more-reports.json ...]
import { readFileSync, writeFileSync } from 'node:fs';
import { basename } from 'node:path';
import { fileURLToPath } from 'node:url';

export const durationsPath = fileURLToPath(new URL('./playwright-spec-durations.json', import.meta.url));

/** Seconds per spec file name, summed over every test result (retries included) in one JSON report. */
export function specDurations(report) {
  const seconds = {};
  const visit = (suite, file) => {
    const owner = suite.file ? basename(suite.file) : file;
    for (const spec of suite.specs ?? []) {
      for (const test of spec.tests ?? []) {
        for (const result of test.results ?? []) {
          seconds[owner] = (seconds[owner] ?? 0) + result.duration / 1000;
        }
      }
    }
    for (const child of suite.suites ?? []) visit(child, owner);
  };
  for (const suite of report.suites ?? []) visit(suite, undefined);
  return seconds;
}

/** Each spec's mean duration, in whole seconds, over the reports that ran it. */
export function meanDurations(reports) {
  const runs = reports.map(specDurations);
  const names = [...new Set(runs.flatMap(Object.keys))].sort();
  return Object.fromEntries(names.map(name => {
    const seen = runs.filter(run => name in run).map(run => run[name]);
    return [name, Math.round(seen.reduce((sum, s) => sum + s, 0) / seen.length)];
  }));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  const reportPaths = process.argv.slice(2);
  if (reportPaths.length === 0) throw new Error('usage: record-spec-durations.mjs <merged JSON report>...');
  const durations = meanDurations(reportPaths.map(path => JSON.parse(readFileSync(path, 'utf8'))));
  writeFileSync(durationsPath, JSON.stringify(durations, null, 2) + '\n');
  const total = Object.values(durations).reduce((sum, s) => sum + s, 0);
  console.log(`Recorded ${Object.keys(durations).length} spec files, ${total} test-seconds, to ${durationsPath}`);
}
