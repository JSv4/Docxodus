"""Apply the unchanged oracle's findings to already-passing baseline cases.

Newly delivered pages of a formerly failed conversion change measured coverage;
retain that oracle finding, while permitting the mandatory failure to be fixed.
No tolerance or score adjustment is applied.
"""
import sys
from bench.common import read_json
from bench.runner import compare

before, after = map(read_json, sys.argv[1:])
findings = compare(before, after, 0)
passing = {case['id'] for case in before['cases'] if all(check['passed'] for check in case['checks'])}
regressions = [finding for finding in findings if any(finding.startswith(case + ':') or finding.startswith(case + ' page ') for case in passing)]
failed = [case['id'] for case in after['cases'] if not all(check['passed'] for check in case['checks'])]
if failed:
    regressions.append('Mandatory failures: ' + ', '.join(failed))
print('\n'.join(regressions) or f'All {len(after["cases"])} mandatory cases pass; no measurement regressed in {len(passing)} passing baseline cases.')
sys.exit(bool(regressions))
