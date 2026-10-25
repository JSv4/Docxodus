"""Protect separate provider results; accept Word adjudication only with live proof.

The original LibreOffice oracle, references, scores and findings remain intact.
All Word metrics are protected at zero tolerance. Only the exact reviewed cases
may use Word as layout authority, and only if both renderer revisions produce
identical rasters from their original and Word-capture input packages in this run.
"""
import argparse
from pathlib import Path
import sys

from bench.common import contained, digest, identity, read_json, validate_registry, write_json
from bench.runner import compare, validate_bundle

METRICS = ('score', 'ssim', 'inkF1', 'precision', 'recall')


def mandatory_failures(result):
    return [c['id'] for c in result['cases'] if c['status'] != 'measured' or
            not c['checks'] or not all(check['passed'] for check in c['checks'])]


def decide(lo_before, lo_after, lo_main, word_before, word_after, policy):
    if policy['schema'] != 1 or policy['referenceTrack'] != 'word-online' or policy['layoutProfile'] != 'default':
        raise ValueError('Unknown provider authority policy')
    for result in (lo_before, lo_after, lo_main):
        if result['track'] != 'libreoffice' or result['contract'].get('layoutProfile', 'default') != 'default':
            raise ValueError('Expected direct LibreOffice reference and default renderer layout')
    for result in (word_before, word_after):
        if result['track'] != policy['referenceTrack'] or result['contract'].get('layoutProfile') != policy['layoutProfile']:
            raise ValueError('Expected registered Word reference and explicit renderer layout')
    lo_findings = compare(lo_before, lo_after, 0)
    compare(lo_main, lo_after, 0)  # Validate the original-input main run's shared contract.
    word_findings = compare(word_before, word_after, 0)
    failures = []
    for label, result in [('LibreOffice candidate', lo_after), ('original-input main', lo_main),
                          ('Word main', word_before), ('Word candidate', word_after)]:
        bad = mandatory_failures(result)
        if bad:
            failures.append(label + ' mandatory failures: ' + ', '.join(bad))
    mappings = policy['cases']
    originals = [m['originalCase'] for m in mappings]
    proxies = [m['wordCase'] for m in mappings]
    if len(set(originals)) != len(originals) or len(set(proxies)) != len(proxies):
        raise ValueError('Duplicate provider authority mapping')
    if any(set(proxies) != {c['id'] for c in r['cases']} for r in (word_before, word_after)):
        raise ValueError('Word authority requires exact registered case coverage')
    lo_records = [{c['id']: c for c in r['cases']} for r in (lo_main, lo_after)]
    word_records = [{c['id']: c for c in r['cases']} for r in (word_before, word_after)]
    equivalence = []
    for mapping in mappings:
        original, proxy = mapping['originalCase'], mapping['wordCase']
        for label, actual, captured in zip(('main', 'candidate'), lo_records, word_records):
            if original not in actual:
                raise ValueError('Unknown original authority case: ' + original)
            a, b = actual[original]['pages'], captured[proxy]['pages']
            same = (bool(a) and len(a) == len(b) and
                    all(x.get('raster', {}).get('candidateSha256') and
                        x['raster']['candidateSha256'] == y.get('raster', {}).get('candidateSha256')
                        for x, y in zip(a, b)))
            equivalence.append({'originalCase': original, 'wordCase': proxy, 'renderer': label,
                                'originalAndCaptureInputRastersIdentical': bool(same)})
            if not same:
                failures.append(f'{original}: {label} capture input differs from original CI raster')
        candidate = word_records[1][proxy]
        if candidate.get('score', -1) < policy['minimumWordScore']:
            failures.append(proxy + ': below required Word score')
        for label, captured in zip(('main', 'candidate'), word_records):
            if not captured[proxy]['pages'] or any(any(key not in p for key in METRICS) for p in captured[proxy]['pages']):
                failures.append(proxy + ': missing required Word ' + label + ' page metrics')
    word_ok = not word_findings and not failures
    # Preserve the existing protection for every passing historical baseline
    # case. A reviewed Word case is adjudicated only after all Word safeguards pass.
    passing = {c['id'] for c in lo_before['cases'] if c['checks'] and all(x['passed'] for x in c['checks'])}
    accepted = set(originals) if word_ok else set()
    findings_by_case = {cid: [f for f in lo_findings if f.startswith(cid + ':') or f.startswith(cid + ' page ')]
                        for cid in sorted(passing)}
    protected_findings = [f for cid, findings in findings_by_case.items() if cid not in accepted for f in findings]
    failures += word_findings + protected_findings
    return {'schema': 1, 'layoutProfile': policy['layoutProfile'],
            'referenceTracks': {'libreoffice': {'allFindings': lo_findings, 'protectedFindings': protected_findings,
                                              'baselineSummary': lo_before['summary'], 'candidateSummary': lo_after['summary']},
                                'word-online': {'allFindings': word_findings, 'baselineSummary': word_before['summary'],
                                                'candidateSummary': word_after['summary']}},
            'inputRasterEquivalence': equivalence, 'wordAuthoritativeCases': sorted(accepted),
            'reviewedLibreOfficeProviderDifferences': {cid: findings_by_case.get(cid, []) for cid in sorted(accepted)},
            'failures': failures, 'passed': not failures}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for name in ('libreoffice-baseline', 'libreoffice-candidate', 'libreoffice-main', 'word-baseline',
                 'word-candidate', 'word-registry', 'word-reference', 'libreoffice-registry',
                 'policy', 'capture-pack', 'main-commit', 'candidate-commit', 'out'):
        parser.add_argument('--' + name, required=True)
    args = parser.parse_args()
    policy = read_json(args.policy)
    if digest(args.capture_pack) != policy['capturePack']['sha256'] or Path(args.capture_pack).stat().st_size != policy['capturePack']['bytes']:
        raise ValueError('Immutable capture pack changed')
    wr = validate_registry(read_json(args.word_registry))
    bundle = read_json(args.word_reference)
    paths = [args.libreoffice_baseline, args.libreoffice_candidate, args.libreoffice_main,
             args.word_baseline, args.word_candidate]
    lb, lc, lm, wb, wc = map(read_json, paths)
    for result, expected in [(lb, policy['libreOfficeBaselineCommit']), (lc, args.candidate_commit),
                             (lm, args.main_commit), (wb, args.main_commit), (wc, args.candidate_commit)]:
        if result['candidate'] != expected:
            raise ValueError('Renderer source commit differs from CI revision')
    for result in (wb, wc):
        if result['contract']['suiteDigest'] != identity(wr) or result['contract']['referenceDigest'] != identity(bundle):
            raise ValueError('Word run does not bind the immutable suite and reference bundle')
    validate_bundle(bundle, Path(args.word_reference).resolve().parent, wr,
                    {'contract': wc['contract']['environment']})
    if set(bundle['cases']) != {c['id'] for c in wr['cases']} or not all(c['eligible'] for c in bundle['cases'].values()):
        raise ValueError('Word authority requires every eligible registered reference')
    original_registry_path = Path(args.libreoffice_registry).resolve()
    original_registry = {c['id']: c for c in read_json(original_registry_path)['cases']}
    captured_registry = {c['id']: c for c in wr['cases']}
    for mapping in policy['cases']:
        original, captured = original_registry[mapping['originalCase']], captured_registry[mapping['wordCase']]
        if (original['sha256'] != mapping['originalInputSha256'] or
                digest(contained(original_registry_path.parent.parent, original['input'])) != mapping['originalInputSha256'] or
                captured['sha256'] != mapping['wordInputSha256'] or
                captured['originalCase']['id'] != mapping['originalCase'] or
                captured['originalCase']['inputSha256'] != mapping['originalInputSha256'] or
                captured['originalCase']['documentAndStylePropertyTreesEqual'] is not True):
            raise ValueError('Original/captured input mapping changed')
    result = decide(lb, lc, lm, wb, wc, policy)
    labels = ('libreofficeBaseline', 'libreofficeCandidate', 'libreofficeMain', 'wordBaseline',
              'wordCandidate', 'policy', 'wordRegistry', 'wordReference')
    result['evidenceSha256'] = {label: digest(p) for label, p in
                              zip(labels, paths + [args.policy, args.word_registry, args.word_reference])}
    result['capturePackSha256'] = digest(args.capture_pack)
    write_json(args.out, result)
    for track, evidence in result['referenceTracks'].items():
        print(track + ': ' + str(len(evidence['allFindings'])) + ' raw findings')
    print('\n'.join(result['failures']) if result['failures'] else
          f'Word: all metrics preserved; {len(result["wordAuthoritativeCases"])} cases adjudicated with exact raster proof. '
          'LibreOffice: all other passing baseline cases protected; all mandatory candidate checks pass.')
    return int(not result['passed'])


if __name__ == '__main__':
    try:
        sys.exit(main())
    except Exception as error:
        print(f'Provider gate error: {type(error).__name__}: {error}', file=sys.stderr)
        sys.exit(2)
