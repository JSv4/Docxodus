"""Require a live re-measurement to reproduce a cached result exactly.

A pull request compares its candidate against baseline and main results that an earlier run on
main measured and cached. This canary re-measures the baseline live against the same cached
reference and fails unless every field of the result is identical, apart from the PDF byte hashes:
a PDF embeds per-run metadata, so its bytes differ between runs while every measured value matches.
"""
import json
import sys

# Fields that legitimately differ between two runs of the same renderer in the same environment.
PER_RUN = {('cases', '*', 'adapter', 'pdfSha256')}


def leaves(value, path=()):
    if isinstance(value, dict):
        for key, child in value.items():
            yield from leaves(child, path + (key,))
    elif isinstance(value, list):
        for index, child in enumerate(value):
            yield from leaves(child, path + (index,))
    else:
        yield path, value


def per_run(path):
    pattern = tuple('*' if isinstance(part, int) else part for part in path)
    return pattern in PER_RUN


def differences(cached, live):
    a, b = dict(leaves(cached)), dict(leaves(live))
    return sorted((p for p in a.keys() | b.keys() if not per_run(p) and a.get(p, KeyError) != b.get(p, KeyError)),
                  key=lambda p: [str(x) for x in p])


def main(cached_path, live_path):
    with open(cached_path) as cached, open(live_path) as live:
        found = differences(json.load(cached), json.load(live))
    if not found:
        print('Live re-measurement reproduces the cached result exactly.')
        return 0
    print(f'Cached result is not reproducible in this environment ({len(found)} differing fields):')
    for path in found[:40]:
        print('  ' + '.'.join(str(part) for part in path))
    print('Re-run the workflow from main (workflow_dispatch, mode=warm) or bump the cache version.')
    return 1


if __name__ == '__main__':
    sys.exit(main(*sys.argv[1:]))
