"""Print the cache key component that identifies this runner's measurement environment.

Cached references and results may only be compared with live measurements taken in the same
environment. The key combines what the pinned benches themselves bind into every result contract
(tool versions and hashes, and the hash of every font file each suite resolves) with what they do
not: the LibreOffice reference build, the runner image, and the installed measurement packages.
Any change to any of them produces a different key, so a pull request restores nothing and
measures everything live.
"""
import argparse
import hashlib
import json
import os
import subprocess
import sys

PACKAGES = ['libreoffice-writer', 'poppler-utils', 'qpdf', 'fontconfig', 'fonts-liberation2']

# Run inside each bench checkout: both ship a package named `bench`, so each is imported in its own
# interpreter. The registry is the CI suite for the LibreOffice oracle and the capture-pack manifest
# for the Word harness, exactly as `bench run` resolves them.
FINGERPRINT = '''
import json, sys
from bench.common import identity, read_json, validate_registry
from bench.environment import environment
manifest, suite = sys.argv[1], sys.argv[2]
registry = validate_registry(read_json(manifest) if manifest else None)
if suite:
    from bench.suites import select_suite
    registry = select_suite(registry, suite)
print(identity(environment([f for c in registry['cases'] for f in c['fonts']])['contract']))
'''


def compose(components):
    """The key for a set of named components: any change to any value changes the key."""
    return hashlib.sha256(json.dumps(components, sort_keys=True).encode()).hexdigest()


def bench_fingerprint(checkout, manifest='', suite=''):
    return subprocess.run([sys.executable, '-c', FINGERPRINT, manifest, suite], cwd=checkout,
                          capture_output=True, text=True, check=True).stdout.strip()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--oracle', required=True)
    parser.add_argument('--suite', required=True)
    parser.add_argument('--word-oracle', required=True)
    parser.add_argument('--word-manifest', required=True, help='Relative to --word-oracle')
    args = parser.parse_args()
    packages = subprocess.run(['dpkg-query', '-W', '-f', '${Package}=${Version}\n', *PACKAGES],
                              capture_output=True, text=True, check=True).stdout.split()
    components = {
        'libreofficeEnvironment': bench_fingerprint(args.oracle, suite=args.suite),
        'wordEnvironment': bench_fingerprint(args.word_oracle, manifest=args.word_manifest),
        'libreofficeReference': subprocess.run([os.environ['BENCH_LIBREOFFICE'], '--version'],
                                               capture_output=True, text=True, check=True).stdout.strip(),
        'image': [os.environ.get('ImageOS', ''), os.environ.get('ImageVersion', '')],
        'packages': packages,
    }
    print(json.dumps(components, indent=2, sort_keys=True), file=sys.stderr)
    print(compose(components))


if __name__ == '__main__':
    main()
