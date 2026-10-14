#!/bin/bash
# Runs the packaging validation that publish.yml performs at release time, without
# publishing anything.
#
# Why this exists: publish.yml only triggers on a `v*` tag, so nothing exercised its
# steps until release day — the most expensive moment to discover breakage. Two bugs
# reached the v10.0.0 tag that way, both latent for weeks:
#
#   * @docxodus/export's package-boundary check assumed a flat dist/. The verified font
#     runtime (#529) added dist/fonts/, which index.js imports, so the check rejected a
#     package that was correct.
#   * publish.yml never set kernel.apparmor_restrict_unprivileged_userns=0, so Chromium
#     could not start its sandbox and npm-export's tests failed.
#
# publish.yml and ci.yml both call this script, so the packaging contract is asserted on
# every pull request and there is one definition of it rather than two that drift.
#
# Deliberately browser-free: both boundary checks are `npm pack --dry-run`, so this adds
# well under a minute to a PR. npm-export's browser suite is covered by playwright.yml,
# which already installs Chromium and permits its user-namespace sandbox.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

# Skip `npm ci` when node_modules is already present and the caller asked us to
# (CI installs separately; a local run usually has one).
SKIP_INSTALL="${RELEASE_PREFLIGHT_SKIP_INSTALL:-0}"

install_if_needed() {
  local dir="$1"
  if [ "$SKIP_INSTALL" = "1" ] && [ -d "$dir/node_modules" ]; then
    echo "  (using existing node_modules)"
    return
  fi
  (cd "$dir" && npm ci)
}

echo "==> docxodus (npm/)"
install_if_needed "$REPO_ROOT/npm"
# The boundary check inspects dist/, so the bundles have to exist first. A tag build
# rebuilds anyway; this keeps the script correct when run on its own.
if [ ! -f "$REPO_ROOT/npm/dist/index.js" ]; then
  echo "  dist/ missing — building"
  (cd "$REPO_ROOT/npm" && npm run build)
fi
(cd "$REPO_ROOT/npm" && npm run test:package-boundary)

echo "==> @docxodus/export (npm-export/)"
install_if_needed "$REPO_ROOT/npm-export"
(cd "$REPO_ROOT/npm-export" && npm run build && npm run test:package-boundary)

echo
echo "Release preflight passed: both packages pack within their declared boundaries."
