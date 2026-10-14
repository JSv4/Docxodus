#!/bin/bash
# Runs the epic #435 MCP acceptance smoke — the workflow and its reopen validation —
# against a scratch copy of TestFiles/NVCA-Model-COI.docx.
#
# Why this exists: both fixtures were checked in with revision assertions nobody
# re-ran, and they sat wrong on main for as long as it took someone to run the README
# by hand (#687). The fixtures are generated, so this script also regenerates them and
# fails if the committed JSON differs — a stale fixture is caught at the same moment as
# a broken engine.
#
# The gate is a superset of the assertions: mcp_probe.py exits nonzero for a transport
# error, an MCP tool error, an edit result with success:false, a batch that reports
# failed/partial, or any failed `expect`/`expectMembers` assertion. On top of that this
# script re-checks the two properties the runner reports rather than enforces — the five
# calls that must fail closed, and the byte-exact transaction replay.
#
# Usage:  scripts/mcp-smoke.sh [path/to/docxodus-mcp]
set -euo pipefail

cd "$(dirname "$0")/.."

SERVER="${1:-tools/mcp-server/bin/Release/net10.0/docxodus-mcp}"
if [[ ! -x "$SERVER" ]]; then
  echo "mcp-smoke: server not found at $SERVER" >&2
  echo "  build it with: dotnet build tools/mcp-server/mcpserver.csproj -c Release" >&2
  exit 1
fi

python3 tools/mcp-server/smoke/build_epic_435_fixtures.py --check

# The run SAVES. Pointing the storage root at TestFiles/ would overwrite a committed
# fixture, so it gets a scratch copy that goes away with the directory.
WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT
export DOCXODUS_STORAGE_ROOT="$WORK"
cp TestFiles/NVCA-Model-COI.docx "$WORK/local.docx"

run() {
  local fixture="$1" trace="$WORK/$2"
  python3 tools/mcp-server/smoke/mcp_probe.py \
    --calls "tools/mcp-server/smoke/$fixture" \
    --trace "$trace" \
    --quiet-server -- "$SERVER" >"$WORK/$2.summary"
  cat "$WORK/$2.summary"
  python3 - "$WORK/$2.summary" "$3" "$4" <<'PY'
import json
import sys

summary = json.load(open(sys.argv[1]))
expected = {"expectedFailures": int(sys.argv[2]), "replayComparisons": int(sys.argv[3])}
wrong = {k: (summary.get(k), v) for k, v in expected.items() if summary.get(k) != v}
if wrong:
    for key, (actual, want) in wrong.items():
        print(f"mcp-smoke: {key} = {actual}, expected {want}", file=sys.stderr)
    sys.exit(1)
PY
}

# The refusals and the replay are properties under test, not incidental counts: an
# engine that stopped refusing, or a fixture that stopped exercising the retry, would
# otherwise pass a run with zero failed assertions.
run epic-435-workflow.json workflow-trace.json 5 1
run epic-435-validation.json validation-trace.json 0 0

echo "mcp-smoke: OK"
