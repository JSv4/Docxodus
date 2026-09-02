#!/bin/bash
# Re-record wasm/DocxodusWasm/docxodus.aotprofile, the method list that the WASM build
# compiles ahead of time (RunAOTCompilation + AOTProfilePath in DocxodusWasm.csproj).
# Everything not in the profile runs on the interpreter/jiterpreter, so a stale profile
# costs steady-state speed, never correctness. Re-run it when the hot paths move — a new
# engine stage, a renamed hot class — and check the wire size the final build prints.
#
# Three steps: a profiler build (AOT off — AOT-compiled methods are invisible to the
# profiler — with the Mono AOT profiler linked in), a browser run of the representative
# workload that dumps the profile (npm/tests/aot-profile-record.spec.ts), and a rebuild
# of the shipped configuration so dist/wasm never holds the profiler flavour.
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

echo "==> 1/3 profiler build (interpreter + AOT profiler)"
"$SCRIPT_DIR/build-wasm.sh" -p:RunAOTCompilation=false -p:WasmProfilers=aot

echo "==> 2/3 recording the profile over the representative workload"
cd "$REPO_ROOT/npm"
DOCXODUS_RECORD_AOT_PROFILE=1 npm test -- aot-profile-record.spec.ts --project=chromium --reporter=line

echo "==> 3/3 rebuilding the shipped configuration"
npm run build:wasm
ls -la "$REPO_ROOT/wasm/DocxodusWasm/docxodus.aotprofile"
