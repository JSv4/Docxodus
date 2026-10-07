#!/usr/bin/env bash
# Allocation regression check (issue #976): run the DocxDiff stress and complex-form benchmarks
# against BASE_REF's library and against this checkout's, then compare allocations.
#
#   scripts/perf-run.sh BASE_REF OUT_DIR
#
# Both sides run THIS checkout's harness code: it is copied into a worktree of BASE_REF, where
# its ../../Docxodus project reference picks up the base library. So the only difference between
# the two runs is the library under test, on the same machine, back to back. Allocation repeats
# to well under 1% between runs of one build, which is what makes it gateable; time is reported
# only. If the harness cannot build or run against the base library (a PR that changed an API it
# uses), the comparison is skipped with a notice rather than failed.
set -euo pipefail

base_ref=$1
out=$(mkdir -p "$2" && cd "$2" && pwd)
repo=$(git rev-parse --show-toplevel)
doc="$repo/TestFiles/NVCA-Model-COI.docx"
harnesses=(docxdiff-stress complex-form-doc)

base_tree="$(mktemp -d)/base"
git -C "$repo" worktree add --detach --quiet "$base_tree" "$base_ref"
trap 'git -C "$repo" worktree remove --force "$base_tree"' EXIT
for h in "${harnesses[@]}"; do
  rm -rf "${base_tree:?}/benchmarks/$h"
  cp -r "$repo/benchmarks/$h" "$base_tree/benchmarks/$h"
  rm -rf "$base_tree/benchmarks/$h/bin" "$base_tree/benchmarks/$h/obj"
done

stress() { dotnet run -c Release --project "$1/benchmarks/docxdiff-stress" -- "$doc" --iterations 3 --warmup 1 --stats-json "$out/$2-stress.json"; }
form() { dotnet run -c Release --project "$1/benchmarks/complex-form-doc" -- "$doc" --stats-json "$out/$2-form.json"; }

# The base side only needs to produce numbers; a failed invariant check there is main's problem,
# not this PR's, so only a missing stats file counts as "could not run".
if ! stress "$base_tree" base > "$out/base.log" 2>&1 \
  || { form "$base_tree" base >> "$out/base.log" 2>&1 || true; [ ! -s "$out/base-form.json" ]; }; then
  echo "::notice title=Perf comparison skipped::The benchmark harness could not build or run against the base commit's library (an API change?). See base.log in the perf-results artifact."
  tail -n 40 "$out/base.log"
  exit 0
fi

# The head side runs for real: the complex-form harness exits non-zero if any of its redline
# invariants (accept-all = revised, reject-all = baseline, no new schema findings) fails.
stress "$repo" head | tee "$out/head.log"
form "$repo" head | tee -a "$out/head.log"

status=0
python3 "$repo/scripts/perf-compare.py" "$out/base-stress.json" "$out/head-stress.json" \
  --label "DocxDiff stress harness (NVCA COI)" || status=1
python3 "$repo/scripts/perf-compare.py" "$out/base-form.json" "$out/head-form.json" \
  --label "Complex-form benchmark (NVCA COI)" || status=1
exit $status
