#!/usr/bin/env bash
# Coverage-guided AFL++ campaign via SharpFuzz IL instrumentation, then a full-oracle
# replay of every discovered queue entry.
# usage: ./run-afl.sh [duration-seconds=2400] [secondaries=nproc-2]
# needs: apt install afl++    and    dotnet tool install --global SharpFuzz.CommandLine
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
command -v afl-fuzz >/dev/null || { echo "afl-fuzz not found (apt install afl++)"; exit 1; }
SHARPFUZZ=$(command -v sharpfuzz || echo "$HOME/.dotnet/tools/sharpfuzz")
[ -x "$SHARPFUZZ" ] || { echo "sharpfuzz not found (dotnet tool install --global SharpFuzz.CommandLine)"; exit 1; }
DURATION=${1:-2400}
SEC=${2:-$(( $(nproc) - 2 ))}
WORK="$HERE/work"
mkdir -p "$WORK/afl-logs"
python3 "$HERE/make_seeds.py" "$WORK/seeds" "$HERE/../../TestFiles"
dotnet build "$HERE/afl/aflharness.csproj" -c Release -v q
BIN="$HERE/afl/bin/Release/net10.0"
if ! "$SHARPFUZZ" "$BIN/Docxodus.dll" 2> "$WORK/sharpfuzz.err"; then
  if grep -qi "already instrumented" "$WORK/sharpfuzz.err"; then
    echo "Docxodus.dll already instrumented; continuing"
  else
    cat "$WORK/sharpfuzz.err"; exit 1
  fi
fi
export AFL_SKIP_BIN_CHECK=1 AFL_I_DONT_CARE_ABOUT_MISSING_CRASHES=1 AFL_NO_UI=1 AFL_SKIP_CPUFREQ=1
timeout "$DURATION" afl-fuzz -i "$WORK/seeds" -o "$WORK/afl-out" -t 10000 -m none \
  -x "$HERE/afl.dict" -M main -- dotnet "$BIN/aflharness.dll" \
  > "$WORK/afl-logs/main.log" 2>&1 &
sleep 2
SCHED=(fast explore exploit coe quad rare)
for i in $(seq 1 "$SEC"); do
  X=""; [ $((i % 2)) -eq 0 ] && X="-x $HERE/afl.dict"
  timeout "$DURATION" afl-fuzz -i "$WORK/seeds" -o "$WORK/afl-out" -t 10000 -m none \
    -p "${SCHED[$(( (i-1) % 6 ))]}" $X -S "s$i" -- dotnet "$BIN/aflharness.dll" \
    > "$WORK/afl-logs/s$i.log" 2>&1 &
done
wait || true
echo "afl campaign complete:"
grep -H -E 'execs_done|saved_crashes|saved_hangs' "$WORK"/afl-out/*/fuzzer_stats || true
echo "replaying the discovered frontier through the full oracle:"
dotnet build "$HERE/replay/replay.csproj" -c Release -v q
dotnet "$HERE/replay/bin/Release/net10.0/replay.dll" "$WORK"/afl-out/*/queue "$WORK/seeds"
