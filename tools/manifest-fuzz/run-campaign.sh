#!/usr/bin/env bash
# Feedback-driven fuzzing campaign for PackageManifestGenerator.
# usage: ./run-campaign.sh [duration-seconds=1500] [workers=nproc-1]
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
DURATION=${1:-1500}
WORKERS=${2:-$(( $(nproc) - 1 ))}
WORK="$HERE/work"
mkdir -p "$WORK/logs"
python3 "$HERE/make_seeds.py" "$WORK/seeds" "$HERE/../../TestFiles"
dotnet build "$HERE/fuzzer/fuzzer.csproj" -c Release -v q
BIN="$HERE/fuzzer/bin/Release/net10.0/fuzzer.dll"
DEADLINE=$(( $(date +%s) + DURATION ))
for i in $(seq 1 "$WORKERS"); do
 (
  while :; do
    left=$(( DEADLINE - $(date +%s) )); [ "$left" -lt 10 ] && break
    dotnet "$BIN" --seeds "$WORK/seeds" --out "$WORK/out" --id "$i" \
      --seconds "$left" --rngseed $(( i * 7919 + RANDOM )) \
      >> "$WORK/logs/w$i.log" 2>&1 || echo "worker $i restarted" >> "$WORK/logs/w$i.log"
  done
 ) &
done
wait
echo "campaign complete:"
grep -h DONE "$WORK"/logs/w*.log | tail -n "$WORKERS" || true
c=$(ls "$WORK/out/crashes" 2>/dev/null | wc -l)
b=$(ls "$WORK/out/bugs" 2>/dev/null | wc -l)
h=$(ls "$WORK/out/hangs" 2>/dev/null | wc -l)
echo "crashes=$c oracle-bugs=$b hangs=$h (repros under $WORK/out)"
[ "$c$b$h" = "000" ]
