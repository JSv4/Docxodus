#!/usr/bin/env python3
"""Compare benchmark --stats-json output from a PR's base commit and its head.

Usage: perf-compare.py BASE.json HEAD.json [--label NAME] [--max-alloc-growth 0.10]

Each file maps a case name to {"MedianMs": float, "AllocBytes": float}. The gate is on
allocation, which repeats to well under 1% between runs of the same build; wall time on a
shared CI runner does not, so it is reported but never fails the check. A case fails when
it allocates more than --max-alloc-growth above base AND at least 1 MiB more, so tiny
cases cannot trip on rounding. Prints a Markdown table, appends it to $GITHUB_STEP_SUMMARY
when set, and exits 1 if any case regressed.
"""

import argparse
import json
import os
import sys

MIB = 1024 * 1024


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("base")
    parser.add_argument("head")
    parser.add_argument("--label", default="benchmark")
    parser.add_argument("--max-alloc-growth", type=float, default=0.10)
    args = parser.parse_args()

    with open(args.base, encoding="utf-8") as f:
        base = json.load(f)
    with open(args.head, encoding="utf-8") as f:
        head = json.load(f)

    lines = [
        f"### {args.label}",
        "",
        "| case | alloc base | alloc head | alloc Δ | time base | time head | time Δ | |",
        "|---|---:|---:|---:|---:|---:|---:|---|",
    ]
    regressions = []
    for case in sorted(set(base) | set(head)):
        if case not in head:
            lines.append(f"| `{case}` | | | | | | | removed |")
            continue
        h = head[case]
        if case not in base:
            lines.append(f"| `{case}` | | {h['AllocBytes'] / MIB:.1f} MiB | | | {h['MedianMs']:.0f} ms | | new |")
            continue
        b = base[case]
        alloc_ratio = h["AllocBytes"] / b["AllocBytes"] if b["AllocBytes"] else 1.0
        time_ratio = h["MedianMs"] / b["MedianMs"] if b["MedianMs"] else 1.0
        grew = h["AllocBytes"] - b["AllocBytes"]
        failed = alloc_ratio > 1 + args.max_alloc_growth and grew >= MIB
        if failed:
            regressions.append(case)
        lines.append(
            f"| `{case}` | {b['AllocBytes'] / MIB:.1f} MiB | {h['AllocBytes'] / MIB:.1f} MiB | {alloc_ratio - 1:+.1%} "
            f"| {b['MedianMs']:.0f} ms | {h['MedianMs']:.0f} ms | {time_ratio - 1:+.1%} | {'❌ allocation' if failed else ''} |")

    lines.append("")
    if regressions:
        lines.append(
            f"**{len(regressions)} case(s) allocate more than {args.max_alloc_growth:.0%} above base:** "
            + ", ".join(f"`{c}`" for c in regressions))
    else:
        lines.append(f"No case allocates more than {args.max_alloc_growth:.0%} above base. "
                     "Time is informational: shared runners are too noisy to gate on it.")
    report = "\n".join(lines) + "\n"
    print(report)
    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        with open(summary, "a", encoding="utf-8") as f:
            f.write(report + "\n")
    return 1 if regressions else 0


if __name__ == "__main__":
    sys.exit(main())
