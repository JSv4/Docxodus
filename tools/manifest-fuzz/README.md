# manifest-fuzz

Fuzzing harnesses for `PackageManifestGenerator` — the hand-written ZIP/OPC/CFB describer
behind package manifests (`docs/architecture/package_manifests.md`). A bespoke parser of
untrusted bytes earns its trust by surviving adversarial input at scale, so re-run a campaign
whenever `Docxodus/Verification/` changes.

## The oracle

The manifest contract, enforced on every arbitrary input under both default and
adversarially-tiny `PackageManifestOptions`:

1. `Generate()` never throws (malformed/encrypted input becomes findings, not exceptions);
2. repeated generation is byte-identical (`ToJsonBytes()` compared);
3. the caller's input buffer is never mutated;
4. canonical JSON always parses; `ToJson(indented: true)` never throws;
5. no input hangs the generator (20 s watchdog) or kills the process.

## Layout

| Path | What it is |
|---|---|
| `fuzzer/` | Self-contained feedback-driven havoc fuzzer. No external tooling needed. Corpus evolution is guided by a semantic coverage proxy: every novel `packageKind` × finding-code-set × digest-shape combination is kept, since each finding code maps to a distinct parser branch. Checks the full oracle. |
| `afl/` | AFL++ persistent-mode harness (SharpFuzz IL instrumentation, true edge coverage). Checks invariant 1; the replay step adds the rest. |
| `replay/` | Runs every input in the given directories through the full oracle under both option profiles. |
| `make_seeds.py` | Generates the structure-aware seed corpus: rich OPC packages (ZIP64, duplicate entries, encryption flags, directory entries, hostile part names, DTDs, truncated/prepended/appended archives), a spec-correct CFB compound file with `EncryptedPackage`/`EncryptionInfo` streams, raw non-ZIP inputs, and the smallest committed `TestFiles` fixtures. |
| `afl.dict` | ZIP/OPC/CFB token dictionary for AFL++. |

All outputs (seeds, corpora, logs, repros) go to `work/`, which is gitignored.

## Running

```bash
./run-campaign.sh                 # havoc campaign: 25 min, nproc-1 workers
./run-campaign.sh 300 4           # quick pass: 5 min, 4 workers
./run-afl.sh                      # AFL++ fleet + frontier replay: 40 min
```

`run-afl.sh` needs `afl-fuzz` (`apt install afl++`) and the SharpFuzz CLI
(`dotnet tool install --global SharpFuzz.CommandLine`). Both scripts exit non-zero if any
oracle violation was recorded; repros land under `work/out/{crashes,bugs,hangs}` and
`work/afl-out/*/crashes`.

## Validate the detector before trusting a clean run

A zero-crash result is only meaningful if the pipeline can prove it would have caught a bug.
Plant a crash in `afl/Program.cs` (e.g. throw when the input starts with three known bytes),
give the seed corpus an input one bit-flip away from the trigger, and confirm afl-fuzz saves a
crash within seconds. Note that the planted branch lives in the *uninstrumented* harness, so
AFL's coverage-equivalent seed trimming can destroy a distant trigger prefix — keep the control
seed adjacent to the trigger. Real defects in the instrumented `Docxodus.dll` are not subject
to that caveat.

## Baseline evidence

Campaign of 2026-08-18 against the initial implementation (`da044f9` and ancestors), before
any of this had run in anger:

- havoc campaign: **180.4 M executions** (7 workers × 25 min), ~98 k distinct manifest
  shapes, ~1.4 M deep determinism/non-mutation checks — zero violations;
- AFL++ fleet: **303.5 M executions** (7 instances × 40 min, varied power schedules),
  ~1,430 distinct execution paths, `pending_favs` exhausted (converged) — zero crashes,
  zero hangs;
- full-oracle replay of all 24,049 frontier inputs under both option profiles — zero failures;
- detector positive control: planted crash found and saved in 140 executions.
