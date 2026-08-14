# The Docxodus puzzle eval

A level is a pair of documents and a budget: **transform `start` into `target`, in at most
`par` tool calls, using only the agent editing surface.** A level is solved when
`DocxDiff` between the player's document and the target returns **zero revisions** — the
scoring function is the comparison engine itself, so a solve is exact rather than judged.

## Why this exists

The arcade proved the *addressing model*: anchors survive a save/reopen round-trip, and the
editor's reconciler stays incremental while a host mutates the session behind it. It proved
nothing about the surface we actually ask agents to drive, because it drives `raw.replaceXml`
— the escape hatch for callers who want to hand-author OOXML.

This pack asks the question a buyer asks: **can an agent, given only the grouped tools and
anchor addressing, reach a specified document state?** The answer is a number, and the
failures are the interesting output — a tool description an agent misreads, an error message
it cannot recover from, an affordance that turns one edit into ten.

## Level format

`eval/puzzles/<id>/level.json`:

| Field | Meaning |
|-------|---------|
| `id`, `title` | Identity. |
| `par` | Minimum mutating tool calls a competent solve takes. The score is calls-used vs par. |
| `brief` | What the player is told. Written as a task, never as a list of operations. |
| `start`, `target` | Paragraph lists, built into DOCX by the harness. Text, so a level diffs in review. |
| `reference` | A worked solution at par, addressed by content rather than by anchor id. |

Documents are declared as paragraph lists rather than committed as binaries so a level stays
readable, and so the harness and any external runner build byte-identical fixtures from the
same source.

### Content addressing

`reference` steps never name an anchor id — ids are minted per build and a player has to
discover them anyway. Steps address blocks the way a solver must:

```json
{ "op": "moveBlock", "find": "Governing Law", "relativeTo": "Indemnification", "position": "before" }
```

`find` matches the first body block whose text contains the string. That is deliberately the
same two-step shape the real surface imposes — search, then act on what you found — so a
reference solution cannot accidentally be easier than the task.

## Running it

The reference solutions run in the .NET suite, which is what keeps par honest:

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~PuzzleEval"
```

Those tests assert three things per level: the reference solution reaches zero revisions
against the target, it does so within par, and the *starting* document does not already score
as solved. The third is the one that catches a level whose target was built wrong — without
it, an empty solution would pass.

## Playing a level with an agent

The levels are transport-agnostic; the harness that scores them is the same one that scores a
model. Point a model at the MCP server with the level's `brief` and its start document, let it
work, then score its output against `target.docx` with `DocxDiff`. The existing stdio runner
(`tools/mcp-server/smoke/mcp_probe.py`) already speaks the protocol and captures variables
between calls, so a model transcript replays without new tooling.

Report a run as `solved/total` plus calls-vs-par, and keep the failures: a level that no model
solves is a finding about the surface, not about the model.
