# Session op descriptions

Status: every family described and checked (stages 1 and 2, issue #982). Generation (stage 3) is tracked in #1027.

## The problem

A session op reaches callers through seven hand-written layers, kept in lockstep by hand:

| Layer | File | What it repeats |
|---|---|---|
| Facade | `Docxodus/Internal/DocxSessionOps.cs` | the op and its parameters |
| Wire JSON | `Docxodus/Internal/DocxSessionJson.cs` | the result shape |
| WASM bridge | `wasm/DocxodusWasm/DocxSessionBridge.cs` | a `[JSExport]` shim per op |
| npm | `npm/src/types.ts`, `npm/src/session.ts` | types and a wrapper per op |
| stdio host | `tools/python-host/Dispatcher.cs` | an op name and argument parsing |
| Python | `python/src/docx_scalpel/{types,session}.py` | types and a wrapper per op |
| MCP | `tools/mcp-server/{ToolCatalog,Dispatcher}.cs` | a JSON schema and argument parsing |

These layers are where drift shows up. Recent examples: per-transport defaults (#960), argument names that differ between dispatchers (#1014), wire logic that only one transport had (#960), and an editor that kept its own copy of the result shape (#969).

## The design

Describe each op **once**, as data. Then **check** every layer against the description, and later **generate** the repetitive layers from it.

A description lives in `tools/op-descriptions/<family>.json`. For each op it records:

- `op`: the canonical name, and `facade`, the `DocxSessionOps` method that owns it.
- `args`: each argument's canonical wire name, its type, whether it is required, and the `default` the facade applies when it is not.
- `result`: the result type.
- `transports`: each layer's name for the op (the WASM export, npm method, Python method, stdio op, MCP tool and action). MCP may list `aliases`, further tool/action pairs that call the same op with the same arguments. A Python wrapper that builds its arguments in a helper names it in `argsFrom` (`file:[Class.]function`).
- Recorded divergences, per transport:
  - `argNames`: the transport spells an argument differently from its canonical name (the MCP name, where the stdio host and MCP differ);
  - `defaults`: the transport supplies its own default for an argument;
  - `requires`: the transport requires an argument the facade treats as optional;
  - `flattens` (MCP): the tool takes an object argument's fields as separate top-level properties, recorded as field to MCP property;
  - `absent`: the transport does not expose the op at all, with the reason.

`tools/op-descriptions/not-described.json` lists the `DocxSessionOps` methods deliberately left out, each with its reason: the lifecycle, transaction, preview and delivery-evidence entry points, which are stateful or composite, plus a few internal or not separately exposed methods.

Divergences are recorded so that they stay visible and can't grow. They are not the target. Each one should eventually be removed in the facade, the way #960 removed per-transport defaults, and then deleted from the file.

### Why data, and why a checked-in file

The options considered in #982 were:

- **C# attributes on the facade methods, read by a source generator.** This keeps the description beside the code. But TypeScript and Python can't read it, and the MCP schema text (descriptions, enums) would not fit in attributes.
- **A checked-in schema file plus a generator script.** Every language can read it, it diffs cleanly in review, and it can be adopted one family at a time. This is the chosen option.
- **`System.Text.Json` source-generated contracts for the wire types.** This is complementary. It would replace the hand-built `StringBuilder` result JSON. It does not describe ops or arguments.

## Adoption, one family at a time

**Stage 1: describe and check (landed for comments).** `Docxodus.Tests/SessionOpDescriptionDriftTests.cs` reads the descriptions and fails on any of these:

- A public `DocxSessionOps` method is neither described nor listed in `not-described.json` (`OPD000`), so a new op cannot bypass the description.
- A description is malformed: a divergence names an unknown argument, or a transport is neither named nor recorded absent with a reason (`OPD006`).
- The facade method is missing (reflection).
- The MCP tool's schema does not advertise the action or one of the arguments, under its recorded name.
- The WASM bridge does not export the op.
- The npm wrapper does not call that export.
- The Python wrapper does not send the stdio op with each required argument's recorded name.
- The stdio host or the MCP server, called for real on a seeded session, refuses the described arguments.
- The stdio host or the MCP server accepts a call that omits a required argument. The exception is an MCP default the description records, which must still be in effect.

The static checks run over every described op in every family; the checks that call the stdio host and the MCP server for real run on the comment family. Writing the comment family's description surfaced four differences between the stdio host and MCP that no test had pinned. #1014 removed them: the comment is `commentAnchorId` on every transport (the stdio host still accepts `parentAnchorId` and `anchorId` as deprecated aliases), and `DocxSessionOps` owns the `markdown` and `resolved` defaults. The description now records those defaults on the arguments themselves, and the family has no divergences left.

**Stage 2: describe the remaining families (landed).** Every family is described: about 150 ops in 19 files, from `text` and `tables` to `rendering`. Describing them recorded 186 divergences, filed by kind: argument names that differ between the stdio host and MCP (#1023), defaults supplied by a transport instead of the facade (#1024), object arguments MCP spreads into top-level properties (#1025), and transports that do not expose an op (#1026). Each is to be removed in the facade, as #960 and #1014 did, and then deleted from the file.

**Stage 3: generate.** Once a family's description is complete and its divergences resolved, generate the layers that are pure repetition:

- the stdio host's argument parsing and the MCP argument parsing, both into calls on the facade;
- the MCP `ToolCatalog` schema properties, which take their prose from a `description` field to be added to the file;
- the WASM `[JSExport]` shims;
- the TypeScript and Python type declarations and thin wrappers.

A generator emits into checked-in files, which keeps review diffs readable and builds offline. A drift test, the stage-1 test reused, fails when the generated output is stale.

### What stays hand-written

These open questions from #982 are answered as follows:

- **Ergonomics stay hand-written.** That covers Python docstrings, TypeScript overloads and convenience methods that compose several ops. The generated layer is the thin wrapper; hand-written code calls it.
- **Grouped-intent MCP tools stay.** The MCP surface is one tool per intent with an `action` argument, while the stdio host uses one op per name. The description maps both; it does not flatten one into the other.
- **Composite and stateful ops stay hand-written.** That means sessions, transactions, previews and history. They are not one facade call with flat arguments, and are out of scope for generation.

## Changing a described op

Change the facade first, as always. Then update the family's description in the same change. If the change adds a transport-specific spelling or default, record it under `argNames` or `defaults`, and say why in the pull request.
