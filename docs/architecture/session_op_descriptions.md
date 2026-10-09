# Session op descriptions

Status: every family described and checked (stages 1 and 2, issue #982). Generation (stage 3, issue #1027) has landed for the comment family; the other families can adopt it once their divergences are resolved.

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
- `args`: each argument's canonical wire name, its type, whether it is required, and the `default` the facade applies when it is not. `nullable` marks a required argument that may be sent as null.
- `result`: the result type.
- `transports`: each layer's name for the op (the WASM export, npm method, Python method, stdio op, MCP tool and action). MCP may list `aliases`, further tool/action pairs that call the same op with the same arguments. A Python wrapper that builds its arguments in a helper names it in `argsFrom` (`file:[Class.]function`).
- Recorded divergences, per transport:
  - `argNames`: the transport spells an argument differently from its canonical name (the MCP name, where the stdio host and MCP differ);
  - `defaults`: the transport supplies its own default for an argument;
  - `requires`: the transport requires an argument the facade treats as optional;
  - `flatAliases` (MCP): deprecated flat top-level properties that the tool still accepts for fields of an object argument it takes nested, recorded as field to MCP property. These are not a divergence in what reaches the facade. `McpNestedOptionsTests` checks them against `tools/mcp-server/ObjectArguments.cs`;
  - `absent`: the transport does not expose the op at all, with the reason. A reason says why the op is kept off and, where one exists, the route a caller takes instead.
  - `route` (MCP, beside `absent`): the op is reached through a selector on a read tool rather than an action, recorded as `{ "tool": "docxodus_get_content", "selector": { "format": "version" } }`. Only one-to-one routes are recorded: the selector value takes the op's arguments unchanged and returns the facade's result unwrapped. `SessionOpRouteTests` checks each route: the tool's schema advertises the selector value, and calling the tool with it returns what the facade returns. A route is recorded beside `absent`, not in place of it, because the drift test's MCP checks expect an action.

A family that is generated (stage 3) also carries:

- `"generate": true` at the top level.
- `mcp.<tool>.properties.<name>.description`: the prose for each argument property the tool's schema advertises, in the order it advertises them. It is per tool, not per argument, because the schema merges every action's arguments into one property list ("add/reply: comment author").
- `transports.wasm.doc`: the XML summary of the `[JSExport]` shim.
- `transports.stdio.deprecatedAliases`: an old spelling of an argument the stdio host still accepts, as argument to old name. This is not a divergence: every transport takes the canonical name, and the alias only keeps old callers working.

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

**Stage 2: describe the remaining families (landed).** Every family is described: about 150 ops in 19 files, from `text` and `tables` to `rendering`. Describing them recorded 186 divergences, filed by kind: argument names that differ between the stdio host and MCP (#1023), defaults supplied by a transport instead of the facade (#1024), object arguments MCP spreads into top-level properties (#1025), and transports that do not expose an op (#1026). Each is to be removed in the facade, as #960 and #1014 did, and then deleted from the file. #1023 removed the argument names: the stdio host takes MCP's name for every argument, keeps its old spelling as a deprecated alias (`DocxSessionJson.AliasedArgument` and its typed readers), and no description records `argNames` any more. `StdioArgumentNameTests` calls the host for each renamed argument and checks it took effect, since the live-call checks here cover only the comment family. #1024 removed the per-transport defaults: every default a transport used to supply is now a named constant in `DocxSessionOps` (`DefaultInsertPosition`, `DefaultListFormat`, `DefaultProjectionDepth` and the rest), each transport passes an omitted argument through (as null, or as the WASM bridge's empty string), and the description records the default on the argument itself. Two supplied values were not defaults but refusals in disguise, because the session refuses the empty value, so `setImageDimensions`' `dimensions` and `repairRevisions`' `repairs` are required on every transport. `setImageMetadata`'s `altText` and `title` are required but `nullable`: a null removes the value, and the stdio host used to read an omitted one as null, so leaving one out silently cleared it. No description records `defaults` or `requires` any more. `FacadeDefaultTests` calls the stdio host and the MCP server with each defaulted argument they expose left out and checks the edit or answer matches the facade's constant; the npm-only `renderBlock` defaults are covered by `npm/tests/facade-defaults.spec.ts`. #1025 removed the spread object arguments. MCP takes the nested `options`, `rule` or `spec` object and parses it with the same `DocxSessionJson` parser as the stdio host. The flat properties remain as deprecated aliases, recorded as `flatAliases`, and no description records `flattens` any more. `McpNestedOptionsTests` calls MCP for each of the fifteen ops and checks the option took effect. #1026 settled the absences. The omissions are exposed: page setup, page numbering and the first-page/odd-even switches on MCP (`docxodus_create`), the horizontal rule on the stdio host and Python, and `listNotes` on npm, Python, the stdio host and MCP. `SessionTransportGapTests` calls each new exposure and checks its effect. Every remaining `absent` now gives a specific reason: browser-editor internals (block lists, editor and preview renders, move targets), MCP reads reached through a `docxodus_get_content` format or a `docxodus_search` mode, raw XML kept off the agent surface, and ops whose job another call already does. Two one-to-one selector routes, `getVersion` (`format=version`) and `getSemanticChanges` (`format=semantic_changes`), are recorded as `route` and checked. `project` (`format=markdown` without `anchorId`) also maps one to one and is the next to record.

**Stage 3: generate (landed for comments).** Once a family's description is complete and its divergences resolved, the layers that are pure repetition are generated from it. `tools/op-codegen` is a small C# console project (not in `Docxodus.sln`; the test project references it). For every family marked `"generate": true` it writes:

| Output | Where | What it holds |
|---|---|---|
| stdio parsing | `tools/python-host/Generated/<Family>Ops.cs` | `IsGenerated<Family>Op` and `DispatchGenerated<Family>Op`: each op's arguments read from the request and passed to the facade, deprecated aliases included |
| MCP parsing | `tools/mcp-server/Generated/<Family>Ops.cs` | `RunGenerated<Family>Action` (arguments into the facade call) and `ValidateGenerated<Family>Arguments` (the batch-step check that runs before any step does) |
| MCP schema | the same file, class `GeneratedMcpSchema` | per tool, `<Tool>Actions` (the action enum members) and `<Tool>Properties` (the argument properties with their prose) |
| WASM shims | `wasm/DocxodusWasm/DocxSessionBridge.cs`, between `// BEGIN GENERATED <family>` and `// END GENERATED <family>` | one `[JSExport]` method per op |

Both dispatchers are `partial` classes, so the generated code calls each transport's own argument helpers (`Str`, `OptStr`, `OptBool`, the span parsers) and throws its own exception type with its own messages. The WASM shims sit in a marked region of the hand-written bridge instead of a file of their own because `SessionOpDescriptionDriftTests` reads the bridge's exports from that one file.

Run it after changing a generated family's description:

```bash
dotnet run --project tools/op-codegen              # rewrite the generated files
dotnet run --project tools/op-codegen -- --check   # exit 1 if any is stale
```

`GeneratedSessionPlumbingTests` runs the generator in memory and fails when a checked-in file differs from its output (`OPG001`), so a description change without a regenerate, or a hand edit of generated code, fails the build's tests. It also checks that a description change does change the output (`OPG003`), that the description lists arguments in the facade's parameter order with matching types (`OPG004`), that no MCP schema names a property twice (`OPG005`), that each tool's schema carries its generated actions and properties in order (`OPG009`), and that no generated code lives outside the generator's outputs: no orphaned generated file, and no generated WASM shim defined outside its region (`OPG010`).

The generator relies on a few rules, and refuses a description that breaks them rather than guessing:

- **Arguments are positional.** The description lists them in the facade's parameter order (after `handle`, leaving out `preconditions`). The WASM shim takes them in the same order, which is the order the npm wrapper passes them.
- **Types it knows.** `anchor`, `string` and `markdown` are strings; `boolean`; and `charSpan`, an optional `{ start, length }` object. Anything else fails generation with the argument named, and supporting it means adding one reader per transport to `Emitters.cs`.
- **Defaults belong to the facade.** An optional argument is passed to the facade as null when absent. On WASM, where JavaScript passes every argument, an optional string with no default arrives as `""` and becomes null; one with a default passes through for the facade to apply.
- **One name, several ops.** When two ops share a stdio op or an MCP action, as `addComment` and `addCommentToRevision` share `add_comment` and `docxodus_comment add`, the caller picks one by naming exactly one required argument that the others lack (`anchorId` or `revisionId`). Naming none or both is refused, and so is naming another op's optional argument (`span` with `revisionId`).
- **No divergences.** A family marked `generate` cannot record `argNames`, `defaults`, `requires`, `flattens`, `absent` or `aliases` on any transport, because generated code has one spelling and one default per argument.

What stays hand-written for a generated family:

- The grouped-intent tool shell in the MCP dispatcher: which tool routes to the family, the `Guarded` precondition wrapper, the result envelope (`docxodus_comment list` returns `{"comments": [...]}` where the stdio host returns the bare array), and the tool's lead schema properties (`sessionId`, `preconditions`), into which `ToolCatalog` splices the generated fragments.
- The mutation lists both dispatchers keep: the stdio host's `IsMutation` and `IsBatchableMutation`, the MCP batch allowlist, and the MCP `MutationTarget` precondition inference. The description does not yet say which ops mutate.
- The TypeScript and Python types and wrappers. They are not generated yet; the drift test still checks them against the description.

**Adopting the next family:**

1. Resolve the family's recorded divergences in the facade (#1023–#1026), and delete them from its description.
2. Add `"generate": true`, the `mcp.<tool>.properties` prose (moved verbatim from `ToolCatalog`), each op's `wasm.doc`, and any stdio `deprecatedAliases`. Order the ops so that the MCP action enum comes out as it is today.
3. Add an empty `// BEGIN GENERATED <family>` / `// END GENERATED <family>` region to `DocxSessionBridge.cs` where the family's shims are, delete the hand-written shims, and run the generator.
4. In the stdio dispatcher, replace the family's cases with `_ when IsGenerated<Family>Op(op) => DispatchGenerated<Family>Op(op, args)`. In the MCP dispatcher, have the tool shell call `RunGenerated<Family>Action`, and route the family's batch-step validation cases to `ValidateGenerated<Family>Arguments`. In `ToolCatalog`, make the tool's schema a `$$"""` literal and splice in `{{GeneratedMcpSchema.<Tool>Actions}}` and `{{GeneratedMcpSchema.<Tool>Properties}}`.
5. Delete the hand-written parsing the generated code replaced. Run the generator again with `--check`, the drift tests and the transport tests.

If a family's ops are spread over an MCP tool whose other actions are hand-written, splice the generated action members and properties next to the hand-written ones; `OPG005` catches a property both sides advertise. A stdio op or MCP tool shared by two generated families is not supported yet (each family's dispatch method throws on names it does not know), and the generator refuses it.

### What stays hand-written

These open questions from #982 are answered as follows:

- **Ergonomics stay hand-written.** That covers Python docstrings, TypeScript overloads and convenience methods that compose several ops. The generated layer is the thin wrapper; hand-written code calls it.
- **Grouped-intent MCP tools stay.** The MCP surface is one tool per intent with an `action` argument, while the stdio host uses one op per name. The description maps both; it does not flatten one into the other.
- **Composite and stateful ops stay hand-written.** That means sessions, transactions, previews and history. They are not one facade call with flat arguments, and are out of scope for generation.

## Changing a described op

Change the facade first, as always. Then update the family's description in the same change. If the change adds a transport-specific spelling or default, record it under `argNames` or `defaults`, and say why in the pull request.
