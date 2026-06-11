# Document IR — M1.3 Registries, List Resolution, Effective Formats, Remaining Scopes

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development. Executed by the controller as sequenced subagent tasks.

**Goal:** Fill the M1.1 placeholders: real style/numbering/theme registries, fully-resolved `IrListInfo`, non-destructive effective-format resolution, and the headers/footers/footnotes/endnotes/comments scopes — completing the IR's coverage of everything the markdown projection consumes, so M1.4's port has no missing substrate.

**Baseline:** `feat/document-ir` @ c8ddc84 — IR suite 100 green, corpus 668/668, M1.2 complete.

**Spec note:** registry type shapes below are new (spec §7.1 only names the types); the spec gets a §7.5 "Registries" section at milestone close reflecting what lands.

## Task 1: Real registries

`IrRegistries.cs` replaces placeholders (keep `Empty` members for tests):

```csharp
internal sealed record IrStyle(string Id, string? Name, string? BasedOn, string Type, bool IsDefault)
{ public XElement? PPr { get; init; } public XElement? RPr { get; init; } }   // cloned, equality-excluded via doc note

internal sealed record IrStyleRegistry(IReadOnlyDictionary<string, IrStyle> Styles,
                                       string? DefaultParagraphStyleId, XElement? DocDefaultsPPr, XElement? DocDefaultsRPr);

internal sealed record IrNumLevel(int Ilvl, string NumberFormat, int? Start, string? LvlText)
{ public XElement? PPr { get; init; } }
internal sealed record IrAbstractNum(int AbstractNumId, IReadOnlyDictionary<int, IrNumLevel> Levels);
internal sealed record IrNum(int NumId, int AbstractNumId, IReadOnlyDictionary<int, int> StartOverrides);
internal sealed record IrNumberingRegistry(IReadOnlyDictionary<int, IrNum> Nums,
                                           IReadOnlyDictionary<int, IrAbstractNum> AbstractNums);

internal sealed record IrThemeFonts(string? MajorAscii, string? MinorAscii);
```

Reader populates them from StyleDefinitionsPart / NumberingDefinitionsPart / ThemePart (all tolerantly: missing part → Empty; malformed entries skipped). `numStyleLink`/`styleLink` indirection: resolve one level if cheap, else record as-is with a TODO. Registries documented reference-equal (consistent with IrDocument's dictionary policy). Tests: registry facts from a programmatic doc with styles + numbering + theme.

## Task 2: Full IrListInfo resolution + projection parity

Reader resolves each paragraph's list facts through the registry: direct `numPr` OR style-chain `numPr` (basedOn walk, cycle-guarded, depth 16 — same discipline as `IsListItem`) → `IrListInfo(numId, abstractNumId, ilvl, numberFormat-from-level, startOverride-from-num, fromStyle)`. numId=0 = "no list" per OOXML (treat as no membership). JSON list object gains numberFormat/startOverride/fromStyle fields. Parity test: for fixtures with lists, compare against `DocxSession.GetListMembership` facts (open the same bytes in a DocxSession; match numId/abstractNumId/ilvl/format/start-override/from-style per anchor). Snapshot regen (list objects gain fields; abstractNumId now resolved — hash-neutral since IrListInfo isn't hashed… VERIFY: IrListInfo is NOT part of ContentHash/FormatFingerprint today (numPr rides in the unmodeled digest) — assert hashes unchanged in the snapshot diff).

## Task 3: Effective format resolution (non-destructive)

New `Docxodus/Ir/IrEffectiveFormats.cs`: `internal sealed class IrEffectiveFormats(IrDocument doc)` exposing `IrParaFormat ResolveParagraph(IrParagraph p)` and `IrRunFormat ResolveRun(IrParagraph p, IrRunFormat direct)`. Cascade: docDefaults → style chain (basedOn, root-first application) → direct; toggle properties follow last-writer-wins at this fidelity tier (document divergences from full OOXML toggle-XOR semantics as a known M1.4+ refinement with TODO). Per-styleId memo cache; thread-safe enough via lock or Lazy (document choice). UnmodeledDigest of effective records = direct record's digest (document: effective resolution only covers modeled fields). Tests: hand-computed expectations over a programmatic style chain (docDefault size 20 → style bold+size 24 → derived style italic basedOn it → paragraph direct override size 28 ⇒ effective bold+italic+28); FormattingAssembler cross-check on one simple fixture if practical, else documented skip.

## Task 4: Remaining scopes — headers/footers/notes/comments (+ N15 record-half)

Reader honors `IrScopes` flags: HeadersFooters → enumerate `main.HeaderParts`/`FooterParts` in the SAME order as the projection (`hdr1`/`ftr1`… naming, `BuildAnchorIndex` is the reference), each walked by the existing block walker into `IrHeaderFooter(scopeName, kind-from-sectPr-references-else-Default, scope)`. Notes → footnotes/endnotes parts into `IrNoteStore` (note id → IrScope of that note's blocks), skipping boilerplate separator/continuation notes exactly like the projection (`IsBoilerplateNote` — make internal and reuse). Comments → `IrCommentStore`: per comment, metadata (author/initials/date) + blocks; N15 record-half: comment range targets recorded as `IrCommentTarget(blockAnchor, startChar, endChar)` — char offsets tracked during paragraph inline accumulation; ranges spanning blocks → one target per touched block (spec open question #2 resolved this way; update spec §12). Scope names fn/en/cmt; anchors in those scopes use the matching scope string. AnchorIndex covers all scopes (collision-checked). Corpus totality re-verified; JSON gains scopes sections (document-level becomes `{"scopes":[…]}` — writer + snapshot regen with review; body-only fixtures byte-stable except the wrapper). Tests: scope naming parity vs projection on a header/footer fixture; note store contents; comment target offsets.

## Exit criteria

- Registries populated; IrListInfo fully resolved with GetListMembership parity test green.
- Effective-format resolver with hand-computed cascade tests.
- All scopes readable; corpus totality 668/668 with `IrScopes.All`; full suite green.
- Spec updated: §7.5 registries, §12 open-question resolutions (comment targets per-block).

## Outcome

All four tasks landed; exit criteria met.

- Task 1 — registries populated (`f43ae32`): IrStyle/Numbering/ThemeFonts resolved tolerantly (missing part → `Empty`, malformed skipped, first-wins duplicates).
- Task 2 — full IrListInfo resolution with `GetListMembership` parity (`7dee824`).
- Task 3 — non-destructive effective-format resolver via the style cascade (`fc3f224`).
- Task 4 — header/footer/note/comment scopes + N15 record-half comment targets (`246a8cc`).

Spec updated: §7.5 Registries documents the landed shapes; §3 file list corrected to reality.

Known deferred items (carried to M1.4+):

- `numStyleLink` indirection — abstractNums carrying only a `w:numStyleLink` resolve to empty `Levels`; link not chased.
- Toggle-property semantics — effective resolution uses last-writer-wins, not full OOXML toggle-XOR.
- Direct `w:asciiTheme` — theme-font references on direct run props are not theme-resolved in the effective record.
- Cross-field comment-range test — comment ranges interacting with field cached results not yet covered by a dedicated test.
- `IrReader` partial-class split — the reader remains a single file (N1–N15 inline); splitting deferred to M1.4.
