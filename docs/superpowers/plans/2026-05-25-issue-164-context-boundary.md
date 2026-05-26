# Issue #164 — Boundary-Aware ContextChars for Grep / FindPlaceholders

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make `TextMatch.ContextBefore` / `ContextAfter` reliably disambiguate adjacent placeholders by (a) widening the default char window from 40 → 80, and (b) adding a `ContextBoundary` enum that snaps the context window to the nearest natural boundary (`Bracket` / `Sentence` / `Comma` / `Char`). Both knobs propagate through `Grep`, `GrepCrossBlock`, and `FindPlaceholders`.

**Architecture:** A new `ContextBoundary` enum plus a private `WalkContext` helper that walks outward from the match by character, stopping at either the `contextChars` cap or the nearest boundary character per the enum. Default mode `Char` matches existing behavior so the only observable change without an explicit boundary opt-in is the wider window. The enum threads through the shared `Docxodus.Internal/DocxSessionOps` core (introduced by PR #168) so the WASM bridge and the Python stdio NDJSON host get it together.

**Tech Stack:** .NET 8, C#, Open XML SDK 3.x, xUnit. Shared `Docxodus.Internal/` core. WASM `[JSExport]` bridge + Python stdio NDJSON dispatcher. npm TypeScript wrapper. Playwright for browser tests.

**Branch:** `feat/164-context-boundary`

---

## Task 1: Create feature branch

**Files:** none (git state only)

- [ ] **Step 1.1: Confirm clean tree on main**

```bash
cd /home/jman/Code/Docxodus
git status
git log --oneline -3
```

Expected: working tree clean (untracked `.claude/` is fine), on `main`, most recent commit is the PR #172 merge or later.

- [ ] **Step 1.2: Create the feature branch**

```bash
git checkout -b feat/164-context-boundary
```

- [ ] **Step 1.3: Commit the plan**

```bash
git add docs/superpowers/plans/2026-05-25-issue-164-context-boundary.md
git commit -m "docs: plan for issue #164 (boundary-aware ContextChars)"
```

---

## Task 2: `ContextBoundary` enum + Grep boundary support (failing tests)

This unit lands the entire .NET-side `Grep` change at once: the new `ContextBoundary` enum, the shared `WalkContext` helper, the updated `Grep` signature, and the widened default. We test the two most-distinct boundary modes (`Char` baseline, `Bracket` — the dominant template-fill use case) here; `Sentence` and `Comma` modes get their own tests in Task 4 but reuse the same helper.

**Files:**
- Test: `Docxodus.Tests/DocxSessionTests.cs` — append `DS250`, `DS251`, `DS252`.

- [ ] **Step 2.1: Add a test-fixture helper for adjacent-placeholder paragraphs**

In `Docxodus.Tests/DocxSessionTests.cs`, near the other `Build*` helpers (search for `BuildDocWithBracketPlaceholders` from #163 to find the neighborhood), add:

```csharp
internal static byte[] BuildDocWithAdjacentPlaceholders()
{
    // One paragraph, three blanks back-to-back like the NVCA address line —
    // the canonical case where 40-char context windows pull in the next blank's
    // text and confuse keyword-based picker logic.
    using var ms = new MemoryStream();
    using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
    {
        var main = doc.AddMainDocumentPart();
        main.Document = new Document(new Body(
            new Paragraph(new Run(new Text(
                "The address is [STREET], in the City of [CITY], County of [COUNTY], in the State of Delaware.")))));
    }
    return ms.ToArray();
}
```

- [ ] **Step 2.2: Add three failing tests**

```csharp
[Fact]
public void DS250_Grep_DefaultContextCharsIsEighty()
{
    // After issue #164, the default ContextChars on Grep is 80 (was 40). Test
    // by grepping a paragraph longer than 40 chars surrounding the match —
    // ContextBefore/After should reach further than they did before.
    using var session = new DocxSession(BuildDocWithAdjacentPlaceholders());
    var matches = session.Grep(@"\[STREET\]");
    var m = Assert.Single(matches);

    // "The address is " is 15 chars — well within 40. Pre-#164 was 15 chars max.
    // Post-#164 default 80 → still 15 chars (the paragraph just isn't that long).
    // The interesting half is After: "in the City of [CITY], County of [COUNTY], in the State of Delaware."
    // is 68 chars. Pre-#164: capped at 40. Post-#164: full 68 visible.
    Assert.True(m.ContextAfter.Length > 40,
        $"ContextAfter should exceed 40 chars under widened default; got {m.ContextAfter.Length}: '{m.ContextAfter}'");
}

[Fact]
public void DS251_Grep_BoundaryCharMatchesPreviousBehavior()
{
    // Char boundary mode is the new default and should reproduce the legacy
    // "just truncate at contextChars" behavior bit-for-bit when contextChars
    // is set to the legacy 40.
    using var session = new DocxSession(BuildDocWithAdjacentPlaceholders());
    var matches = session.Grep(@"\[CITY\]",
        scope: ProjectionScopes.Body, contextChars: 40, boundary: ContextBoundary.Char);
    var m = Assert.Single(matches);

    // Char mode at contextChars=40 must not look past 40 chars in either direction.
    Assert.True(m.ContextBefore.Length <= 40);
    Assert.True(m.ContextAfter.Length <= 40);
}

[Fact]
public void DS252_Grep_BoundaryBracketStopsAtNearestBracket()
{
    // Bracket boundary mode is the dominant template-fill use case: each
    // placeholder's context window must not include any OTHER placeholder's
    // brackets — so a keyword check like ContextAfter.Contains("County of")
    // applied to [CITY]'s match does not accidentally trip on text that
    // belongs to the next blank.
    using var session = new DocxSession(BuildDocWithAdjacentPlaceholders());
    var matches = session.Grep(@"\[CITY\]",
        scope: ProjectionScopes.Body, contextChars: 80, boundary: ContextBoundary.Bracket);
    var m = Assert.Single(matches);

    // Neither ContextBefore nor ContextAfter may contain "[" or "]"
    // (those characters mark the boundary the walker must stop at).
    Assert.DoesNotContain('[', m.ContextBefore);
    Assert.DoesNotContain(']', m.ContextBefore);
    Assert.DoesNotContain('[', m.ContextAfter);
    Assert.DoesNotContain(']', m.ContextAfter);

    // And the actual content should be the text between the previous bracket
    // and this match (Before), and between this match and the next bracket (After).
    Assert.Equal(", in the City of ", m.ContextBefore);
    Assert.Equal(", County of ", m.ContextAfter);
}
```

- [ ] **Step 2.3: Run, verify all three fail**

```bash
cd /home/jman/Code/Docxodus
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS250|FullyQualifiedName~DS251|FullyQualifiedName~DS252" 2>&1 | tail -15
```

Expected: build error `'ContextBoundary' could not be found`.

---

## Task 3: `ContextBoundary` enum + Grep boundary support (implementation)

**Files:**
- Modify: `Docxodus/DocxSession.cs` — add `ContextBoundary` enum, `WalkContext` helper, update `Grep` and `GrepCrossBlock`.

- [ ] **Step 3.1: Add the `ContextBoundary` enum**

In `Docxodus/DocxSession.cs`, add the enum near the other public types like `WhitespaceMode` (search the file for `public enum WhitespaceMode` to find the right neighborhood):

```csharp
/// <summary>
/// Controls where <see cref="DocxSession.Grep"/> stops walking outward when
/// computing <see cref="TextMatch.ContextBefore"/> / <see cref="TextMatch.ContextAfter"/>.
/// The default <see cref="Char"/> just truncates at <c>contextChars</c>; the other
/// modes additionally stop at a natural-language boundary so the returned context
/// is unambiguously *this* match's surroundings, not text that belongs to an
/// adjacent placeholder or sibling sentence.
/// </summary>
public enum ContextBoundary
{
    /// <summary>No natural boundary; truncate at <c>contextChars</c> chars in each direction.
    /// Matches legacy behavior. This is the default.</summary>
    Char = 0,

    /// <summary>Stop at the nearest <c>'['</c> or <c>']'</c>. The dominant
    /// template-fill case: each placeholder's context is unambiguously its own,
    /// even when multiple placeholders crowd into one sentence.</summary>
    Bracket = 1,

    /// <summary>Stop at the nearest sentence-terminator (<c>. ! ? : ;</c>). Useful
    /// for callers building LLM prompts that want a self-contained snippet per match.</summary>
    Sentence = 2,

    /// <summary>Stop at the nearest comma. Useful for matches inside enumerations
    /// (<c>"X, Y, Z"</c>) where adjacent items are unambiguous siblings.</summary>
    Comma = 3,
}
```

- [ ] **Step 3.2: Add a private `WalkContext` helper**

In `Docxodus/DocxSession.cs`, add a private helper near the existing private statics (search for `private static string ExtractFormatting` or `private static bool ScopeMatches` to find the right neighborhood):

```csharp
/// <summary>
/// Walks outward from a match span by character, stopping at either the
/// <c>contextChars</c> cap or the nearest character that qualifies as a
/// boundary under <paramref name="boundary"/>. Returns the <c>(before, after)</c>
/// text slices. Used by both <see cref="Grep"/> and <see cref="GrepCrossBlock"/>.
/// </summary>
private static (string Before, string After) WalkContext(
    string text, int matchStart, int matchLength, int contextChars, ContextBoundary boundary)
{
    int matchEnd = matchStart + matchLength;

    int leftCap = Math.Max(0, matchStart - contextChars);
    int leftStop = matchStart;
    while (leftStop > leftCap)
    {
        if (IsBoundary(text[leftStop - 1], boundary)) break;
        leftStop--;
    }

    int rightCap = Math.Min(text.Length, matchEnd + contextChars);
    int rightStop = matchEnd;
    while (rightStop < rightCap)
    {
        if (IsBoundary(text[rightStop], boundary)) break;
        rightStop++;
    }

    return (text.Substring(leftStop, matchStart - leftStop),
            text.Substring(matchEnd, rightStop - matchEnd));
}

private static bool IsBoundary(char c, ContextBoundary mode) => mode switch
{
    ContextBoundary.Char => false,
    ContextBoundary.Bracket => c is '[' or ']',
    ContextBoundary.Sentence => c is '.' or '!' or '?' or ':' or ';',
    ContextBoundary.Comma => c is ',',
    _ => false,
};
```

- [ ] **Step 3.3: Update `Grep` signature and context computation**

In `Docxodus/DocxSession.cs::Grep` (around line 421), change the signature to add `boundary` and update the default `contextChars`:

```csharp
public IReadOnlyList<TextMatch> Grep(
    string pattern,
    System.Text.RegularExpressions.RegexOptions options = System.Text.RegularExpressions.RegexOptions.None,
    ProjectionScopes scope = ProjectionScopes.Body,
    int contextChars = 80,                               // CHANGED: was 40
    WhitespaceMode whitespace = WhitespaceMode.Preserve,
    ContextBoundary boundary = ContextBoundary.Char)     // NEW
{
```

Then replace the existing context-computation block (lines 494-497, the `ctxStart`/`ctxBefore`/`ctxEnd`/`ctxAfter` calculation) with:

```csharp
var (ctxBefore, ctxAfter) = WalkContext(map.FlatText, m.Index, m.Length, contextChars, boundary);
```

- [ ] **Step 3.4: Update `GrepCrossBlock` similarly**

In the same file at around line 544, update the `GrepCrossBlock` signature:

```csharp
public IReadOnlyList<CrossBlockMatch> GrepCrossBlock(
    string pattern,
    System.Text.RegularExpressions.RegexOptions options = System.Text.RegularExpressions.RegexOptions.None,
    ProjectionScopes scope = ProjectionScopes.Body,
    int contextChars = 80,                               // CHANGED: was 40
    WhitespaceMode whitespace = WhitespaceMode.Preserve,
    ContextBoundary boundary = ContextBoundary.Char)     // NEW
{
```

And replace the context-computation block (lines 655-657) with:

```csharp
var (ctxBefore, ctxAfter) = WalkContext(concat, m.Index, m.Length, contextChars, boundary);
```

- [ ] **Step 3.5: Run the three new tests, verify they pass**

```bash
cd /home/jman/Code/Docxodus
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS250|FullyQualifiedName~DS251|FullyQualifiedName~DS252" 2>&1 | tail -10
```

Expected: 3/3 pass.

- [ ] **Step 3.6: Run the full DocxSession + WmlToMarkdownConverter test suites for regressions**

```bash
cd /home/jman/Code/Docxodus
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DocxSessionTests|FullyQualifiedName~WmlToMarkdownConverterTests" 2>&1 | tail -5
```

Expected: all green. If a test failed because it asserted `ContextBefore.Length` was exactly 40 (now 80) or asserted exact context strings that are now longer, the assertion was over-tight — update it to use `Contains` / `EndsWith` / `StartsWith` instead of `==`. Do NOT roll back the default widening unless multiple tests break in ways that suggest the change was wrong (a single test fixture relying on the old length is just a test-side assumption to fix).

- [ ] **Step 3.7: Commit**

```bash
git add Docxodus/DocxSession.cs Docxodus.Tests/DocxSessionTests.cs
git commit -m "feat(session): add ContextBoundary + widen default contextChars to 80 (#164)"
```

---

## Task 4: Sentence + Comma modes (failing tests)

**Files:**
- Test: `Docxodus.Tests/DocxSessionTests.cs` — append `DS253`, `DS254`.

- [ ] **Step 4.1: Add the two tests**

```csharp
[Fact]
public void DS253_Grep_BoundarySentenceStopsAtTerminator()
{
    // Sentence mode stops at . ! ? : ;
    // Multi-sentence paragraph; assert ContextBefore stops at the previous '.'.
    using var session = new DocxSession(BuildDocWithSentences());
    var matches = session.Grep(@"\[FILL\]",
        scope: ProjectionScopes.Body, contextChars: 200, boundary: ContextBoundary.Sentence);
    var m = Assert.Single(matches);

    Assert.DoesNotContain('.', m.ContextBefore);
    Assert.DoesNotContain('.', m.ContextAfter);
    Assert.DoesNotContain(';', m.ContextBefore);
    Assert.DoesNotContain(';', m.ContextAfter);
}

[Fact]
public void DS254_Grep_BoundaryCommaStopsAtComma()
{
    // Comma mode for matches inside enumerations.
    using var session = new DocxSession(BuildDocWithCommaSeparatedItems());
    var matches = session.Grep(@"\[B\]",
        scope: ProjectionScopes.Body, contextChars: 200, boundary: ContextBoundary.Comma);
    var m = Assert.Single(matches);

    Assert.DoesNotContain(',', m.ContextBefore);
    Assert.DoesNotContain(',', m.ContextAfter);
}
```

- [ ] **Step 4.2: Add the two test-fixture helpers**

Next to the other `Build*` helpers:

```csharp
internal static byte[] BuildDocWithSentences()
{
    using var ms = new MemoryStream();
    using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
    {
        var main = doc.AddMainDocumentPart();
        main.Document = new Document(new Body(
            new Paragraph(new Run(new Text(
                "This is the first sentence. Here is some text with a [FILL] placeholder. And a third sentence after.")))));
    }
    return ms.ToArray();
}

internal static byte[] BuildDocWithCommaSeparatedItems()
{
    using var ms = new MemoryStream();
    using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
    {
        var main = doc.AddMainDocumentPart();
        main.Document = new Document(new Body(
            new Paragraph(new Run(new Text(
                "Item one is [A], item two is [B], and item three is [C].")))));
    }
    return ms.ToArray();
}
```

- [ ] **Step 4.3: Run, verify both pass (the `WalkContext` helper from Task 3 already supports these modes)**

```bash
cd /home/jman/Code/Docxodus
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS253|FullyQualifiedName~DS254" 2>&1 | tail -10
```

Expected: 2/2 pass. These tests don't need additional implementation work — the Sentence/Comma branches in `IsBoundary` from Step 3.2 already handle them. The reason they're in their own task is to keep the commit boundary clean ("add boundary modes" separate from "exercise all modes").

- [ ] **Step 4.4: Commit**

```bash
git add Docxodus.Tests/DocxSessionTests.cs
git commit -m "test(session): exercise Sentence and Comma context boundaries (#164)"
```

---

## Task 5: `FindPlaceholders` accepts `contextChars` + `boundary` (failing test + impl)

`FindPlaceholders` currently calls `Grep` internally with the default `contextChars`. After #164, callers should be able to specify both `contextChars` and `boundary` directly on `FindPlaceholders` rather than dropping down to `Grep`.

**Files:**
- Modify: `Docxodus/DocxSession.cs::FindPlaceholders`.
- Test: `Docxodus.Tests/DocxSessionTests.cs` — append `DS255`.

- [ ] **Step 5.1: Add the failing test**

```csharp
[Fact]
public void DS255_FindPlaceholders_AcceptsBoundaryAndContextChars()
{
    using var session = new DocxSession(BuildDocWithAdjacentPlaceholders());

    // FindPlaceholders should plumb the boundary parameter through to the
    // internal Grep call so picker code can rely on bracket-bounded context
    // without dropping down to Grep manually.
    var placeholders = session.FindPlaceholders(
        PlaceholderKinds.All,
        ProjectionScopes.Body,
        contextChars: 80,
        boundary: ContextBoundary.Bracket);

    // We have three placeholders ([STREET], [CITY], [COUNTY]); none of their
    // contexts should contain bracket characters from sibling placeholders.
    Assert.Equal(3, placeholders.Count);
    foreach (var p in placeholders)
    {
        Assert.DoesNotContain('[', p.Match.ContextBefore);
        Assert.DoesNotContain(']', p.Match.ContextBefore);
        Assert.DoesNotContain('[', p.Match.ContextAfter);
        Assert.DoesNotContain(']', p.Match.ContextAfter);
    }
}
```

- [ ] **Step 5.2: Run, verify it fails**

```bash
cd /home/jman/Code/Docxodus
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS255" 2>&1 | tail -10
```

Expected: build error `'FindPlaceholders' has no overload that takes 4 arguments`.

- [ ] **Step 5.3: Add the parameters to `FindPlaceholders`**

In `Docxodus/DocxSession.cs::FindPlaceholders` (around line 1082), add two parameters and pass them through to the internal `Grep` call:

```csharp
public IReadOnlyList<TemplatePlaceholder> FindPlaceholders(
    PlaceholderKinds kinds = PlaceholderKinds.All,
    ProjectionScopes scope = ProjectionScopes.Body,
    int contextChars = 80,                             // NEW
    ContextBoundary boundary = ContextBoundary.Char)   // NEW
{
    ThrowIfDisposed();
    if (kinds == 0) return Array.Empty<TemplatePlaceholder>();

    // Single bracket-or-dollar-bracket scan; classify by content after the match.
    var matches = Grep(@"\$?\[[^\[\]]+\]",
        System.Text.RegularExpressions.RegexOptions.None, scope,
        contextChars, WhitespaceMode.Preserve, boundary);   // CHANGED: pass new params
    // ...rest unchanged
```

- [ ] **Step 5.4: Update `FillPlaceholders` to thread the same options into `FindPlaceholders`**

`FillPlaceholders` (from #163) calls `FindPlaceholders` internally and may benefit from the same boundary semantics. Update `FillOptions` to include the new params with the same defaults:

```csharp
public sealed record FillOptions
{
    public PlaceholderKinds Kinds { get; init; } = PlaceholderKinds.BlankFill | PlaceholderKinds.Instruction;
    public ProjectionScopes Scope { get; init; } = ProjectionScopes.Body;
    public int MaxPasses { get; init; } = 8;
    public bool PreserveDollarPrefix { get; init; } = true;

    /// <summary>Threaded through to <see cref="DocxSession.FindPlaceholders"/> calls
    /// inside the multi-pass loop. Default 80 (matches the new Grep default).</summary>
    public int ContextChars { get; init; } = 80;                                    // NEW

    /// <summary>Boundary mode for the per-match context windows the picker sees.
    /// Default <see cref="ContextBoundary.Char"/> (legacy truncate-at-contextChars).
    /// Pickers that rely on bracket-bounded context can opt into
    /// <see cref="ContextBoundary.Bracket"/> for unambiguous per-placeholder context.</summary>
    public ContextBoundary Boundary { get; init; } = ContextBoundary.Char;          // NEW
}
```

And update the `FillPlaceholders` body's call to `FindPlaceholders`:

```csharp
var placeholders = FindPlaceholders(opts.Kinds, opts.Scope, opts.ContextChars, opts.Boundary)
    .OrderByDescending(p => p.Match.EnclosingAnchor.Anchor.Id, StringComparer.Ordinal)
    .ThenByDescending(p => p.Match.Span.Start)
    .ToList();
```

- [ ] **Step 5.5: Run the new test and a regression of FindPlaceholders/FillPlaceholders tests**

```bash
cd /home/jman/Code/Docxodus
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS255|FullyQualifiedName~DS120|FullyQualifiedName~DS121|FullyQualifiedName~DS122|FullyQualifiedName~DS123|FullyQualifiedName~DS124|FullyQualifiedName~DS125|FullyQualifiedName~DS126|FullyQualifiedName~DS232|FullyQualifiedName~DS233|FullyQualifiedName~DS240|FullyQualifiedName~DS241|FullyQualifiedName~DS242|FullyQualifiedName~DS243|FullyQualifiedName~DS244|FullyQualifiedName~DS245|FullyQualifiedName~DS246" 2>&1 | tail -5
```

Expected: all green.

- [ ] **Step 5.6: Commit**

```bash
git add Docxodus/DocxSession.cs Docxodus.Tests/DocxSessionTests.cs
git commit -m "feat(session): FindPlaceholders + FillOptions accept contextChars/boundary (#164)"
```

---

## Task 6: Shared-core + WASM bridge + Python stdio plumbing

**Files:**
- Modify: `Docxodus/Internal/DocxSessionOps.cs` — add `ContextBoundary` to `Grep`, `GrepCrossBlock`, `FindPlaceholders` ops.
- Modify: `wasm/DocxodusWasm/DocxSessionBridge.cs` — `ParseGrepOptions` parses `boundary` field; bridge methods pass through.
- Modify: `tools/python-host/Dispatcher.cs` — extract boundary from args for relevant ops.

- [ ] **Step 6.1: Update `DocxSessionOps`**

In `Docxodus/Internal/DocxSessionOps.cs`, find the `Grep` and `GrepCrossBlock` methods (around lines 30-38) and update their signatures to accept a `ContextBoundary` parameter. Also update `FindPlaceholders`:

```csharp
public static string Grep(int handle, string pattern, RegexOptions regexOpts,
    ProjectionScopes scope, int contextChars, WhitespaceMode whitespace, ContextBoundary boundary) =>
    DocxSessionJson.SerializeMatches(
        SessionRegistry.Get(handle).Grep(pattern, regexOpts, scope, contextChars, whitespace, boundary));

public static string GrepCrossBlock(int handle, string pattern, RegexOptions regexOpts,
    ProjectionScopes scope, int contextChars, WhitespaceMode whitespace, ContextBoundary boundary) =>
    DocxSessionJson.SerializeCrossBlockMatches(
        SessionRegistry.Get(handle).GrepCrossBlock(pattern, regexOpts, scope, contextChars, whitespace, boundary));

public static string FindPlaceholders(int handle, PlaceholderKinds kinds, ProjectionScopes scope,
    int contextChars, ContextBoundary boundary) =>
    DocxSessionJson.SerializePlaceholders(
        SessionRegistry.Get(handle).FindPlaceholders(kinds, scope, contextChars, boundary));
```

- [ ] **Step 6.2: Update the WASM bridge**

In `wasm/DocxodusWasm/DocxSessionBridge.cs`:

(a) Update `ParseGrepOptions` to parse a new `boundary` field (around line 296):

```csharp
private static void ParseGrepOptions(string optionsJson, out RegexOptions regexOpts,
    out ProjectionScopes scope, out int contextChars, out WhitespaceMode whitespace,
    out ContextBoundary boundary)
{
    regexOpts = RegexOptions.None;
    scope = ProjectionScopes.Body;
    contextChars = 80;                                  // CHANGED: matches new .NET default
    whitespace = WhitespaceMode.Preserve;
    boundary = ContextBoundary.Char;                    // NEW

    if (string.IsNullOrEmpty(optionsJson)) return;
    using var doc = JsonDocument.Parse(optionsJson);
    var root = doc.RootElement;

    if (root.TryGetProperty("regexOptions", out var ro) && ro.ValueKind == JsonValueKind.Number)
        regexOpts = (RegexOptions)ro.GetInt32();
    if (root.TryGetProperty("scope", out var sc) && sc.ValueKind == JsonValueKind.Number)
        scope = (ProjectionScopes)sc.GetInt32();
    if (root.TryGetProperty("contextChars", out var c) && c.ValueKind == JsonValueKind.Number)
        contextChars = c.GetInt32();
    if (root.TryGetProperty("whitespace", out var w) && w.ValueKind == JsonValueKind.Number)
        whitespace = (WhitespaceMode)w.GetInt32();
    if (root.TryGetProperty("boundary", out var b) && b.ValueKind == JsonValueKind.Number)
        boundary = (ContextBoundary)b.GetInt32();
}
```

(b) Update the `Grep` and `GrepCrossBlock` JSExports to pass the new param through (around lines 116, 128):

```csharp
[JSExport]
public static string Grep(int h, string pattern, string optionsJson)
{
    ParseGrepOptions(optionsJson, out var regexOpts, out var scope, out var contextChars, out var whitespace, out var boundary);
    return DocxSessionOps.Grep(h, pattern, regexOpts, scope, contextChars, whitespace, boundary);
}

[JSExport]
public static string GrepCrossBlock(int h, string pattern, string optionsJson)
{
    ParseGrepOptions(optionsJson, out var regexOpts, out var scope, out var contextChars, out var whitespace, out var boundary);
    return DocxSessionOps.GrepCrossBlock(h, pattern, regexOpts, scope, contextChars, whitespace, boundary);
}
```

(c) Update the `FindPlaceholders` JSExport to accept and pass `contextChars` + `boundary` (around line 170). The current signature is `FindPlaceholders(int h, int kinds, int scope)` — extend it. Since this is a public JSExport, an additive change is fine but the npm side will need to pass the new args. Make the new args optional from the wire side by accepting an options JSON instead:

```csharp
[JSExport]
public static string FindPlaceholders(int h, int kinds, int scope, int contextChars, int boundary) =>
    DocxSessionOps.FindPlaceholders(h, (PlaceholderKinds)kinds, (ProjectionScopes)scope, contextChars, (ContextBoundary)boundary);
```

This changes the wire arity — npm callers will need to pass the two new positional ints. Defaults at the npm side (`contextChars = 80, boundary = 0`) will be applied in the TS wrapper (Task 8 handles that).

- [ ] **Step 6.3: Update the Python stdio dispatcher**

In `tools/python-host/Dispatcher.cs`, find the `grep`, `grep_cross_block`, and `find_placeholders` ops (search for `"grep"` and `"find_placeholders"`). Add `contextChars` and `boundary` extraction:

```csharp
"grep" => DocxSessionOps.Grep(
    Handle(args), Str(args, "pattern"),
    (RegexOptions)IntOr(args, "regexOptions", 0),
    (ProjectionScopes)IntOr(args, "scope", 1),
    IntOr(args, "contextChars", 80),
    (WhitespaceMode)IntOr(args, "whitespace", 0),
    (ContextBoundary)IntOr(args, "boundary", 0)),    // NEW

"grep_cross_block" => DocxSessionOps.GrepCrossBlock(
    Handle(args), Str(args, "pattern"),
    (RegexOptions)IntOr(args, "regexOptions", 0),
    (ProjectionScopes)IntOr(args, "scope", 1),
    IntOr(args, "contextChars", 80),
    (WhitespaceMode)IntOr(args, "whitespace", 0),
    (ContextBoundary)IntOr(args, "boundary", 0)),    // NEW

"find_placeholders" => DocxSessionOps.FindPlaceholders(
    Handle(args),
    (PlaceholderKinds)IntOr(args, "kinds", 7),
    (ProjectionScopes)IntOr(args, "scope", 1),
    IntOr(args, "contextChars", 80),                 // NEW
    (ContextBoundary)IntOr(args, "boundary", 0)),    // NEW
```

If `IntOr(args, "name", default)` doesn't exist in the dispatcher's helpers, look at how the existing ops extract optional ints. The existing `grep` op likely already uses a pattern that handles missing args; mirror it. If the existing entries hardcode `IntOr` differently, match that style.

- [ ] **Step 6.4: Build the WASM target + verify**

```bash
cd /home/jman/Code/Docxodus
./scripts/build-wasm.sh 2>&1 | tail -5
```

Expected: clean build.

- [ ] **Step 6.5: Commit**

```bash
git add Docxodus/Internal/DocxSessionOps.cs wasm/DocxodusWasm/DocxSessionBridge.cs tools/python-host/Dispatcher.cs
git commit -m "feat(bridge): wire ContextBoundary through ops, WASM bridge, Python stdio (#164)"
```

---

## Task 7: npm wrapper — `GrepOptions.boundary`, `ContextBoundary` enum

**Files:**
- Modify: `npm/src/types.ts` — extend `GrepOptions`; add `ContextBoundary` const enum.
- Modify: `npm/src/session.ts` — `findPlaceholders` accepts new options; pass through.
- Modify: `npm/src/index.ts` — re-export `ContextBoundary`.

- [ ] **Step 7.1: Add `ContextBoundary` const enum in `types.ts`**

In `npm/src/types.ts`, near `PlaceholderKinds` (around line 902):

```typescript
/**
 * Numeric flag layout matching the .NET `ContextBoundary` enum. Controls
 * how `Grep` / `GrepCrossBlock` / `FindPlaceholders` decide where to stop
 * walking outward when computing `TextMatch.contextBefore` / `contextAfter`.
 *
 * - `Char` (default) — truncate at `contextChars`. Matches legacy behavior.
 * - `Bracket` — stop at `[` or `]`. Use for template fills: each placeholder's
 *   context is unambiguously its own even when multiple placeholders crowd
 *   into one sentence.
 * - `Sentence` — stop at `.`, `!`, `?`, `:`, `;`.
 * - `Comma` — stop at `,`. For matches inside enumerations.
 */
export const ContextBoundary = {
  Char: 0,
  Bracket: 1,
  Sentence: 2,
  Comma: 3,
} as const;
```

- [ ] **Step 7.2: Extend `GrepOptions`**

In `npm/src/types.ts` (around line 927), extend the interface:

```typescript
export interface GrepOptions {
  regexOptions?: number;
  scope?: number;
  contextChars?: number;
  whitespace?: number;

  /**
   * Boundary mode for context computation. Use `ContextBoundary.Bracket` for
   * template-fill picker code that needs unambiguous per-placeholder context;
   * default `ContextBoundary.Char` (0) matches legacy truncate-at-contextChars
   * behavior.
   */
  boundary?: number;     // NEW
}
```

- [ ] **Step 7.3: Update the WASM bridge declaration**

In `npm/src/types.ts`, find the `DocxodusWasmExports.DocxSessionBridge.FindPlaceholders` declaration (search for `FindPlaceholders`). The new bridge signature takes 5 ints — update the TS declaration to match:

```typescript
FindPlaceholders: (handle: number, kinds: number, scope: number, contextChars: number, boundary: number) => string;
```

- [ ] **Step 7.4: Update `findPlaceholders` wrapper to accept options**

In `npm/src/session.ts`, the current `findPlaceholders` signature is:

```typescript
findPlaceholders(kinds: number = PlaceholderKinds.All, scope: number = 1): TemplatePlaceholder[]
```

Extend it to accept `contextChars` and `boundary`:

```typescript
findPlaceholders(
  kinds: number = PlaceholderKinds.All,
  scope: number = 1,
  contextChars: number = 80,                          // NEW (matches new .NET default)
  boundary: number = ContextBoundary.Char,            // NEW
): TemplatePlaceholder[] {
  return JSON.parse(
    this.wasm.FindPlaceholders(this.handle, kinds, scope, contextChars, boundary)
  ) as TemplatePlaceholder[];
}
```

Add `ContextBoundary` to the type imports at the top of `session.ts`.

- [ ] **Step 7.5: Update `FillOptions` TS type**

In `npm/src/types.ts`, find `FillOptions` (added in #163, around line 930) and extend it:

```typescript
export interface FillOptions {
  kinds?: number;
  scope?: number;
  maxPasses?: number;
  preserveDollarPrefix?: boolean;
  contextChars?: number;       // NEW
  boundary?: number;           // NEW
}
```

- [ ] **Step 7.6: Update `fillPlaceholders` TS impl to pass new options through**

In `npm/src/session.ts`, find the `fillPlaceholders` method (added in #163). It calls `this.findPlaceholders(kinds, scope)` — update that call to pass the two new options:

```typescript
const contextChars = opts.contextChars ?? 80;
const boundary = opts.boundary ?? ContextBoundary.Char;
// ...
for (let pass = 1; pass <= maxPasses; pass++) {
  const placeholders = this.findPlaceholders(kinds, scope, contextChars, boundary)
    .sort((a, b) => { /* unchanged */ });
  // ...
}
```

- [ ] **Step 7.7: Re-export `ContextBoundary` from `index.ts`**

In `npm/src/index.ts`, find the existing re-exports (search for `PlaceholderKinds` to find the right block) and add `ContextBoundary`:

```typescript
export { PlaceholderKinds, ContextBoundary } from "./types.js";
```

- [ ] **Step 7.8: Build and type-check**

```bash
cd /home/jman/Code/Docxodus/npm
npm run build 2>&1 | tail -5
npx tsc --noEmit 2>&1 | tail -10
```

Expected: both clean.

- [ ] **Step 7.9: Commit**

```bash
cd /home/jman/Code/Docxodus
git add npm/src/
git commit -m "feat(npm): expose ContextBoundary + boundary option on Grep/FindPlaceholders/FillOptions (#164)"
```

---

## Task 8: Playwright integration test

**Files:**
- Create: `npm/tests/context-boundary.spec.ts`

- [ ] **Step 8.1: Write the spec**

Inspect existing specs for harness pattern:

```bash
ls /home/jman/Code/Docxodus/npm/tests/ | head -10
```

The best model is `fill-placeholders.spec.ts` (from #163) — it uses inline base64 fixtures and the raw `DocxSessionBridge`. Mirror that approach.

Create `npm/tests/context-boundary.spec.ts` with these assertion areas:

1. **Default ContextChars is now 80.** Grep a placeholder in a paragraph whose surrounding text is > 40 chars; assert `contextBefore.length > 40` or `contextAfter.length > 40` (using a fixture where this is the case).
2. **`boundary: Bracket` stops at `[`/`]`.** Use the adjacent-placeholders fixture (`[STREET], in the City of [CITY], County of [COUNTY]`); grep `\[CITY\]`; assert `contextBefore` doesn't contain `[` or `]`.
3. **`boundary: Sentence` stops at sentence terminators.** Use a multi-sentence fixture; grep a placeholder; assert `contextBefore` doesn't contain `.`.
4. **`boundary: Comma` stops at commas.** Use the comma-separated fixture; assert `contextBefore` doesn't contain `,`.
5. **`findPlaceholders` honors `boundary` parameter end-to-end.** Use the adjacent-placeholders fixture; call `findPlaceholders(PlaceholderKinds.All, 1, 80, ContextBoundary.Bracket)`; assert every placeholder's `match.contextBefore`/`contextAfter` are bracket-free.

For fixture bytes: build them once via a small .NET program (mirroring the `BuildDocWith*` helpers in `DocxSessionTests.cs`), or inline a single shared fixture that satisfies multiple assertions (the `BuildDocWithAdjacentPlaceholders` paragraph covers assertions 1, 2, and 5; the multi-sentence paragraph for 3; the comma list for 4).

- [ ] **Step 8.2: Run the spec**

```bash
cd /home/jman/Code/Docxodus/npm
npm test -- --grep "context-boundary" 2>&1 | tail -15
```

Expected: green.

- [ ] **Step 8.3: Commit**

```bash
cd /home/jman/Code/Docxodus
git add npm/tests/context-boundary.spec.ts
git commit -m "test(npm): ContextBoundary integration — default+Bracket/Sentence/Comma (#164)"
```

---

## Task 9: Documentation

**Files:**
- `CHANGELOG.md`
- `docs/architecture/docx_mutation_api.md`
- `~/Code/Docxodus-Agents/TOOL_SURFACE.md` (out-of-repo)
- `~/Code/Docxodus-Agents/PITFALLS.md` (out-of-repo)

- [ ] **Step 9.1: CHANGELOG entry**

Under `## [Unreleased]` / `### Added`:

```markdown
- **`ContextBoundary` enum + widened default `contextChars`** (issue #164). `DocxSession.Grep`, `GrepCrossBlock`, and `FindPlaceholders` now accept a `ContextBoundary` parameter that controls where the context-computation walker stops: `Char` (default, legacy truncate-at-N behavior), `Bracket` (stop at `[`/`]` — the dominant template-fill case for unambiguous per-placeholder context), `Sentence` (stop at `.!?:;`), `Comma` (stop at `,`). Default `contextChars` widened from 40 → 80 across all three methods so plain `.Contains` checks have enough text to disambiguate without the agent dropping into boundary mode. `FillOptions` gains `ContextChars` + `Boundary` fields threaded into the internal `FindPlaceholders` calls. Shared core (`DocxSessionOps.Grep` / `GrepCrossBlock` / `FindPlaceholders`) propagates the new param so both the WASM bridge and the Python stdio NDJSON host pick it up (npm `GrepOptions.boundary`, exported `ContextBoundary` const). Tests: `DS250`–`DS255`, Playwright `context-boundary.spec.ts`.
```

- [ ] **Step 9.2: `docs/architecture/docx_mutation_api.md`**

Find the `## Grep — cross-run text search` section. After the existing "Performance" or "Known limits" subsection, add:

```markdown
### Context boundary modes

`Grep` / `GrepCrossBlock` / `FindPlaceholders` accept a `ContextBoundary` parameter
that decides where the context-computation walker stops:

| Mode | Stops at | Use when |
|---|---|---|
| `Char` (default) | nothing — truncate at `contextChars` | legacy callers, free-form text where boundaries are noisy |
| `Bracket` | `[`, `]` | template fills with adjacent placeholders — each `ContextBefore`/`ContextAfter` is guaranteed to belong to this match only |
| `Sentence` | `.`, `!`, `?`, `:`, `;` | LLM prompt-building where each snippet should be a self-contained sentence |
| `Comma` | `,` | matches inside enumerations |

The default `contextChars` widened from 40 → 80 in #164. Combined with `Bracket`
mode this lets a template-fill picker use plain `.Contains` / `EndsWith` checks
without cross-pollution from adjacent placeholders:

\`\`\`csharp
var matches = session.Grep(@"\[CITY\]",
    scope: ProjectionScopes.Body,
    contextChars: 80,
    boundary: ContextBoundary.Bracket);
// matches[0].ContextBefore guaranteed bracket-free
\`\`\`
```

(Use real triple backticks in the file, not the escaped form above.)

- [ ] **Step 9.3: Update agent guide — `TOOL_SURFACE.md`**

In `~/Code/Docxodus-Agents/TOOL_SURFACE.md`, find the "Discovery — finding what to edit" table. Update the `Grep` and `FindPlaceholders` rows to mention the new boundary option:

```markdown
| Find every literal or regex match across the doc | `session.Grep(pattern, opts?, scope?, contextChars=80, whitespace?, boundary=Char)` | `IReadOnlyList<TextMatch>` — set `boundary=Bracket` for unambiguous per-placeholder context |
```

And under "Anchor shapes you'll see" or in a new tip section near the Grep discussion, add:

```markdown
**Tip — boundary-aware context.** When picker logic uses `.Contains` / `.EndsWith`
on `ContextBefore` / `ContextAfter`, pass `boundary: ContextBoundary.Bracket`
(numeric `1`) to guarantee the context window doesn't bleed into adjacent
placeholders. The default `Char` mode (legacy) truncates at `contextChars` but
makes no boundary guarantee.
```

- [ ] **Step 9.4: Update agent guide — `PITFALLS.md`**

In `~/Code/Docxodus-Agents/PITFALLS.md` §1 (Context-window keyword search is too fuzzy), add a "Fixed in 6.1.0 (#164)" callout at the top:

```markdown
> **Fixed in Docxodus 6.1.0 (#164).** `Grep` / `FindPlaceholders` now accept a
> `boundary: ContextBoundary.Bracket` option that snaps the context window to
> the nearest `[` or `]`, guaranteeing that adjacent placeholders' text can't
> bleed into this match's context. The default `contextChars` also widened
> from 40 → 80. This section is preserved for agents working against
> Docxodus 6.0.x or earlier, or using `boundary: Char` (the legacy default).
```

- [ ] **Step 9.5: Commit**

```bash
cd /home/jman/Code/Docxodus
git add CHANGELOG.md docs/architecture/docx_mutation_api.md
git commit -m "docs: ContextBoundary modes + widened contextChars default (#164)"
```

`~/Code/Docxodus-Agents/` is not a git repo — files saved in place.

---

## Task 10: Final verification + PR

- [ ] **Step 10.1: Force a clean .NET rebuild and run the full test suite**

```bash
cd /home/jman/Code/Docxodus
rm -rf Docxodus/bin Docxodus/obj Docxodus.Tests/bin Docxodus.Tests/obj
dotnet build Docxodus.Tests/Docxodus.Tests.csproj 2>&1 | tail -5
dotnet test Docxodus.Tests/Docxodus.Tests.csproj 2>&1 | tail -10
```

Expected: 0 failed.

- [ ] **Step 10.2: npm build + Playwright**

```bash
cd /home/jman/Code/Docxodus/npm
npm run build 2>&1 | tail -5
npm test 2>&1 | tail -15
```

Expected: all Playwright tests green, including `context-boundary.spec.ts`, `fill-placeholders.spec.ts`, `anchor-introspection.spec.ts`.

- [ ] **Step 10.3: Release build (warnings-as-errors)**

Build individual non-WASM projects (per CLAUDE.md, building the solution in Release leaks WASM-mode DLLs):

```bash
cd /home/jman/Code/Docxodus
rm -rf Docxodus/bin Docxodus/obj Docxodus.Tests/bin Docxodus.Tests/obj
dotnet build -c Release Docxodus/Docxodus.csproj 2>&1 | tail -5
dotnet build -c Release Docxodus.Tests/Docxodus.Tests.csproj 2>&1 | tail -5
```

Expected: clean. (The `tools/python-host/pyhost.csproj` Release-build failure is the pre-existing main-branch issue tracked by #173 — not a regression.)

- [ ] **Step 10.4: Push and open PR**

```bash
git push -u origin feat/164-context-boundary 2>&1 | tail -5

gh pr create --title "feat: boundary-aware ContextChars for Grep / FindPlaceholders" --body "$(cat <<'EOF'
Closes #164.

## Summary
- New `ContextBoundary` enum (`Char`, `Bracket`, `Sentence`, `Comma`) on `Grep` / `GrepCrossBlock` / `FindPlaceholders` controls where the context-computation walker stops.
- Default `contextChars` widened from 40 → 80 across all three methods. Plain `.Contains` checks now have enough text to disambiguate without dropping into boundary mode.
- `FillOptions` gains `ContextChars` + `Boundary` fields threaded into internal `FindPlaceholders` calls.
- Shared `DocxSessionOps` propagates the new param to both the WASM bridge and the Python stdio NDJSON host (`grep`, `grep_cross_block`, `find_placeholders` ops).
- npm wrapper: `GrepOptions.boundary`, `ContextBoundary` const enum, threaded through `fillPlaceholders`.

## Test plan
- [x] `DS250` — default `contextChars` is 80 (wider than legacy 40).
- [x] `DS251` — `Boundary.Char` matches legacy truncate-at-N behavior.
- [x] `DS252` — `Boundary.Bracket` excludes `[`/`]` from context.
- [x] `DS253` — `Boundary.Sentence` excludes sentence terminators.
- [x] `DS254` — `Boundary.Comma` excludes commas.
- [x] `DS255` — `FindPlaceholders` end-to-end accepts and propagates boundary.
- [x] Playwright `context-boundary.spec.ts` — full npm round trip.
- [x] Full `dotnet test` passes.
- [x] `npm run build && npm test` passes.
- [x] Release build clean (modulo pre-existing #173).

## Note for reviewers
- `findPlaceholders` JSExport gains two new positional ints (`contextChars`, `boundary`) — this is a wire-format change that npm wrapper carries via additional default args. Stale npm bundles will break against the new bridge; rebuild required.
- `Grep` / `GrepCrossBlock` JSExports take the new `boundary` via the existing `optionsJson` blob — additive, no wire-format break.
EOF
)" 2>&1 | tail -5
```

Expected: PR URL printed.

---

## Self-Review

**Spec coverage:**
- §1 (widen default to 80): Tasks 3 (Grep), 3 (GrepCrossBlock), 5 (FindPlaceholders), 6 (bridge default), 7 (npm default). ✓
- §2 (`ContextBoundary` enum + boundary-aware walker): Tasks 2-4 (.NET impl + 4 mode tests), 5 (FindPlaceholders propagation), 6 (bridge plumbing), 7 (npm enum + types), 8 (Playwright). ✓
- §3 (same params on `FindPlaceholders`): Task 5. ✓
- Acceptance criteria — "Default `ContextChars` raised to 80; existing tests pass without changes": Step 3.6 (regression run + guidance for over-tight assertions). ✓
- Acceptance criteria — "`ContextBoundary` enum exposed via `GrepOptions` and `FindPlaceholders`": Tasks 7 (npm) + 3/5 (.NET). ✓
- Acceptance criteria — "Three adjacent `[_]` produce three context values that don't contain 'city of' in Bracket mode": Task 2 (DS252) + Task 5 (DS255). ✓
- Acceptance criteria — Docs updated: Task 9. ✓

**Placeholder scan:** No `TODO`, no `Similar to Task N`, no bare "add error handling." Every code block is real C# / TS / shell.

**Type consistency:**
- `ContextBoundary` enum members: `Char=0, Bracket=1, Sentence=2, Comma=3` across .NET and TS. ✓
- `FindPlaceholders` JSExport wire arity (`int h, int kinds, int scope, int contextChars, int boundary`) matches the TS `DocxodusWasmExports.FindPlaceholders` declaration. ✓
- `FillOptions.ContextChars` / `.Boundary` named consistently in C# (PascalCase) and TS (camelCase). ✓

One naming note worth flagging:
- The `Boundary` parameter name on .NET methods is consistent with the enum name; I chose lowercase `boundary` for the TS option name to match other camelCase fields in `GrepOptions` / `FillOptions`. Same value space, different casing per language convention.

---

**Plan complete and saved to `docs/superpowers/plans/2026-05-25-issue-164-context-boundary.md`.**

Two execution options:

1. **Subagent-Driven (recommended)** — same flow as #162/#163 (which both shipped cleanly this way).
2. **Inline Execution** — execute tasks here with checkpoints.
