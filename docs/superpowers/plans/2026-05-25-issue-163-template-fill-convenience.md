# Issue #163 — Template-Fill Convenience

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Land three related `DocxSession` template-fill ergonomics: a `ReplaceInner` overload that strips brackets while preserving any prefix/suffix outside them, an enriched `TemplatePlaceholder.AlternativeKinds` field that catches long-clause-with-blanks misclassifications, and a `FillPlaceholders(picker, options?)` convenience that bundles reverse-offset ordering, `$`-prefix preservation, and multi-pass nested-bracket iteration into one call.

**Architecture:** `ReplaceInner` is a thin overload on top of the existing `ReplaceMatch` that decomposes a match's text by bracket position. The classifier change is purely additive (new field on the record + extra heuristic that runs when the primary classification is borderline). `FillPlaceholders` is a control loop on top of `FindPlaceholders` + `ReplaceMatch` that iterates until no more placeholders match the requested kinds, stopping when a pass makes zero state changes. All three ship cross-stack: .NET → tests → WASM bridge → npm wrapper → Playwright spec → docs.

**Tech Stack:** .NET 8, C#, Open XML SDK 3.x, xUnit. WASM `[JSExport]` bridge, npm TypeScript wrapper. Playwright for browser tests.

**Branch:** `feat/163-fill-placeholders`

---

## Task 1: Create feature branch

**Files:** none (git state only)

- [ ] **Step 1.1: Confirm clean tree on main**

```bash
cd /home/jman/Code/Docxodus
git status
git log --oneline -3
```

Expected: working tree clean (untracked `.claude/` is fine), on `main`, most recent commit is the #169 merge.

- [ ] **Step 1.2: Create the feature branch**

```bash
git checkout -b feat/163-fill-placeholders
```

Expected: `Switched to a new branch 'feat/163-fill-placeholders'`.

- [ ] **Step 1.3: Commit the plan**

```bash
git add docs/superpowers/plans/2026-05-25-issue-163-template-fill-convenience.md
git commit -m "docs: plan for issue #163 (template-fill convenience)"
```

---

## Task 2: `ReplaceInner` overload (failing test)

`ReplaceInner` strips the `[…]` portion of a match's text and substitutes the new inner content, **preserving any prefix or suffix outside the brackets** in the match. The canonical use case: `FindPlaceholders` returns matches like `$[___]` (the regex `\$?\[…\]` captures the `$`); `ReplaceInner(match, "0.20")` yields `$0.20`, not `0.20`.

**Files:**
- Test: `Docxodus.Tests/DocxSessionTests.cs` — append `DS230_ReplaceInner_StripsBracketsKeepsPrefix` and `DS231_ReplaceInner_NoBracketsReturnsError`.

- [ ] **Step 2.1: Add a test-fixture helper next to the existing `BuildDS*` helpers**

Search for an existing helper:

```bash
grep -nE "internal static byte\[\] BuildDS00" /home/jman/Code/Docxodus/Docxodus.Tests/DocxSessionTests.cs | head -5
```

Add this helper alongside them (use unqualified Open XML SDK types like the surrounding helpers — `BuildDS001_SimpleTwoParagraphs` is a good style reference):

```csharp
internal static byte[] BuildDocWithBracketPlaceholders()
{
    using var ms = new MemoryStream();
    using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
    {
        var main = doc.AddMainDocumentPart();
        main.Document = new Document(new Body(
            new Paragraph(new Run(new Text("The name of this corporation is [_____]."))),
            new Paragraph(new Run(new Text("The price per share is $[___]."))),
            new Paragraph(new Run(new Text("Located at [____________]."))),
            new Paragraph(new Run(new Text("[Notwithstanding the foregoing, additional terms may apply.]"))),
            new Paragraph(new Run(new Text("[outer [inner] clause]")))));
    }
    return ms.ToArray();
}
```

- [ ] **Step 2.2: Add the two failing tests**

```csharp
[Fact]
public void DS230_ReplaceInner_StripsBracketsKeepsPrefix()
{
    using var session = new DocxSession(BuildDocWithBracketPlaceholders());

    // The "$[___]" placeholder includes the leading "$" in the match (regex \$?\[…\]).
    // ReplaceInner must keep the "$" prefix and substitute only the bracketed portion.
    var dollarMatch = session.FindPlaceholders(PlaceholderKinds.BlankFill)
        .Single(p => p.Match.Text.StartsWith("$["));

    var r = session.ReplaceInner(dollarMatch.Match, "0.20");
    Assert.True(r.Success);

    // After replacement, the paragraph should read "The price per share is $0.20."
    var anchorId = dollarMatch.Match.EnclosingAnchor.Anchor.Id;
    var info = session.GetAnchorInfo(anchorId);
    Assert.NotNull(info);
    Assert.Equal("The price per share is $0.20.", info!.TextPreview);
}

[Fact]
public void DS231_ReplaceInner_NoBracketsReturnsError()
{
    using var session = new DocxSession(BuildDocWithBracketPlaceholders());

    // Pick any match, then hand-craft a TextMatch whose Text has no brackets.
    var p = session.FindPlaceholders(PlaceholderKinds.BlankFill).First();
    var bogus = new TextMatch
    {
        Text = "no brackets here",
        EnclosingAnchor = p.Match.EnclosingAnchor,
        Span = p.Match.Span,
        Fragments = p.Match.Fragments,
        ContextBefore = "",
        ContextAfter = "",
        Groups = Array.Empty<string>(),
    };

    var r = session.ReplaceInner(bogus, "anything");
    Assert.False(r.Success);
    Assert.Equal(EditErrorCode.MalformedMarkdown, r.Error?.Code);
}
```

- [ ] **Step 2.3: Run and verify both fail**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS230_ReplaceInner_StripsBracketsKeepsPrefix|FullyQualifiedName~DS231_ReplaceInner_NoBracketsReturnsError" 2>&1 | tail -15
```

Expected: build error `'DocxSession' does not contain a definition for 'ReplaceInner'`.

---

## Task 3: `ReplaceInner` overload (implementation)

**Files:**
- Modify: `Docxodus/DocxSession.cs` — add `ReplaceInner` next to `ReplaceMatch` (around line 985).

- [ ] **Step 3.1: Add `ReplaceInner` directly below `ReplaceMatch`**

Locate `public EditResult ReplaceMatch(TextMatch match, string replace)` in `Docxodus/DocxSession.cs` (around line 985). Add this method immediately after:

```csharp
/// <summary>
/// Replace the bracketed portion of a <see cref="TextMatch"/> with <paramref name="newInner"/>,
/// preserving any prefix or suffix outside the brackets. Designed for
/// <see cref="FindPlaceholders"/> matches like <c>$[___]</c> where the regex
/// <c>\$?\[…\]</c> captures the leading <c>$</c>: <c>ReplaceInner(match, "0.20")</c>
/// yields <c>$0.20</c> (not <c>0.20</c>). For matches without any prefix/suffix,
/// this is equivalent to <see cref="ReplaceMatch"/> with the new inner value.
/// Returns <see cref="EditErrorCode.MalformedMarkdown"/> if the match text does
/// not contain balanced brackets.
/// </summary>
public EditResult ReplaceInner(TextMatch match, string newInner)
{
    if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
    if (match is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "match is null");

    int lb = match.Text.IndexOf('[');
    int rb = match.Text.LastIndexOf(']');
    if (lb < 0 || rb <= lb)
        return EditResult.Fail(EditErrorCode.MalformedMarkdown,
            $"match text has no balanced brackets: '{match.Text}'");

    var prefix = match.Text[..lb];
    var suffix = match.Text[(rb + 1)..];
    return ReplaceMatch(match, prefix + newInner + suffix);
}
```

- [ ] **Step 3.2: Run the two new tests and verify they pass**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS230_ReplaceInner_StripsBracketsKeepsPrefix|FullyQualifiedName~DS231_ReplaceInner_NoBracketsReturnsError" 2>&1 | tail -10
```

Expected: 2/2 pass.

- [ ] **Step 3.3: Run the full DocxSession test class to ensure no regressions**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DocxSessionTests" 2>&1 | tail -5
```

Expected: all green.

- [ ] **Step 3.4: Commit**

```bash
git add Docxodus/DocxSession.cs Docxodus.Tests/DocxSessionTests.cs
git commit -m "feat(session): add ReplaceInner overload that preserves bracket prefix/suffix (#163)"
```

---

## Task 4: Smarter classifier — `AlternativeKinds` field (failing test)

A long bracketed clause that happens to contain `_______` inside (e.g. `[that would result in at least $_______ in gross proceeds]`) is classified `BlankFill` today (the rule is "≥ 2 underscores inside"). The classifier can be tightened by exposing **all** plausible classifications via a new `AlternativeKinds` field on `TemplatePlaceholder`. Primary `Kind` keeps the current heuristic; `AlternativeKinds` is populated when secondary classifiers also match.

**Files:**
- Test: `Docxodus.Tests/DocxSessionTests.cs` — append `DS232_Classifier_LongClauseWithUnderscoresExposesAlternativeKinds` and `DS233_Classifier_SimpleBlankFillEmptyAlternativeKinds`.

- [ ] **Step 4.1: Add failing tests**

```csharp
[Fact]
public void DS232_Classifier_LongClauseWithUnderscoresExposesAlternativeKinds()
{
    // Long bracketed clauses that contain embedded underscores are ambiguous —
    // strictly speaking they're an AlternativeClause (a real clause), but the
    // current classifier rule fires BlankFill because >= 2 underscores are present.
    // Surface BOTH classifications so callers can decide.
    using var session = new DocxSession(BuildDocWithLongClauseBlank());

    var p = session.FindPlaceholders().Single();
    Assert.Equal(PlaceholderKind.BlankFill, p.Kind);              // primary stays BlankFill (back-compat)
    Assert.Contains(PlaceholderKind.AlternativeClause, p.AlternativeKinds);
}

[Fact]
public void DS233_Classifier_SimpleBlankFillEmptyAlternativeKinds()
{
    // Plain "[_____]" placeholders are unambiguous — AlternativeKinds should be empty.
    using var session = new DocxSession(BuildDocWithBracketPlaceholders());
    var simpleBlankFill = session.FindPlaceholders(PlaceholderKinds.BlankFill)
        .Single(p => p.Match.Text == "[_____]");
    Assert.Empty(simpleBlankFill.AlternativeKinds);
}
```

- [ ] **Step 4.2: Add `BuildDocWithLongClauseBlank` helper**

Next to the other `Build*` helpers in `DocxSessionTests.cs`:

```csharp
internal static byte[] BuildDocWithLongClauseBlank()
{
    using var ms = new MemoryStream();
    using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
    {
        var main = doc.AddMainDocumentPart();
        main.Document = new Document(new Body(
            new Paragraph(new Run(new Text(
                "[that would result in at least $_______ in gross proceeds, including or excluding proceeds previously received]")))));
    }
    return ms.ToArray();
}
```

- [ ] **Step 4.3: Run, verify both fail**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS232_Classifier_LongClauseWithUnderscoresExposesAlternativeKinds|FullyQualifiedName~DS233_Classifier_SimpleBlankFillEmptyAlternativeKinds" 2>&1 | tail -15
```

Expected: build error `'TemplatePlaceholder' does not contain a definition for 'AlternativeKinds'`.

---

## Task 5: Smarter classifier — `AlternativeKinds` field (implementation)

**Files:**
- Modify: `Docxodus/DocxSession.cs` — `TemplatePlaceholder` record at line 222, and `Classify`/`FindPlaceholders` at line 1082+.

- [ ] **Step 5.1: Add `AlternativeKinds` to `TemplatePlaceholder`**

Find the existing record at `Docxodus/DocxSession.cs:222`:

```csharp
public sealed record TemplatePlaceholder
{
    required public TextMatch Match { get; init; }
    required public PlaceholderKind Kind { get; init; }

    /// <summary>For <see cref="PlaceholderKind.Instruction"/>: the inner text with
    /// surrounding brackets/asterisks stripped (e.g. <c>"[insert percentage]"</c> →
    /// <c>"insert percentage"</c>; <c>"[*specify name*]"</c> → <c>"specify name"</c>).
    /// <c>null</c> for other kinds.</summary>
    public string? Hint { get; init; }

    /// <summary>
    /// Additional plausible classifications when the primary <see cref="Kind"/> is
    /// borderline. Empty by default; populated when a secondary heuristic also
    /// matches the placeholder text. The classic case is a long bracketed clause
    /// that happens to contain a <c>_______</c> blank: primary <see cref="Kind"/>
    /// is <see cref="PlaceholderKind.BlankFill"/> for back-compat, with
    /// <see cref="PlaceholderKind.AlternativeClause"/> in <c>AlternativeKinds</c>
    /// so callers can detect the ambiguity and treat the placeholder as a clause
    /// (strip brackets, then fill the inner blank).
    /// </summary>
    public IReadOnlyList<PlaceholderKind> AlternativeKinds { get; init; } = Array.Empty<PlaceholderKind>();
}
```

- [ ] **Step 5.2: Update `Classify` to compute alternatives**

In `Docxodus/DocxSession.cs::FindPlaceholders` (around line 1082), modify the `Classify` local function. It currently returns `PlaceholderKind?`. Change it to return a tuple — primary + alternatives — and update the call site accordingly:

```csharp
// Replace the existing local Classify function (lines 1108-1128) with:
static (PlaceholderKind? Primary, IReadOnlyList<PlaceholderKind> Alternatives) Classify(string text)
{
    var inner = text.StartsWith('$') ? text[2..^1] : text[1..^1];

    // BlankFill: 2+ underscores anywhere inside (so "[__]" director-count slots,
    // "[___ times]" unit-suffix slots, and "[________ __, 20__]" date-shaped
    // slots all qualify). Tighter than "any underscore" to avoid false positives
    // on quoted identifiers like "[a_b]". Trade-off in writeup at the FindPlaceholders
    // section of docs/architecture/docx_mutation_api.md.
    bool isBlankFill = inner.Count(c => c == '_') >= 2;

    // Instruction: italicized (asterisk-wrapped) text, or starts with the
    // drafter verbs "insert" / "specify". Conservative leading-word check
    // so general prose in brackets doesn't mis-classify.
    bool isInstruction = false;
    if (inner.StartsWith('*') && inner.EndsWith('*') && inner.Length > 2) isInstruction = true;
    else
    {
        var firstWord = inner.TakeWhile(char.IsLetter).ToArray();
        var w = new string(firstWord).ToLowerInvariant();
        if (w is "insert" or "specify") isInstruction = true;
    }

    // Secondary classification: long-clause-with-blanks. When BlankFill fires but
    // the inner text reads like a multi-word clause (4+ spaces between words),
    // the placeholder is plausibly an AlternativeClause with an embedded blank.
    // Caller can detect via AlternativeKinds and strip the outer brackets, then
    // separately fill the inner _______ run.
    bool looksClause = inner.Count(c => c == ' ') >= 4;

    // Primary classification keeps the original priority order:
    //   BlankFill → Instruction → AlternativeClause
    if (isBlankFill)
    {
        var alts = looksClause ? new[] { PlaceholderKind.AlternativeClause } : Array.Empty<PlaceholderKind>();
        return (PlaceholderKind.BlankFill, alts);
    }
    if (isInstruction)
        return (PlaceholderKind.Instruction, Array.Empty<PlaceholderKind>());
    return (PlaceholderKind.AlternativeClause, Array.Empty<PlaceholderKind>());
}
```

- [ ] **Step 5.3: Update the `FindPlaceholders` loop to populate the new field**

Around line 1094-1106 in the same method, update the loop:

```csharp
foreach (var m in matches)
{
    var (classified, alternatives) = Classify(m.Text);
    if (classified is not PlaceholderKind kind) continue;
    if (!kinds.HasFlag(KindToFlag(kind))) continue;
    results.Add(new TemplatePlaceholder
    {
        Match = m,
        Kind = kind,
        Hint = kind == PlaceholderKind.Instruction ? ExtractHint(m.Text) : null,
        AlternativeKinds = alternatives,
    });
}
```

- [ ] **Step 5.4: Run new tests + ensure existing FindPlaceholders tests still pass**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS232|FullyQualifiedName~DS233|FullyQualifiedName~DS120|FullyQualifiedName~DS121|FullyQualifiedName~DS122|FullyQualifiedName~DS123|FullyQualifiedName~DS124|FullyQualifiedName~DS125|FullyQualifiedName~DS126" 2>&1 | tail -10
```

Expected: all green. `DS120`–`DS126` are the existing `FindPlaceholders` tests (per CHANGELOG); their assertions on `Kind` and `Hint` are unchanged.

- [ ] **Step 5.5: Commit**

```bash
git add Docxodus/DocxSession.cs Docxodus.Tests/DocxSessionTests.cs
git commit -m "feat(session): expose AlternativeKinds on TemplatePlaceholder for borderline classifications (#163)"
```

---

## Task 6: `FillPlaceholders` convenience (failing tests)

The picker-driven loop that every template-fill agent re-implements. Bundles reverse-offset ordering, `$`-prefix preservation, multi-pass iteration, and skip/fill accounting.

**Files:**
- Test: `Docxodus.Tests/DocxSessionTests.cs` — append `DS240`–`DS244` (5 tests).

- [ ] **Step 6.1: Add the failing tests**

```csharp
[Fact]
public void DS240_FillPlaceholders_BlankFillPickerReplacesValue()
{
    using var session = new DocxSession(BuildDocWithBracketPlaceholders());
    var result = session.FillPlaceholders(p =>
        p.Kind == PlaceholderKind.BlankFill && p.Match.Text == "[_____]"
            ? "ACME, INC."
            : null);
    Assert.True(result.Filled >= 1);
    Assert.Empty(result.Errors);

    // Verify the first paragraph now contains "ACME, INC."
    var projection = session.Project();
    var firstPara = projection.AnchorIndex.Values
        .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h");
    Assert.Contains("ACME, INC.", firstPara.TextPreview);
}

[Fact]
public void DS241_FillPlaceholders_PickerReturningNullSkips()
{
    using var session = new DocxSession(BuildDocWithBracketPlaceholders());
    var result = session.FillPlaceholders(p => null);   // skip everything
    Assert.Equal(0, result.Filled);
    Assert.True(result.Skipped > 0);
    Assert.NotEmpty(result.Unfilled);
}

[Fact]
public void DS242_FillPlaceholders_PreservesDollarPrefix()
{
    using var session = new DocxSession(BuildDocWithBracketPlaceholders());
    var result = session.FillPlaceholders(p =>
        p.Match.Text.Contains("$[___]") ? "0.20" : null);
    Assert.Equal(1, result.Filled);

    // The dollar-prefixed paragraph should now read "$0.20", not "0.20" or "$$0.20".
    var projection = session.Project();
    var allMd = projection.Markdown;
    Assert.Contains("$0.20", allMd);
    Assert.DoesNotContain("$$0.20", allMd);
}

[Fact]
public void DS243_FillPlaceholders_DollarPrefixDisabled()
{
    using var session = new DocxSession(BuildDocWithBracketPlaceholders());
    var options = new FillOptions { PreserveDollarPrefix = false };
    var result = session.FillPlaceholders(
        p => p.Match.Text.Contains("$[___]") ? "0.20" : null,
        options);
    Assert.Equal(1, result.Filled);

    // With PreserveDollarPrefix=false, the "$" is overwritten by the replacement.
    var projection = session.Project();
    Assert.Contains("share is 0.20.", projection.Markdown);
    Assert.DoesNotContain("$0.20", projection.Markdown);
}

[Fact]
public void DS244_FillPlaceholders_AlternativeClauseMultiPassStripsNestedBrackets()
{
    // The "[outer [inner] clause]" placeholder is nested; FindPlaceholders returns
    // INNERMOST first, so stripping has to iterate. FillPlaceholders should converge
    // after multiple passes when picker strips brackets uniformly.
    using var session = new DocxSession(BuildDocWithBracketPlaceholders());
    var options = new FillOptions { Kinds = PlaceholderKinds.AlternativeClause };
    var result = session.FillPlaceholders(p =>
    {
        // Strip [ and ] from the match's bracketed portion, preserve any prefix/suffix.
        var t = p.Match.Text;
        int lb = t.IndexOf('['), rb = t.LastIndexOf(']');
        if (lb < 0 || rb <= lb) return null;
        return t[..lb] + t[(lb + 1)..rb] + t[(rb + 1)..];
    }, options);

    Assert.True(result.Passes >= 2);

    // After convergence, no AlternativeClause matches remain.
    var leftover = session.FindPlaceholders(PlaceholderKinds.AlternativeClause);
    Assert.Empty(leftover);
}
```

- [ ] **Step 6.2: Run, verify they fail**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS240_FillPlaceholders_BlankFillPickerReplacesValue|FullyQualifiedName~DS241_FillPlaceholders_PickerReturningNullSkips|FullyQualifiedName~DS242_FillPlaceholders_PreservesDollarPrefix|FullyQualifiedName~DS243_FillPlaceholders_DollarPrefixDisabled|FullyQualifiedName~DS244_FillPlaceholders_AlternativeClauseMultiPassStripsNestedBrackets" 2>&1 | tail -15
```

Expected: build errors — `FillOptions`, `FillPlaceholders`, `BulkEditResult` not defined.

---

## Task 7: `FillPlaceholders` convenience (implementation)

**Files:**
- Modify: `Docxodus/DocxSession.cs` — add `FillOptions`, `BulkEditResult`, and `FillPlaceholders` next to `FindPlaceholders` (around line 1082).

- [ ] **Step 7.1: Add `FillOptions` and `BulkEditResult` records**

Just above `FindPlaceholders` (around the existing `ReplaceOptions` record block — search the file for `public sealed record ReplaceOptions`) add:

```csharp
/// <summary>
/// Options for <see cref="DocxSession.FillPlaceholders"/>.
/// </summary>
public sealed record FillOptions
{
    /// <summary>Which placeholder kinds to fill. Defaults to <c>BlankFill | Instruction</c>
    /// — the kinds with a single replacement value. Add <c>AlternativeClause</c> to
    /// run the picker on bracketed clauses too (e.g. for bracket-stripping).</summary>
    public PlaceholderKinds Kinds { get; init; } = PlaceholderKinds.BlankFill | PlaceholderKinds.Instruction;

    /// <summary>Which package parts to scan. Defaults to body.</summary>
    public ProjectionScopes Scope { get; init; } = ProjectionScopes.Body;

    /// <summary>Maximum iteration passes. <see cref="FindPlaceholders"/> returns
    /// innermost brackets only; stripping one layer can surface a previously-nested
    /// outer layer, so multi-pass iteration is sometimes needed. The default of 8
    /// is a safety cap against infinite loops on adversarial input. Set higher if
    /// you have deeply-nested templates.</summary>
    public int MaxPasses { get; init; } = 8;

    /// <summary>When <c>true</c> (default), if the placeholder match text starts
    /// with <c>"$"</c> (the regex <c>\$?\[…\]</c> captured a leading dollar sign)
    /// and the picker's return value does not start with <c>"$"</c>, the dollar
    /// is preserved by prepending it to the replacement. Set to <c>false</c> if
    /// you want full control over the replacement and to overwrite the <c>$</c>.</summary>
    public bool PreserveDollarPrefix { get; init; } = true;
}

/// <summary>
/// Aggregate result envelope returned by <see cref="DocxSession.FillPlaceholders"/>.
/// </summary>
public sealed record BulkEditResult
{
    /// <summary>Number of placeholders filled by the picker.</summary>
    public int Filled { get; init; }

    /// <summary>Number of placeholders for which the picker returned <c>null</c>
    /// (counted once per placeholder, in the first pass that saw it).</summary>
    public int Skipped { get; init; }

    /// <summary>How many iteration passes ran. <c>1</c> means a single pass
    /// converged; higher values mean multi-pass nested-bracket stripping.</summary>
    public int Passes { get; init; }

    /// <summary>Placeholders the picker returned <c>null</c> for.</summary>
    public IReadOnlyList<TemplatePlaceholder> Unfilled { get; init; } = Array.Empty<TemplatePlaceholder>();

    /// <summary>Per-replacement failures. Populated when <see cref="ReplaceMatch"/>
    /// returned <c>Success = false</c> for an attempted fill.</summary>
    public IReadOnlyList<EditError> Errors { get; init; } = Array.Empty<EditError>();
}
```

- [ ] **Step 7.2: Add `FillPlaceholders` method**

Add this method directly after `FindPlaceholders` in `DocxSession.cs` (around line 1146, after the existing local-helper closures):

```csharp
/// <summary>
/// Picker-driven template fill. For every placeholder matching
/// <see cref="FillOptions.Kinds"/>, calls <paramref name="picker"/>; if the picker
/// returns a non-null string, the placeholder is replaced (with optional
/// <c>$</c>-prefix preservation per <see cref="FillOptions.PreserveDollarPrefix"/>).
/// Iterates until no more placeholders match (or until <see cref="FillOptions.MaxPasses"/>
/// is reached, or a pass makes zero state changes) — important when
/// <see cref="FillOptions.Kinds"/> includes <see cref="PlaceholderKinds.AlternativeClause"/>
/// and the doc has nested brackets that surface only after the inner ones are stripped.
/// Replacements within a paragraph are applied in reverse-offset order automatically.
/// </summary>
public BulkEditResult FillPlaceholders(
    Func<TemplatePlaceholder, string?> picker,
    FillOptions? options = null)
{
    ThrowIfDisposed();
    ArgumentNullException.ThrowIfNull(picker);
    var opts = options ?? new FillOptions();

    int filled = 0;
    int passes = 0;
    var errors = new List<EditError>();
    var unfilled = new List<TemplatePlaceholder>();
    var seenSkipKeys = new HashSet<(string AnchorId, int Start, int Length)>();

    for (passes = 1; passes <= opts.MaxPasses; passes++)
    {
        var placeholders = FindPlaceholders(opts.Kinds, opts.Scope)
            .OrderByDescending(p => p.Match.EnclosingAnchor.Anchor.Id, StringComparer.Ordinal)
            .ThenByDescending(p => p.Match.Span.Start)
            .ToList();
        if (placeholders.Count == 0) break;

        int passChanges = 0;
        foreach (var p in placeholders)
        {
            var pick = picker(p);
            if (pick is null)
            {
                // Count each skip exactly once per placeholder lifetime.
                var key = (p.Match.EnclosingAnchor.Anchor.Id, p.Match.Span.Start, p.Match.Span.Length);
                if (seenSkipKeys.Add(key))
                    unfilled.Add(p);
                continue;
            }

            if (opts.PreserveDollarPrefix && p.Match.Text.StartsWith("$") && !pick.StartsWith("$"))
                pick = "$" + pick;

            var r = ReplaceMatch(p.Match, pick);
            if (r.Success)
            {
                filled++;
                passChanges++;
            }
            else if (r.Error is { } err)
            {
                errors.Add(err);
            }
        }

        // If this pass made no changes, the picker is steady-state — stop iterating.
        if (passChanges == 0) break;
    }

    return new BulkEditResult
    {
        Filled = filled,
        Skipped = unfilled.Count,
        Passes = passes > opts.MaxPasses ? opts.MaxPasses : passes,
        Unfilled = unfilled,
        Errors = errors,
    };
}
```

- [ ] **Step 7.3: Run the new tests, verify they pass**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS240|FullyQualifiedName~DS241|FullyQualifiedName~DS242|FullyQualifiedName~DS243|FullyQualifiedName~DS244" 2>&1 | tail -10
```

Expected: 5/5 pass.

- [ ] **Step 7.4: Run the full DocxSession suite for regressions**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DocxSessionTests" 2>&1 | tail -5
```

Expected: all green.

- [ ] **Step 7.5: Commit**

```bash
git add Docxodus/DocxSession.cs Docxodus.Tests/DocxSessionTests.cs
git commit -m "feat(session): FillPlaceholders convenience with picker-driven multi-pass fill (#163)"
```

---

## Task 8: WASM bridge — expose `ReplaceInner`, `AlternativeKinds`, and `FillPlaceholders` is NOT bridged

`FillPlaceholders` takes a `Func<TemplatePlaceholder, string?>` — there's no clean way to pass a C# delegate across the JS/WASM boundary. The npm wrapper will re-implement the same loop on top of the existing `findPlaceholders` + `replaceMatch` + new `replaceInner` JSExports. This is consistent with how other higher-order operations are handled across the bridge (the .NET side has a single combined method; the npm side composes from primitives).

What we DO need to bridge:
- `ReplaceInner` JSExport (one new method).
- `AlternativeKinds` field in the `TemplatePlaceholder` JSON serialization (additive — existing JSON consumers ignore unknown keys).

**Files:**
- Modify: `wasm/DocxodusWasm/DocxSessionBridge.cs` — add `ReplaceInner` JSExport, update `SerializePlaceholders` to emit `alternativeKinds`.

- [ ] **Step 8.1: Add `ReplaceInner` JSExport**

In `wasm/DocxodusWasm/DocxSessionBridge.cs`, find `ReplaceMatch` JSExport (search for `public static string ReplaceMatch` or similar — likely near the other Replace* exports around line 175-200). Add directly after:

```csharp
/// <summary>
/// Bridge for <see cref="DocxSession.ReplaceInner"/>. Takes a JSON-serialized
/// <see cref="TextMatch"/> (same shape as Grep results) plus the new inner
/// content; returns the standard EditResult JSON.
/// </summary>
[JSExport]
public static string ReplaceInner(int h, string matchJson, string newInner)
{
    TextMatch? match;
    try { match = DeserializeTextMatch(matchJson); }
    catch (JsonException) { return Serialize(EditResult.Fail(EditErrorCode.MalformedMarkdown, "malformed match JSON")); }
    if (match is null) return Serialize(EditResult.Fail(EditErrorCode.AnchorNotFound, "match is null"));
    return Serialize(Get(h).ReplaceInner(match, newInner));
}
```

If a `DeserializeTextMatch` helper does not yet exist, you'll likely find the npm side passes match JSON directly to `replaceMatch`. Look at how `ReplaceMatch` does this (search the file). The pattern may be to pass `(anchor, spanStart, spanLength, newInner)` directly instead of a TextMatch JSON blob. Inspect the existing `ReplaceMatch` JSExport to see the convention.

**If the existing convention is `(anchor, spanStart, spanLength, replace)`**, use that shape instead:

```csharp
[JSExport]
public static string ReplaceInner(int h, string matchText, string anchor, int spanStart, int spanLength, string newInner)
{
    // Reconstruct a minimal TextMatch from the wire params; bracket math operates
    // on matchText alone, while the span addresses the live element for ReplaceMatch.
    int lb = matchText.IndexOf('[');
    int rb = matchText.LastIndexOf(']');
    if (lb < 0 || rb <= lb)
        return Serialize(EditResult.Fail(EditErrorCode.MalformedMarkdown,
            $"match text has no balanced brackets: '{matchText}'", anchor));
    var prefix = matchText[..lb];
    var suffix = matchText[(rb + 1)..];
    return Serialize(Get(h).ReplaceTextAtSpan(anchor, spanStart, spanLength, prefix + newInner + suffix));
}
```

Inspect `npm/src/session.ts::replaceMatch` to confirm which convention is in use, then pick the matching pattern.

- [ ] **Step 8.2: Update `SerializePlaceholders` to emit `alternativeKinds`**

Find `SerializePlaceholders` in `DocxSessionBridge.cs` (search the file). It currently emits `{kind, hint, match}`. Add `alternativeKinds` as a JSON array of enum strings:

```csharp
// Inside the existing SerializePlaceholders loop, after emitting "hint":
sb.Append(",\"alternativeKinds\":[");
for (int i = 0; i < p.AlternativeKinds.Count; i++)
{
    if (i > 0) sb.Append(',');
    sb.Append(JsonString(KindToString(p.AlternativeKinds[i])));
}
sb.Append(']');
```

If a `KindToString` helper does not exist in the bridge file, look at how `kind` itself is serialized in the existing method — there's already a mapping from `PlaceholderKind` enum to the snake_case string union (`"blank_fill"`, `"alternative_clause"`, `"instruction"`). Reuse that helper.

- [ ] **Step 8.3: Build the WASM target**

```bash
./scripts/build-wasm.sh 2>&1 | tail -5
```

Expected: clean build.

- [ ] **Step 8.4: Commit**

```bash
git add wasm/DocxodusWasm/DocxSessionBridge.cs
git commit -m "feat(wasm): expose ReplaceInner + alternativeKinds on placeholder JSON (#163)"
```

---

## Task 9: npm wrapper — `replaceInner`, `alternativeKinds`, `fillPlaceholders`

The npm side gets the primitive `replaceInner` from the new JSExport, the additive `alternativeKinds` field on `TemplatePlaceholder`, and a TypeScript-side implementation of `fillPlaceholders` that re-implements the .NET control loop. Same name, same shape — agents on either platform see the same API.

**Files:**
- Modify: `npm/src/types.ts` — add `alternativeKinds`, `FillOptions`, `BulkEditResult`; declare new JSExport.
- Modify: `npm/src/session.ts` — add `replaceInner` and `fillPlaceholders` methods.

- [ ] **Step 9.1: Update `TemplatePlaceholder` type**

In `npm/src/types.ts` around line 909:

```typescript
export interface TemplatePlaceholder {
  kind: PlaceholderKind;
  /** For `instruction` placeholders: the inner text with surrounding brackets/asterisks stripped. */
  hint?: string;
  match: TextMatch;
  /**
   * Additional plausible classifications when the primary `kind` is borderline.
   * Empty by default. The classic case is a long bracketed clause that happens
   * to contain a `_______` blank: primary `kind` stays `"blank_fill"`
   * (back-compat) and `alternativeKinds` contains `"alternative_clause"`.
   */
  alternativeKinds: PlaceholderKind[];
}
```

- [ ] **Step 9.2: Add `FillOptions` and `BulkEditResult` types**

Below the placeholder interfaces in `npm/src/types.ts`:

```typescript
/**
 * Options for {@link DocxSession.fillPlaceholders}.
 */
export interface FillOptions {
  /** Which placeholder kinds to fill. Defaults to `BlankFill | Instruction`. */
  kinds?: number;
  /** Which package parts to scan. Defaults to body (1). */
  scope?: number;
  /** Max iteration passes for multi-pass nested-bracket scenarios. Default 8. */
  maxPasses?: number;
  /** When the match starts with `$` and the picker's return value doesn't,
   *  preserve the `$` by prepending it. Default true. */
  preserveDollarPrefix?: boolean;
}

/**
 * Aggregate result returned by {@link DocxSession.fillPlaceholders}.
 */
export interface BulkEditResult {
  filled: number;
  skipped: number;
  passes: number;
  unfilled: TemplatePlaceholder[];
  errors: EditError[];
}
```

- [ ] **Step 9.3: Declare the new bridge method on `DocxodusWasmExports`**

In `npm/src/types.ts`, find the `DocxodusWasmExports` interface (search for `FindPlaceholders` to locate it — should be near line 691). Add the new bridge entry next to it. The exact signature depends on which convention you adopted in Task 8.1 — match it:

If the JSExport signature is `ReplaceInner(int h, string matchText, string anchor, int spanStart, int spanLength, string newInner)`:

```typescript
ReplaceInner: (
  handle: number,
  matchText: string,
  anchor: string,
  spanStart: number,
  spanLength: number,
  newInner: string,
) => string;
```

- [ ] **Step 9.4: Add `replaceInner` wrapper on `DocxSession`**

In `npm/src/session.ts`, find `replaceMatch` (around line 88-110 based on prior reads). Add `replaceInner` directly after:

```typescript
/**
 * Replace the bracketed portion of a `TextMatch` with `newInner`, preserving any
 * prefix or suffix outside the brackets. Designed for `findPlaceholders` matches
 * like `$[___]` where the regex `\$?\[…\]` captures a leading `$`:
 * `replaceInner(match, "0.20")` yields `$0.20`, not `0.20`.
 *
 * Returns `MalformedMarkdown` if the match text does not contain balanced brackets.
 */
replaceInner(match: TextMatch, newInner: string): EditResult {
  return JSON.parse(this.wasm.ReplaceInner(
    this.handle,
    match.text,
    match.enclosingAnchor.id,
    match.span.start,
    match.span.length,
    newInner,
  )) as EditResult;
}
```

- [ ] **Step 9.5: Add `fillPlaceholders` wrapper on `DocxSession`**

After `replaceInner`, add:

```typescript
/**
 * Picker-driven template fill. For every placeholder matching `options.kinds`,
 * calls `picker`; if the picker returns a non-null string, the placeholder is
 * replaced (with optional `$`-prefix preservation). Iterates until no more
 * placeholders match (or `maxPasses` is reached, or a pass makes zero changes)
 * — handles nested brackets that surface only after the inner ones are stripped.
 *
 * The TypeScript implementation mirrors the .NET `DocxSession.FillPlaceholders`
 * exactly. The picker is invoked synchronously inside the WASM worker; do not
 * use it for async work or for picker logic that needs to await fetched data.
 */
fillPlaceholders(
  picker: (p: TemplatePlaceholder) => string | null | undefined,
  options?: FillOptions,
): BulkEditResult {
  const opts = options ?? {};
  const kinds = opts.kinds ?? (PlaceholderKinds.BlankFill | PlaceholderKinds.Instruction);
  const scope = opts.scope ?? 1; // Body
  const maxPasses = opts.maxPasses ?? 8;
  const preserveDollarPrefix = opts.preserveDollarPrefix ?? true;

  let filled = 0;
  let passes = 0;
  const errors: EditError[] = [];
  const unfilled: TemplatePlaceholder[] = [];
  const seenSkipKeys = new Set<string>();

  for (passes = 1; passes <= maxPasses; passes++) {
    const placeholders = this.findPlaceholders(kinds, scope)
      .sort((a, b) => {
        const cmp = b.match.enclosingAnchor.id.localeCompare(a.match.enclosingAnchor.id);
        if (cmp !== 0) return cmp;
        return b.match.span.start - a.match.span.start;
      });
    if (placeholders.length === 0) break;

    let passChanges = 0;
    for (const p of placeholders) {
      const pick = picker(p);
      if (pick == null) {
        const key = `${p.match.enclosingAnchor.id}:${p.match.span.start}:${p.match.span.length}`;
        if (!seenSkipKeys.has(key)) {
          seenSkipKeys.add(key);
          unfilled.push(p);
        }
        continue;
      }

      let replacement = pick;
      if (preserveDollarPrefix && p.match.text.startsWith("$") && !replacement.startsWith("$")) {
        replacement = "$" + replacement;
      }

      const r = this.replaceMatch(p.match, replacement);
      if (r.success) {
        filled++;
        passChanges++;
      } else if (r.error) {
        errors.push(r.error);
      }
    }

    if (passChanges === 0) break;
  }

  return {
    filled,
    skipped: unfilled.length,
    passes: Math.min(passes, maxPasses),
    unfilled,
    errors,
  };
}
```

Add `TemplatePlaceholder`, `TextMatch`, `EditError`, `FillOptions`, `BulkEditResult`, `PlaceholderKinds` to the imports at the top of `session.ts` if not already present.

- [ ] **Step 9.6: Re-export new types from `index.ts`**

In `npm/src/index.ts`, find the public type re-export block (search for `export type {` or `AnchorInfo` which Unit E added). Append:

```typescript
export type { FillOptions, BulkEditResult } from "./types.js";
```

- [ ] **Step 9.7: Build the npm package and type-check**

```bash
cd /home/jman/Code/Docxodus/npm
npm run build 2>&1 | tail -5
npx tsc --noEmit 2>&1 | tail -10
```

Expected: both clean (no errors, no warnings).

- [ ] **Step 9.8: Commit**

```bash
cd /home/jman/Code/Docxodus
git add npm/src/
git commit -m "feat(npm): replaceInner + fillPlaceholders + alternativeKinds (#163)"
```

---

## Task 10: Playwright integration test

**Files:**
- Create: `npm/tests/fill-placeholders.spec.ts`

- [ ] **Step 10.1: Inspect a similar existing spec**

```bash
ls /home/jman/Code/Docxodus/npm/tests/ | head -15
```

Read `anchor-introspection.spec.ts` (added in #162) end-to-end — it's the most recent and closest in style. Mirror its harness setup, fixture-load pattern, and try/finally cleanup idiom.

- [ ] **Step 10.2: Write the spec**

Create `npm/tests/fill-placeholders.spec.ts` with these assertion blocks (mirror the harness pattern from `anchor-introspection.spec.ts`):

1. **`fillPlaceholders` picker fills a single BlankFill.** Load a fixture with one `[_____]`, picker returns a literal, assert `result.filled === 1` and the doc text reflects the fill.
2. **`fillPlaceholders` picker returning null skips.** Same fixture, picker always returns `null`, assert `result.filled === 0`, `result.unfilled.length > 0`.
3. **`fillPlaceholders` preserves `$` prefix.** Fixture with `$[___]`, picker returns plain `"0.20"`, doc ends up with `$0.20`.
4. **`fillPlaceholders` multi-pass strips nested brackets.** Fixture with `[outer [inner] clause]`, picker strips outer brackets, assert `result.passes >= 2` and no AlternativeClause matches remain in `findPlaceholders` after.
5. **`replaceInner` directly.** Match a `$[___]` placeholder, call `replaceInner(match, "0.20")`, assert the result text contains `$0.20`.
6. **`alternativeKinds` populated for long-clause-with-blanks.** Construct or fetch a fixture that has `[...some long sentence with $_______ in it...]`, call `findPlaceholders`, assert the placeholder's `kind === "blank_fill"` AND `alternativeKinds` includes `"alternative_clause"`.

You'll need a fixture; the simplest path is to construct it inline via the existing test-harness DOCX-builder pattern (the `anchor-introspection.spec.ts` uses an existing TestFiles fixture; check if any has bracket placeholders, otherwise see the harness's `buildDocFromXml`-style helper).

- [ ] **Step 10.3: Run the spec**

```bash
cd /home/jman/Code/Docxodus/npm && npm test -- --grep "fill-placeholders" 2>&1 | tail -15
```

(Use `npm test --` not `npx playwright test` because the harness copy happens in the `pretest` script. Pitfall §10 of `~/Code/Docxodus-Agents/PITFALLS.md`.)

Expected: green.

- [ ] **Step 10.4: Commit**

```bash
cd /home/jman/Code/Docxodus
git add npm/tests/fill-placeholders.spec.ts
git commit -m "test(npm): fillPlaceholders + replaceInner + alternativeKinds integration (#163)"
```

---

## Task 11: Documentation + #162 minor cleanup

Three docs files in this repo, two agent-guide files outside it, plus three small cleanups carried over from #162 reviewer notes.

**Files:**
- `CHANGELOG.md`
- `docs/architecture/docx_mutation_api.md`
- `~/Code/Docxodus-Agents/TOOL_SURFACE.md` (out-of-repo)
- `~/Code/Docxodus-Agents/PITFALLS.md` (out-of-repo)
- `Docxodus.Tests/DocxSessionTests.cs` — DS220 extension (#162 carryover)
- `Docxodus/WmlToMarkdownConverter.cs` — drop redundant comment (#162 carryover)

- [ ] **Step 11.1: CHANGELOG entry**

Open `CHANGELOG.md`, find `## [Unreleased]` and add under `### Added`:

```markdown
- **Template-fill convenience — `FillPlaceholders`, `ReplaceInner`, `AlternativeKinds`** (issue #163). `DocxSession.FillPlaceholders(picker, options?)` bundles the three foot-guns every template-fill agent re-implements: reverse-offset ordering across matches within a paragraph, `$`-prefix preservation (`$[___]` → `$0.20` instead of `0.20`), and multi-pass iteration for nested AlternativeClause brackets. Returns a `BulkEditResult` with `Filled` / `Skipped` / `Passes` counts plus per-failure error and unfilled-placeholder lists. New `DocxSession.ReplaceInner(match, newInner)` overload replaces only the bracketed portion of a match, preserving any prefix or suffix outside it — the canonical use case for `$[___]` matches where the regex `\$?\[…\]` captured a leading `$`. `TemplatePlaceholder.AlternativeKinds` is a new additive field listing secondary classifications when the primary `Kind` is borderline (e.g. a long bracketed clause containing a `_______` blank: primary `Kind` stays `BlankFill` for back-compat, with `AlternativeClause` in `AlternativeKinds`). WASM bridge: new `ReplaceInner` JSExport; placeholder JSON gains an `alternativeKinds` array. npm wrapper: `session.replaceInner(match, newInner)`, `session.fillPlaceholders(picker, options?)` (TS-side mirror of the .NET control loop), new `FillOptions` and `BulkEditResult` types. Tests: `DS230`–`DS233` (ReplaceInner + AlternativeKinds), `DS240`–`DS244` (FillPlaceholders), Playwright `fill-placeholders.spec.ts`.
```

- [ ] **Step 11.2: `docs/architecture/docx_mutation_api.md`**

Find the existing `FindPlaceholders — template-slot enumeration` section. After the existing recipe code block, add a new subsection:

```markdown
### `FillPlaceholders` — picker-driven multi-pass fill

The 5-line recipe above is now a first-class call:

```csharp
var summary = session.FillPlaceholders(p => p.Kind switch
{
    PlaceholderKind.Instruction when p.Hint?.Contains("price") == true => "1.50",
    PlaceholderKind.BlankFill when p.Match.ContextBefore.TrimEnd().EndsWith("name is") => "ACME, INC.",
    _ => null
});
// summary.Filled / .Skipped / .Passes / .Unfilled / .Errors
```

What `FillPlaceholders` does internally that the recipe doesn't:

- **Reverse-offset ordering.** Earlier-offset matches in the same paragraph go stale once a later edit lands; `FillPlaceholders` sorts every pass by `(anchorId desc, span.Start desc)` so each block's matches are applied right-to-left.
- **`$`-prefix preservation.** The placeholder regex `\$?\[…\]` captures `$[___]` including the leading `$`. With `FillOptions.PreserveDollarPrefix = true` (default), the picker's return value gets `$` prepended when needed so `"0.20"` lands as `$0.20`, not `0.20`.
- **Multi-pass iteration.** `FindPlaceholders` returns innermost brackets only; stripping the inner can surface a previously-nested outer. The loop re-finds placeholders each pass and stops when a pass makes zero changes (or `MaxPasses` — default 8 — is hit).

The picker is invoked once per placeholder per pass; return `null` to skip. `BulkEditResult.Unfilled` lists every placeholder the picker said `null` to (deduplicated across passes).

### `ReplaceInner` — strip brackets while preserving prefix/suffix

```csharp
// match.Text = "$[___]"
session.ReplaceInner(match, "0.20");
// paragraph now contains "$0.20" (the leading $ outside the brackets survives).
```

`ReplaceInner` parses the brackets out of `match.Text` and substitutes the new inner for everything between (and including) `[` and `]`, then dispatches to `ReplaceMatch` with the recomposed string. Returns `MalformedMarkdown` if the match text has no balanced brackets.

### `TemplatePlaceholder.AlternativeKinds`

When the primary classification is borderline, secondary classifications are exposed via the `AlternativeKinds` list. The current borderline case: a `BlankFill` whose inner text contains 4+ spaces (i.e. reads like a multi-word clause that happens to contain a `_______`). Primary `Kind` stays `BlankFill` for back-compat; `AlternativeKinds` lists `AlternativeClause` so callers can detect the ambiguity.
```

- [ ] **Step 11.3: Update agent guide — `TOOL_SURFACE.md`**

In `~/Code/Docxodus-Agents/TOOL_SURFACE.md`, find the "Edit — text content (Tier A)" table. Replace the row for `Replace a specific TextMatch …` with two rows:

```markdown
| Replace a specific `TextMatch` returned by `Grep`/`FindPlaceholders` | `session.ReplaceMatch(textMatch, replace)` |
| Replace only the bracketed portion of a match, preserving prefix/suffix (e.g. `$[___]` → `$0.20`) | `session.ReplaceInner(textMatch, newInner)` |
```

Then find the section that recommends the "5-line template-fill recipe" (likely under template-fill or placeholders) and add a callout pointing at `FillPlaceholders`:

```markdown
**Prefer `session.FillPlaceholders(picker, options?)` for the whole loop** — it
bundles reverse-offset ordering, `$`-prefix preservation, and multi-pass nested
bracket iteration. The recipe below is now a one-line picker callback.
```

- [ ] **Step 11.4: Update agent guide — `PITFALLS.md`**

In `~/Code/Docxodus-Agents/PITFALLS.md`, add "Fixed in Docxodus 6.1.0 (#163)" callouts at the tops of §3 (the `$`-prefix preservation pitfall), §4 (BlankFill classifier permissiveness), §5 (innermost-bracket-only iteration), and §6 (reverse-offset ordering) — these are exactly the four foot-guns `FillPlaceholders` now bundles away.

For each section, insert a blockquote right after the `## §N. <title>` heading, before the `**Symptom.**` line:

```markdown
> **Fixed in Docxodus 6.1.0 (#163) for template fills.** `session.FillPlaceholders(picker, options?)`
> handles this automatically; or use `session.ReplaceInner(match, newInner)` for direct `$[…]` cases.
> This section is preserved for agents working against Docxodus 6.0.x or earlier, or doing
> custom edit loops outside `FillPlaceholders`.
```

- [ ] **Step 11.5: #162 carryover — extend DS220 to cover dedup + null-string skip**

In `Docxodus.Tests/DocxSessionTests.cs`, locate `DS220_GetAnchorInfosBulkLookup`. Add this companion test directly after it:

```csharp
[Fact]
public void DS220b_GetAnchorInfos_DedupesAndSkipsNullIds()
{
    using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
    var projection = session.Project();
    var realId = projection.AnchorIndex.Values
        .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h" or "li")
        .Anchor.Id;

    // Mixed input: same real id twice, a null, and an unknown id.
    string?[] ids = new string?[] { realId, realId, null, "p:body:unknown-id-xyz" };
    var bulk = session.GetAnchorInfos(ids.Where(s => s is not null).Select(s => s!));

    // Note: null strings are silently filtered by the caller's Where clause,
    // since the public API takes IEnumerable<string> (non-null). The dedup
    // assertion still holds: realId appears twice in the input enumerable;
    // bulk has exactly one entry for it.
    Assert.Equal(2, bulk.Count);
    Assert.True(bulk.ContainsKey(realId));
    Assert.Null(bulk["p:body:unknown-id-xyz"]);

    // And explicitly: passing null strings via reflection-style is also tolerated
    // by the implementation, even though the static type is non-nullable.
    IEnumerable<string> withNulls = new[] { realId, null!, realId };
    var bulkWithNulls = session.GetAnchorInfos(withNulls);
    Assert.Single(bulkWithNulls);                   // dedup: one entry
    Assert.True(bulkWithNulls.ContainsKey(realId));
}
```

- [ ] **Step 11.6: #162 carryover — drop redundant comment in `ComputeTextPreview`**

In `Docxodus/WmlToMarkdownConverter.cs::ComputeTextPreview`, find the line:

```csharp
        ? text.Substring(0, TextPreviewMaxLength) + "…"   // "…"
```

Remove the trailing `// "…"` comment (the literal echoes itself; the comment is noise):

```csharp
        ? text.Substring(0, TextPreviewMaxLength) + "…"
```

- [ ] **Step 11.7: Run DS220b + ensure nothing else broke**

```bash
cd /home/jman/Code/Docxodus
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS220" 2>&1 | tail -5
```

Expected: both pass.

- [ ] **Step 11.8: Commit (combined docs + carryover)**

```bash
git add CHANGELOG.md docs/architecture/docx_mutation_api.md Docxodus.Tests/DocxSessionTests.cs Docxodus/WmlToMarkdownConverter.cs
git commit -m "docs: template-fill convenience + minor #162 cleanups (#163)"
```

Note: `~/Code/Docxodus-Agents/` is not a git repo (per #162 Unit G), so its files are saved in place — no commit needed there.

---

## Task 12: Final verification + PR

- [ ] **Step 12.1: Full .NET test suite**

After `./scripts/build-wasm.sh` runs in Task 8, the cached `Docxodus.dll` is WASM-mode. Force a clean rebuild before testing:

```bash
cd /home/jman/Code/Docxodus
dotnet clean Docxodus.sln 2>&1 | tail -3
# If clean isn't enough (#162 Unit H note), nuke bin/obj:
rm -rf Docxodus/bin Docxodus/obj Docxodus.Tests/bin Docxodus.Tests/obj
dotnet build Docxodus.sln 2>&1 | tail -5
dotnet test Docxodus.Tests/Docxodus.Tests.csproj 2>&1 | tail -10
```

Expected: all tests pass; 0 failed.

- [ ] **Step 12.2: npm build + Playwright**

```bash
cd /home/jman/Code/Docxodus/npm
npm run build 2>&1 | tail -5
npm test 2>&1 | tail -15
```

Expected: all green, including the new `fill-placeholders.spec.ts` and the unchanged `anchor-introspection.spec.ts`.

- [ ] **Step 12.3: Release build (warnings-as-errors)**

```bash
cd /home/jman/Code/Docxodus
dotnet build -c Release Docxodus.sln 2>&1 | tail -10
```

Expected: clean. Pre-existing legacy warnings are OK; any new warning in files this branch touched needs to be investigated.

- [ ] **Step 12.4: Push and open PR (links via `Closes #163`)**

```bash
git push -u origin feat/163-fill-placeholders 2>&1 | tail -5

gh pr create --title "feat: template-fill convenience — FillPlaceholders, ReplaceInner, AlternativeKinds" --body "$(cat <<'EOF'
Closes #163.

## Summary
- `DocxSession.FillPlaceholders(picker, options?)` — picker-driven multi-pass fill that bundles reverse-offset ordering, `$`-prefix preservation, and nested-bracket iteration.
- `DocxSession.ReplaceInner(match, newInner)` — replaces only the bracketed portion of a match, preserving prefix/suffix (so `$[___]` → `$0.20` not `0.20`).
- `TemplatePlaceholder.AlternativeKinds` — additive field exposing secondary classifications when the primary `Kind` is borderline (e.g. long clause containing a blank).
- npm wrapper: same shape — `session.replaceInner(...)`, `session.fillPlaceholders(...)`, `alternativeKinds` on placeholder results.
- Folds in three minor #162 review carryovers: DS220 dedup test extension, "single pass" docstring tighten, redundant comment removal.

## Test plan
- [x] `DS230`–`DS231` — `ReplaceInner` strip + error path.
- [x] `DS232`–`DS233` — `AlternativeKinds` populated for borderline / empty for simple.
- [x] `DS240`–`DS244` — `FillPlaceholders` picker, null-skip, `$`-prefix preserve / disable, multi-pass nested-bracket.
- [x] `DS220b` — `GetAnchorInfos` dedup + null-id handling (#162 carryover).
- [x] Playwright `fill-placeholders.spec.ts` — all 6 assertion areas.
- [x] Full `dotnet test` passes.
- [x] `npm run build && npm test` passes.
- [x] Release build clean.
EOF
)" 2>&1 | tail -5
```

Expected: PR URL printed.

---

## Self-Review

**Spec coverage** (against issue #163's three pieces):
- §1 `FillPlaceholders` convenience: Tasks 6+7 (.NET) + Task 9 step 9.5 (npm). ✓
- §1 sub-feature: reverse-offset ordering, `$`-prefix preservation, multi-pass nested iteration — all three covered in Task 7's implementation with explicit test coverage in DS242, DS243, DS244. ✓
- §2 `ReplaceInner` overload: Tasks 2+3 (.NET) + Task 9 step 9.4 (npm). ✓
- §3 Smarter classifier — chose `AlternativeKinds` (option (b) from the issue body) over `Confidence` field (option (a)) because it preserves all candidate classifications without losing information. Tasks 4+5. ✓
- WASM bridge: Task 8 propagates `ReplaceInner` + `alternativeKinds` (decision documented in Task 8: `FillPlaceholders` is not bridged — re-implemented in npm because picker delegates don't cross the JS/WASM boundary cleanly).
- Docs: Task 11. ✓
- Carryover items from #162 (DS220 extension, comment cleanup): Task 11 steps 11.5–11.6. ✓

**Placeholder scan:** no `TODO`, no `Similar to Task N`, no naked "add error handling." Every code block is real C# / TS / shell.

**Type consistency:** verified `FillOptions`, `BulkEditResult`, `TemplatePlaceholder.AlternativeKinds` are spelled the same in .NET (PascalCase) and TS (camelCase) layers. `MaxPasses` not `MaxBracketStripPasses` — corrected from the issue body which used the longer name; the shorter name matches the file's existing convention (e.g., `MaxReplacements` in `ReplaceOptions`).

One spec deviation worth flagging:
- Issue body proposes `FillOptions.MaxBracketStripPasses`. Plan uses `MaxPasses` because the option governs all iteration, not just bracket-stripping (it caps `FillPlaceholders` whether the picker is filling blanks or stripping clauses). Naming the option after one of its use cases would be misleading.

---

**Plan complete and saved to `docs/superpowers/plans/2026-05-25-issue-163-template-fill-convenience.md`.**

Two execution options:

1. **Subagent-Driven (recommended)** — same flow as #162. Fresh subagent per task, two-stage review (spec then code quality), continuous execution.
2. **Inline Execution** — execute tasks here with checkpoints. Faster on the small ones but consumes my context window.
