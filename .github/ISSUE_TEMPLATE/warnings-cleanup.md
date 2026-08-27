---
name: Technical Debt - Warnings Cleanup
about: Track warnings that need to be resolved
title: "Fix compiler warnings and re-enable TreatWarningsAsErrors"
labels: technical-debt, good-first-issue
---

## Summary

`Directory.Build.props` sets `TreatWarningsAsErrors=true` for Release, but `Docxodus.csproj`
and `Docxodus.Tests.csproj` both override it to `false` — the inherited OpenXmlPowerTools core
carries too much legacy debt to fail the build on. The CLI tools, MCP server, python-host and
WASM project do inherit the strict setting. The goal is to clear the two overrides.

## Current suppressions

`Docxodus/Docxodus.csproj`:
```xml
<TreatWarningsAsErrors>false</TreatWarningsAsErrors>
<NoWarn>$(NoWarn);CS8073;CA2200;CS8632</NoWarn>
```

`Docxodus.Tests/Docxodus.Tests.csproj`:
```xml
<TreatWarningsAsErrors>false</TreatWarningsAsErrors>
<NoWarn>$(NoWarn);xUnit1012;xUnit2020</NoWarn>
```

## Current baseline

Measure with a clean rebuild — an incremental build reports zero because nothing recompiles:

```bash
dotnet build Docxodus/Docxodus.csproj --no-incremental
dotnet build Docxodus.Tests/Docxodus.Tests.csproj --no-incremental
```

The library builds with **134 warnings**, the test project with **808**. Don't add to either
baseline. The mix is dominated by StyleCop file-header rules (`SA1633`, `SA1636`) and
using-directive ordering (`SA1206`); the genuinely interesting remainder is a handful of
nullable-flow warnings (`CS8600`, `CS8604`) and `CA2022` (ignored `Stream.Read` return value).

## Where the debt lives

The library is `<Nullable>enable</Nullable>` (issue #13); the exception is the inherited core,
where legacy files carry an explicit `#nullable disable` header. List them with:

```bash
grep -l "^#nullable disable" Docxodus/*.cs
```

When substantially refactoring one of those files, consider removing its header and fixing that
file's warnings. `CS8632` stays in `NoWarn` deliberately: with the project context enabled it can
only fire inside the opted-out files, where some inert `?` annotations are kept for the day each
file migrates.

## Goal

Once a file's warnings are fixed:
1. Remove its `#nullable disable` header.
2. When the whole project is clean, drop `<TreatWarningsAsErrors>false</TreatWarningsAsErrors>`
   and the `<NoWarn>` entries.
3. Let `Directory.Build.props` handle warning-as-error for Release builds.
