---
name: Technical Debt - Warnings Cleanup
about: Track warnings that need to be resolved
title: "Fix test-project warnings and re-enable TreatWarningsAsErrors"
labels: technical-debt, good-first-issue
---

## Summary

`Directory.Build.props` sets `TreatWarningsAsErrors=true` for Release. The core library builds
warning-free and inherits it. `Docxodus.Tests.csproj` still overrides it to `false`; the goal is to
clear that override too.

## Current suppressions

`Docxodus.Tests/Docxodus.Tests.csproj`:
```xml
<TreatWarningsAsErrors>false</TreatWarningsAsErrors>
<NoWarn>$(NoWarn);xUnit1012;xUnit2020</NoWarn>
```

## Current baseline

Measure with a clean rebuild — an incremental build reports zero because nothing recompiles:

```bash
dotnet build Docxodus.Tests/Docxodus.Tests.csproj --no-incremental
```

The test project builds with **434 warnings**, almost all nullable-flow warnings in test code
(`CS8602`, `CS8600`, `CS8604`), plus a few StyleCop layout rules (`SA1010`, `SA1134`, `SA1211`).
Don't add to that number.

## Goal

Fix the warnings file by file; when the project is clean, drop
`<TreatWarningsAsErrors>false</TreatWarningsAsErrors>` and let `Directory.Build.props` handle
warning-as-error for Release builds.
