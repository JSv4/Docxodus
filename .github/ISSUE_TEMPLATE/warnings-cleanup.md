---
name: Technical Debt - Warnings Cleanup
about: Track warnings that need to be resolved
title: "Fix compiler warnings and re-enable TreatWarningsAsErrors"
labels: technical-debt, good-first-issue
---

## Summary

The `OpenXmlPowerTools` library currently has `TreatWarningsAsErrors` disabled and several warning codes suppressed due to legacy code issues. These should be fixed to improve code quality.

## Current Suppressions

In `OpenXmlPowerTools/OpenXmlPowerTools.csproj`:
```xml
<TreatWarningsAsErrors>false</TreatWarningsAsErrors>
<NoWarn>$(NoWarn);CS8073;CA2200</NoWarn>
```

## Warnings to Fix

### CS8632 - Nullable annotations without context
Files with `?` nullable annotations but project has `<Nullable>disable</Nullable>`:
- `PtOpenXmlUtil.cs`
- `WmlToHtmlConverter.cs`
- `WmlComparer.cs`

**Fix**: Either enable nullable for the project and fix all nullable warnings, or remove the `?` annotations.

### CS8073 - Comparison always true/false
`PresentationBuilder.cs` has comparisons of `IdPartPair` (a struct) to null, which is always true.

**Fix**: Remove unnecessary null checks or change the logic.

### CA2200 - Re-throwing changes stack information
Files using `throw e;` instead of `throw;`:
- `DocumentBuilder.cs`
- `PresentationBuilder.cs`
- `SpreadsheetWriter.cs`

**Fix**: Change `throw e;` to `throw;` to preserve stack trace.

## Goal

Once all warnings are fixed:
1. Remove `<TreatWarningsAsErrors>false</TreatWarningsAsErrors>`
2. Remove `<NoWarn>` entries
3. Let `Directory.Build.props` handle warning-as-error for Release builds

## Related

Part of the .NET 8 migration effort.
