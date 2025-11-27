# Docxodus WASM NPM Package - Progress Tracker

**Started:** 2025-11-27
**Status:** In Progress

---

## Current Phase: Phase 1 - Setup & WASM Compatibility Verification

---

## Task Checklist

### Phase 1: Setup & Compatibility
- [x] **Verify SkiaSharp WASM compatibility** ← DONE (pending browser test)
  - Created minimal test project at `/home/jman/Code/docxodus-wasm-test/`
  - SkiaSharp.NativeAssets.WebAssembly compiles to WASM
  - WASM file size: 6.1MB (includes SkiaSharp native library)
  - **IMPORTANT**: Requires manual `NativeFileReference` for standalone WASM
  - Test server: `http://localhost:8080`
- [x] Create WASM wrapper structure in Redliner repo
- [x] Setup .NET WASM project with proper configuration
- [x] Test Docxodus compiles to WASM with all dependencies ← DONE!

### Phase 2: WASM Wrapper Layer
- [x] Implement DocumentConverter.cs with JSExport (HTML conversion)
- [x] Implement DocumentComparer.cs with JSExport (document comparison)
- [x] Create main.js bootstrap
- [x] Test basic WASM functionality in browser ✅ WORKING!

### Phase 3: NPM Package
- [x] Create TypeScript type definitions (src/types.ts)
- [x] Implement main entry point (src/index.ts)
- [x] Implement React hooks (src/react.ts)
- [x] Configure package.json with proper exports

### Phase 4: Build Pipeline
- [x] Create build-wasm.sh script
- [x] Test full build pipeline (Release build: 37MB total)

### Phase 5: Testing & Examples
- [x] Create Playwright test suite using existing .NET fixtures
- [x] 32 tests passing (HTML conversion, comparison, tracked changes)
- [ ] Create React demo app (optional)

### Phase 6: Documentation & Release
- [x] Write comprehensive README
- [x] Update CHANGELOG
- [ ] Publish to npm
- [ ] Create GitHub release

---

## Session Notes

### 2025-11-27 - Initial Planning

**Completed:**
1. Explored Redliner/Docxodus codebase structure
2. Identified key APIs to expose:
   - `WmlComparer.Compare()` - document comparison
   - `WmlToHtmlConverter.ConvertToHtml()` - HTML conversion
3. Researched WASM compatibility:
   - SkiaSharp has official WASM support: `SkiaSharp.NativeAssets.WebAssembly 3.119.1`
   - DocumentFormat.OpenXml 3.x targets .NET 8.0 (should work)
4. Created comprehensive plan in `WASM_NPM_PACKAGE_PLAN.md`

**Key Findings:**
- Docxodus is .NET 8.0 with OpenXML SDK 3.2.0
- Uses SkiaSharp 2.88.9 for cross-platform graphics
- Main risk: SkiaSharp WASM compatibility with image handling

**Next Steps:**
- Create minimal WASM test project with SkiaSharp
- Verify it compiles and runs in browser
- Then proceed to full repo setup

### 2025-11-27 - SkiaSharp WASM Compatibility Test

**Completed:**
1. Installed .NET WASM workloads:
   - `wasm-experimental` - For browser-wasm RuntimeIdentifier
   - `wasm-tools` - For Emscripten toolchain
2. Created test project at `/home/jman/Code/docxodus-wasm-test/SkiaSharpWasmTest/`
3. Configured csproj for browser-wasm:
   - `<RuntimeIdentifier>browser-wasm</RuntimeIdentifier>`
   - `<AllowUnsafeBlocks>true</AllowUnsafeBlocks>`
   - `<InvariantGlobalization>true</InvariantGlobalization>`
4. Added JSExport test methods for SkiaSharp operations
5. Successfully compiled with SkiaSharp native library linked

**CRITICAL DISCOVERY:**
The `SkiaSharp.NativeAssets.WebAssembly` package does NOT auto-link for standalone .NET 8 WASM.
It only auto-links for Uno Platform projects (checks for `$(IsUnoHead)`).

**Solution:** Must manually add NativeFileReference in csproj:
```xml
<ItemGroup>
  <NativeFileReference Include="$(NuGetPackageRoot)skiasharp.nativeassets.webassembly/2.88.9/build/netstandard1.0/libSkiaSharp.a/3.1.34/simd,st/libSkiaSharp.a" />
</ItemGroup>
```

Choose library variant based on:
- Emscripten version (3.1.34 matches .NET 8.0.22)
- Threading mode: `st` (single-threaded) or `mt` (multi-threaded)
- SIMD support: `simd` or no SIMD

**Output:**
- dotnet.native.wasm: 6.1MB (with SkiaSharp native)
- All managed assemblies in _framework/ folder
- Test server running at http://localhost:8080

**Next:**
- Manual browser test to verify runtime execution
- Create WASM wrapper in Redliner repo

---

## Technical Details

### Target Bundle Sizes (Estimated)
| Component | Uncompressed | Brotli |
|-----------|--------------|--------|
| dotnet.native.wasm | 4-6 MB | 1.5-2 MB |
| Managed assemblies | 2-4 MB | 0.5-1 MB |
| SkiaSharp WASM | 2-3 MB | 0.8-1 MB |
| **Total** | **8-13 MB** | **3-4 MB** |

### APIs to Expose
```typescript
// HTML Conversion
convertDocxToHtml(file: File | Uint8Array, options?: ConversionOptions): Promise<string>

// Document Comparison
compareDocuments(original, modified, options?): Promise<Uint8Array>  // Returns redlined DOCX
compareDocumentsToHtml(original, modified, options?): Promise<string>  // Returns redlined HTML
getRevisions(comparedDoc: Uint8Array): Promise<Revision[]>
```

### Repository Structure (in Redliner repo)
```
Redliner/
├── Docxodus/                # Existing library
├── Docxodus.Tests/          # Existing tests
├── tools/                   # Existing CLI tools
├── wasm/                    # NEW: WASM wrapper project
│   └── DocxodusWasm/
│       ├── DocxodusWasm.csproj
│       ├── Program.cs
│       ├── DocumentConverter.cs
│       └── DocumentComparer.cs
├── npm/                     # NEW: NPM package
│   ├── src/
│   │   ├── index.ts
│   │   ├── types.ts
│   │   └── react.ts
│   ├── dist/wasm/           # Built WASM files
│   └── package.json
└── examples/
    └── react-demo/
```

---

### 2025-11-27 - JSON Serialization Fix

**Issue:**
Browser showed error: "Reflection-based serialization has been disabled for this application"
This happens because WASM builds are trimmed and reflection-based JSON serialization is disabled.

**Solution:**
Implemented System.Text.Json source generators (AOT-safe serialization):
1. Created `JsonContext.cs` with `DocxodusJsonContext : JsonSerializerContext`
2. Added `[JsonSerializable]` attributes for all DTO types
3. Created typed DTO classes: `ErrorResponse`, `VersionInfo`, `RevisionsResponse`, `RevisionInfo`
4. Updated all `JsonSerializer.Serialize()` calls to use source generator context

**Result:**
Build succeeded. Server running at http://localhost:8081 for browser testing.

---

## Blockers & Issues

_None yet_

---

## References

- [SkiaSharp.NativeAssets.WebAssembly](https://www.nuget.org/packages/SkiaSharp.NativeAssets.WebAssembly/)
- [DocumentFormat.OpenXml 3.3.0](https://www.nuget.org/packages/DocumentFormat.OpenXml/)
- [.NET 8 WASM Browser](https://learn.microsoft.com/en-us/aspnet/core/blazor/webassembly-build-tools-and-aot)
- Main plan: `WASM_NPM_PACKAGE_PLAN.md`
