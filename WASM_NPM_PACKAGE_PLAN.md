# Docxodus WASM NPM Package Plan

## Overview

Create a new npm package that wraps Docxodus functionality (document comparison and HTML conversion) as a WebAssembly library callable from JavaScript/TypeScript/React applications.

**Target Repository:** New repo (e.g., `docxodus-wasm` or `redlines-js`)
**Core Features to Expose:**
1. Document comparison (redlining) - `WmlComparer.Compare()`
2. DOCX to HTML conversion - `WmlToHtmlConverter.ConvertToHtml()`

---

## Phase 1: Project Setup & WASM Compatibility Verification

### 1.1 Create New Repository Structure

```
docxodus-wasm/
├── dotnet/                          # .NET WASM project
│   ├── DocxodusWasm/               # WASM wrapper project
│   │   ├── DocxodusWasm.csproj
│   │   ├── Program.cs              # Entry point
│   │   ├── DocumentConverter.cs    # [JSExport] methods for HTML conversion
│   │   ├── DocumentComparer.cs     # [JSExport] methods for comparison
│   │   └── main.js                 # WASM bootstrap
│   └── DocxodusWasm.sln
├── npm/                             # NPM package
│   ├── src/
│   │   ├── index.ts                # Main entry point
│   │   ├── types.ts                # TypeScript definitions
│   │   ├── react.ts                # React hooks
│   │   └── loader.ts               # WASM loader utilities
│   ├── dist/                       # Built output
│   │   ├── wasm/                   # WASM files from .NET build
│   │   └── ...                     # Compiled TS
│   ├── package.json
│   ├── tsconfig.json
│   ├── vite.config.ts
│   └── README.md
├── examples/                        # Example applications
│   ├── react-demo/
│   └── vanilla-js-demo/
├── scripts/                         # Build scripts
│   ├── build-wasm.sh
│   └── copy-wasm.js
├── .github/
│   └── workflows/
│       └── publish.yml             # CI/CD for npm publish
├── README.md
└── LICENSE
```

### 1.2 Install Required .NET Workloads

```bash
# Install the experimental WASM workload
dotnet workload install wasm-experimental

# Verify installation
dotnet workload list
```

### 1.3 Create WASM Project Configuration

**DocxodusWasm.csproj:**
```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
    <RuntimeIdentifier>browser-wasm</RuntimeIdentifier>
    <OutputType>Exe</OutputType>
    <AllowUnsafeBlocks>true</AllowUnsafeBlocks>
    <WasmMainJSPath>main.js</WasmMainJSPath>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>

    <!-- Size optimization -->
    <PublishTrimmed>true</PublishTrimmed>
    <TrimMode>full</TrimMode>
    <InvariantGlobalization>true</InvariantGlobalization>

    <!-- Disable unnecessary features -->
    <DebuggerSupport>false</DebuggerSupport>
    <EventSourceSupport>false</EventSourceSupport>
    <UseSystemResourceKeys>true</UseSystemResourceKeys>
  </PropertyGroup>

  <ItemGroup>
    <!-- Reference Docxodus via NuGet or local project -->
    <PackageReference Include="Docxodus" Version="1.0.0" />
    <!-- OR for local development: -->
    <!-- <ProjectReference Include="../../Redliner/Docxodus/Docxodus.csproj" /> -->

    <!-- WASM-specific SkiaSharp -->
    <PackageReference Include="SkiaSharp.Views.Blazor" Version="3.119.1" />
    <PackageReference Include="SkiaSharp.NativeAssets.WebAssembly" Version="3.119.1" />
  </ItemGroup>

  <ItemGroup>
    <WasmExtraFilesToDeploy Include="main.js" />
  </ItemGroup>
</Project>
```

---

## Phase 2: Implement WASM Wrapper Layer

### 2.1 Create JSExport Entry Points

**DocumentConverter.cs** - HTML Conversion:
```csharp
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using Docxodus;
using DocumentFormat.OpenXml.Packaging;

[SupportedOSPlatform("browser")]
public partial class DocumentConverter
{
    [JSExport]
    public static string ConvertDocxToHtml(byte[] docxBytes)
    {
        if (docxBytes == null || docxBytes.Length == 0)
            return "{\"error\": \"No document data provided\"}";

        try
        {
            var doc = new WmlDocument("document.docx", docxBytes);
            var settings = new WmlToHtmlConverterSettings
            {
                PageTitle = "Converted Document",
                FabricateCssClasses = true,
                CssClassPrefix = "docx-"
            };

            using var ms = new MemoryStream(docxBytes);
            using var wordDoc = WordprocessingDocument.Open(ms, false);
            var html = WmlToHtmlConverter.ConvertToHtml(wordDoc, settings);
            return html.ToString();
        }
        catch (Exception ex)
        {
            return $"{{\"error\": \"{EscapeJson(ex.Message)}\"}}";
        }
    }

    [JSExport]
    public static string ConvertDocxToHtmlWithOptions(
        byte[] docxBytes,
        string pageTitle,
        string cssPrefix,
        bool fabricateClasses,
        string additionalCss)
    {
        try
        {
            var settings = new WmlToHtmlConverterSettings
            {
                PageTitle = pageTitle ?? "Document",
                CssClassPrefix = cssPrefix ?? "docx-",
                FabricateCssClasses = fabricateClasses,
                AdditionalCss = additionalCss ?? ""
            };

            using var ms = new MemoryStream(docxBytes);
            using var wordDoc = WordprocessingDocument.Open(ms, false);
            var html = WmlToHtmlConverter.ConvertToHtml(wordDoc, settings);
            return html.ToString();
        }
        catch (Exception ex)
        {
            return $"{{\"error\": \"{EscapeJson(ex.Message)}\"}}";
        }
    }

    private static string EscapeJson(string s) =>
        s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n");
}
```

**DocumentComparer.cs** - Document Comparison:
```csharp
using System.Runtime.InteropServices.JavaScript;
using System.Runtime.Versioning;
using Docxodus;

[SupportedOSPlatform("browser")]
public partial class DocumentComparer
{
    [JSExport]
    public static byte[] CompareDocuments(
        byte[] originalBytes,
        byte[] modifiedBytes,
        string authorName)
    {
        try
        {
            var original = new WmlDocument("original.docx", originalBytes);
            var modified = new WmlDocument("modified.docx", modifiedBytes);

            var settings = new WmlComparerSettings
            {
                AuthorForRevisions = authorName ?? "Docxodus",
                DetailThreshold = 0.15
            };

            var result = WmlComparer.Compare(original, modified, settings);
            return result.DocumentByteArray;
        }
        catch (Exception ex)
        {
            // Return empty array on error - caller should check length
            Console.WriteLine($"Comparison error: {ex.Message}");
            return Array.Empty<byte>();
        }
    }

    [JSExport]
    public static string CompareDocumentsToHtml(
        byte[] originalBytes,
        byte[] modifiedBytes,
        string authorName)
    {
        try
        {
            var original = new WmlDocument("original.docx", originalBytes);
            var modified = new WmlDocument("modified.docx", modifiedBytes);

            var settings = new WmlComparerSettings
            {
                AuthorForRevisions = authorName ?? "Docxodus"
            };

            var result = WmlComparer.Compare(original, modified, settings);

            // Convert comparison result to HTML
            var htmlSettings = new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = true,
                CssClassPrefix = "redline-"
            };

            using var ms = new MemoryStream(result.DocumentByteArray);
            using var wordDoc = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
            var html = WmlToHtmlConverter.ConvertToHtml(wordDoc, htmlSettings);
            return html.ToString();
        }
        catch (Exception ex)
        {
            return $"{{\"error\": \"{ex.Message.Replace("\"", "\\\"")}\"}}";
        }
    }

    [JSExport]
    public static string GetRevisionsJson(byte[] comparedDocBytes)
    {
        try
        {
            var doc = new WmlDocument("compared.docx", comparedDocBytes);
            var revisions = WmlComparer.GetRevisions(doc, new WmlComparerSettings());

            var json = System.Text.Json.JsonSerializer.Serialize(revisions.Select(r => new
            {
                Author = r.Author,
                Date = r.Date,
                RevisionType = r.RevisionType.ToString(),
                Text = r.Text
            }));

            return json;
        }
        catch (Exception ex)
        {
            return $"{{\"error\": \"{ex.Message.Replace("\"", "\\\"")}\"}}";
        }
    }
}
```

**Program.cs** - Entry Point:
```csharp
using System.Runtime.InteropServices.JavaScript;

Console.WriteLine("Docxodus WASM Library Initialized");

// Keep the runtime alive
await Task.Delay(-1);
```

### 2.2 Create main.js Bootstrap

```javascript
// main.js - WASM bootstrap file
export async function createDocxodusRuntime() {
    const { dotnet } = await import('./dotnet.js');

    const runtime = await dotnet
        .withDiagnosticTracing(false)
        .create();

    const config = runtime.getConfig();
    const exports = await runtime.getAssemblyExports(config.mainAssemblyName);

    return {
        DocumentConverter: exports.DocumentConverter,
        DocumentComparer: exports.DocumentComparer,
        runtime
    };
}
```

---

## Phase 3: Create NPM Package Wrapper

### 3.1 TypeScript Type Definitions

**src/types.ts:**
```typescript
export interface ConversionOptions {
  pageTitle?: string;
  cssPrefix?: string;
  fabricateClasses?: boolean;
  additionalCss?: string;
}

export interface ComparisonOptions {
  authorName?: string;
}

export interface Revision {
  Author: string;
  Date: string;
  RevisionType: string;
  Text: string;
}

export interface DocxodusExports {
  DocumentConverter: {
    ConvertDocxToHtml(docxBytes: Uint8Array): string;
    ConvertDocxToHtmlWithOptions(
      docxBytes: Uint8Array,
      pageTitle: string,
      cssPrefix: string,
      fabricateClasses: boolean,
      additionalCss: string
    ): string;
  };
  DocumentComparer: {
    CompareDocuments(
      originalBytes: Uint8Array,
      modifiedBytes: Uint8Array,
      authorName: string
    ): Uint8Array;
    CompareDocumentsToHtml(
      originalBytes: Uint8Array,
      modifiedBytes: Uint8Array,
      authorName: string
    ): string;
    GetRevisionsJson(comparedDocBytes: Uint8Array): string;
  };
}

export interface InitOptions {
  wasmPath?: string;
  diagnosticTracing?: boolean;
}
```

### 3.2 Main Entry Point

**src/index.ts:**
```typescript
import type { DocxodusExports, InitOptions, ConversionOptions, ComparisonOptions, Revision } from './types';

export type { ConversionOptions, ComparisonOptions, Revision, InitOptions };

let docxodusInstance: DocxodusExports | null = null;
let initPromise: Promise<DocxodusExports> | null = null;

export async function init(options: InitOptions = {}): Promise<void> {
  if (docxodusInstance) return;
  if (initPromise) {
    await initPromise;
    return;
  }

  const wasmPath = options.wasmPath || '/wasm/dotnet.js';

  initPromise = (async () => {
    const module = await import(/* @vite-ignore */ wasmPath);
    const runtime = await module.dotnet
      .withDiagnosticTracing(options.diagnosticTracing ?? false)
      .create();

    const config = runtime.getConfig();
    const exports = await runtime.getAssemblyExports(config.mainAssemblyName);
    docxodusInstance = exports as DocxodusExports;
    return docxodusInstance;
  })();

  await initPromise;
}

function getExports(): DocxodusExports {
  if (!docxodusInstance) {
    throw new Error('Docxodus not initialized. Call init() first.');
  }
  return docxodusInstance;
}

// === HTML Conversion APIs ===

export async function convertDocxToHtml(
  file: File | Uint8Array,
  options: ConversionOptions = {}
): Promise<string> {
  const exports = getExports();
  const bytes = file instanceof File
    ? new Uint8Array(await file.arrayBuffer())
    : file;

  if (Object.keys(options).length === 0) {
    return exports.DocumentConverter.ConvertDocxToHtml(bytes);
  }

  return exports.DocumentConverter.ConvertDocxToHtmlWithOptions(
    bytes,
    options.pageTitle || 'Document',
    options.cssPrefix || 'docx-',
    options.fabricateClasses ?? true,
    options.additionalCss || ''
  );
}

// === Document Comparison APIs ===

export async function compareDocuments(
  original: File | Uint8Array,
  modified: File | Uint8Array,
  options: ComparisonOptions = {}
): Promise<Uint8Array> {
  const exports = getExports();

  const originalBytes = original instanceof File
    ? new Uint8Array(await original.arrayBuffer())
    : original;
  const modifiedBytes = modified instanceof File
    ? new Uint8Array(await modified.arrayBuffer())
    : modified;

  return exports.DocumentComparer.CompareDocuments(
    originalBytes,
    modifiedBytes,
    options.authorName || 'Docxodus'
  );
}

export async function compareDocumentsToHtml(
  original: File | Uint8Array,
  modified: File | Uint8Array,
  options: ComparisonOptions = {}
): Promise<string> {
  const exports = getExports();

  const originalBytes = original instanceof File
    ? new Uint8Array(await original.arrayBuffer())
    : original;
  const modifiedBytes = modified instanceof File
    ? new Uint8Array(await modified.arrayBuffer())
    : modified;

  return exports.DocumentComparer.CompareDocumentsToHtml(
    originalBytes,
    modifiedBytes,
    options.authorName || 'Docxodus'
  );
}

export async function getRevisions(comparedDoc: Uint8Array): Promise<Revision[]> {
  const exports = getExports();
  const json = exports.DocumentComparer.GetRevisionsJson(comparedDoc);

  if (json.startsWith('{"error"')) {
    throw new Error(JSON.parse(json).error);
  }

  return JSON.parse(json);
}

export function isInitialized(): boolean {
  return docxodusInstance !== null;
}
```

### 3.3 React Hook

**src/react.ts:**
```typescript
import { useState, useEffect, useCallback } from 'react';
import { init, convertDocxToHtml, compareDocuments, compareDocumentsToHtml, isInitialized } from './index';
import type { InitOptions, ConversionOptions, ComparisonOptions } from './types';

export interface UseDocxodusResult {
  convertToHtml: ((file: File | Uint8Array, options?: ConversionOptions) => Promise<string>) | null;
  compare: ((original: File | Uint8Array, modified: File | Uint8Array, options?: ComparisonOptions) => Promise<Uint8Array>) | null;
  compareToHtml: ((original: File | Uint8Array, modified: File | Uint8Array, options?: ComparisonOptions) => Promise<string>) | null;
  loading: boolean;
  error: Error | null;
  initialized: boolean;
}

export function useDocxodus(options: InitOptions = {}): UseDocxodusResult {
  const [loading, setLoading] = useState(!isInitialized());
  const [error, setError] = useState<Error | null>(null);
  const [initialized, setInitialized] = useState(isInitialized());

  useEffect(() => {
    if (initialized) return;

    let cancelled = false;

    const initWasm = async () => {
      try {
        setLoading(true);
        await init(options);
        if (!cancelled) {
          setInitialized(true);
          setError(null);
        }
      } catch (err) {
        if (!cancelled) {
          setError(err instanceof Error ? err : new Error('Init failed'));
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    };

    initWasm();
    return () => { cancelled = true; };
  }, [initialized]);

  const convertToHtml = useCallback(
    async (file: File | Uint8Array, opts?: ConversionOptions) => {
      if (!initialized) throw new Error('Not initialized');
      return convertDocxToHtml(file, opts);
    },
    [initialized]
  );

  const compare = useCallback(
    async (original: File | Uint8Array, modified: File | Uint8Array, opts?: ComparisonOptions) => {
      if (!initialized) throw new Error('Not initialized');
      return compareDocuments(original, modified, opts);
    },
    [initialized]
  );

  const compareToHtml = useCallback(
    async (original: File | Uint8Array, modified: File | Uint8Array, opts?: ComparisonOptions) => {
      if (!initialized) throw new Error('Not initialized');
      return compareDocumentsToHtml(original, modified, opts);
    },
    [initialized]
  );

  return {
    convertToHtml: initialized ? convertToHtml : null,
    compare: initialized ? compare : null,
    compareToHtml: initialized ? compareToHtml : null,
    loading,
    error,
    initialized
  };
}
```

### 3.4 Package Configuration

**package.json:**
```json
{
  "name": "@redlines/docxodus",
  "version": "1.0.0",
  "description": "Client-side DOCX comparison and HTML conversion using WebAssembly",
  "type": "module",
  "main": "./dist/index.js",
  "module": "./dist/index.js",
  "types": "./dist/index.d.ts",
  "exports": {
    ".": {
      "types": "./dist/index.d.ts",
      "import": "./dist/index.js"
    },
    "./react": {
      "types": "./dist/react.d.ts",
      "import": "./dist/react.js"
    },
    "./wasm/*": "./dist/wasm/*"
  },
  "files": [
    "dist",
    "README.md",
    "LICENSE"
  ],
  "scripts": {
    "build:wasm": "cd ../dotnet && dotnet publish DocxodusWasm -c Release",
    "copy:wasm": "node scripts/copy-wasm.js",
    "build:ts": "tsc && vite build",
    "build": "npm run build:wasm && npm run copy:wasm && npm run build:ts",
    "prepublishOnly": "npm run build"
  },
  "peerDependencies": {
    "react": ">=16.8.0"
  },
  "peerDependenciesMeta": {
    "react": {
      "optional": true
    }
  },
  "devDependencies": {
    "@types/node": "^20.0.0",
    "@types/react": "^18.0.0",
    "react": "^18.0.0",
    "typescript": "^5.0.0",
    "vite": "^5.0.0"
  },
  "keywords": [
    "docx",
    "word",
    "openxml",
    "html",
    "comparison",
    "redline",
    "track-changes",
    "wasm",
    "webassembly"
  ],
  "repository": {
    "type": "git",
    "url": "https://github.com/JSv4/docxodus-wasm"
  },
  "license": "MIT"
}
```

---

## Phase 4: Build & Distribution Pipeline

### 4.1 Build Scripts

**scripts/build-wasm.sh:**
```bash
#!/bin/bash
set -e

cd "$(dirname "$0")/../dotnet"

echo "Building Docxodus WASM..."
dotnet publish DocxodusWasm -c Release \
  /p:PublishTrimmed=true \
  /p:TrimMode=full \
  /p:InvariantGlobalization=true

echo "WASM build complete!"
```

**scripts/copy-wasm.js:**
```javascript
import fs from 'fs/promises';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const sourceDir = path.join(__dirname, '../dotnet/DocxodusWasm/bin/Release/net8.0/browser-wasm/AppBundle');
const destDir = path.join(__dirname, '../npm/dist/wasm');

async function copyWasmFiles() {
  await fs.mkdir(destDir, { recursive: true });

  const files = ['dotnet.js', 'dotnet.native.wasm', 'dotnet.native.js'];

  for (const file of files) {
    await fs.copyFile(path.join(sourceDir, file), path.join(destDir, file));
    console.log(`Copied ${file}`);
  }

  // Copy managed assemblies
  const managedSrc = path.join(sourceDir, 'managed');
  const managedDest = path.join(destDir, 'managed');
  await fs.mkdir(managedDest, { recursive: true });

  const dlls = await fs.readdir(managedSrc);
  for (const dll of dlls.filter(f => f.endsWith('.dll'))) {
    await fs.copyFile(path.join(managedSrc, dll), path.join(managedDest, dll));
  }

  console.log(`Copied ${dlls.length} managed assemblies`);
}

copyWasmFiles().catch(console.error);
```

### 4.2 GitHub Actions CI/CD

**.github/workflows/publish.yml:**
```yaml
name: Build and Publish

on:
  release:
    types: [created]
  push:
    branches: [main]
  pull_request:
    branches: [main]

jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - uses: actions/setup-dotnet@v4
        with:
          dotnet-version: '8.0.x'

      - uses: actions/setup-node@v4
        with:
          node-version: '20'
          registry-url: 'https://registry.npmjs.org'

      - name: Install .NET WASM workload
        run: dotnet workload install wasm-experimental

      - name: Build WASM
        run: npm run build:wasm
        working-directory: npm

      - name: Copy WASM files
        run: npm run copy:wasm
        working-directory: npm

      - name: Install npm dependencies
        run: npm ci
        working-directory: npm

      - name: Build TypeScript
        run: npm run build:ts
        working-directory: npm

      - name: Publish to npm
        if: github.event_name == 'release'
        run: npm publish --access public
        working-directory: npm
        env:
          NODE_AUTH_TOKEN: ${{ secrets.NPM_TOKEN }}
```

---

## Phase 5: Testing & Examples

### 5.1 Create Example React App

**examples/react-demo/src/App.tsx:**
```tsx
import React, { useState } from 'react';
import { useDocxodus } from '@redlines/docxodus/react';
import DOMPurify from 'dompurify';

function App() {
  const { convertToHtml, compareToHtml, loading, error } = useDocxodus({
    wasmPath: '/wasm/dotnet.js'
  });

  const [html, setHtml] = useState('');
  const [mode, setMode] = useState<'convert' | 'compare'>('convert');

  const handleConvert = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file || !convertToHtml) return;

    const result = await convertToHtml(file);
    setHtml(DOMPurify.sanitize(result));
  };

  const handleCompare = async (files: FileList) => {
    if (files.length < 2 || !compareToHtml) return;

    const result = await compareToHtml(files[0], files[1], {
      authorName: 'Demo User'
    });
    setHtml(DOMPurify.sanitize(result));
  };

  if (loading) return <div>Loading Docxodus WASM...</div>;
  if (error) return <div>Error: {error.message}</div>;

  return (
    <div className="app">
      <h1>Docxodus WASM Demo</h1>

      <div className="controls">
        <button onClick={() => setMode('convert')}>Convert</button>
        <button onClick={() => setMode('compare')}>Compare</button>
      </div>

      {mode === 'convert' && (
        <input type="file" accept=".docx" onChange={handleConvert} />
      )}

      {mode === 'compare' && (
        <input
          type="file"
          accept=".docx"
          multiple
          onChange={(e) => e.target.files && handleCompare(e.target.files)}
        />
      )}

      {html && (
        <div
          className="document-content"
          dangerouslySetInnerHTML={{ __html: html }}
        />
      )}
    </div>
  );
}

export default App;
```

---

## Phase 6: Documentation

### 6.1 README Structure

1. **Installation** - npm install instructions
2. **Setup** - Copy WASM files to public directory
3. **Quick Start** - Basic usage examples
4. **API Reference**
   - `init(options)` - Initialize WASM runtime
   - `convertDocxToHtml(file, options)` - Convert DOCX to HTML
   - `compareDocuments(original, modified, options)` - Compare and get redlined DOCX
   - `compareDocumentsToHtml(original, modified, options)` - Compare and get HTML
   - `getRevisions(comparedDoc)` - Extract revision list
5. **React Integration** - useDocxodus hook
6. **Bundler Configuration** - Vite, Webpack, Next.js setup
7. **Bundle Size** - Expected sizes (3-8 MB compressed)
8. **Browser Support** - Chrome 91+, Firefox 89+, Safari 15+, Edge 91+

---

## Implementation Checklist

### Phase 1: Setup
- [ ] Create new GitHub repository
- [ ] Initialize directory structure
- [ ] Install wasm-experimental workload
- [ ] Create .NET WASM project with proper config
- [ ] Test that Docxodus compiles to WASM (verify SkiaSharp WASM compatibility)

### Phase 2: WASM Layer
- [ ] Implement DocumentConverter.cs with JSExport
- [ ] Implement DocumentComparer.cs with JSExport
- [ ] Create main.js bootstrap
- [ ] Test basic WASM functionality in browser

### Phase 3: NPM Package
- [ ] Create TypeScript types
- [ ] Implement main entry point (index.ts)
- [ ] Implement React hook (react.ts)
- [ ] Configure package.json with proper exports
- [ ] Create build scripts

### Phase 4: Build Pipeline
- [ ] Create build-wasm.sh script
- [ ] Create copy-wasm.js script
- [ ] Setup GitHub Actions workflow
- [ ] Test full build pipeline

### Phase 5: Testing
- [ ] Create React demo app
- [ ] Create vanilla JS demo
- [ ] Test with various DOCX files
- [ ] Test comparison with tracked changes

### Phase 6: Documentation & Release
- [ ] Write comprehensive README
- [ ] Create API documentation
- [ ] Publish to npm
- [ ] Create GitHub release

---

## Known Risks & Mitigations

| Risk | Mitigation |
|------|------------|
| SkiaSharp WASM compatibility issues | Test early; may need to conditionally disable image processing |
| Large bundle size (5-15MB) | Enable trimming, InvariantGlobalization, Brotli compression |
| Memory limits in browser | Stream processing, chunked file handling for large docs |
| OpenXML ZIP memory explosion | Document size limits, lazy loading strategies |

---

## Estimated Bundle Sizes

| Component | Uncompressed | Brotli Compressed |
|-----------|--------------|-------------------|
| dotnet.native.wasm | ~4-6 MB | ~1.5-2 MB |
| Managed assemblies | ~2-4 MB | ~0.5-1 MB |
| SkiaSharp WASM | ~2-3 MB | ~0.8-1 MB |
| **Total** | **8-13 MB** | **3-4 MB** |

---

## Alternative Approaches Considered

1. **Blazor Component** - Rejected: Too heavy, requires Blazor runtime
2. **Server-side API** - Rejected: User wants client-side processing
3. **Emscripten direct** - Rejected: .NET 8 WASM is more maintainable
4. **WebWorker isolation** - Could add later: Would prevent UI blocking

---

## Next Steps

1. **Verify SkiaSharp WASM works** - Create minimal test project
2. **Publish Docxodus to NuGet** - Makes referencing easier
3. **Start with HTML conversion** - Simpler API, fewer dependencies
4. **Add comparison second** - More complex, validate architecture first
