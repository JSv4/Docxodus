#!/bin/bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
WASM_PROJECT="$REPO_ROOT/wasm/DocxodusWasm"
NPM_DIR="$REPO_ROOT/npm"
WASM_DIST="$NPM_DIR/dist/wasm"

echo "Building Docxodus WASM..."
echo "Project: $WASM_PROJECT"
echo "Output: $WASM_DIST"

# Build in Release mode for smaller size
cd "$WASM_PROJECT"
dotnet build -c Release

# Source AppBundle location
APPBUNDLE="$WASM_PROJECT/bin/Release/net8.0/browser-wasm/AppBundle"

if [ ! -d "$APPBUNDLE" ]; then
    echo "Error: AppBundle not found at $APPBUNDLE"
    echo "Trying Debug build..."
    APPBUNDLE="$WASM_PROJECT/bin/Debug/net8.0/browser-wasm/AppBundle"
fi

if [ ! -d "$APPBUNDLE" ]; then
    echo "Error: AppBundle not found"
    exit 1
fi

echo "AppBundle found at: $APPBUNDLE"

# Clean and create destination
rm -rf "$WASM_DIST"
mkdir -p "$WASM_DIST"

# Copy the _framework directory (contains all WASM and JS files)
echo "Copying _framework..."
cp -r "$APPBUNDLE/_framework" "$WASM_DIST/"

# Copy main.js
echo "Copying main.js..."
cp "$WASM_PROJECT/main.js" "$WASM_DIST/"

# Copy index.html for testing
cp "$WASM_PROJECT/index.html" "$WASM_DIST/"

# Report sizes
echo ""
echo "Build complete! File sizes:"
echo "----------------------------"
du -sh "$WASM_DIST/_framework/"*.wasm 2>/dev/null | head -10
echo ""
echo "Total WASM directory size:"
du -sh "$WASM_DIST"
