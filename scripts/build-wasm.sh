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

# Publish in Release mode for trimming and smaller size
cd "$WASM_PROJECT"
dotnet publish -c Release

# Source AppBundle location (publish output differs from build)
APPBUNDLE="$WASM_PROJECT/bin/Release/net8.0/browser-wasm/AppBundle"

if [ ! -d "$APPBUNDLE" ]; then
    echo "Error: AppBundle not found at $APPBUNDLE"
    echo "Checking publish output location..."
    APPBUNDLE="$WASM_PROJECT/bin/Release/net8.0/browser-wasm/publish/wwwroot"
fi

if [ ! -d "$APPBUNDLE" ]; then
    echo "Trying alternative publish location..."
    APPBUNDLE="$WASM_PROJECT/bin/Release/net8.0/browser-wasm/native"
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

# Patch dotnet.js for cross-origin CDN compatibility
# The .NET WASM runtime uses credentials:"same-origin" which conflicts with CDN's CORS wildcard
# (Access-Control-Allow-Origin: * cannot be used with credentials)
echo "Patching dotnet.js for CDN compatibility..."
if [[ "$OSTYPE" == "darwin"* ]]; then
    # macOS sed requires empty string for -i
    sed -i '' 's/credentials:"same-origin"/credentials:"omit"/g' "$WASM_DIST/_framework/dotnet.js"
else
    sed -i 's/credentials:"same-origin"/credentials:"omit"/g' "$WASM_DIST/_framework/dotnet.js"
fi

# Report sizes
echo ""
echo "Build complete! File sizes:"
echo "----------------------------"

# Check for webcil files first (trimmed output uses .wasm extension but may be smaller)
if ls "$WASM_DIST/_framework/"*.wasm 1>/dev/null 2>&1; then
    echo "Largest WASM files:"
    du -h "$WASM_DIST/_framework/"*.wasm 2>/dev/null | sort -rh | head -10
fi

# Check for Brotli compressed files
if ls "$WASM_DIST/_framework/"*.br 1>/dev/null 2>&1; then
    echo ""
    echo "Brotli compressed files available (.br):"
    du -sh "$WASM_DIST/_framework/"*.br 2>/dev/null | head -5
fi

# Check for gzip compressed files
if ls "$WASM_DIST/_framework/"*.gz 1>/dev/null 2>&1; then
    echo ""
    echo "Gzip compressed files available (.gz):"
    du -sh "$WASM_DIST/_framework/"*.gz 2>/dev/null | head -5
fi

echo ""
echo "Total file count:"
find "$WASM_DIST/_framework" -type f | wc -l

echo ""
echo "Total WASM directory size:"
du -sh "$WASM_DIST"
