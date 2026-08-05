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
APPBUNDLE="$WASM_PROJECT/bin/Release/net10.0/browser-wasm/AppBundle"

if [ ! -d "$APPBUNDLE" ]; then
    echo "Error: AppBundle not found at $APPBUNDLE"
    echo "Checking publish output location..."
    APPBUNDLE="$WASM_PROJECT/bin/Release/net10.0/browser-wasm/publish/wwwroot"
fi

if [ ! -d "$APPBUNDLE" ]; then
    echo "Trying alternative publish location..."
    APPBUNDLE="$WASM_PROJECT/bin/Release/net10.0/browser-wasm/native"
fi

if [ ! -d "$APPBUNDLE" ]; then
    echo "Error: AppBundle not found"
    exit 1
fi

echo "AppBundle found at: $APPBUNDLE"

# Clean and create destination
rm -rf "$WASM_DIST"
mkdir -p "$WASM_DIST"

# Copy the _framework directory (contains all WASM and JS files).
# Debug artifacts (.map source maps, .symbols) are excluded — they are dead weight
# in the npm package; use a Debug build when you need them.
echo "Copying _framework (excluding debug artifacts)..."
mkdir -p "$WASM_DIST/_framework"
find "$APPBUNDLE/_framework" -maxdepth 1 -type f \
    ! -name "*.map" ! -name "*.symbols" \
    -exec cp {} "$WASM_DIST/_framework/" \;

# Copy main.js
echo "Copying main.js..."
cp "$WASM_PROJECT/main.js" "$WASM_DIST/"

# Copy index.html for testing
cp "$WASM_PROJECT/index.html" "$WASM_DIST/"

# Patch dotnet.js and dotnet.native.js for cross-origin CDN compatibility
# The .NET WASM runtime uses credentials:"same-origin" which conflicts with CDN's CORS wildcard
# (Access-Control-Allow-Origin: * cannot be used with credentials)
# Both files make fetch requests and both need to be patched.
echo "Patching dotnet.js and dotnet.native.js for CDN compatibility..."
if [[ "$OSTYPE" == "darwin"* ]]; then
    # macOS sed requires empty string for -i
    sed -i '' 's/credentials:"same-origin"/credentials:"omit"/g' "$WASM_DIST/_framework/dotnet.js"
    sed -i '' 's/credentials:"same-origin"/credentials:"omit"/g' "$WASM_DIST/_framework/dotnet.native.js"
else
    sed -i 's/credentials:"same-origin"/credentials:"omit"/g' "$WASM_DIST/_framework/dotnet.js"
    sed -i 's/credentials:"same-origin"/credentials:"omit"/g' "$WASM_DIST/_framework/dotnet.native.js"
fi

# Verify the patches were applied
echo "Verifying patches..."
if grep -q 'credentials:"same-origin"' "$WASM_DIST/_framework/dotnet.js" 2>/dev/null; then
    echo "WARNING: dotnet.js still contains credentials:same-origin"
fi
if grep -q 'credentials:"same-origin"' "$WASM_DIST/_framework/dotnet.native.js" 2>/dev/null; then
    echo "WARNING: dotnet.native.js still contains credentials:same-origin"
fi

# Some .NET WASM SDK builds emit per-asset "integrity" in dotnet.boot.js, but the
# published loader (dotnet.js) only applies SRI when the field is named "hash".
# Normalize so asset fetches get integrity checks (matches published npm packages).
BOOT_JS="$WASM_DIST/_framework/dotnet.boot.js"
if [[ -f "$BOOT_JS" ]] && grep -q '"integrity"' "$BOOT_JS"; then
    echo "Normalizing dotnet.boot.js asset integrity fields to hash..."
    if [[ "$OSTYPE" == "darwin"* ]]; then
        sed -i '' 's/"integrity"\([[:space:]]*:[[:space:]]*"sha256-\)/"hash"\1/g' "$BOOT_JS"
    else
        sed -i 's/"integrity"\([[:space:]]*:[[:space:]]*"sha256-\)/"hash"\1/g' "$BOOT_JS"
    fi
fi

# Precompress every framework asset with Brotli (quality 11) so hosts that support
# content negotiation (nginx brotli_static, Caddy precompressed, Netlify, Vercel,
# Cloudflare Pages) can serve ~3.3 MB over the wire instead of ~13 MB. The .br
# siblings ship in the npm package; hosts that ignore them serve the raw files
# exactly as before. gzip is deliberately NOT precompressed — gzip-capable hosts
# compress on the fly, while brotli-11 is too slow for that.
echo ""
echo "Precompressing framework assets (brotli -11)..."
node -e '
const fs = require("fs"), zlib = require("zlib"), path = require("path");
const dir = process.argv[1];
let raw = 0, br = 0, n = 0;
for (const f of fs.readdirSync(dir)) {
  const p = path.join(dir, f);
  if (!fs.statSync(p).isFile() || f.endsWith(".br")) continue;
  const buf = fs.readFileSync(p);
  const c = zlib.brotliCompressSync(buf, { params: {
    [zlib.constants.BROTLI_PARAM_QUALITY]: 11,
    [zlib.constants.BROTLI_PARAM_SIZE_HINT]: buf.length } });
  fs.writeFileSync(p + ".br", c);
  raw += buf.length; br += c.length; n++;
}
console.log(`  ${n} assets: ${(raw/1048576).toFixed(2)} MB raw -> ${(br/1048576).toFixed(2)} MB brotli`);
fs.writeFileSync(path.join(dir, ".wire-size"), String(br));
' "$WASM_DIST/_framework"

# Report sizes
echo ""
echo "Build complete! File sizes:"
echo "----------------------------"
echo "Largest WASM files:"
du -h "$WASM_DIST/_framework/"*.wasm 2>/dev/null | sort -rh | head -10

echo ""
echo "Total file count:"
find "$WASM_DIST/_framework" -type f | wc -l

echo ""
echo "Total WASM directory size:"
du -sh "$WASM_DIST"

# Wire-size budget gate. The brotli total is what a negotiation-capable host actually
# sends to boot the runtime. Budget 4.0 MB (measured ~3.3 MB after trimming) — if this
# trips, something re-rooted an assembly or a dependency grew; see
# docs/architecture/wasm-packaging.md before raising it.
WIRE_BUDGET_BYTES=$((4 * 1024 * 1024))
WIRE_BYTES=$(cat "$WASM_DIST/_framework/.wire-size")
rm -f "$WASM_DIST/_framework/.wire-size"
echo ""
echo "Wire size (brotli total): $((WIRE_BYTES / 1024)) KB (budget: $((WIRE_BUDGET_BYTES / 1024)) KB)"
if [ "$WIRE_BYTES" -gt "$WIRE_BUDGET_BYTES" ]; then
    echo "ERROR: WASM wire payload exceeds the ${WIRE_BUDGET_BYTES}-byte budget."
    echo "Did an assembly get re-rooted (TrimmerRootAssembly) or a dependency grow?"
    exit 1
fi
