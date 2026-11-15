#!/bin/bash
# Packs the MCP server (tools/mcp-server) as the NuGet package Docxodus.McpServer and checks
# the result is what NuGet.org and MCP clients need:
#
#   * the package is both a .NET tool (`dotnet tool install`, `dnx`) and an MCP server
#     (NuGet.org's packagetype=mcpserver search);
#   * it carries .mcp/server.json, and the version in that manifest is the package's own.
#     NuGet.org builds the client configuration it shows from that file, so a manifest left
#     at the in-tree placeholder would tell every client to install a version that does not
#     exist.
#
# The in-tree manifest keeps the placeholder version 0.0.0, like the npm manifests; this
# script packs a stamped copy and leaves the working tree untouched.
#
# publish.yml and ci.yml both call this script, so the packaging is exercised on every pull
# request instead of for the first time on release day.
#
# Usage: scripts/pack-mcp-server.sh <version> <output-dir>
set -euo pipefail

VERSION="${1:?usage: pack-mcp-server.sh <version> <output-dir>}"
OUTPUT="${2:?usage: pack-mcp-server.sh <version> <output-dir>}"
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PROJECT="$REPO_ROOT/tools/mcp-server"

mkdir -p "$OUTPUT"
OUTPUT="$(cd "$OUTPUT" && pwd)"
WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

VERSION="$VERSION" python3 - "$PROJECT/.mcp/server.json" "$WORK/server.json" <<'PY'
import json, os, sys
source, target = sys.argv[1:3]
version = os.environ["VERSION"]
with open(source, encoding="utf-8") as f:
    manifest = json.load(f)
manifest["version"] = version
for package in manifest["packages"]:
    package["version"] = version
with open(target, "w", encoding="utf-8", newline="\n") as f:
    json.dump(manifest, f, indent=2)
    f.write("\n")
PY

dotnet pack "$PROJECT/mcpserver.csproj" -c Release -o "$OUTPUT" \
  -p:PackageVersion="$VERSION" -p:McpServerManifest="$WORK/server.json"

PACKAGE="$OUTPUT/Docxodus.McpServer.$VERSION.nupkg"
VERSION="$VERSION" python3 - "$PACKAGE" <<'PY'
import json, os, re, sys, zipfile
version = os.environ["VERSION"]
with zipfile.ZipFile(sys.argv[1]) as package:
    names = package.namelist()
    nuspec = package.read(next(n for n in names if n.endswith(".nuspec"))).decode("utf-8")
    assert ".mcp/server.json" in names, "the package has no .mcp/server.json"
    manifest = json.loads(package.read(".mcp/server.json"))
types = set(re.findall(r'<packageType name="([^"]+)"', nuspec))
assert {"McpServer", "DotnetTool"} <= types, f"package types are {sorted(types)}"
assert f"<version>{version}</version>" in nuspec, "the nuspec version is not the requested one"
entry = manifest["packages"][0]
assert manifest["version"] == entry["version"] == version, "server.json was not stamped"
assert entry["registryType"] == "nuget" and entry["identifier"] == "Docxodus.McpServer"
assert "README.md" in names, "the package has no README"
print(f"Docxodus.McpServer {version}: package types {sorted(types)}, server.json stamped.")
PY
