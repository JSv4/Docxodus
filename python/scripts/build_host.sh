#!/usr/bin/env bash
#
# Publish a self-contained docxodus-pyhost binary for one RID and stage it
# under python/vendor/<rid>/ where pyproject.toml picks it up as wheel data.
#
# Usage:
#   scripts/build_host.sh                    # current host RID, ReadyToRun on
#   scripts/build_host.sh linux-x64          # cross-RID
#   scripts/build_host.sh osx-arm64 --no-r2r # disable PublishReadyToRun (smaller wheel)
#
# Per-RID this produces ~25 MB compressed. cibuildwheel invokes us once per
# matrix entry; the .NET 8 runtime is bundled inside the single-file binary
# and extracts itself on first launch.

set -euo pipefail

# Locate the Docxodus repo root from this script's location: python/scripts/.
script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
python_dir="$(cd "$script_dir/.." && pwd)"
repo_root="$(cd "$python_dir/.." && pwd)"

rid="${1:-}"
ready_to_run="true"
for arg in "$@"; do
    case "$arg" in
        --no-r2r) ready_to_run="false" ;;
    esac
done

if [[ -z "$rid" ]]; then
    case "$(uname -s):$(uname -m)" in
        Linux:x86_64)    rid="linux-x64" ;;
        Linux:aarch64)   rid="linux-arm64" ;;
        Darwin:x86_64)   rid="osx-x64" ;;
        Darwin:arm64)    rid="osx-arm64" ;;
        MINGW*:*|MSYS*:*|CYGWIN*:*) rid="win-x64" ;;
        *) echo "could not infer RID from $(uname -sm); pass one explicitly" >&2; exit 1 ;;
    esac
fi

out_dir="$python_dir/vendor/$rid"
csproj="$repo_root/tools/python-host/pyhost.csproj"

if [[ ! -f "$csproj" ]]; then
    echo "expected $csproj — is this a Docxodus monorepo clone?" >&2
    exit 1
fi

echo "Publishing docxodus-pyhost for $rid → $out_dir"
rm -rf "$out_dir"
mkdir -p "$out_dir"

dotnet publish "$csproj" \
    -c Release \
    -r "$rid" \
    --self-contained true \
    -p:PublishSingleFile=true \
    -p:PublishReadyToRun="$ready_to_run" \
    -p:IncludeNativeLibrariesForSelfExtract=true \
    -o "$out_dir"

binary_name="docxodus-pyhost"
case "$rid" in
    win-*) binary_name="docxodus-pyhost.exe" ;;
esac

if [[ ! -f "$out_dir/$binary_name" ]]; then
    echo "publish completed but $out_dir/$binary_name is missing" >&2
    exit 1
fi

size_mb=$(du -m "$out_dir/$binary_name" | cut -f1)
echo "✓ $out_dir/$binary_name (${size_mb} MB)"
