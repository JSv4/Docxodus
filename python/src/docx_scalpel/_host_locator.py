"""Locate the ``docxodus-pyhost`` binary on disk.

Resolution order:
1. ``DOCXODUS_HOST`` env var — overrides everything. Useful when iterating on
   the host without rebuilding the wheel.
2. Bundled binary at ``docx_scalpel/_bin/docxodus-pyhost[.exe]`` — populated by
   ``scripts/build_host.sh`` at wheel-build time.
3. Dev fallback — walk up from this file looking for the Docxodus monorepo's
   ``tools/python-host/bin/{Release,Debug}/net10.0/docxodus-pyhost[.exe]``. Lets
   ``pip install -e .`` work straight from a clone without copying binaries.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

from .errors import DocxodusHostNotFoundError

_HOST_EXE = "docxodus-pyhost.exe" if sys.platform == "win32" else "docxodus-pyhost"


def find_host() -> Path:
    """Return the absolute path of the ``docxodus-pyhost`` binary to spawn.

    Raises :class:`DocxodusHostNotFoundError` if none of the resolution
    strategies find a binary.
    """
    env = os.environ.get("DOCXODUS_HOST")
    if env:
        p = Path(env).expanduser()
        if p.is_file():
            return p.resolve()
        raise DocxodusHostNotFoundError(
            f"DOCXODUS_HOST is set to {env!r} but no file exists there"
        )

    bundled = Path(__file__).resolve().parent / "_bin" / _HOST_EXE
    if bundled.is_file():
        return bundled

    dev = _find_dev_binary()
    if dev is not None:
        return dev

    raise DocxodusHostNotFoundError(
        f"could not locate {_HOST_EXE}. Tried:\n"
        f"  1. $DOCXODUS_HOST (unset)\n"
        f"  2. bundled at {bundled}\n"
        "  3. dev fallback under any ancestor's tools/python-host/bin/{Release,Debug}/net10.0/\n"
        "Fix: set DOCXODUS_HOST=/path/to/docxodus-pyhost, or run "
        "`dotnet build tools/python-host/pyhost.csproj` from the Docxodus repo root, "
        "or install a wheel that ships a bundled binary."
    )


def _find_dev_binary() -> Path | None:
    """Walk up from this file looking for the Docxodus monorepo's built binary."""
    start = Path(__file__).resolve()
    for parent in start.parents:
        for config in ("Release", "Debug"):
            candidate = parent / "tools" / "python-host" / "bin" / config / "net10.0" / _HOST_EXE
            if candidate.is_file():
                return candidate
        # Once we find the monorepo root, no need to keep walking.
        if (parent / "Docxodus.sln").is_file():
            return None
    return None
