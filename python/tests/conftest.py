"""Shared pytest fixtures for docx-scalpel.

Test data is the same ``TestFiles/`` corpus the C# tests use — Python and .NET
exercise byte-identical inputs so a divergence on either side is detectable.
"""

from __future__ import annotations

import logging
from pathlib import Path

import pytest

# python/tests/conftest.py → python/tests → python → Docxodus repo root → TestFiles
TEST_FILES_DIR = Path(__file__).resolve().parents[2] / "TestFiles"


@pytest.fixture(scope="session")
def test_files_dir() -> Path:
    if not TEST_FILES_DIR.is_dir():
        pytest.skip(f"TestFiles/ not found at {TEST_FILES_DIR}; need a Docxodus monorepo clone")
    return TEST_FILES_DIR


@pytest.fixture(scope="session")
def tour_plan_bytes(test_files_dir: Path) -> bytes:
    """The HC001 5-day tour plan template — the C# smoke-test fixture."""
    return (test_files_dir / "HC001-5DayTourPlanTemplate.docx").read_bytes()


@pytest.fixture(autouse=True)
def _route_host_logs(caplog: pytest.LogCaptureFixture) -> None:
    """Pipe `docx_scalpel.host` logs into pytest's caplog so failures show host stderr."""
    caplog.set_level(logging.DEBUG, logger="docx_scalpel.host")
