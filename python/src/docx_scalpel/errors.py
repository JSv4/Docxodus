"""Exception types raised by docx-scalpel.

Transport-layer failures (host crash, malformed JSON, unknown op) raise.
Business-level outcomes (anchor not found, malformed markdown) return as
``EditResult(success=False, error=EditError(...))`` — never raise.
"""

from __future__ import annotations


class DocxScalpelError(Exception):
    """Base class for all docx-scalpel exceptions."""


class DocxodusTransportError(DocxScalpelError):
    """Raised when the NDJSON transport to ``docxodus-pyhost`` fails.

    Covers: host process exit, malformed reply, response id mismatch,
    host-reported ``ok: false`` envelopes (``unknown_op``, ``malformed_request``,
    ``invalid_argument``, ``internal_error``).
    """

    def __init__(self, message: str, *, code: str | None = None, trace: str | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.trace = trace


class DocxodusHostNotFoundError(DocxScalpelError):
    """Raised when ``_host_locator`` can't find ``docxodus-pyhost`` on disk.

    Resolution order is documented in ``_host_locator.find_host()``.
    """
