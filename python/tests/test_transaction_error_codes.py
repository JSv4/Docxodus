"""Wire-string parity for the MCP mutation-transaction ``EditErrorCode`` members.

These codes are produced by ``tools/mcp-server`` and generated from the C# enum by
``EnumToSnake``. ``docx-scalpel`` itself has no transaction surface (idempotent retries
are MCP-only), but it decodes every ``EditError`` on the wire, so a client that talks to
the MCP server through another path must still be able to name them.
"""

from __future__ import annotations

from pathlib import Path

from docx_scalpel.enums import EditErrorCode

TRANSACTION_CODES = {
    "invalid_transaction": "INVALID_TRANSACTION",
    "transaction_conflict": "TRANSACTION_CONFLICT",
    "transaction_result_evicted": "TRANSACTION_RESULT_EVICTED",
    "transaction_incomplete": "TRANSACTION_INCOMPLETE",
}


def test_transaction_error_codes_round_trip_from_their_wire_strings() -> None:
    for wire, member_name in TRANSACTION_CODES.items():
        member = getattr(EditErrorCode, member_name)
        assert member.value == wire
        # ``_missing_`` degrades an unknown code to INTERNAL_ERROR, so a member that was
        # never added would decode silently. Decoding the wire string and demanding the
        # exact member is what actually catches an absent or drifted code.
        assert EditErrorCode(wire) is member


def test_transaction_error_codes_are_declared_in_this_checkout() -> None:
    # Guards against an unrelated installed copy of docx_scalpel satisfying the import
    # above: the source of record for this repository must declare them too.
    source = (
        Path(__file__).resolve().parents[1] / "src" / "docx_scalpel" / "enums.py"
    ).read_text(encoding="utf-8")
    for wire, member_name in TRANSACTION_CODES.items():
        assert f'{member_name} = "{wire}"' in source
