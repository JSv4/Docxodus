from __future__ import annotations

import pytest

from docx_scalpel.types import SemanticValue, SemanticValueKind


@pytest.mark.parametrize("value", [-(2**53 - 1), 2**53 - 1])
def test_semantic_integer_accepts_javascript_safe_boundaries(value: int) -> None:
    parsed = SemanticValue._from_wire({"kind": "integer", "value": value})

    assert parsed.kind is SemanticValueKind.INTEGER
    assert parsed.value == value


@pytest.mark.parametrize("value", [-(2**53), 2**53, True, 1.5, "1"])
def test_semantic_integer_rejects_nonportable_values(value: object) -> None:
    with pytest.raises(ValueError, match="semantic integer"):
        SemanticValue._from_wire({"kind": "integer", "value": value})
