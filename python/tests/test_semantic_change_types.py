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


def _object_value() -> SemanticValue:
    return SemanticValue._from_wire(
        {"kind": "object", "value": {"styleId": {"kind": "string", "value": "Heading1"}}}
    )


def test_object_values_are_hashable() -> None:
    parsed = _object_value()

    assert parsed in {parsed}
    assert len({parsed, _object_value()}) == 1


def test_object_values_reject_interior_mutation() -> None:
    parsed = _object_value()

    with pytest.raises(TypeError):
        parsed.value["styleId"] = SemanticValue._from_wire(  # type: ignore[index]
            {"kind": "string", "value": "smuggled"}
        )


def test_object_values_read_like_mappings() -> None:
    parsed = _object_value()

    assert parsed.value["styleId"].value == "Heading1"  # type: ignore[index]
    assert list(parsed.value) == ["styleId"]  # type: ignore[arg-type]
    assert len(parsed.value) == 1  # type: ignore[arg-type]


@pytest.mark.parametrize(
    "payload",
    [
        {"kind": "object"},
        {"kind": "object", "value": None},
        {"kind": "array"},
        {"kind": "array", "value": None},
        {"kind": "string"},
        {"kind": "string", "value": None},
        {"kind": "boolean"},
        {"kind": "boolean", "value": "true"},
        {"kind": "digest", "value": "aa"},
        {"kind": "digest", "algorithm": "SHA-256"},
        {"kind": "digest", "algorithm": "", "value": "aa"},
        {"kind": "digest", "algorithm": "SHA-256", "value": ""},
    ],
)
def test_schema_required_members_are_enforced(payload: dict) -> None:
    with pytest.raises(ValueError):
        SemanticValue._from_wire(payload)
