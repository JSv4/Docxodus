"""Wire-decoding regressions for native content controls."""

from __future__ import annotations

from docx_scalpel import ContentControlBindingInfo, ContentControlInfo


def _wire() -> dict[str, object]:
    return {
        "anchorId": "sdt:body:abc",
        "type": "plain_text",
        "placement": "inline",
        "nativeId": "17",
        "isBound": True,
        "owningPartUri": "/word/document.xml",
        "scope": "body",
        "depth": 0,
        "hasValidNativeId": True,
        "hasDuplicateNativeId": False,
        "canMutate": False,
        "canDetachTargetBinding": True,
    }


def test_content_control_empty_binding_object_is_decoded_by_key_presence() -> None:
    wire = _wire()
    wire["binding"] = {}

    decoded = ContentControlInfo._from_wire(wire)

    assert decoded.binding == ContentControlBindingInfo(None, None, None)
    assert decoded.is_bound is True


def test_content_control_null_or_absent_binding_decodes_as_none() -> None:
    absent = ContentControlInfo._from_wire(_wire())
    wire = _wire()
    wire["binding"] = None
    explicit_null = ContentControlInfo._from_wire(wire)

    assert absent.binding is None
    assert explicit_null.binding is None
