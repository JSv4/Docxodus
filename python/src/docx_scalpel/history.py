"""Host-owned exact version history through the existing local .NET process.

No new network/transit layer. File-backed histories survive client and process restarts;
memory histories are private to their open handle. Accepted sequence, not wall time,
orders content. Package-boundary calls are not a fine-grained typing recorder.
"""

from __future__ import annotations

import base64
from dataclasses import dataclass, field
from pathlib import Path
from types import MappingProxyType, TracebackType
from typing import Any, Mapping

from ._transport import _Transport
from .errors import DocxScalpelError
from .types import VerificationDigest


class DocxHistoryError(DocxScalpelError):
    """A typed history-domain failure (for example StaleHead or PayloadMismatch)."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


def _position(value: int) -> str:
    if isinstance(value, bool) or not isinstance(value, int) or not 0 <= value <= 9223372036854775807:
        raise ValueError("History positions must be nonnegative Int64 integers")
    return str(value)


@dataclass(frozen=True, slots=True)
class HistoryBlobReference:
    digest: VerificationDigest
    length: int

    @classmethod
    def _from_wire(cls, data: Mapping[str, Any]) -> HistoryBlobReference:
        return cls(VerificationDigest._from_wire(data["digest"]), data["length"])

    def to_wire(self) -> dict[str, Any]:
        return {"digest": {"algorithm": self.digest.algorithm, "value": self.digest.value}, "length": self.length}


def _reference(data: Mapping[str, Any] | None) -> HistoryBlobReference | None:
    return None if data is None else HistoryBlobReference._from_wire(data)


@dataclass(frozen=True, slots=True)
class HistoryHead:
    revision: int
    state: HistoryBlobReference

    @classmethod
    def _from_wire(cls, data: Mapping[str, Any]) -> HistoryHead:
        return cls(int(data["revision"]), HistoryBlobReference._from_wire(data["state"]))

    def to_wire(self) -> dict[str, Any]:
        return {"revision": _position(self.revision), "state": self.state.to_wire()}


@dataclass(frozen=True, slots=True)
class DocxSnapshotReference:
    blob: HistoryBlobReference
    content_digest: VerificationDigest

    @classmethod
    def _from_wire(cls, data: Mapping[str, Any]) -> DocxSnapshotReference:
        return cls(HistoryBlobReference._from_wire(data["blob"]), VerificationDigest._from_wire(data["contentDigest"]))


@dataclass(frozen=True, slots=True)
class DocxVersionMetadata:
    author: str
    created_at: str
    label: str | None = None
    message: str | None = None
    application_metadata: Mapping[str, str] = field(default_factory=dict)

    def __post_init__(self) -> None:
        object.__setattr__(self, "application_metadata", MappingProxyType(dict(self.application_metadata)))

    @classmethod
    def _from_wire(cls, data: Mapping[str, Any]) -> DocxVersionMetadata:
        return cls(data["author"], data["createdAt"], data.get("label"), data.get("message"), data["applicationMetadata"])

    def to_wire(self) -> dict[str, Any]:
        return {"author": self.author, "createdAt": self.created_at, "label": self.label,
                "message": self.message, "applicationMetadata": dict(self.application_metadata)}


@dataclass(frozen=True, slots=True)
class DocxVersionRecord:
    document_id: str
    metadata: DocxVersionMetadata
    nonce: str
    parent: HistoryBlobReference | None
    restored_from: HistoryBlobReference | None
    sequence: int
    snapshot: DocxSnapshotReference

    @classmethod
    def _from_wire(cls, data: Mapping[str, Any]) -> DocxVersionRecord:
        return cls(data["documentId"], DocxVersionMetadata._from_wire(data["metadata"]), data["nonce"],
                   _reference(data["parent"]), _reference(data["restoredFrom"]), int(data["sequence"]),
                   DocxSnapshotReference._from_wire(data["snapshot"]))


@dataclass(frozen=True, slots=True)
class DocxStoredVersion:
    id: HistoryBlobReference
    record: DocxVersionRecord

    @classmethod
    def _from_wire(cls, data: Mapping[str, Any]) -> DocxStoredVersion:
        return cls(HistoryBlobReference._from_wire(data["id"]), DocxVersionRecord._from_wire(data["record"]))


@dataclass(frozen=True, slots=True)
class DocxHistoryState:
    commit: HistoryBlobReference | None
    document_id: str
    epoch: int
    initial_snapshot: DocxSnapshotReference
    sequence: int
    snapshot: DocxSnapshotReference
    version: HistoryBlobReference

    @classmethod
    def _from_wire(cls, data: Mapping[str, Any]) -> DocxHistoryState:
        return cls(_reference(data["commit"]), data["documentId"], int(data["epoch"]),
                   DocxSnapshotReference._from_wire(data["initialSnapshot"]), int(data["sequence"]),
                   DocxSnapshotReference._from_wire(data["snapshot"]), HistoryBlobReference._from_wire(data["version"]))


@dataclass(frozen=True, slots=True)
class DocxHistoryView:
    head: HistoryHead
    state: DocxHistoryState
    version: DocxStoredVersion

    @classmethod
    def _from_wire(cls, data: Mapping[str, Any]) -> DocxHistoryView:
        return cls(HistoryHead._from_wire(data["head"]), DocxHistoryState._from_wire(data["state"]),
                   DocxStoredVersion._from_wire(data["version"]))


@dataclass(frozen=True, slots=True)
class DocxVersionPage:
    versions: tuple[DocxStoredVersion, ...]
    next: HistoryBlobReference | None


class DocxHistoryClient:
    """Use open_history(); a client is bound to its original subprocess, never a reused handle."""

    def __init__(self, transport: _Transport, handle: int) -> None:
        self._transport = transport
        self._handle = handle
        self._closed = False

    def __enter__(self) -> DocxHistoryClient:
        if self._closed:
            raise DocxHistoryError("Closed", "History client is closed")
        return self

    def __exit__(self, exc_type: type[BaseException] | None, exc: BaseException | None,
                 tb: TracebackType | None) -> None:
        self.close()

    def close(self) -> None:
        if not self._closed:
            self._transport.call("close_history", {"handle": self._handle})
            self._closed = True

    def _call(self, operation: str, document_id: str, *, docx_bytes: bytes | None = None, **fields: Any) -> dict[str, Any]:
        if self._closed:
            raise DocxHistoryError("Closed", "History client is closed")
        args: dict[str, Any] = {"handle": self._handle,
                               "request": {"schemaVersion": 1, "operation": operation, "documentId": document_id, **fields}}
        if docx_bytes is not None:
            if len(docx_bytes) > 256 * 1024 * 1024:
                raise ValueError("DOCX input exceeds the byte limit")
            args["docxB64"] = base64.b64encode(docx_bytes).decode("ascii")
        result = self._transport.call("history", args)
        if not result["success"]:
            raise DocxHistoryError(result["errorCode"], result["message"])
        return result

    def read(self, document_id: str) -> DocxHistoryView | None:
        value = self._call("read", document_id)["view"]
        return None if value is None else DocxHistoryView._from_wire(value)

    def create_version(self, document_id: str, expected_head: HistoryHead | None, docx_bytes: bytes,
                       metadata: DocxVersionMetadata) -> DocxHistoryView:
        return DocxHistoryView._from_wire(self._call("create", document_id, docx_bytes=docx_bytes,
            expectedHead=None if expected_head is None else expected_head.to_wire(), metadata=metadata.to_wire())["view"])

    def list_versions(self, document_id: str, cursor: HistoryBlobReference | None = None, limit: int = 25) -> DocxVersionPage:
        page = self._call("list", document_id, versionId=None if cursor is None else cursor.to_wire(), limit=limit)["page"]
        return DocxVersionPage(tuple(DocxStoredVersion._from_wire(version) for version in page["versions"]), _reference(page["next"]))

    def get_version(self, document_id: str, version_id: HistoryBlobReference) -> DocxStoredVersion:
        return DocxStoredVersion._from_wire(self._call("get", document_id, versionId=version_id.to_wire())["version"])

    def export_version(self, document_id: str, version_id: HistoryBlobReference) -> bytes:
        return base64.b64decode(self._call("export", document_id, versionId=version_id.to_wire())["bytes"], validate=True)

    def materialize(self, document_id: str, sequence: int, max_entries_to_scan: int = 10_000) -> bytes:
        return base64.b64decode(self._call("materialize", document_id, sequence=_position(sequence),
                                         maxEntriesToScan=max_entries_to_scan)["bytes"], validate=True)

    def replay(self, document_id: str, sequence: int, max_entries_to_scan: int = 10_000) -> bytes:
        return base64.b64decode(self._call("replay", document_id, sequence=_position(sequence),
                                         maxEntriesToScan=max_entries_to_scan)["bytes"], validate=True)

    def resolve_sequence_at_time(self, document_id: str, cutoff: str, max_entries_to_scan: int = 10_000) -> int:
        return int(self._call("resolveTime", document_id, cutoff=cutoff, maxEntriesToScan=max_entries_to_scan)["sequence"])

    def restore_version(self, document_id: str, expected_head: HistoryHead, version_id: HistoryBlobReference,
                        metadata: DocxVersionMetadata) -> DocxHistoryView:
        return DocxHistoryView._from_wire(self._call("restore", document_id, expectedHead=expected_head.to_wire(),
                                                   versionId=version_id.to_wire(), metadata=metadata.to_wire())["view"])


def open_history(root: str | Path | None = None) -> DocxHistoryClient:
    """Open a host-owned filesystem history, or an isolated ephemeral memory history when None.

    The caller controls permissions/retention of root/blobs and root/heads. Closing never
    deletes storage. Reopen after a process restart to obtain a new valid client handle.
    """
    transport = _Transport.get()
    handle = transport.call("open_history", {"root": None if root is None else str(Path(root).resolve())})
    return DocxHistoryClient(transport, handle)
