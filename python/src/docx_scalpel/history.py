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

MAX_HISTORY_ARCHIVE_BYTES = 64 * 1024 * 1024


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
class HistoryRequestIdentity:
    id: str
    fingerprint: VerificationDigest

    @classmethod
    def _from_wire(cls, data: Mapping[str, Any]) -> HistoryRequestIdentity:
        return cls(data["id"], VerificationDigest._from_wire(data["fingerprint"]))


@dataclass(frozen=True, slots=True)
class HistoryRequestJournal:
    document_id: str
    revision: int
    index: HistoryBlobReference | None
    current: HistoryRequestIdentity | None

    @classmethod
    def _from_wire(cls, data: Mapping[str, Any]) -> HistoryRequestJournal:
        return cls(data["documentId"], int(data["revision"]), _reference(data["index"]),
                   None if data["current"] is None else HistoryRequestIdentity._from_wire(data["current"]))


@dataclass(frozen=True, slots=True)
class DocxHistoryState:
    commit: HistoryBlobReference | None
    document_id: str
    epoch: int
    initial_snapshot: DocxSnapshotReference
    sequence: int
    snapshot: DocxSnapshotReference
    version: HistoryBlobReference
    requests: HistoryRequestJournal | None = None
    parent_publication: HistoryHead | None = None
    operation: HistoryBlobReference | None = None

    @classmethod
    def _from_wire(cls, data: Mapping[str, Any]) -> DocxHistoryState:
        return cls(_reference(data["commit"]), data["documentId"], int(data["epoch"]),
                   DocxSnapshotReference._from_wire(data["initialSnapshot"]), int(data["sequence"]),
                   DocxSnapshotReference._from_wire(data["snapshot"]), HistoryBlobReference._from_wire(data["version"]),
                   None if data.get("requests") is None else HistoryRequestJournal._from_wire(data["requests"]),
                   None if data.get("parentPublication") is None else HistoryHead._from_wire(data["parentPublication"]),
                   _reference(data.get("operation")))


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


@dataclass(frozen=True, slots=True)
class PackageHistoryCommit:
    after: DocxSnapshotReference
    before: DocxSnapshotReference
    contribution: HistoryBlobReference | None
    document_id: str
    epoch: int
    kind: str
    parent: HistoryBlobReference | None
    sequence: int
    version: HistoryBlobReference

    @classmethod
    def _from_wire(cls, data: Mapping[str, Any]) -> PackageHistoryCommit:
        return cls(DocxSnapshotReference._from_wire(data["after"]), DocxSnapshotReference._from_wire(data["before"]),
                   _reference(data["contribution"]), data["documentId"], int(data["epoch"]), data["kind"],
                   _reference(data["parent"]), int(data["sequence"]), HistoryBlobReference._from_wire(data["version"]))


@dataclass(frozen=True, slots=True)
class DocxHistoryLogEntry:
    id: HistoryBlobReference
    commit: PackageHistoryCommit
    metadata: DocxVersionMetadata


@dataclass(frozen=True, slots=True)
class DocxHistoryUpdate:
    after: HistoryHead | None
    view: DocxHistoryView
    entries: tuple[DocxHistoryLogEntry, ...]
    reset: bool


@dataclass(frozen=True, slots=True)
class DocxHistoryArchiveInfo:
    document_id: str
    head: HistoryHead
    blob_count: int
    total_blob_bytes: int

    @classmethod
    def _from_wire(cls, data: Mapping[str, Any]) -> DocxHistoryArchiveInfo:
        return cls(data["documentId"], HistoryHead._from_wire(data["head"]), data["blobCount"], int(data["totalBlobBytes"]))


@dataclass(frozen=True, slots=True)
class DocxHistoryImportResult:
    archive: DocxHistoryArchiveInfo
    view: DocxHistoryView
    already_present: bool


@dataclass(frozen=True, slots=True)
class DocxTextSplice:
    part_uri: str
    text_node: int
    offset: int
    delete_count: int
    insert: str


def _splice(data: Mapping[str, Any] | None) -> DocxTextSplice | None:
    return None if data is None else DocxTextSplice(data["partUri"], data["textNode"], data["offset"], data["deleteCount"], data["insert"])


@dataclass(frozen=True, slots=True)
class DocxOperationRequest:
    request_id: str
    base: HistoryHead
    kind: str
    metadata: DocxVersionMetadata
    text: DocxTextSplice | None
    read_parts: tuple[str, ...]
    resolves: HistoryBlobReference | None


@dataclass(frozen=True, slots=True)
class DocxOperationInput:
    document_id: str
    request: DocxOperationRequest
    candidate: HistoryBlobReference | None


@dataclass(frozen=True, slots=True)
class DocxOperationRecord:
    document_id: str
    input: HistoryBlobReference
    parent: HistoryBlobReference | None
    revision: int
    before: HistoryHead
    proposed_snapshot: DocxSnapshotReference
    after_snapshot: DocxSnapshotReference
    version: HistoryBlobReference
    content_commit: HistoryBlobReference | None
    status: str
    conflict: str | None
    applied_text: DocxTextSplice | None


@dataclass(frozen=True, slots=True)
class DocxStoredOperation:
    id: HistoryBlobReference
    record: DocxOperationRecord
    input: DocxOperationInput

    @classmethod
    def _from_wire(cls, data: Mapping[str, Any]) -> DocxStoredOperation:
        record = data["record"]
        request = data["input"]["request"]
        return cls(HistoryBlobReference._from_wire(data["id"]), DocxOperationRecord(
            record["documentId"], HistoryBlobReference._from_wire(record["input"]), _reference(record["parent"]),
            int(record["revision"]), HistoryHead._from_wire(record["before"]),
            DocxSnapshotReference._from_wire(record["proposedSnapshot"]), DocxSnapshotReference._from_wire(record["afterSnapshot"]),
            HistoryBlobReference._from_wire(record["version"]), _reference(record["contentCommit"]), record["status"],
            record["conflict"], _splice(record["appliedText"])), DocxOperationInput(data["input"]["documentId"],
            DocxOperationRequest(request["requestId"], HistoryHead._from_wire(request["base"]), request["kind"],
                DocxVersionMetadata._from_wire(request["metadata"]), _splice(request["text"]), tuple(request["readParts"]),
                _reference(request["resolves"])), _reference(data["input"]["candidate"])))


@dataclass(frozen=True, slots=True)
class DocxOperationUpdate:
    after: HistoryHead | None
    view: DocxHistoryView
    operations: tuple[DocxStoredOperation, ...]


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

    def document(self, document_id: str) -> DocxHistoryDocument:
        """Bind a stable identity without creating a version; this client owns the lifetime."""
        return DocxHistoryDocument(self, document_id)

    def export_history_archive(self, document_id: str) -> bytes:
        return base64.b64decode(self._call("exportArchive", document_id)["bytes"], validate=True)

    def import_history_archive(self, data: bytes) -> DocxHistoryImportResult:
        """Import the original identity/head, never overwrite a different local history."""
        _archive_size(data)
        result = self._call("importArchive", "", docx_bytes=data)["import"]
        return DocxHistoryImportResult(DocxHistoryArchiveInfo._from_wire(result["archive"]),
                                      DocxHistoryView._from_wire(result["view"]), result["alreadyPresent"])

    def read_changes_since(self, document_id: str, after: HistoryHead | None, max_entries_to_scan: int = 10_000) -> DocxHistoryUpdate:
        """Host-triggered validated tail; first join starts at the latest checkpoint."""
        update = self._call("updates", document_id, expectedHead=None if after is None else after.to_wire(),
                            maxEntriesToScan=max_entries_to_scan)["update"]
        return DocxHistoryUpdate(None if update["after"] is None else HistoryHead._from_wire(update["after"]),
            DocxHistoryView._from_wire(update["view"]), tuple(DocxHistoryLogEntry(
                HistoryBlobReference._from_wire(entry["id"]), PackageHistoryCommit._from_wire(entry["commit"]),
                DocxVersionMetadata._from_wire(entry["metadata"])) for entry in update["entries"]), update["reset"])

    def read_operations_since(self, document_id: str, after: HistoryHead | None, max_entries_to_scan: int = 10_000) -> DocxOperationUpdate:
        value = self._call("operations", document_id, expectedHead=None if after is None else after.to_wire(),
                           maxEntriesToScan=max_entries_to_scan)["operationUpdate"]
        return DocxOperationUpdate(None if value["after"] is None else HistoryHead._from_wire(value["after"]),
            DocxHistoryView._from_wire(value["view"]), tuple(DocxStoredOperation._from_wire(op) for op in value["operations"]))

    def get_operation(self, document_id: str, operation_id: HistoryBlobReference) -> DocxStoredOperation:
        return DocxStoredOperation._from_wire(self._call("getOperation", document_id, operationId=operation_id.to_wire())["operation"])

    def export_operation_proposal(self, document_id: str, operation_id: HistoryBlobReference) -> bytes:
        return base64.b64decode(self._call("exportOperationProposal", document_id, operationId=operation_id.to_wire())["bytes"], validate=True)

    def compare_versions(self, document_id: str, before_version_id: HistoryBlobReference, after_version_id: HistoryBlobReference) -> bytes:
        return base64.b64decode(self._call("compare", document_id, beforeVersionId=before_version_id.to_wire(),
            afterVersionId=after_version_id.to_wire())["bytes"], validate=True)

    def create_version(self, document_id: str, expected_head: HistoryHead | None, docx_bytes: bytes,
                       metadata: DocxVersionMetadata, *, request_id: str | None = None) -> DocxHistoryView:
        """With request_id, retry the original captured bytes/metadata/head to recover its original result."""
        return DocxHistoryView._from_wire(self._call("create", document_id, docx_bytes=docx_bytes,
            expectedHead=None if expected_head is None else expected_head.to_wire(), metadata=metadata.to_wire(),
            **({} if request_id is None else {"requestId": request_id}))["view"])

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
                        metadata: DocxVersionMetadata, *, request_id: str | None = None) -> DocxHistoryView:
        return DocxHistoryView._from_wire(self._call("restore", document_id, expectedHead=expected_head.to_wire(),
            versionId=version_id.to_wire(), metadata=metadata.to_wire(),
            **({} if request_id is None else {"requestId": request_id}))["view"])


class DocxHistoryReader:
    """Document-scoped reads; no methods publish or change a live editor."""

    def __init__(self, client: DocxHistoryClient, document_id: str) -> None:
        self._client = client
        self.document_id = document_id

    def read(self) -> DocxHistoryView | None:
        return self._client.read(self.document_id)

    def list_versions(self, cursor: HistoryBlobReference | None = None, limit: int = 25) -> DocxVersionPage:
        return self._client.list_versions(self.document_id, cursor, limit)

    def get_version(self, version_id: HistoryBlobReference) -> DocxStoredVersion:
        return self._client.get_version(self.document_id, version_id)

    def export_docx(self, version_id: HistoryBlobReference | None = None) -> bytes:
        """Exact selected snapshot, or latest captured once; adds no external history."""
        return base64.b64decode(self._client._call("exportDocx", self.document_id,
            versionId=None if version_id is None else version_id.to_wire())["bytes"], validate=True)

    def export_history_archive(self) -> bytes:
        return self._client.export_history_archive(self.document_id)

    def materialize(self, sequence: int, max_entries_to_scan: int = 10_000) -> bytes:
        return self._client.materialize(self.document_id, sequence, max_entries_to_scan)

    def replay(self, sequence: int, max_entries_to_scan: int = 10_000) -> bytes:
        return self._client.replay(self.document_id, sequence, max_entries_to_scan)

    def resolve_sequence_at_time(self, cutoff: str, max_entries_to_scan: int = 10_000) -> int:
        return self._client.resolve_sequence_at_time(self.document_id, cutoff, max_entries_to_scan)

    def read_changes_since(self, after: HistoryHead | None, max_entries_to_scan: int = 10_000) -> DocxHistoryUpdate:
        return self._client.read_changes_since(self.document_id, after, max_entries_to_scan)

    def read_operations_since(self, after: HistoryHead | None, max_entries_to_scan: int = 10_000) -> DocxOperationUpdate:
        return self._client.read_operations_since(self.document_id, after, max_entries_to_scan)

    def get_operation(self, operation_id: HistoryBlobReference) -> DocxStoredOperation:
        return self._client.get_operation(self.document_id, operation_id)

    def export_operation_proposal(self, operation_id: HistoryBlobReference) -> bytes:
        return self._client.export_operation_proposal(self.document_id, operation_id)

    def compare_versions(self, before_version_id: HistoryBlobReference, after_version_id: HistoryBlobReference) -> bytes:
        """Redlined DOCX through the existing DocxCompare accepted-input revision policy."""
        return self._client.compare_versions(self.document_id, before_version_id, after_version_id)


class DocxHistoryDocument(DocxHistoryReader):
    """Explicit checkpoint/restore controls; persist request IDs and exact inputs before calling."""

    def create_version(self, expected_head: HistoryHead | None, data: bytes, metadata: DocxVersionMetadata,
                       *, request_id: str) -> DocxHistoryView:
        _request_id(request_id)
        return self._client.create_version(self.document_id, expected_head, data, metadata, request_id=request_id)

    def restore_version(self, expected_head: HistoryHead, version_id: HistoryBlobReference, metadata: DocxVersionMetadata,
                        *, request_id: str) -> DocxHistoryView:
        _request_id(request_id)
        return self._client.restore_version(self.document_id, expected_head, version_id, metadata, request_id=request_id)


class DocxHistoryArchive(DocxHistoryReader):
    """Standalone readonly file. Use open_history_archive() and close or a with block."""

    def __init__(self, client: DocxHistoryClient, info: DocxHistoryArchiveInfo) -> None:
        super().__init__(client, info.document_id)
        self.info = info

    def __enter__(self) -> DocxHistoryArchive:
        self._client.__enter__()
        return self

    def __exit__(self, exc_type: type[BaseException] | None, exc: BaseException | None,
                 tb: TracebackType | None) -> None:
        self.close()

    def close(self) -> None:
        self._client.close()


def _archive_size(data: bytes) -> None:
    if len(data) > MAX_HISTORY_ARCHIVE_BYTES:
        raise DocxHistoryError("ResourceLimit", "History archive exceeds 64 MiB")


def _request_id(value: str) -> None:
    if not isinstance(value, str) or not value.strip():
        raise DocxHistoryError("InvalidRequest", "A durable request_id is required")


def open_history_archive(data: bytes) -> DocxHistoryArchive:
    """Open independently of storage, read-only; byte bindings limit the archive to 64 MiB."""
    _archive_size(data)
    transport = _Transport.get()
    result = transport.call("open_history_archive", {"docxB64": base64.b64encode(data).decode("ascii")})
    if not result["success"]:
        raise DocxHistoryError(result["errorCode"], result["message"])
    return DocxHistoryArchive(DocxHistoryClient(transport, result["handle"]), DocxHistoryArchiveInfo._from_wire(result["archive"]))


def open_history(root: str | Path | None = None) -> DocxHistoryClient:
    """Open a host-owned filesystem history, or an isolated ephemeral memory history when None.

    The caller controls permissions/retention of root/blobs and root/heads. Closing never
    deletes storage. Reopen after a process restart to obtain a new valid client handle.
    """
    transport = _Transport.get()
    handle = transport.call("open_history", {"root": None if root is None else str(Path(root).resolve())})
    return DocxHistoryClient(transport, handle)
