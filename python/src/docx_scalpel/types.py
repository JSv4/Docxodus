"""Frozen dataclasses mirroring the C# value types in ``Docxodus``.

Wire keys are camelCase (matching the WASM bridge so the JSON shapes are
interchangeable between TypeScript and Python clients). Decoders translate to
snake_case Python fields.

Each type exposes a private ``_from_wire(d)`` classmethod that takes the parsed
JSON dict and returns an instance. The leading underscore marks these as
transport-internal: callers should never need to invoke them — every public
``DocxSession`` method that returns one of these types already decodes for you.
Encoders are simple dict-builders on the encode side (``to_wire()``) and live
where they're used in ``session.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Callable, Generic, Mapping, Sequence, TypeVar

from .enums import (
    AnchorIdRendering,
    AnchorRenderMode,
    ConflictResolution,
    ContextBoundary,
    DocxDiffFormatComparison,
    DocxDiffRevisionGranularity,
    DocxDiffRevisionType,
    DeliverableArtifactAvailability,
    DeliverableArtifactRole,
    DeliverableCheckStatus,
    DeliverableFindingCategory,
    DeliverableFindingDisposition,
    DeliverablePackageChangeKind,
    DeliverableSemanticChangeFamily,
    DeliverableVerificationDecision,
    DeliverableVerificationMode,
    DeliveryArtifactVerificationStatus,
    EditErrorCode,
    EmptyParagraphMode,
    HeaderFooterKind,
    HyperlinkKind,
    LineSpacingRule,
    MutationBatchMode,
    PackageContentTypeDeclarationKind,
    PackageContentTypeSource,
    PackageKind,
    PackageRelationshipTargetMode,
    ParagraphAlignment,
    PlaceholderKind,
    PlaceholderKinds,
    ProjectionScopes,
    RedlinePackageDivergenceKind,
    RedlineProofDirection,
    RedlineRevisionDisposition,
    TableAnchorEntityKind,
    TableRenderMode,
    TableRowHeightRule,
    TableVerticalMergeRole,
    TrackedChangeMode,
    VerificationFindingSeverity,
    WhitespaceMode,
)

__all__ = [
    "Anchor",
    "CharSpan",
    "FormatOp",
    "ParagraphBorderEdge",
    "ParagraphFormatOp",
    "EditError",
    "EditResult",
    "MarkdownPatch",
    "AnchorTarget",
    "AnchorInfo",
    "PreconditionTarget",
    "PreconditionFailure",
    "TextRangePrecondition",
    "MutationPreconditions",
    "MutationBatchStep",
    "MutationBatchStepResult",
    "MutationBatchFailure",
    "MutationBatchChangeSet",
    "MutationBatchResult",
    "BlockMetadata",
    "BulkEditResult",
    "ChangeLocation",
    "FillOptions",
    "FindOptions",
    "HeaderFooterRef",
    "HtmlOptions",
    "ListMembership",
    "NumberFormat",
    "PageCitation",
    "PageCitationRequest",
    "PageMap",
    "PageMapFragment",
    "PageMapPage",
    "PageMapRect",
    "PageMapRegistrationResult",
    "PageMapStatus",
    "ParagraphFormatting",
    "RunFormattingInfo",
    "TableStyleFormatting",
    "StyleInfo",
    "InlineSpan",
    "FormattingInspection",
    "RunFormatting",
    "RunFragment",
    "SectionInfo",
    "TextMatch",
    "BlockSlice",
    "CrossBlockMatch",
    "TemplatePlaceholder",
    "MarkdownProjection",
    "DocxSessionSettings",
    "PackageAnnotationCounts",
    "PackageContentTypeDeclaration",
    "PackageManifest",
    "PackageManifestEntry",
    "PackageManifestFacts",
    "PackageRelationship",
    "PackageRevisionCounts",
    "VerificationDigest",
    "VerificationFinding",
    "DeliverablePackageIdentity",
    "DeliverableCheckResult",
    "DeliverableFinding",
    "DeliverablePackageChange",
    "DeliverableSemanticChange",
    "DeliverableSemanticDelta",
    "DeliverableArtifactMetadata",
    "DeliverableVerificationResult",
    "RedlineReversibilityProof",
    "RedlineProofPathResult",
    "RedlineProofPackageIdentity",
    "RedlineProofFinding",
    "RedlineRevisionClassification",
    "RedlineRevisionIdentity",
    "RedlineModeledSemanticComparison",
    "RedlinePackageDivergence",
    "WmlToMarkdownConverterSettings",
    "DocumentAnnotation",
    "AnnotationUpdate",
    "DocumentRange",
    "HyperlinkInfo",
    "ImageBinaryFormat",
    "ImageMarkupKind",
    "ImagePlacement",
    "ImageWrapMode",
    "ImageWrapSide",
    "ImageHorizontalReference",
    "ImageVerticalReference",
    "ImageHorizontalAlignment",
    "ImageVerticalAlignment",
    "FloatingImageLayout",
    "ImageInsertOptions",
    "ImageDimensions",
    "ImageFormatCapability",
    "ImageOccurrence",
    "ImageCapabilities",
    "ContentControlType",
    "ContentControlPlacement",
    "ContentControlBindingPolicy",
    "ContentControlFillOptions",
    "ContentControlBindingInfo",
    "ContentControlInfo",
    "BookmarkRangeSegment",
    "BookmarkInfo",
    "EditSummary",
    "ReplaceOptions",
    "DocxDiffSettings",
    "DocxDiffRevision",
    "DocxDiffFormatChange",
    "SemanticChangeOperation",
    "SemanticChangeFamily",
    "SemanticValueKind",
    "SemanticValue",
    "SemanticChange",
    "SemanticChangeSet",
    "DocxDiffReviewer",
    "DocxDiffConsolidateSettings",
    "DocxDiffConflictCompetitor",
    "DocxDiffConflict",
    "DocxDiffConsolidatedRevision",
    "TableInsertOptions",
    "TableBorderSpec",
    "TableRowOptions",
    "TableCellMetadata",
    "TableRowMetadata",
    "TableColumnMetadata",
    "TableMetadata",
    "TableMetadataResult",
    "TableCellResolutionResult",
    "TableAnchorLocation",
    "RetainedTableAnchor",
    "TableAnchorMapping",
]


# ---------------------------------------------------------------------------
# Package verification manifests
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class VerificationDigest:
    """Algorithm-labelled lower-case hexadecimal digest."""

    algorithm: str
    value: str

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "VerificationDigest":
        return cls(algorithm=str(d["algorithm"]), value=str(d["value"]))


@dataclass(frozen=True, slots=True)
class ChangeLocation:
    """Stable package location attached to a verification finding."""

    entry_uri: str | None = None
    owner_uri: str | None = None
    relationship_id: str | None = None
    target_uri: str | None = None
    property_path: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "ChangeLocation":
        return cls(
            entry_uri=d.get("entryUri"),
            owner_uri=d.get("ownerUri"),
            relationship_id=d.get("relationshipId"),
            target_uri=d.get("targetUri"),
            property_path=d.get("propertyPath"),
        )


@dataclass(frozen=True, slots=True)
class VerificationFinding:
    code: str
    severity: VerificationFindingSeverity
    message: str
    location: ChangeLocation | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "VerificationFinding":
        location = d.get("location")
        return cls(
            code=str(d["code"]),
            severity=VerificationFindingSeverity(str(d["severity"])),
            message=str(d["message"]),
            location=ChangeLocation._from_wire(location) if isinstance(location, Mapping) else None,
        )


@dataclass(frozen=True, slots=True)
class PackageManifestEntry:
    uri: str
    occurrence: int
    content_type: str | None
    content_type_source: PackageContentTypeSource
    size: int
    compressed_size: int
    raw_bytes_digest: VerificationDigest | None
    normalized_xml_digest: VerificationDigest | None
    is_xml: bool
    is_encrypted: bool | None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "PackageManifestEntry":
        raw = d.get("rawBytesDigest")
        normalized = d.get("normalizedXmlDigest")
        return cls(
            uri=str(d["uri"]),
            occurrence=int(d["occurrence"]),
            content_type=d.get("contentType"),
            content_type_source=PackageContentTypeSource(str(d["contentTypeSource"])),
            size=int(d["size"]),
            compressed_size=int(d["compressedSize"]),
            raw_bytes_digest=VerificationDigest._from_wire(raw) if isinstance(raw, Mapping) else None,
            normalized_xml_digest=(
                VerificationDigest._from_wire(normalized)
                if isinstance(normalized, Mapping)
                else None
            ),
            is_xml=bool(d["isXml"]),
            is_encrypted=(
                bool(d["isEncrypted"]) if d.get("isEncrypted") is not None else None
            ),
        )


@dataclass(frozen=True, slots=True)
class PackageContentTypeDeclaration:
    kind: PackageContentTypeDeclarationKind
    key: str
    content_type: str
    occurrence: int

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "PackageContentTypeDeclaration":
        return cls(
            kind=PackageContentTypeDeclarationKind(str(d["kind"])),
            key=str(d["key"]),
            content_type=str(d["contentType"]),
            occurrence=int(d["occurrence"]),
        )


@dataclass(frozen=True, slots=True)
class PackageRelationship:
    owner_uri: str
    id: str
    type: str
    target: str
    target_mode: PackageRelationshipTargetMode
    resolved_target_uri: str | None
    is_target_present: bool | None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "PackageRelationship":
        present = d.get("isTargetPresent")
        return cls(
            owner_uri=str(d["ownerUri"]),
            id=str(d["id"]),
            type=str(d["type"]),
            target=str(d["target"]),
            target_mode=PackageRelationshipTargetMode(str(d["targetMode"])),
            resolved_target_uri=d.get("resolvedTargetUri"),
            is_target_present=bool(present) if present is not None else None,
        )


@dataclass(frozen=True, slots=True)
class PackageRevisionCounts:
    insertions: int = 0
    deletions: int = 0
    move_from: int = 0
    move_to: int = 0
    property_changes: int = 0
    run_property_changes: int = 0
    structural_changes: int = 0
    other_changes: int = 0
    total: int = 0

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "PackageRevisionCounts":
        return cls(
            insertions=int(d.get("insertions", 0)),
            deletions=int(d.get("deletions", 0)),
            move_from=int(d.get("moveFrom", 0)),
            move_to=int(d.get("moveTo", 0)),
            property_changes=int(d.get("propertyChanges", 0)),
            run_property_changes=int(d.get("runPropertyChanges", 0)),
            structural_changes=int(d.get("structuralChanges", 0)),
            other_changes=int(d.get("otherChanges", 0)),
            total=int(d.get("total", 0)),
        )


@dataclass(frozen=True, slots=True)
class PackageAnnotationCounts:
    comments: int = 0
    comment_replies: int = 0
    threaded_comment_metadata: int = 0
    resolved_comments: int = 0
    people: int = 0
    docxodus_annotations: int = 0

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "PackageAnnotationCounts":
        return cls(
            comments=int(d.get("comments", 0)),
            comment_replies=int(d.get("commentReplies", 0)),
            threaded_comment_metadata=int(d.get("threadedCommentMetadata", 0)),
            resolved_comments=int(d.get("resolvedComments", 0)),
            people=int(d.get("people", 0)),
            docxodus_annotations=int(d.get("docxodusAnnotations", 0)),
        )


@dataclass(frozen=True, slots=True)
class PackageManifestFacts:
    main_document_uri: str | None
    is_strict_ooxml: bool
    is_macro_enabled: bool
    has_core_properties: bool
    has_extended_properties: bool
    has_custom_properties: bool
    section_count: int
    paragraph_count: int
    table_count: int
    header_part_count: int
    footer_part_count: int
    footnote_count: int
    endnote_count: int
    style_count: int
    numbering_definition_count: int
    theme_part_count: int
    media_part_count: int
    custom_xml_part_count: int
    drawing_count: int
    alt_chunk_count: int
    field_count: int
    revisions: PackageRevisionCounts
    annotations: PackageAnnotationCounts

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "PackageManifestFacts":
        return cls(
            main_document_uri=d.get("mainDocumentUri"),
            is_strict_ooxml=bool(d["isStrictOoxml"]),
            is_macro_enabled=bool(d["isMacroEnabled"]),
            has_core_properties=bool(d["hasCoreProperties"]),
            has_extended_properties=bool(d["hasExtendedProperties"]),
            has_custom_properties=bool(d["hasCustomProperties"]),
            section_count=int(d["sectionCount"]),
            paragraph_count=int(d["paragraphCount"]),
            table_count=int(d["tableCount"]),
            header_part_count=int(d["headerPartCount"]),
            footer_part_count=int(d["footerPartCount"]),
            footnote_count=int(d["footnoteCount"]),
            endnote_count=int(d["endnoteCount"]),
            style_count=int(d["styleCount"]),
            numbering_definition_count=int(d["numberingDefinitionCount"]),
            theme_part_count=int(d["themePartCount"]),
            media_part_count=int(d["mediaPartCount"]),
            custom_xml_part_count=int(d["customXmlPartCount"]),
            drawing_count=int(d["drawingCount"]),
            alt_chunk_count=int(d["altChunkCount"]),
            field_count=int(d["fieldCount"]),
            revisions=PackageRevisionCounts._from_wire(d.get("revisions", {})),
            annotations=PackageAnnotationCounts._from_wire(d.get("annotations", {})),
        )


@dataclass(frozen=True, slots=True)
class PackageManifestInspectionLimits:
    """Lowered safety ceilings applied while inspecting an untrusted package.

    Each unset field keeps the engine default. Supplying a limit constrains the inspection
    itself, rather than inspecting under the defaults and rejecting only afterwards.
    """

    opc_entries: int | None = None
    expanded_opc_bytes: int | None = None
    xml_part_bytes: int | None = None
    opc_uri_characters: int | None = None
    opc_compression_ratio: float | None = None

    def _to_wire(self) -> dict[str, int | float]:
        wire: dict[str, int | float] = {}
        if self.opc_entries is not None:
            wire["opcEntries"] = self.opc_entries
        if self.expanded_opc_bytes is not None:
            wire["expandedOpcBytes"] = self.expanded_opc_bytes
        if self.xml_part_bytes is not None:
            wire["xmlPartBytes"] = self.xml_part_bytes
        if self.opc_uri_characters is not None:
            wire["opcUriCharacters"] = self.opc_uri_characters
        if self.opc_compression_ratio is not None:
            wire["opcCompressionRatio"] = self.opc_compression_ratio
        return wire


@dataclass(frozen=True, slots=True)
class PackageManifest:
    """Deterministic schema-v1 description of a DOCX/OPC package."""

    schema: str
    schema_version: int
    package_kind: PackageKind
    is_valid: bool
    raw_package_bytes_digest: VerificationDigest
    ordered_opc_content_digest: VerificationDigest | None
    normalized_semantic_digest: VerificationDigest | None
    entries: tuple[PackageManifestEntry, ...]
    content_types: tuple[PackageContentTypeDeclaration, ...]
    relationships: tuple[PackageRelationship, ...]
    facts: PackageManifestFacts
    findings: tuple[VerificationFinding, ...]

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "PackageManifest":
        ordered = d.get("orderedOpcContentDigest")
        semantic = d.get("normalizedSemanticDigest")
        return cls(
            schema=str(d["schema"]),
            schema_version=int(d["schemaVersion"]),
            package_kind=PackageKind(str(d["packageKind"])),
            is_valid=bool(d["isValid"]),
            raw_package_bytes_digest=VerificationDigest._from_wire(
                d["rawPackageBytesDigest"]
            ),
            ordered_opc_content_digest=(
                VerificationDigest._from_wire(ordered)
                if isinstance(ordered, Mapping)
                else None
            ),
            normalized_semantic_digest=(
                VerificationDigest._from_wire(semantic)
                if isinstance(semantic, Mapping)
                else None
            ),
            entries=tuple(PackageManifestEntry._from_wire(x) for x in d.get("entries", ())),
            content_types=tuple(
                PackageContentTypeDeclaration._from_wire(x)
                for x in d.get("contentTypes", ())
            ),
            relationships=tuple(
                PackageRelationship._from_wire(x) for x in d.get("relationships", ())
            ),
            facts=PackageManifestFacts._from_wire(d["facts"]),
            findings=tuple(VerificationFinding._from_wire(x) for x in d.get("findings", ())),
        )


# ---------------------------------------------------------------------------
# Deliverable verification
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class DeliverablePackageIdentity:
    package_kind: str
    manifest_valid: bool
    raw_package_bytes_digest: VerificationDigest
    ordered_opc_content_digest: VerificationDigest | None = None
    normalized_semantic_digest: VerificationDigest | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DeliverablePackageIdentity":
        ordered = d.get("orderedOpcContentDigest")
        semantic = d.get("normalizedSemanticDigest")
        return cls(
            package_kind=str(d["packageKind"]),
            manifest_valid=bool(d["manifestValid"]),
            raw_package_bytes_digest=VerificationDigest._from_wire(
                d["rawPackageBytesDigest"]
            ),
            ordered_opc_content_digest=(
                VerificationDigest._from_wire(ordered)
                if isinstance(ordered, Mapping)
                else None
            ),
            normalized_semantic_digest=(
                VerificationDigest._from_wire(semantic)
                if isinstance(semantic, Mapping)
                else None
            ),
        )


@dataclass(frozen=True, slots=True)
class DeliverableCheckResult:
    check: str
    status: DeliverableCheckStatus
    finding_count: int
    diagnostic: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DeliverableCheckResult":
        return cls(
            check=str(d["check"]),
            status=DeliverableCheckStatus(str(d["status"])),
            finding_count=int(d["findingCount"]),
            diagnostic=d.get("diagnostic"),
        )


@dataclass(frozen=True, slots=True)
class DeliverableFinding:
    finding_id: str
    code: str
    category: DeliverableFindingCategory
    severity: VerificationFindingSeverity
    disposition: DeliverableFindingDisposition
    blocks_delivery: bool
    message: str
    owning_part_uri: str
    remediation: str
    location: ChangeLocation | None = None
    anchor_id: str | None = None
    scope: str | None = None
    x_path: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DeliverableFinding":
        location = d.get("location")
        return cls(
            finding_id=str(d["findingId"]),
            code=str(d["code"]),
            category=DeliverableFindingCategory(str(d["category"])),
            severity=VerificationFindingSeverity(str(d["severity"])),
            disposition=DeliverableFindingDisposition(str(d["disposition"])),
            blocks_delivery=bool(d["blocksDelivery"]),
            message=str(d["message"]),
            owning_part_uri=str(d["owningPartUri"]),
            remediation=str(d["remediation"]),
            location=(
                ChangeLocation._from_wire(location)
                if isinstance(location, Mapping)
                else None
            ),
            anchor_id=d.get("anchorId"),
            scope=d.get("scope"),
            x_path=d.get("xPath"),
        )


@dataclass(frozen=True, slots=True)
class DeliverablePackageChange:
    change_id: str
    kind: DeliverablePackageChangeKind
    location: ChangeLocation
    before_digest: VerificationDigest | None = None
    after_digest: VerificationDigest | None = None
    before_value: str | None = None
    after_value: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DeliverablePackageChange":
        before = d.get("beforeDigest")
        after = d.get("afterDigest")
        return cls(
            change_id=str(d["changeId"]),
            kind=DeliverablePackageChangeKind(str(d["kind"])),
            location=ChangeLocation._from_wire(d["location"]),
            before_digest=(
                VerificationDigest._from_wire(before)
                if isinstance(before, Mapping)
                else None
            ),
            after_digest=(
                VerificationDigest._from_wire(after)
                if isinstance(after, Mapping)
                else None
            ),
            before_value=d.get("beforeValue"),
            after_value=d.get("afterValue"),
        )


@dataclass(frozen=True, slots=True)
class DeliverableSemanticChange:
    change_id: str
    fingerprint: str
    operation: "SemanticChangeOperation"
    family: DeliverableSemanticChangeFamily
    part_uri: str
    path: str
    left_anchor: str | None = None
    right_anchor: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DeliverableSemanticChange":
        return cls(
            change_id=str(d["changeId"]),
            fingerprint=str(d["fingerprint"]),
            operation=SemanticChangeOperation(str(d["operation"])),
            family=DeliverableSemanticChangeFamily(str(d["family"])),
            part_uri=str(d["partUri"]),
            path=str(d["path"]),
            left_anchor=d.get("leftAnchor"),
            right_anchor=d.get("rightAnchor"),
        )


@dataclass(frozen=True, slots=True)
class DeliverableSemanticDelta:
    schema: str
    schema_version: int
    change_count: int
    canonical_digest: VerificationDigest
    changes: tuple[DeliverableSemanticChange, ...]

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DeliverableSemanticDelta":
        schema = str(d.get("schema", ""))
        schema_version = int(d.get("schemaVersion", 0))
        if schema != "docxodus.semantic-changes" or schema_version != 1:
            raise ValueError(
                f"unsupported deliverable semantic schema {schema!r} version {schema_version}"
            )
        changes = tuple(
            DeliverableSemanticChange._from_wire(change)
            for change in d.get("changes", ())
        )
        change_count = int(d.get("changeCount", len(changes)))
        if change_count != len(changes):
            raise ValueError(
                f"deliverable semantic count {change_count} does not match "
                f"{len(changes)} entries"
            )
        return cls(
            schema=schema,
            schema_version=schema_version,
            change_count=change_count,
            canonical_digest=VerificationDigest._from_wire(d["canonicalDigest"]),
            changes=changes,
        )


@dataclass(frozen=True, slots=True)
class DeliverableArtifactMetadata:
    artifact_id: str
    role: DeliverableArtifactRole
    media_type: str
    availability: DeliverableArtifactAvailability
    render_diagnostic_count: int
    byte_length: int | None = None
    digest: VerificationDigest | None = None
    unavailable_reason: str | None = None
    page_count: int | None = None
    renderer_fingerprint: str | None = None
    source_package_digest: VerificationDigest | None = None
    page_map_digest: VerificationDigest | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DeliverableArtifactMetadata":
        digest = d.get("digest")
        source = d.get("sourcePackageDigest")
        page_map = d.get("pageMapDigest")
        return cls(
            artifact_id=str(d["artifactId"]),
            role=DeliverableArtifactRole(str(d["role"])),
            media_type=str(d["mediaType"]),
            availability=DeliverableArtifactAvailability(str(d["availability"])),
            render_diagnostic_count=int(d["renderDiagnosticCount"]),
            byte_length=(
                int(d["byteLength"]) if d.get("byteLength") is not None else None
            ),
            digest=(
                VerificationDigest._from_wire(digest)
                if isinstance(digest, Mapping)
                else None
            ),
            unavailable_reason=d.get("unavailableReason"),
            page_count=(
                int(d["pageCount"]) if d.get("pageCount") is not None else None
            ),
            renderer_fingerprint=d.get("rendererFingerprint"),
            source_package_digest=(
                VerificationDigest._from_wire(source)
                if isinstance(source, Mapping)
                else None
            ),
            page_map_digest=(
                VerificationDigest._from_wire(page_map)
                if isinstance(page_map, Mapping)
                else None
            ),
        )


@dataclass(frozen=True, slots=True)
class DeliveryArtifactVerification:
    """One recorded artifact's independent re-hash verdict from receipt verification."""

    artifact_id: str
    status: DeliveryArtifactVerificationStatus
    expected_length: int | None = None
    actual_length: int | None = None
    expected_digest: VerificationDigest | None = None
    actual_digest: VerificationDigest | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DeliveryArtifactVerification":
        expected_digest = d.get("expectedDigest")
        actual_digest = d.get("actualDigest")
        return cls(
            artifact_id=str(d["artifactId"]),
            status=DeliveryArtifactVerificationStatus(str(d["status"])),
            expected_length=(
                int(d["expectedLength"]) if d.get("expectedLength") is not None else None
            ),
            actual_length=(
                int(d["actualLength"]) if d.get("actualLength") is not None else None
            ),
            expected_digest=(
                VerificationDigest._from_wire(expected_digest)
                if isinstance(expected_digest, Mapping)
                else None
            ),
            actual_digest=(
                VerificationDigest._from_wire(actual_digest)
                if isinstance(actual_digest, Mapping)
                else None
            ),
        )


@dataclass(frozen=True, slots=True)
class DeliveryReceiptVerificationResult:
    """Portable delivery change receipt verification verdict.

    Mirrors the .NET ``DeliveryReceiptVerificationResult`` through the shared
    string-in/string-out facade wire shape (issue #520): the envelope booleans, one
    :class:`DeliveryArtifactVerification` per recorded artifact, and free-form findings.
    """

    is_valid: bool
    receipt_digest_valid: bool
    contract_valid: bool
    citation_bindings_valid: bool
    artifacts: tuple[DeliveryArtifactVerification, ...]
    findings: tuple[str, ...]

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DeliveryReceiptVerificationResult":
        return cls(
            is_valid=bool(d["isValid"]),
            receipt_digest_valid=bool(d["receiptDigestValid"]),
            contract_valid=bool(d["contractValid"]),
            citation_bindings_valid=bool(d["citationBindingsValid"]),
            artifacts=tuple(
                DeliveryArtifactVerification._from_wire(item)
                for item in d.get("artifacts", ())
            ),
            findings=tuple(str(item) for item in d.get("findings", ())),
        )


@dataclass(frozen=True, slots=True)
class DeliverableVerificationResult:
    schema: str
    schema_version: int
    mode: DeliverableVerificationMode
    decision: DeliverableVerificationDecision
    analysis_completed: bool
    baseline_compared: bool
    deliverable_package: DeliverablePackageIdentity
    checks: tuple[DeliverableCheckResult, ...]
    findings: tuple[DeliverableFinding, ...]
    resolved_findings: tuple[DeliverableFinding, ...]
    package_changes: tuple[DeliverablePackageChange, ...]
    companion_artifacts: tuple[DeliverableArtifactMetadata, ...]
    baseline_package: DeliverablePackageIdentity | None = None
    semantic_delta: DeliverableSemanticDelta | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DeliverableVerificationResult":
        schema = str(d.get("schema", ""))
        schema_version = int(d.get("schemaVersion", 0))
        expected = "https://docxodus.dev/schemas/verification/deliverable-verification/v1"
        if schema != expected or schema_version != 1:
            raise ValueError(
                f"unsupported deliverable-verification schema {schema!r} "
                f"version {schema_version}"
            )
        baseline = d.get("baselinePackage")
        semantic = d.get("semanticDelta")
        return cls(
            schema=schema,
            schema_version=schema_version,
            mode=DeliverableVerificationMode(str(d["mode"])),
            decision=DeliverableVerificationDecision(str(d["decision"])),
            analysis_completed=bool(d["analysisCompleted"]),
            baseline_compared=bool(d["baselineCompared"]),
            deliverable_package=DeliverablePackageIdentity._from_wire(
                d["deliverablePackage"]
            ),
            checks=tuple(
                DeliverableCheckResult._from_wire(item)
                for item in d.get("checks", ())
            ),
            findings=tuple(
                DeliverableFinding._from_wire(item)
                for item in d.get("findings", ())
            ),
            resolved_findings=tuple(
                DeliverableFinding._from_wire(item)
                for item in d.get("resolvedFindings", ())
            ),
            package_changes=tuple(
                DeliverablePackageChange._from_wire(item)
                for item in d.get("packageChanges", ())
            ),
            companion_artifacts=tuple(
                DeliverableArtifactMetadata._from_wire(item)
                for item in d.get("companionArtifacts", ())
            ),
            baseline_package=(
                DeliverablePackageIdentity._from_wire(baseline)
                if isinstance(baseline, Mapping)
                else None
            ),
            semantic_delta=(
                DeliverableSemanticDelta._from_wire(semantic)
                if isinstance(semantic, Mapping)
                else None
            ),
        )


# ---------------------------------------------------------------------------
# Redline reversibility proof
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class RedlineProofPackageIdentity:
    """Input or output package identity recorded by the proof."""

    raw_package_bytes_digest: VerificationDigest
    ordered_opc_content_digest: VerificationDigest | None = None
    normalized_whole_package_digest: VerificationDigest | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RedlineProofPackageIdentity":
        ordered = d.get("orderedOpcContentDigest")
        normalized = d.get("normalizedWholePackageDigest")
        return cls(
            raw_package_bytes_digest=VerificationDigest._from_wire(
                d["rawPackageBytesDigest"]
            ),
            ordered_opc_content_digest=(
                VerificationDigest._from_wire(ordered)
                if isinstance(ordered, Mapping)
                else None
            ),
            normalized_whole_package_digest=(
                VerificationDigest._from_wire(normalized)
                if isinstance(normalized, Mapping)
                else None
            ),
        )


@dataclass(frozen=True, slots=True)
class RedlineRevisionIdentity:
    """A stable, part-qualified identity for one native Word revision.

    ``family`` and ``resolution_status`` stay plain strings. The vocabulary is the
    same one ``RevisionListEntry`` reads off the live markup, but the proof wire is
    camelCase (``contentInsert``) where the session listing is snake_case
    (``content_insert``), so multi-word values from the two surfaces do not compare
    equal as strings.
    """

    id: str
    part_uri: str = ""
    scope: str = ""
    type: str = ""
    family: str = "unsupported"
    constituent_ids: tuple[str, ...] = ()
    constituent_keys: tuple[str, ...] = ()
    author: str = "unknown"
    date: str | None = None
    date_utc: str | None = None
    text: str = ""
    anchor_id: str | None = None
    affected_anchor_ids: tuple[str, ...] = ()
    resolution_status: str = "unsupported"
    diagnostic: RevisionDiagnostic | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RedlineRevisionIdentity":
        diagnostic = d.get("diagnostic")
        return cls(
            id=str(d["id"]),
            part_uri=d.get("partUri", ""),
            scope=d.get("scope", ""),
            type=d.get("type", ""),
            family=d.get("family", "unsupported"),
            constituent_ids=tuple(d.get("constituentIds", ())),
            constituent_keys=tuple(d.get("constituentKeys", ())),
            author=d.get("author", "unknown"),
            date=d.get("date"),
            date_utc=d.get("dateUtc"),
            text=d.get("text", ""),
            anchor_id=d.get("anchorId"),
            affected_anchor_ids=tuple(d.get("affectedAnchorIds", ())),
            resolution_status=d.get("resolutionStatus", "unsupported"),
            diagnostic=(
                RevisionDiagnostic._from_wire(diagnostic)
                if isinstance(diagnostic, Mapping)
                else None
            ),
        )


@dataclass(frozen=True, slots=True)
class RedlineRevisionClassification:
    """Classification of a baseline / intended-final / redline revision identity triple."""

    disposition: RedlineRevisionDisposition
    reason: str
    baseline: RedlineRevisionIdentity | None = None
    intended_final: RedlineRevisionIdentity | None = None
    redline: RedlineRevisionIdentity | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RedlineRevisionClassification":
        def identity(key: str) -> RedlineRevisionIdentity | None:
            value = d.get(key)
            return (
                RedlineRevisionIdentity._from_wire(value)
                if isinstance(value, Mapping)
                else None
            )

        return cls(
            disposition=RedlineRevisionDisposition(str(d["disposition"])),
            reason=d.get("reason", ""),
            baseline=identity("baseline"),
            intended_final=identity("intendedFinal"),
            redline=identity("redline"),
        )


@dataclass(frozen=True, slots=True)
class RedlineModeledSemanticComparison:
    """Modeled semantic comparison for one proof path.

    ``available`` and ``change_count`` are explicit so an empty modeled change set is
    never mistaken for complete package equality.
    """

    available: bool
    equivalent: bool | None = None
    schema: str | None = None
    change_count: int | None = None
    diagnostic: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RedlineModeledSemanticComparison":
        equivalent = d.get("equivalent")
        change_count = d.get("changeCount")
        return cls(
            available=bool(d["available"]),
            equivalent=None if equivalent is None else bool(equivalent),
            schema=d.get("schema"),
            change_count=None if change_count is None else int(change_count),
            diagnostic=d.get("diagnostic"),
        )


@dataclass(frozen=True, slots=True)
class RedlinePackageDivergence:
    """One added, removed, or modified package entry on a proof path."""

    kind: RedlinePackageDivergenceKind
    part_uri: str
    occurrence: int = 0
    anchor_id: str | None = None
    applicable_revision_ids: tuple[str, ...] = ()
    expected_raw_digest: VerificationDigest | None = None
    actual_raw_digest: VerificationDigest | None = None
    expected_normalized_digest: VerificationDigest | None = None
    actual_normalized_digest: VerificationDigest | None = None
    has_modeled_semantic_change: bool = False
    unknown_or_unmodeled: bool = False

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RedlinePackageDivergence":
        def digest(key: str) -> VerificationDigest | None:
            value = d.get(key)
            return (
                VerificationDigest._from_wire(value)
                if isinstance(value, Mapping)
                else None
            )

        return cls(
            kind=RedlinePackageDivergenceKind(str(d["kind"])),
            part_uri=str(d["partUri"]),
            occurrence=int(d.get("occurrence", 0)),
            anchor_id=d.get("anchorId"),
            applicable_revision_ids=tuple(d.get("applicableRevisionIds", ())),
            expected_raw_digest=digest("expectedRawDigest"),
            actual_raw_digest=digest("actualRawDigest"),
            expected_normalized_digest=digest("expectedNormalizedDigest"),
            actual_normalized_digest=digest("actualNormalizedDigest"),
            has_modeled_semantic_change=bool(d.get("hasModeledSemanticChange", False)),
            unknown_or_unmodeled=bool(d.get("unknownOrUnmodeled", False)),
        )


@dataclass(frozen=True, slots=True)
class RedlineProofFinding:
    """A structured, actionable proof finding."""

    code: str
    severity: VerificationFindingSeverity
    message: str
    direction: RedlineProofDirection | None = None
    location: ChangeLocation | None = None
    anchor_id: str | None = None
    revision_ids: tuple[str, ...] = ()
    remediation: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RedlineProofFinding":
        direction = d.get("direction")
        location = d.get("location")
        return cls(
            code=str(d["code"]),
            severity=VerificationFindingSeverity(str(d["severity"])),
            message=str(d["message"]),
            direction=(
                RedlineProofDirection(str(direction)) if direction is not None else None
            ),
            location=(
                ChangeLocation._from_wire(location)
                if isinstance(location, Mapping)
                else None
            ),
            anchor_id=d.get("anchorId"),
            revision_ids=tuple(d.get("revisionIds", ())),
            remediation=d.get("remediation"),
        )


@dataclass(frozen=True, slots=True)
class RedlineProofPathResult:
    """Result of accepting or rejecting only the generated revision set."""

    direction: RedlineProofDirection
    completed: bool
    equivalent: bool
    expected_package: RedlineProofPackageIdentity
    modeled_semantic: RedlineModeledSemanticComparison
    requested_revision_ids: tuple[str, ...] = ()
    resolved_revision_ids: tuple[str, ...] = ()
    implicitly_resolved_revision_ids: tuple[str, ...] = ()
    surviving_pre_existing_revisions: tuple[RedlineRevisionIdentity, ...] = ()
    pre_existing_revisions_preserved: bool = False
    normalized_whole_package_equivalent: bool = False
    ordered_opc_content_equivalent: bool = False
    exact_package_bytes_equivalent: bool = False
    divergence_analysis_completed: bool = False
    actual_package: RedlineProofPackageIdentity | None = None
    first_divergence: RedlinePackageDivergence | None = None
    divergences: tuple[RedlinePackageDivergence, ...] = ()
    findings: tuple[RedlineProofFinding, ...] = ()

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RedlineProofPathResult":
        actual = d.get("actualPackage")
        first = d.get("firstDivergence")
        return cls(
            direction=RedlineProofDirection(str(d["direction"])),
            completed=bool(d["completed"]),
            equivalent=bool(d["equivalent"]),
            expected_package=RedlineProofPackageIdentity._from_wire(
                d["expectedPackage"]
            ),
            modeled_semantic=RedlineModeledSemanticComparison._from_wire(
                d["modeledSemantic"]
            ),
            requested_revision_ids=tuple(d.get("requestedRevisionIds", ())),
            resolved_revision_ids=tuple(d.get("resolvedRevisionIds", ())),
            implicitly_resolved_revision_ids=tuple(
                d.get("implicitlyResolvedRevisionIds", ())
            ),
            surviving_pre_existing_revisions=tuple(
                RedlineRevisionIdentity._from_wire(item)
                for item in d.get("survivingPreExistingRevisions", ())
            ),
            pre_existing_revisions_preserved=bool(
                d.get("preExistingRevisionsPreserved", False)
            ),
            normalized_whole_package_equivalent=bool(
                d.get("normalizedWholePackageEquivalent", False)
            ),
            ordered_opc_content_equivalent=bool(
                d.get("orderedOpcContentEquivalent", False)
            ),
            exact_package_bytes_equivalent=bool(
                d.get("exactPackageBytesEquivalent", False)
            ),
            divergence_analysis_completed=bool(
                d.get("divergenceAnalysisCompleted", False)
            ),
            actual_package=(
                RedlineProofPackageIdentity._from_wire(actual)
                if isinstance(actual, Mapping)
                else None
            ),
            first_divergence=(
                RedlinePackageDivergence._from_wire(first)
                if isinstance(first, Mapping)
                else None
            ),
            divergences=tuple(
                RedlinePackageDivergence._from_wire(item)
                for item in d.get("divergences", ())
            ),
            findings=tuple(
                RedlineProofFinding._from_wire(item)
                for item in d.get("findings", ())
            ),
        )


@dataclass(frozen=True, slots=True)
class RedlineReversibilityProof:
    """Canonical schema-v1 proof returned by ``prove_redline_reversibility``.

    ``success`` is true only when both paths completed, reached their expected
    document, and left every pre-existing revision intact.
    """

    schema: str
    schema_version: int
    success: bool
    require_exact_package_bytes: bool
    baseline_package: RedlineProofPackageIdentity
    intended_final_package: RedlineProofPackageIdentity
    redline_package: RedlineProofPackageIdentity
    revision_classifications: tuple[RedlineRevisionClassification, ...] = ()
    accept_to_final: RedlineProofPathResult | None = None
    reject_to_baseline: RedlineProofPathResult | None = None
    findings: tuple[RedlineProofFinding, ...] = ()

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RedlineReversibilityProof":
        schema = str(d.get("schema", ""))
        schema_version = int(d.get("schemaVersion", 0))
        expected = (
            "https://docxodus.dev/schemas/verification/redline-reversibility-proof/v1"
        )
        if schema != expected or schema_version != 1:
            raise ValueError(
                f"unsupported redline-reversibility-proof schema {schema!r} "
                f"version {schema_version}"
            )

        def path(key: str) -> RedlineProofPathResult | None:
            value = d.get(key)
            return (
                RedlineProofPathResult._from_wire(value)
                if isinstance(value, Mapping)
                else None
            )

        return cls(
            schema=schema,
            schema_version=schema_version,
            success=bool(d["success"]),
            require_exact_package_bytes=bool(d["requireExactPackageBytes"]),
            baseline_package=RedlineProofPackageIdentity._from_wire(
                d["baselinePackage"]
            ),
            intended_final_package=RedlineProofPackageIdentity._from_wire(
                d["intendedFinalPackage"]
            ),
            redline_package=RedlineProofPackageIdentity._from_wire(d["redlinePackage"]),
            revision_classifications=tuple(
                RedlineRevisionClassification._from_wire(item)
                for item in d.get("revisionClassifications", ())
            ),
            accept_to_final=path("acceptToFinal"),
            reject_to_baseline=path("rejectToBaseline"),
            findings=tuple(
                RedlineProofFinding._from_wire(item)
                for item in d.get("findings", ())
            ),
        )

# ---------------------------------------------------------------------------
# Anchors
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class Anchor:
    """A block-level address into a Docxodus session.

    The ``id`` is the canonical wire form (e.g. ``"p:body:abcd1234..."``) and
    is what every mutation op consumes. ``kind`` / ``scope`` / ``unid`` are the
    decomposed parts.
    """

    id: str
    kind: str
    scope: str
    unid: str

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "Anchor":
        return cls(id=d["id"], kind=d["kind"], scope=d["scope"], unid=d["unid"])


@dataclass(frozen=True, slots=True)
class TableInsertOptions:
    borderless: bool = False
    cell_contents: tuple[str, ...] = ()
    cell_alignment: str | None = None
    column_widths: tuple[int, ...] = ()

    def to_wire(self) -> dict[str, Any]:
        result: dict[str, Any] = {"borderless": self.borderless}
        if self.cell_contents:
            result["cellContents"] = list(self.cell_contents)
        if self.cell_alignment is not None:
            result["cellAlignment"] = self.cell_alignment
        if self.column_widths:
            result["columnWidths"] = list(self.column_widths)
        return result


@dataclass(frozen=True, slots=True)
class TableBorderSpec:
    scope: str = "all"
    style: str | None = None
    size: int | None = None
    color: str | None = None

    def to_wire(self) -> dict[str, Any]:
        return {key: value for key, value in {
            "scope": self.scope, "style": self.style, "size": self.size, "color": self.color,
        }.items() if value is not None}


@dataclass(frozen=True, slots=True)
class TableRowOptions:
    repeat_header: bool | None = None
    allow_break_across_pages: bool | None = None
    height_twips: int | None = None
    height_rule: TableRowHeightRule = TableRowHeightRule.AT_LEAST


@dataclass(frozen=True, slots=True)
class TableCellMetadata:
    anchor: Anchor
    table_anchor_id: str
    row_anchor_id: str
    row_index: int
    column_index: int
    row_span: int
    column_span: int
    vertical_merge: TableVerticalMergeRole
    paragraph_anchors: tuple[Anchor, ...] = ()

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "TableCellMetadata":
        return cls(
            anchor=Anchor._from_wire(d["anchor"]),
            table_anchor_id=d["tableAnchorId"], row_anchor_id=d["rowAnchorId"],
            row_index=int(d["rowIndex"]), column_index=int(d["columnIndex"]),
            row_span=int(d["rowSpan"]), column_span=int(d["columnSpan"]),
            vertical_merge=TableVerticalMergeRole(d.get("verticalMerge", "none")),
            paragraph_anchors=tuple(Anchor._from_wire(a) for a in d.get("paragraphAnchors", ())),
        )


@dataclass(frozen=True, slots=True)
class TableRowMetadata:
    anchor: Anchor
    table_anchor_id: str
    row_index: int
    grid_before: int
    grid_after: int
    cells: tuple[TableCellMetadata, ...] = ()

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "TableRowMetadata":
        return cls(
            anchor=Anchor._from_wire(d["anchor"]), table_anchor_id=d["tableAnchorId"],
            row_index=int(d["rowIndex"]), grid_before=int(d.get("gridBefore", 0)),
            grid_after=int(d.get("gridAfter", 0)),
            cells=tuple(TableCellMetadata._from_wire(c) for c in d.get("cells", ())),
        )


@dataclass(frozen=True, slots=True)
class TableColumnMetadata:
    anchor: Anchor
    table_anchor_id: str
    column_index: int
    width_twips: int
    is_virtual: bool
    cell_anchor_ids: tuple[str, ...] = ()

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "TableColumnMetadata":
        return cls(
            anchor=Anchor._from_wire(d["anchor"]), table_anchor_id=d["tableAnchorId"],
            column_index=int(d["columnIndex"]), width_twips=int(d.get("widthTwips", 0)),
            is_virtual=bool(d.get("isVirtual", False)),
            cell_anchor_ids=tuple(d.get("cellAnchorIds", ())),
        )


@dataclass(frozen=True, slots=True)
class TableMetadata:
    anchor: Anchor
    columns: tuple[TableColumnMetadata, ...] = ()
    rows: tuple[TableRowMetadata, ...] = ()

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "TableMetadata":
        return cls(
            anchor=Anchor._from_wire(d["anchor"]),
            columns=tuple(TableColumnMetadata._from_wire(c) for c in d.get("columns", ())),
            rows=tuple(TableRowMetadata._from_wire(r) for r in d.get("rows", ())),
        )


@dataclass(frozen=True, slots=True)
class TableAnchorLocation:
    anchor: Anchor
    entity_kind: TableAnchorEntityKind
    row_index: int | None = None
    column_index: int | None = None
    row_span: int | None = None
    column_span: int | None = None
    is_virtual: bool = False

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "TableAnchorLocation":
        return cls(
            anchor=Anchor._from_wire(d["anchor"]),
            entity_kind=TableAnchorEntityKind(d["entityKind"]),
            row_index=d.get("rowIndex"), column_index=d.get("columnIndex"),
            row_span=d.get("rowSpan"), column_span=d.get("columnSpan"),
            is_virtual=bool(d.get("isVirtual", False)),
        )


@dataclass(frozen=True, slots=True)
class RetainedTableAnchor:
    before: TableAnchorLocation
    after: TableAnchorLocation

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RetainedTableAnchor":
        return cls(TableAnchorLocation._from_wire(d["before"]), TableAnchorLocation._from_wire(d["after"]))


@dataclass(frozen=True, slots=True)
class TableAnchorMapping:
    retained: tuple[RetainedTableAnchor, ...] = ()
    added: tuple[TableAnchorLocation, ...] = ()
    invalidated: tuple[TableAnchorLocation, ...] = ()

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "TableAnchorMapping":
        return cls(
            retained=tuple(RetainedTableAnchor._from_wire(x) for x in d.get("retained", ())),
            added=tuple(TableAnchorLocation._from_wire(x) for x in d.get("added", ())),
            invalidated=tuple(TableAnchorLocation._from_wire(x) for x in d.get("invalidated", ())),
        )


@dataclass(frozen=True, slots=True)
class AnchorTarget:
    """Search-result anchor with extra metadata (``partUri``, ``textPreview``)."""

    id: str
    kind: str
    scope: str
    unid: str
    part_uri: str
    text_preview: str
    citation: PageCitation | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "AnchorTarget":
        return cls(
            id=d["id"],
            kind=d["kind"],
            scope=d["scope"],
            unid=d["unid"],
            part_uri=d.get("partUri", ""),
            text_preview=d.get("textPreview", ""),
            citation=PageCitation._from_wire(d["citation"]) if d.get("citation") else None,
        )


@dataclass(frozen=True, slots=True)
class AnchorInfo:
    """Minimal anchor metadata returned by ``get_anchor_info`` / ``get_anchor_infos``."""

    id: str
    kind: str
    scope: str
    text_preview: str
    content_hash: str | None = None
    visible_text: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "AnchorInfo":
        return cls(
            id=d["id"],
            kind=d["kind"],
            scope=d["scope"],
            text_preview=d.get("textPreview", ""),
            content_hash=d.get("contentHash"),
            visible_text=d.get("visibleText"),
        )


@dataclass(frozen=True, slots=True)
class PreconditionTarget:
    """Current target metadata returned when an optimistic guard fails."""

    exists: bool
    anchor_id: str | None = None
    kind: str | None = None
    scope: str | None = None
    content_hash: str | None = None
    visible_text: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "PreconditionTarget":
        return cls(
            exists=bool(d.get("exists", False)),
            anchor_id=d.get("anchorId"),
            kind=d.get("kind"),
            scope=d.get("scope"),
            content_hash=d.get("contentHash"),
            visible_text=d.get("visibleText"),
        )


@dataclass(frozen=True, slots=True)
class PreconditionFailure:
    """Expected/actual detail for ``PRECONDITION_FAILED``."""

    condition: str
    expected: Any
    actual: Any
    current_version: int
    current_target: PreconditionTarget | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "PreconditionFailure":
        target = d.get("currentTarget")
        return cls(
            condition=str(d.get("condition", "")),
            expected=d.get("expected"),
            actual=d.get("actual"),
            current_version=int(d.get("currentVersion", 0)),
            current_target=PreconditionTarget._from_wire(target) if target else None,
        )


@dataclass(frozen=True, slots=True)
class TextRangePrecondition:
    start: int
    length: int
    text: str

    def to_wire(self) -> dict[str, Any]:
        return {"start": self.start, "length": self.length, "text": self.text}


@dataclass(frozen=True, slots=True)
class MutationPreconditions:
    """Optional optimistic guards evaluated immediately before a mutation."""

    expected_version: int | None = None
    anchor_id: str | None = None
    expected_content_hash: str | None = None
    expected_text: str | None = None
    expected_text_range: TextRangePrecondition | None = None
    expected_kind: str | None = None
    expected_scope: str | None = None
    expected_match_count: int | None = None

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {}
        if self.expected_version is not None: out["expectedVersion"] = self.expected_version
        if self.anchor_id is not None: out["anchorId"] = self.anchor_id
        if self.expected_content_hash is not None: out["expectedContentHash"] = self.expected_content_hash
        if self.expected_text is not None: out["expectedText"] = self.expected_text
        if self.expected_text_range is not None: out["expectedTextRange"] = self.expected_text_range.to_wire()
        if self.expected_kind is not None: out["expectedKind"] = self.expected_kind
        if self.expected_scope is not None: out["expectedScope"] = self.expected_scope
        if self.expected_match_count is not None: out["expectedMatchCount"] = self.expected_match_count
        return out


class NumberFormat(str, Enum):
    """Six list formats supported by the list write surface and surfaced
    on ``ListMembership.format``. String-valued so the wire JSON round-trips
    transparently."""

    DECIMAL = "decimal"
    UPPER_LETTER = "upperLetter"
    LOWER_LETTER = "lowerLetter"
    UPPER_ROMAN = "upperRoman"
    LOWER_ROMAN = "lowerRoman"
    BULLET = "bullet"

    @classmethod
    def _from_wire(cls, raw: str) -> "NumberFormat":
        try:
            return cls(raw)
        except ValueError:
            return cls.DECIMAL


@dataclass(frozen=True, slots=True)
class ListMembership:
    """Numbering facts for a list-item paragraph."""

    anchor_id: str
    num_id: int
    abstract_num_id: int
    level: int
    format: NumberFormat
    is_auto_numbered: bool
    from_style: bool
    start: int = 1
    start_override: int | None = None
    level_text: str | None = None
    left_indent_twips: int | None = None
    right_indent_twips: int | None = None
    first_line_indent_twips: int | None = None
    hanging_indent_twips: int | None = None
    generated_label: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "ListMembership":
        return cls(
            anchor_id=d["anchorId"],
            num_id=int(d["numId"]),
            abstract_num_id=int(d["abstractNumId"]),
            level=int(d["level"]),
            format=NumberFormat._from_wire(d["format"]),
            is_auto_numbered=bool(d["isAutoNumbered"]),
            from_style=bool(d["fromStyle"]),
            start=int(d.get("start", 1)),
            start_override=int(d["startOverride"]) if "startOverride" in d else None,
            level_text=d.get("levelText"),
            left_indent_twips=(int(d["leftIndentTwips"]) if "leftIndentTwips" in d else None),
            right_indent_twips=(int(d["rightIndentTwips"]) if "rightIndentTwips" in d else None),
            first_line_indent_twips=(int(d["firstLineIndentTwips"]) if "firstLineIndentTwips" in d else None),
            hanging_indent_twips=(int(d["hangingIndentTwips"]) if "hangingIndentTwips" in d else None),
            generated_label=d.get("generatedLabel"),
        )


@dataclass(frozen=True, slots=True)
class BlockMetadata:
    """Block-level structural metadata."""

    anchor_id: str
    kind: str
    scope: str
    has_inline_formatting: bool
    style_id: str | None = None
    style_name: str | None = None
    outline_level: int | None = None
    list: ListMembership | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "BlockMetadata":
        return cls(
            anchor_id=d["anchorId"],
            kind=d["kind"],
            scope=d["scope"],
            has_inline_formatting=bool(d["hasInlineFormatting"]),
            style_id=d.get("styleId"),
            style_name=d.get("styleName"),
            outline_level=int(d["outlineLevel"]) if "outlineLevel" in d else None,
            list=ListMembership._from_wire(d["list"]) if "list" in d else None,
        )


@dataclass(frozen=True, slots=True)
class HeaderFooterRef:
    """A ``w:headerReference``/``w:footerReference`` on a section.

    Carries the story kind the reference supplies and the URI of the part holding it,
    so a caller can map a :class:`HeaderFooterKind` to a part — and thence to that
    part's projection anchors, which carry the same ``part_uri`` — instead of guessing
    from part-collection order, which carries no kind information.
    """

    kind: HeaderFooterKind
    part_uri: str
    #: True when this section declares no reference of ``kind`` and the story is inherited from
    #: the nearest preceding section that does (ECMA-376 §17.6.17).
    inherited: bool = False

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "HeaderFooterRef":
        return cls(
            kind=HeaderFooterKind(d["kind"]),
            part_uri=d["partUri"],
            inherited=bool(d.get("inherited", False)),
        )


@dataclass(frozen=True, slots=True)
class SectionInfo:
    """Page-layout snapshot for the w:sectPr that governs an anchor."""

    anchor_id: str
    section_unid: str
    page_width_twips: int
    page_height_twips: int
    landscape: bool
    margin_top_twips: int
    margin_bottom_twips: int
    margin_left_twips: int
    margin_right_twips: int
    columns: int
    header_part_uris: tuple[str, ...]
    footer_part_uris: tuple[str, ...]
    #: Header references in declaration order, each with its ``w:type``. Describes exactly
    #: the parts :attr:`header_part_uris` lists, plus the kind each supplies.
    header_refs: tuple[HeaderFooterRef, ...] = ()
    #: Footer references in declaration order, each with its ``w:type``.
    footer_refs: tuple[HeaderFooterRef, ...] = ()
    #: The page number this section starts at (``w:pgNumType/@w:start``). ``None`` when the
    #: section continues the previous section's numbering.
    page_number_start: int | None = None
    #: This section's page-number format (``w:pgNumType/@w:fmt``). ``None`` means Word's default
    #: ``1, 2, 3`` — deliberately not reported as ``DECIMAL``, so a caller can tell "inherits"
    #: from "explicitly decimal" and avoid writing an attribute the document never had.
    page_number_format: NumberFormat | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "SectionInfo":
        return cls(
            anchor_id=d.get("anchorId", ""),
            section_unid=d["sectionUnid"],
            page_width_twips=int(d["pageWidthTwips"]),
            page_height_twips=int(d["pageHeightTwips"]),
            landscape=bool(d["landscape"]),
            margin_top_twips=int(d["marginTopTwips"]),
            margin_bottom_twips=int(d["marginBottomTwips"]),
            margin_left_twips=int(d["marginLeftTwips"]),
            margin_right_twips=int(d["marginRightTwips"]),
            columns=int(d["columns"]),
            header_part_uris=tuple(d["headerPartUris"]),
            footer_part_uris=tuple(d["footerPartUris"]),
            # .get: an older host that predates the refs still decodes.
            header_refs=tuple(HeaderFooterRef._from_wire(r) for r in d.get("headerRefs", ())),
            footer_refs=tuple(HeaderFooterRef._from_wire(r) for r in d.get("footerRefs", ())),
            # Omitted by the host when the attribute is absent — "inherit", not a default value.
            page_number_start=(
                int(d["pageNumberStart"]) if d.get("pageNumberStart") is not None else None
            ),
            page_number_format=(
                NumberFormat._from_wire(d["pageNumberFormat"])
                if d.get("pageNumberFormat") is not None
                else None
            ),
        )


# ---------------------------------------------------------------------------
# Spans + formatting
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class CharSpan:
    """Half-open character range within an anchor's plain-text projection."""

    start: int
    length: int

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "CharSpan":
        return cls(start=int(d["start"]), length=int(d["length"]))

    def to_wire(self) -> dict[str, int]:
        return {"start": self.start, "length": self.length}


@dataclass(frozen=True, slots=True)
class DocumentRange:
    """Two-ended, end-exclusive bookmark range; endpoints must share one story part."""

    start_anchor_id: str
    start_offset: int
    end_anchor_id: str
    end_offset: int

    def to_wire(self) -> dict[str, Any]:
        return {
            "startAnchorId": self.start_anchor_id,
            "startOffset": self.start_offset,
            "endAnchorId": self.end_anchor_id,
            "endOffset": self.end_offset,
        }

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DocumentRange":
        return cls(d["startAnchorId"], int(d["startOffset"]), d["endAnchorId"], int(d["endOffset"]))


@dataclass(frozen=True, slots=True)
class HyperlinkInfo:
    id: str
    kind: HyperlinkKind
    owning_part_uri: str
    scope: str
    anchor_id: str
    span: CharSpan
    text: str
    target: str | None = None
    relationship_id: str | None = None
    relationship_is_external: bool | None = None
    is_broken: bool = False

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "HyperlinkInfo":
        return cls(d["id"], HyperlinkKind(d["kind"]), d["owningPartUri"], d["scope"],
                   d["anchorId"], CharSpan._from_wire(d["span"]), d.get("text", ""),
                   d.get("target"), d.get("relationshipId"), d.get("relationshipIsExternal"),
                   bool(d.get("isBroken", False)))


class ImageBinaryFormat(str, Enum):
    UNKNOWN = "unknown"
    PNG = "png"
    JPEG = "jpeg"
    GIF = "gif"
    BMP = "bmp"
    TIFF = "tiff"
    WEBP = "webp"


class ImageMarkupKind(str, Enum):
    MODERN_DRAWING = "modern_drawing"
    LEGACY_VML = "legacy_vml"
    UNSUPPORTED_DRAWING = "unsupported_drawing"


class ImagePlacement(str, Enum):
    INLINE = "inline"
    FLOATING = "floating"


class ImageWrapMode(str, Enum):
    NONE = "none"
    SQUARE = "square"
    TIGHT = "tight"
    THROUGH = "through"
    TOP_AND_BOTTOM = "top_and_bottom"
    UNKNOWN = "unknown"


class ImageWrapSide(str, Enum):
    BOTH_SIDES = "both_sides"
    LEFT = "left"
    RIGHT = "right"
    LARGEST = "largest"
    UNKNOWN = "unknown"


class ImageHorizontalReference(str, Enum):
    PAGE = "page"
    MARGIN = "margin"
    COLUMN = "column"
    CHARACTER = "character"
    UNKNOWN = "unknown"


class ImageVerticalReference(str, Enum):
    PAGE = "page"
    MARGIN = "margin"
    PARAGRAPH = "paragraph"
    LINE = "line"
    UNKNOWN = "unknown"


class ImageHorizontalAlignment(str, Enum):
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    INSIDE = "inside"
    OUTSIDE = "outside"
    UNKNOWN = "unknown"


class ImageVerticalAlignment(str, Enum):
    TOP = "top"
    CENTER = "center"
    BOTTOM = "bottom"
    INSIDE = "inside"
    OUTSIDE = "outside"
    UNKNOWN = "unknown"


@dataclass(frozen=True, slots=True)
class FloatingImageLayout:
    horizontal_relative_from: ImageHorizontalReference = ImageHorizontalReference.COLUMN
    horizontal_offset_emu: int | None = 0
    horizontal_alignment: ImageHorizontalAlignment | None = None
    vertical_relative_from: ImageVerticalReference = ImageVerticalReference.PARAGRAPH
    vertical_offset_emu: int | None = 0
    vertical_alignment: ImageVerticalAlignment | None = None
    wrap_mode: ImageWrapMode = ImageWrapMode.SQUARE
    wrap_side: ImageWrapSide = ImageWrapSide.BOTH_SIDES
    distance_top_emu: int = 0
    distance_bottom_emu: int = 0
    distance_left_emu: int = 0
    distance_right_emu: int = 0
    relative_height: int = 251658240
    behind_document: bool = False
    locked: bool = False
    layout_in_cell: bool = True
    allow_overlap: bool = True
    raw_horizontal_reference: str | None = None
    raw_vertical_reference: str | None = None
    raw_horizontal_position: str | None = None
    raw_vertical_position: str | None = None
    raw_wrap_mode: str | None = None
    raw_wrap_side: str | None = None
    raw_relative_size_horizontal: str | None = None
    raw_relative_size_vertical: str | None = None
    raw_flag_tokens: Mapping[str, str] | None = None

    def to_wire(self) -> dict[str, Any]:
        return {"horizontalRelativeFrom": self.horizontal_relative_from.value,
                "horizontalOffsetEmu": self.horizontal_offset_emu,
                "horizontalAlignment": (self.horizontal_alignment.value
                                         if self.horizontal_alignment is not None else None),
                "verticalRelativeFrom": self.vertical_relative_from.value,
                "verticalOffsetEmu": self.vertical_offset_emu,
                "verticalAlignment": (self.vertical_alignment.value
                                       if self.vertical_alignment is not None else None),
                "wrapMode": self.wrap_mode.value, "wrapSide": self.wrap_side.value,
                "distanceTopEmu": self.distance_top_emu, "distanceBottomEmu": self.distance_bottom_emu,
                "distanceLeftEmu": self.distance_left_emu, "distanceRightEmu": self.distance_right_emu,
                "relativeHeight": self.relative_height, "behindDocument": self.behind_document,
                "locked": self.locked, "layoutInCell": self.layout_in_cell,
                "allowOverlap": self.allow_overlap}

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "FloatingImageLayout":
        horizontal_alignment = d.get("horizontalAlignment")
        vertical_alignment = d.get("verticalAlignment")
        return cls(
            horizontal_relative_from=ImageHorizontalReference(
                d.get("horizontalRelativeFrom", "unknown")),
            horizontal_offset_emu=d.get("horizontalOffsetEmu"),
            horizontal_alignment=(ImageHorizontalAlignment(horizontal_alignment)
                                  if horizontal_alignment is not None else None),
            vertical_relative_from=ImageVerticalReference(
                d.get("verticalRelativeFrom", "unknown")),
            vertical_offset_emu=d.get("verticalOffsetEmu"),
            vertical_alignment=(ImageVerticalAlignment(vertical_alignment)
                                if vertical_alignment is not None else None),
            wrap_mode=ImageWrapMode(d.get("wrapMode", "unknown")),
            wrap_side=ImageWrapSide(d.get("wrapSide", "unknown")),
            distance_top_emu=int(d.get("distanceTopEmu", 0)),
            distance_bottom_emu=int(d.get("distanceBottomEmu", 0)),
            distance_left_emu=int(d.get("distanceLeftEmu", 0)),
            distance_right_emu=int(d.get("distanceRightEmu", 0)),
            relative_height=int(d.get("relativeHeight", 0)),
            behind_document=bool(d.get("behindDocument", False)),
            locked=bool(d.get("locked", False)),
            layout_in_cell=bool(d.get("layoutInCell", True)),
            allow_overlap=bool(d.get("allowOverlap", True)),
            raw_horizontal_reference=d.get("rawHorizontalReference"),
            raw_vertical_reference=d.get("rawVerticalReference"),
            raw_horizontal_position=d.get("rawHorizontalPosition"),
            raw_vertical_position=d.get("rawVerticalPosition"),
            raw_wrap_mode=d.get("rawWrapMode"),
            raw_wrap_side=d.get("rawWrapSide"),
            raw_relative_size_horizontal=d.get("rawRelativeSizeHorizontal"),
            raw_relative_size_vertical=d.get("rawRelativeSizeVertical"),
            raw_flag_tokens=d.get("rawFlagTokens"),
        )


@dataclass(frozen=True, slots=True)
class ImageInsertOptions:
    placement: ImagePlacement = ImagePlacement.INLINE
    width_points: float | None = None
    height_points: float | None = None
    preserve_aspect: bool = True
    alt_text: str | None = None
    title: str | None = None
    floating_layout: FloatingImageLayout | None = None

    def to_wire(self) -> dict[str, Any]:
        result: dict[str, Any] = {"placement": self.placement.value, "preserveAspect": self.preserve_aspect}
        if self.width_points is not None: result["widthPoints"] = self.width_points
        if self.height_points is not None: result["heightPoints"] = self.height_points
        if self.alt_text is not None: result["altText"] = self.alt_text
        if self.title is not None: result["title"] = self.title
        if self.floating_layout is not None: result["floatingLayout"] = self.floating_layout.to_wire()
        return result


@dataclass(frozen=True, slots=True)
class ImageDimensions:
    width_points: float | None = None
    height_points: float | None = None
    preserve_aspect: bool = True

    def to_wire(self) -> dict[str, Any]:
        result: dict[str, Any] = {"preserveAspect": self.preserve_aspect}
        if self.width_points is not None: result["widthPoints"] = self.width_points
        if self.height_points is not None: result["heightPoints"] = self.height_points
        return result


@dataclass(frozen=True, slots=True)
class ImageOccurrence:
    id: str
    markup_kind: ImageMarkupKind
    placement: ImagePlacement | None
    can_mutate: bool
    unsupported_reason: str | None
    owning_part_uri: str
    scope: str
    anchor_id: str
    span: CharSpan
    relationship_id: str | None
    target_part_uri: str | None
    linked_relationship_id: str | None
    linked_target: str | None
    is_embedded: bool
    is_linked: bool
    is_broken: bool
    media_file_name: str | None
    content_type: str | None
    format: ImageBinaryFormat
    content_type_matches_bytes: bool | None
    intrinsic_width_pixels: int | None
    intrinsic_height_pixels: int | None
    rendered_width_points: float | None
    rendered_height_points: float | None
    alt_text: str | None
    title: str | None
    floating_layout: FloatingImageLayout | None
    floating_layout_supported: bool

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "ImageOccurrence":
        return cls(
            id=d["id"],
            markup_kind=ImageMarkupKind(d["markupKind"]),
            placement=(ImagePlacement(d["placement"]) if d.get("placement") else None),
            can_mutate=bool(d["canMutate"]),
            unsupported_reason=d.get("unsupportedReason"),
            owning_part_uri=d["owningPartUri"],
            scope=d["scope"],
            anchor_id=d["anchorId"],
            span=CharSpan._from_wire(d["span"]),
            relationship_id=d.get("relationshipId"),
            target_part_uri=d.get("targetPartUri"),
            linked_relationship_id=d.get("linkedRelationshipId"),
            linked_target=d.get("linkedTarget"),
            is_embedded=bool(d.get("isEmbedded", False)),
            is_linked=bool(d.get("isLinked", False)),
            is_broken=bool(d.get("isBroken", False)),
            media_file_name=d.get("mediaFileName"),
            content_type=d.get("contentType"),
            format=ImageBinaryFormat(d.get("format", "unknown")),
            content_type_matches_bytes=d.get("contentTypeMatchesBytes"),
            intrinsic_width_pixels=d.get("intrinsicWidthPixels"),
            intrinsic_height_pixels=d.get("intrinsicHeightPixels"),
            rendered_width_points=d.get("renderedWidthPoints"),
            rendered_height_points=d.get("renderedHeightPoints"),
            alt_text=d.get("altText"),
            title=d.get("title"),
            floating_layout=(FloatingImageLayout._from_wire(d["floatingLayout"])
                             if "floatingLayout" in d else None),
            floating_layout_supported=bool(d.get("floatingLayoutSupported", False)),
        )


@dataclass(frozen=True, slots=True)
class ImageFormatCapability:
    format: ImageBinaryFormat
    content_type: str
    can_inspect: bool
    can_insert: bool
    can_replace: bool
    limitation: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "ImageFormatCapability":
        return cls(ImageBinaryFormat(d["format"]), d["contentType"],
                   bool(d["canInspect"]), bool(d["canInsert"]),
                   bool(d["canReplace"]), d.get("limitation"))


@dataclass(frozen=True, slots=True)
class ImageCapabilities:
    schema_version: int
    runtime: str
    formats: tuple[ImageFormatCapability, ...]
    operations: tuple[str, ...]
    mutable_wrap_modes: tuple[ImageWrapMode, ...]
    horizontal_references: tuple[ImageHorizontalReference, ...]
    vertical_references: tuple[ImageVerticalReference, ...]
    max_input_bytes: int
    max_rendered_points: float
    default_dpi: float
    uses_header_parsing_only: bool
    accepts_binary_bytes: bool
    supports_network_fetch: bool
    supports_file_io: bool

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "ImageCapabilities":
        return cls(
            schema_version=int(d["schemaVersion"]),
            runtime=d["runtime"],
            formats=tuple(ImageFormatCapability._from_wire(value)
                          for value in d.get("formats", ())),
            operations=tuple(d.get("operations", ())),
            mutable_wrap_modes=tuple(ImageWrapMode(value)
                                     for value in d.get("mutableWrapModes", ())),
            horizontal_references=tuple(ImageHorizontalReference(value)
                                        for value in d.get("horizontalReferences", ())),
            vertical_references=tuple(ImageVerticalReference(value)
                                      for value in d.get("verticalReferences", ())),
            max_input_bytes=int(d["maxInputBytes"]),
            max_rendered_points=float(d["maxRenderedPoints"]),
            default_dpi=float(d["defaultDpi"]),
            uses_header_parsing_only=bool(d["usesHeaderParsingOnly"]),
            accepts_binary_bytes=bool(d["acceptsBinaryBytes"]),
            supports_network_fetch=bool(d["supportsNetworkFetch"]),
            supports_file_io=bool(d["supportsFileIo"]),
        )


class ContentControlType(str, Enum):
    PLAIN_TEXT = "plain_text"
    RICH_TEXT = "rich_text"
    CHECKBOX = "checkbox"
    DATE = "date"
    DROP_DOWN_LIST = "drop_down_list"
    COMBO_BOX = "combo_box"
    PICTURE = "picture"
    REPEATING_SECTION = "repeating_section"
    REPEATING_SECTION_ITEM = "repeating_section_item"
    UNSUPPORTED = "unsupported"


class ContentControlPlacement(str, Enum):
    INLINE = "inline"
    BLOCK = "block"
    ROW = "row"
    CELL = "cell"
    UNKNOWN = "unknown"


class ContentControlBindingPolicy(str, Enum):
    PRESERVE = "preserve"
    DETACH_TARGET = "detach_target"


@dataclass(frozen=True, slots=True)
class ContentControlFillOptions:
    binding_policy: ContentControlBindingPolicy = ContentControlBindingPolicy.PRESERVE

    def to_wire(self) -> dict[str, Any]:
        return {"bindingPolicy": self.binding_policy.value}


@dataclass(frozen=True, slots=True)
class ContentControlBindingInfo:
    store_item_id: str | None = None
    xpath: str | None = None
    prefix_mappings: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "ContentControlBindingInfo":
        return cls(d.get("storeItemId"), d.get("xpath"), d.get("prefixMappings"))


@dataclass(frozen=True, slots=True)
class ContentControlInfo:
    anchor_id: str
    type: ContentControlType
    placement: ContentControlPlacement
    native_id: str | None
    tag: str | None
    alias: str | None
    lock: str | None
    is_showing_placeholder: bool
    is_bound: bool
    binding: ContentControlBindingInfo | None
    owning_part_uri: str
    scope: str
    parent_anchor_id: str | None
    depth: int
    has_valid_native_id: bool
    has_duplicate_native_id: bool
    can_mutate: bool
    can_detach_target_binding: bool
    unsupported_reason: str | None
    text: str
    item_values: tuple[str, ...]

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "ContentControlInfo":
        return cls(
            anchor_id=d["anchorId"], type=ContentControlType(d["type"]),
            placement=ContentControlPlacement(d["placement"]), native_id=d.get("nativeId"),
            tag=d.get("tag"), alias=d.get("alias"), lock=d.get("lock"),
            is_showing_placeholder=bool(d.get("isShowingPlaceholder", False)),
            is_bound=bool(d.get("isBound", False)),
            binding=(ContentControlBindingInfo._from_wire(d["binding"])
                     if "binding" in d and d["binding"] is not None else None),
            owning_part_uri=d["owningPartUri"], scope=d["scope"],
            parent_anchor_id=d.get("parentAnchorId"), depth=int(d.get("depth", 0)),
            has_valid_native_id=bool(d.get("hasValidNativeId", False)),
            has_duplicate_native_id=bool(d.get("hasDuplicateNativeId", False)),
            can_mutate=bool(d.get("canMutate", False)),
            can_detach_target_binding=bool(d.get("canDetachTargetBinding", False)),
            unsupported_reason=d.get("unsupportedReason"), text=d.get("text", ""),
            item_values=tuple(d.get("itemValues", ())),
        )


@dataclass(frozen=True, slots=True)
class BookmarkRangeSegment:
    owning_part_uri: str
    scope: str
    anchor_id: str
    span: CharSpan
    text: str

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "BookmarkRangeSegment":
        return cls(d["owningPartUri"], d["scope"], d["anchorId"],
                   CharSpan._from_wire(d["span"]), d.get("text", ""))


@dataclass(frozen=True, slots=True)
class BookmarkInfo:
    name: str
    bookmark_id: str
    start_part_uri: str
    start_scope: str
    end_part_uri: str | None
    end_scope: str | None
    range: DocumentRange | None
    segments: tuple[BookmarkRangeSegment, ...]
    text: str
    is_paired: bool
    is_managed: bool
    is_valid: bool
    validation_error: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "BookmarkInfo":
        return cls(d["name"], d["bookmarkId"], d["startPartUri"], d["startScope"],
                   d.get("endPartUri"), d.get("endScope"),
                   DocumentRange._from_wire(d["range"]) if d.get("range") else None,
                   tuple(BookmarkRangeSegment._from_wire(s) for s in d.get("segments", ())),
                   d.get("text", ""), bool(d.get("isPaired", False)),
                   bool(d.get("isManaged", False)), bool(d.get("isValid", False)),
                   d.get("validationError"))


@dataclass(frozen=True, slots=True)
class FormatOp:
    """Set of formatting changes to apply.

    Each field is tri-state: ``True`` to turn on, ``False`` to turn off, ``None``
    to leave unchanged. Strings (``color``, ``run_style``) are passed through;
    ``None`` means "don't change", empty string means "clear".
    """

    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strike: bool | None = None
    code: bool | None = None
    color: str | None = None
    run_style: str | None = None

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {}
        if self.bold is not None: out["bold"] = self.bold
        if self.italic is not None: out["italic"] = self.italic
        if self.underline is not None: out["underline"] = self.underline
        if self.strike is not None: out["strike"] = self.strike
        if self.code is not None: out["code"] = self.code
        if self.color is not None: out["color"] = self.color
        if self.run_style is not None: out["runStyle"] = self.run_style
        return out


@dataclass(frozen=True, slots=True)
class ParagraphBorderEdge:
    """One edge of a paragraph border (a ``w:pBdr`` child — ``w:top``/``w:bottom``) for
    :attr:`ParagraphFormatOp.top_border`/:attr:`ParagraphFormatOp.bottom_border`.

    When an edge is set, fields left ``None`` fall back to the host's defaults (style
    "single", size 6 eighths-of-a-point ≈0.75pt, color "auto", space 1pt).
    """

    style: str | None = None
    size: int | None = None
    color: str | None = None
    space: int | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "ParagraphBorderEdge":
        return cls(
            style=d.get("style"),
            size=int(d["size"]) if "size" in d else None,
            color=d.get("color"),
            space=int(d["space"]) if "space" in d else None,
        )

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {}
        if self.style is not None: out["style"] = self.style
        if self.size is not None: out["size"] = self.size
        if self.color is not None: out["color"] = self.color
        if self.space is not None: out["space"] = self.space
        return out


@dataclass(frozen=True, slots=True)
class ParagraphFormatOp:
    """Paragraph-level formatting changes for ``DocxSession.set_paragraph_format``.

    Each field is tri-state: ``None`` leaves that property unchanged. ``clear_borders=True``
    removes the entire ``w:pBdr`` (all paragraph borders) before applying any
    ``top_border``/``bottom_border`` set in this same op — this is how an S-1-style horizontal
    rule paragraph gets (or loses) its bottom border.

    Indent/spacing fields are absolute values in twips (1440 twips = 1 inch; 20 twips = 1pt)
    and must be >= 0. ``first_line_indent`` (``w:ind/@w:firstLine``) and ``hanging_indent``
    (``w:ind/@w:hanging``) share one either/or slot in Word — setting one removes the other,
    and an op carrying both is rejected with ``INVALID_PARAGRAPH_FORMAT``. ``line_spacing``
    (``w:spacing/@w:line``) is measured per ``line_spacing_rule``: 240ths of a line under
    ``AUTO`` (the default when the rule is omitted — 240 = single, 360 = 1.5x, 480 = double),
    twips under ``EXACT``/``AT_LEAST``; a rule without ``line_spacing`` is rejected.
    """

    alignment: ParagraphAlignment | None = None
    indent_delta: int | None = None
    first_line_indent: int | None = None
    hanging_indent: int | None = None
    spacing_before: int | None = None
    spacing_after: int | None = None
    line_spacing: int | None = None
    line_spacing_rule: LineSpacingRule | None = None
    page_break_before: bool | None = None
    top_border: ParagraphBorderEdge | None = None
    bottom_border: ParagraphBorderEdge | None = None
    clear_borders: bool | None = None

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {}
        if self.alignment is not None: out["alignment"] = self.alignment.value
        if self.indent_delta is not None: out["indentDelta"] = self.indent_delta
        if self.first_line_indent is not None: out["firstLineIndent"] = self.first_line_indent
        if self.hanging_indent is not None: out["hangingIndent"] = self.hanging_indent
        if self.spacing_before is not None: out["spacingBefore"] = self.spacing_before
        if self.spacing_after is not None: out["spacingAfter"] = self.spacing_after
        if self.line_spacing is not None: out["lineSpacing"] = self.line_spacing
        if self.line_spacing_rule is not None: out["lineSpacingRule"] = self.line_spacing_rule.value
        if self.page_break_before is not None: out["pageBreakBefore"] = self.page_break_before
        if self.top_border is not None: out["topBorder"] = self.top_border.to_wire()
        if self.bottom_border is not None: out["bottomBorder"] = self.bottom_border.to_wire()
        if self.clear_borders is not None: out["clearBorders"] = self.clear_borders
        return out


@dataclass(frozen=True, slots=True)
class ParagraphFormatting:
    """High-signal paragraph properties at one direct/effective cascade layer."""

    style_id: str | None = None
    alignment: ParagraphAlignment | None = None
    left_indent_twips: int | None = None
    right_indent_twips: int | None = None
    first_line_indent_twips: int | None = None
    hanging_indent_twips: int | None = None
    spacing_before_twips: int | None = None
    spacing_after_twips: int | None = None
    line_spacing: int | None = None
    line_spacing_rule: LineSpacingRule | None = None
    keep_next: bool | None = None
    keep_lines: bool | None = None
    page_break_before: bool | None = None
    outline_level: int | None = None
    shading_fill: str | None = None
    top_border: ParagraphBorderEdge | None = None
    bottom_border: ParagraphBorderEdge | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "ParagraphFormatting":
        return cls(
            style_id=d.get("styleId"),
            alignment=(ParagraphAlignment(d["alignment"]) if "alignment" in d else None),
            left_indent_twips=int(d["leftIndentTwips"]) if "leftIndentTwips" in d else None,
            right_indent_twips=int(d["rightIndentTwips"]) if "rightIndentTwips" in d else None,
            first_line_indent_twips=int(d["firstLineIndentTwips"]) if "firstLineIndentTwips" in d else None,
            hanging_indent_twips=int(d["hangingIndentTwips"]) if "hangingIndentTwips" in d else None,
            spacing_before_twips=int(d["spacingBeforeTwips"]) if "spacingBeforeTwips" in d else None,
            spacing_after_twips=int(d["spacingAfterTwips"]) if "spacingAfterTwips" in d else None,
            line_spacing=int(d["lineSpacing"]) if "lineSpacing" in d else None,
            line_spacing_rule=(LineSpacingRule(d["lineSpacingRule"]) if "lineSpacingRule" in d else None),
            keep_next=d.get("keepNext"),
            keep_lines=d.get("keepLines"),
            page_break_before=d.get("pageBreakBefore"),
            outline_level=int(d["outlineLevel"]) if "outlineLevel" in d else None,
            shading_fill=d.get("shadingFill"),
            top_border=(ParagraphBorderEdge._from_wire(d["topBorder"]) if "topBorder" in d else None),
            bottom_border=(ParagraphBorderEdge._from_wire(d["bottomBorder"]) if "bottomBorder" in d else None),
        )


@dataclass(frozen=True, slots=True)
class RunFormattingInfo:
    """High-signal character properties; absent direct values remain ``None``."""

    style_id: str | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    underline_style: str | None = None
    strike: bool | None = None
    code: bool | None = None
    color: str | None = None
    highlight: str | None = None
    vert_align: str | None = None
    font_size_pts: float | None = None
    font_family: str | None = None
    caps: bool | None = None
    small_caps: bool | None = None
    hidden: bool | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RunFormattingInfo":
        return cls(
            style_id=d.get("styleId"), bold=d.get("bold"), italic=d.get("italic"),
            underline=d.get("underline"), underline_style=d.get("underlineStyle"),
            strike=d.get("strike"), code=d.get("code"), color=d.get("color"),
            highlight=d.get("highlight"), vert_align=d.get("vertAlign"),
            font_size_pts=float(d["fontSizePts"]) if "fontSizePts" in d else None,
            font_family=d.get("fontFamily"), caps=d.get("caps"),
            small_caps=d.get("smallCaps"), hidden=d.get("hidden"),
        )


@dataclass(frozen=True, slots=True)
class TableStyleFormatting:
    alignment: str | None = None
    width_twips: int | None = None
    indent_twips: int | None = None
    layout: str | None = None
    has_borders: bool | None = None
    cell_shading_fill: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "TableStyleFormatting":
        return cls(
            alignment=d.get("alignment"),
            width_twips=int(d["widthTwips"]) if "widthTwips" in d else None,
            indent_twips=int(d["indentTwips"]) if "indentTwips" in d else None,
            layout=d.get("layout"), has_borders=d.get("hasBorders"),
            cell_shading_fill=d.get("cellShadingFill"),
        )


@dataclass(frozen=True, slots=True)
class StyleInfo:
    """One explicit style definition; ``id`` is accepted by style mutations."""

    id: str
    name: str
    type: str
    is_default: bool
    is_custom: bool
    has_latent_exception: bool
    based_on: str | None = None
    next: str | None = None
    ui_priority: int | None = None
    semi_hidden: bool | None = None
    unhide_when_used: bool | None = None
    quick_format: bool | None = None
    locked: bool | None = None
    resolved_paragraph: ParagraphFormatting | None = None
    resolved_run: RunFormattingInfo | None = None
    resolved_table: TableStyleFormatting | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "StyleInfo":
        return cls(
            id=d["id"], name=d["name"], type=d["type"],
            is_default=bool(d["isDefault"]), is_custom=bool(d["isCustom"]),
            has_latent_exception=bool(d["hasLatentException"]),
            based_on=d.get("basedOn"), next=d.get("next"),
            ui_priority=int(d["uiPriority"]) if "uiPriority" in d else None,
            semi_hidden=d.get("semiHidden"), unhide_when_used=d.get("unhideWhenUsed"),
            quick_format=d.get("quickFormat"), locked=d.get("locked"),
            resolved_paragraph=(ParagraphFormatting._from_wire(d["resolvedParagraph"]) if "resolvedParagraph" in d else None),
            resolved_run=(RunFormattingInfo._from_wire(d["resolvedRun"]) if "resolvedRun" in d else None),
            resolved_table=(TableStyleFormatting._from_wire(d["resolvedTable"]) if "resolvedTable" in d else None),
        )


@dataclass(frozen=True, slots=True)
class InlineSpan:
    """Text-bearing run; ``anchor_id`` + ``span`` can be passed to ``apply_format``."""

    anchor_id: str
    run_unid: str
    span: CharSpan
    text: str
    direct: RunFormattingInfo
    effective: RunFormattingInfo
    content_control_anchor_ids: tuple[str, ...] = ()

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "InlineSpan":
        return cls(
            anchor_id=d["anchorId"], run_unid=d["runUnid"],
            span=CharSpan._from_wire(d["span"]), text=d["text"],
            direct=RunFormattingInfo._from_wire(d["direct"]),
            effective=RunFormattingInfo._from_wire(d["effective"]),
            content_control_anchor_ids=tuple(d.get("contentControlAnchorIds", ())),
        )


@dataclass(frozen=True, slots=True)
class FormattingInspection:
    """Explicitly separated direct and effective formatting for one paragraph."""

    anchor_id: str
    direct_paragraph: ParagraphFormatting
    effective_paragraph: ParagraphFormatting
    runs: tuple[InlineSpan, ...]

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "FormattingInspection":
        return cls(
            anchor_id=d["anchorId"],
            direct_paragraph=ParagraphFormatting._from_wire(d["directParagraph"]),
            effective_paragraph=ParagraphFormatting._from_wire(d["effectiveParagraph"]),
            runs=tuple(InlineSpan._from_wire(v) for v in d.get("runs", ())),
        )


@dataclass(frozen=True, slots=True)
class RunFormatting:
    """Resolved run-level formatting for a ``RunFragment``."""

    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    code: bool = False
    color: str | None = None
    hyperlink_url: str | None = None
    run_style: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RunFormatting":
        return cls(
            bold=bool(d.get("bold", False)),
            italic=bool(d.get("italic", False)),
            underline=bool(d.get("underline", False)),
            strike=bool(d.get("strike", False)),
            code=bool(d.get("code", False)),
            color=d.get("color"),
            hyperlink_url=d.get("hyperlinkUrl"),
            run_style=d.get("runStyle"),
        )


@dataclass(frozen=True, slots=True)
class RunFragment:
    """A single contiguous run inside a ``TextMatch`` slice."""

    unid: str
    text: str
    span_in_element: CharSpan
    formatting: RunFormatting

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RunFragment":
        return cls(
            unid=d["unid"],
            text=d["text"],
            span_in_element=CharSpan._from_wire(d["spanInElement"]),
            formatting=RunFormatting._from_wire(d["formatting"]),
        )


# ---------------------------------------------------------------------------
# Search / Grep results
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class PageMapRect:
    x: float
    y: float
    width: float
    height: float

    def to_wire(self) -> dict[str, float]:
        return {"x": self.x, "y": self.y, "width": self.width, "height": self.height}


@dataclass(frozen=True, slots=True)
class PageMapPage:
    page_number: int
    page_in_section: int
    width: float
    height: float
    page_name: str
    section_index: int | None = None

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {
            "pageNumber": self.page_number,
            "pageInSection": self.page_in_section,
            "width": self.width,
            "height": self.height,
        }
        out["pageName"] = self.page_name
        if self.section_index is not None:
            out["sectionIndex"] = self.section_index
        return out

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "PageMapPage":
        return cls(
            page_number=int(d["pageNumber"]),
            page_in_section=int(d["pageInSection"]),
            width=float(d["width"]),
            height=float(d["height"]),
            page_name=d["pageName"],
            section_index=int(d["sectionIndex"]) if d.get("sectionIndex") is not None else None,
        )


@dataclass(frozen=True, slots=True)
class PageMapFragment:
    fragment_id: str
    anchor_id: str
    fragment_index: int
    page_number: int
    geometry: PageMapRect
    story: str
    in_table_cell: bool = False

    def to_wire(self) -> dict[str, Any]:
        return {
            "fragmentId": self.fragment_id,
            "anchorId": self.anchor_id,
            "fragmentIndex": self.fragment_index,
            "pageNumber": self.page_number,
            "geometry": self.geometry.to_wire(),
            "story": self.story,
            "inTableCell": self.in_table_cell,
        }

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "PageMapFragment":
        g = d["geometry"]
        return cls(
            fragment_id=d["fragmentId"],
            anchor_id=d["anchorId"],
            fragment_index=int(d["fragmentIndex"]),
            page_number=int(d["pageNumber"]),
            geometry=PageMapRect(
                float(g["x"]),
                float(g["y"]),
                float(g["width"]),
                float(g["height"]),
            ),
            story=d["story"],
            in_table_cell=bool(d.get("inTableCell", False)),
        )


@dataclass(frozen=True, slots=True)
class PageMap:
    document_version: int
    renderer_fingerprint: str
    pages: tuple[PageMapPage, ...] = ()
    fragments: tuple[PageMapFragment, ...] = ()
    mode: str = "paginated"
    availability: str = "available"
    schema_version: int = 1

    def to_wire(self) -> dict[str, Any]:
        return {
            "schemaVersion": self.schema_version,
            "mode": self.mode,
            "availability": self.availability,
            "documentVersion": self.document_version,
            "rendererFingerprint": self.renderer_fingerprint,
            "pages": [p.to_wire() for p in self.pages],
            "fragments": [f.to_wire() for f in self.fragments],
        }


@dataclass(frozen=True, slots=True)
class PageCitationRequest:
    document_version: int
    renderer_fingerprint: str

    def to_wire(self) -> dict[str, Any]:
        return {
            "documentVersion": self.document_version,
            "rendererFingerprint": self.renderer_fingerprint,
        }


@dataclass(frozen=True, slots=True)
class PageCitation:
    anchor_id: str
    availability: str
    document_version: int
    renderer_fingerprint: str
    pages: tuple[PageMapPage, ...] = ()
    fragments: tuple[PageMapFragment, ...] = ()
    unavailable_reason: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "PageCitation":
        return cls(
            anchor_id=d["anchorId"],
            availability=d["availability"],
            document_version=int(d["documentVersion"]),
            renderer_fingerprint=d.get("rendererFingerprint", ""),
            pages=tuple(PageMapPage._from_wire(p) for p in d.get("pages", ())),
            fragments=tuple(PageMapFragment._from_wire(f) for f in d.get("fragments", ())),
            unavailable_reason=d.get("unavailableReason"),
        )


@dataclass(frozen=True, slots=True)
class PageMapRegistrationResult:
    success: bool
    error: str | None = None
    message: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "PageMapRegistrationResult":
        return cls(bool(d.get("success")), d.get("error"), d.get("message"))


@dataclass(frozen=True, slots=True)
class PageMapStatus:
    availability: str
    document_version: int
    unavailable_reason: str | None = None
    renderer_fingerprint: str | None = None
    mode: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "PageMapStatus":
        return cls(
            d["availability"],
            int(d["documentVersion"]),
            d.get("unavailableReason"),
            d.get("rendererFingerprint"),
            d.get("mode"),
        )


@dataclass(frozen=True, slots=True)
class TextMatch:
    """A single grep / find-by-text match."""

    text: str
    enclosing_anchor: Anchor
    span: CharSpan
    fragments: tuple[RunFragment, ...] = ()
    context_before: str = ""
    context_after: str = ""
    groups: tuple[str, ...] = ()
    citation: PageCitation | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "TextMatch":
        return cls(
            text=d["text"],
            enclosing_anchor=Anchor._from_wire(d["enclosingAnchor"]),
            span=CharSpan._from_wire(d["span"]),
            fragments=tuple(RunFragment._from_wire(f) for f in d.get("fragments", ())),
            context_before=d.get("contextBefore", ""),
            context_after=d.get("contextAfter", ""),
            groups=tuple(d.get("groups", ())),
            citation=PageCitation._from_wire(d["citation"]) if d.get("citation") else None,
        )


@dataclass(frozen=True, slots=True)
class BlockSlice:
    """One block's contribution to a ``CrossBlockMatch``."""

    anchor: Anchor
    span_in_block: CharSpan
    fragments: tuple[RunFragment, ...] = ()

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "BlockSlice":
        return cls(
            anchor=Anchor._from_wire(d["anchor"]),
            span_in_block=CharSpan._from_wire(d["spanInBlock"]),
            fragments=tuple(RunFragment._from_wire(f) for f in d.get("fragments", ())),
        )


@dataclass(frozen=True, slots=True)
class CrossBlockMatch:
    """A grep match that spans multiple adjacent blocks."""

    text: str
    enclosing_anchors: tuple[Anchor, ...]
    slices: tuple[BlockSlice, ...]
    context_before: str = ""
    context_after: str = ""
    groups: tuple[str, ...] = ()
    citations: tuple[PageCitation, ...] = ()

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "CrossBlockMatch":
        return cls(
            text=d["text"],
            enclosing_anchors=tuple(Anchor._from_wire(a) for a in d.get("enclosingAnchors", ())),
            slices=tuple(BlockSlice._from_wire(s) for s in d.get("slices", ())),
            context_before=d.get("contextBefore", ""),
            context_after=d.get("contextAfter", ""),
            groups=tuple(d.get("groups", ())),
            citations=tuple(PageCitation._from_wire(c) for c in d.get("citations", ())),
        )


# ---------------------------------------------------------------------------
# Placeholders
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class TemplatePlaceholder:
    """A classified bracketed region from ``find_placeholders``."""

    kind: PlaceholderKind
    match: TextMatch
    alternative_kinds: tuple[PlaceholderKind, ...] = ()
    hint: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "TemplatePlaceholder":
        return cls(
            kind=PlaceholderKind(d["kind"]),
            match=TextMatch._from_wire(d["match"]),
            alternative_kinds=tuple(
                PlaceholderKind(k) for k in d.get("alternativeKinds", ())
            ),
            hint=d.get("hint"),
        )


# ---------------------------------------------------------------------------
# Mutation results
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class EditError:
    """Business-level mutation failure inside a successful ``EditResult`` envelope."""

    code: EditErrorCode
    message: str
    anchor_id: str | None = None
    precondition: PreconditionFailure | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "EditError":
        return cls(
            code=EditErrorCode(d["code"]),
            message=d.get("message", ""),
            anchor_id=d.get("anchorId"),
            precondition=PreconditionFailure._from_wire(d["precondition"])
            if d.get("precondition") else None,
        )


@dataclass(frozen=True, slots=True)
class TableMetadataResult:
    success: bool
    metadata: TableMetadata | None = None
    error: EditError | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "TableMetadataResult":
        metadata = d.get("metadata")
        error = d.get("error")
        return cls(bool(d.get("success", False)),
            TableMetadata._from_wire(metadata) if metadata else None,
            EditError._from_wire(error) if error else None)


@dataclass(frozen=True, slots=True)
class TableCellResolutionResult:
    success: bool
    cell: TableCellMetadata | None = None
    error: EditError | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "TableCellResolutionResult":
        cell = d.get("cell")
        error = d.get("error")
        return cls(bool(d.get("success", False)),
            TableCellMetadata._from_wire(cell) if cell else None,
            EditError._from_wire(error) if error else None)


@dataclass(frozen=True, slots=True)
class MarkdownPatch:
    """A scoped markdown re-projection produced by a successful mutation."""

    scope_anchor_id: str
    markdown: str

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "MarkdownPatch":
        return cls(
            scope_anchor_id=d["scopeAnchorId"],
            markdown=d.get("markdown", ""),
        )


@dataclass(frozen=True, slots=True)
class EditResult:
    """The typed envelope returned by every mutation op.

    ``success=False`` here is a *normal business outcome* (anchor not found,
    malformed markdown, etc.). Transport failures raise ``DocxodusTransportError``
    instead of returning an ``EditResult``.
    """

    success: bool
    created: tuple[Anchor, ...] = ()
    removed: tuple[Anchor, ...] = ()
    modified: tuple[Anchor, ...] = ()
    patch: MarkdownPatch | None = None
    error: EditError | None = None
    annotation_id: str | None = None
    table_anchors: TableAnchorMapping | None = None
    hyperlink_id: str | None = None
    bookmark_name: str | None = None
    image_id: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "EditResult":
        patch_d = d.get("patch")
        err_d = d.get("error")
        return cls(
            success=bool(d.get("success", False)),
            created=tuple(Anchor._from_wire(a) for a in d.get("created", ())),
            removed=tuple(Anchor._from_wire(a) for a in d.get("removed", ())),
            modified=tuple(Anchor._from_wire(a) for a in d.get("modified", ())),
            patch=MarkdownPatch._from_wire(patch_d) if patch_d else None,
            error=EditError._from_wire(err_d) if err_d else None,
            annotation_id=d.get("annotationId"),
            table_anchors=TableAnchorMapping._from_wire(d["tableAnchors"])
            if d.get("tableAnchors") else None,
            hyperlink_id=d.get("hyperlinkId"),
            bookmark_name=d.get("bookmarkName"),
            image_id=d.get("imageId"),
        )


@dataclass(frozen=True, slots=True)
class MutationBatchStep:
    """One standardized stdio batch operation and its normal method arguments."""

    operation: str
    args: Mapping[str, Any] = field(default_factory=dict)

    def to_wire(self) -> dict[str, Any]:
        return {"operation": self.operation, "args": dict(self.args)}


@dataclass(frozen=True, slots=True)
class MutationBatchStepResult:
    index: int
    tool: str
    action: str
    success: bool
    rolled_back: bool
    results: tuple[EditResult, ...]

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "MutationBatchStepResult":
        return cls(
            index=int(d["index"]),
            tool=str(d.get("tool", "")),
            action=str(d.get("action", "")),
            success=bool(d.get("success", False)),
            rolled_back=bool(d.get("rolledBack", False)),
            results=tuple(EditResult._from_wire(r) for r in d.get("results", ())),
        )


@dataclass(frozen=True, slots=True)
class MutationBatchFailure:
    index: int
    tool: str
    action: str
    error: EditError
    rolled_back: bool

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "MutationBatchFailure":
        return cls(
            index=int(d["index"]),
            tool=str(d.get("tool", "")),
            action=str(d.get("action", "")),
            error=EditError._from_wire(d["error"]),
            rolled_back=bool(d.get("rolledBack", False)),
        )


_BatchItem = TypeVar("_BatchItem")


@dataclass(frozen=True, slots=True)
class MutationBatchChangeSet(Generic[_BatchItem]):
    """Added, removed, and modified semantic objects produced by apply or preview."""

    added: tuple[_BatchItem, ...] = ()
    removed: tuple[_BatchItem, ...] = ()
    modified: tuple[_BatchItem, ...] = ()

    @classmethod
    def _from_wire(
        cls,
        d: Mapping[str, Any] | None,
        decode: Callable[[Mapping[str, Any]], _BatchItem],
    ) -> "MutationBatchChangeSet[_BatchItem]":
        value = d or {}
        return cls(
            added=tuple(decode(item) for item in value.get("added", ())),
            removed=tuple(decode(item) for item in value.get("removed", ())),
            modified=tuple(decode(item) for item in value.get("modified", ())),
        )


@dataclass(frozen=True, slots=True)
class MutationBatchResult:
    mode: MutationBatchMode
    status: str
    success: bool
    rolled_back: bool
    steps: tuple[MutationBatchStepResult, ...]
    failure: MutationBatchFailure | None = None
    preview: bool = False
    base_version: int = 0
    result_version: int = 0
    #: ``None`` — never ``""`` — when the hash could not be computed (the reason is in
    #: ``warnings``). An absent hash must not compare equal to another absent hash.
    package_hash: str | None = None
    revision_changes: MutationBatchChangeSet[RevisionListEntry] = field(
        default_factory=MutationBatchChangeSet
    )
    comment_changes: MutationBatchChangeSet[CommentListEntry] = field(
        default_factory=MutationBatchChangeSet
    )
    annotation_changes: MutationBatchChangeSet[DocumentAnnotation] = field(
        default_factory=MutationBatchChangeSet
    )
    warnings: tuple[str, ...] = ()
    html: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "MutationBatchResult":
        failure = d.get("failure")
        return cls(
            mode=MutationBatchMode(d.get("mode", "atomic")),
            status=str(d.get("status", "failed")),
            success=bool(d.get("success", False)),
            rolled_back=bool(d.get("rolledBack", False)),
            steps=tuple(MutationBatchStepResult._from_wire(s) for s in d.get("steps", ())),
            failure=MutationBatchFailure._from_wire(failure) if failure else None,
            preview=bool(d.get("preview", False)),
            base_version=int(d.get("baseVersion", 0)),
            result_version=int(d.get("resultVersion", 0)),
            package_hash=(
                None if d.get("packageHash") is None else str(d["packageHash"])
            ),
            revision_changes=MutationBatchChangeSet._from_wire(
                d.get("revisionChanges"), RevisionListEntry._from_wire
            ),
            comment_changes=MutationBatchChangeSet._from_wire(
                d.get("commentChanges"), CommentListEntry._from_wire
            ),
            annotation_changes=MutationBatchChangeSet._from_wire(
                d.get("annotationChanges"), DocumentAnnotation._from_wire
            ),
            warnings=tuple(str(value) for value in d.get("warnings", ())),
            html=d.get("html"),
        )


@dataclass(frozen=True, slots=True)
class CommentListEntry:
    """One native Word comment, in comments-part order — see ``Session.list_comments``.

    ``anchor_id`` addresses the definition (kind ``cmt``) for ``update_comment``/
    ``remove_comment``; ``date`` is the raw ``w:date`` attribute string (``None`` when
    absent); ``text`` is the flattened body (paragraphs joined by a space, the
    ``w:annotationRef`` mark excluded). ``parent_anchor_id`` and ``resolved`` are
    populated only when the comment has a ``commentsExtended.xml`` entry; ``None``
    distinguishes a legacy/flat comment from an explicitly reopened one. The numeric
    ``w:id`` is deliberately not surfaced — comments are addressed by anchor everywhere.
    """

    anchor_id: str
    author: str
    initials: str | None = None
    date: str | None = None
    text: str = ""
    parent_anchor_id: str | None = None
    resolved: bool | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "CommentListEntry":
        return cls(
            anchor_id=d["anchorId"],
            author=d.get("author", "unknown"),
            initials=d.get("initials"),
            date=d.get("date"),
            text=d.get("text", ""),
            parent_anchor_id=d.get("parentAnchorId"),
            resolved=d.get("resolved"),
        )


@dataclass(frozen=True, slots=True)
class RevisionDiagnostic:
    code: str
    message: str

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RevisionDiagnostic":
        return cls(code=d.get("code", ""), message=d.get("message", ""))


@dataclass(frozen=True, slots=True)
class RevisionListEntry:
    """One tracked revision read directly off the live markup — see
    ``Session.list_revisions``.

    ``id`` is an opaque, deterministic ``rev2-…`` identity; ``constituent_ids`` exposes
    the native Word ids and ``constituent_keys`` their QName-qualified identities,
    which stay distinct where roles legally reuse a numeric id. ``family`` identifies the exact atomic operation while ``type``
    is its coarse display class. Part/scope and every affected anchor are included.
    Unsafe native topology remains listed through ``resolution_status`` and
    ``diagnostic`` and fails closed when resolution is requested.
    """

    id: str
    type: str
    author: str
    family: str = "unsupported"
    constituent_ids: tuple[str, ...] = ()
    constituent_keys: tuple[str, ...] = ()
    date: str | None = None
    date_utc: str | None = None
    text: str = ""
    part_uri: str = ""
    scope: str = ""
    anchor_id: str | None = None
    affected_anchors: tuple[Anchor, ...] = ()
    resolution_status: str = "unsupported"
    diagnostic: RevisionDiagnostic | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RevisionListEntry":
        return cls(
            id=d["id"],
            type=d.get("type", ""),
            family=d.get("family", "unsupported"),
            constituent_ids=tuple(d.get("constituentIds", ())),
            constituent_keys=tuple(d.get("constituentKeys", ())),
            author=d.get("author", "unknown"),
            date=d.get("date"),
            date_utc=d.get("dateUtc"),
            text=d.get("text", ""),
            part_uri=d.get("partUri", ""),
            scope=d.get("scope", ""),
            anchor_id=d.get("anchorId"),
            affected_anchors=tuple(Anchor._from_wire(a) for a in d.get("affectedAnchors", ())),
            resolution_status=d.get("resolutionStatus", "unsupported"),
            diagnostic=(RevisionDiagnostic._from_wire(d["diagnostic"])
                        if d.get("diagnostic") is not None else None),
        )


# ---------------------------------------------------------------------------
# Projection
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class MarkdownProjection:
    """The markdown + anchor-index pair returned by ``project`` / ``project_anchor``."""

    markdown: str
    anchor_index: Mapping[str, AnchorTarget]
    page_citations: Mapping[str, PageCitation] = field(default_factory=dict)

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "MarkdownProjection":
        idx = d.get("anchorIndex", {}) or {}
        # The wire entries don't repeat the id key (it's the dict key); rebuild
        # the AnchorTarget by injecting the id from the surrounding key.
        decoded: dict[str, AnchorTarget] = {}
        for anchor_id, entry in idx.items():
            decoded[anchor_id] = AnchorTarget(
                id=anchor_id,
                kind=entry.get("kind", ""),
                scope=entry.get("scope", ""),
                unid=entry.get("unid", ""),
                part_uri=entry.get("partUri", ""),
                text_preview=entry.get("textPreview", ""),
            )
        citations = {
            anchor_id: PageCitation._from_wire(citation)
            for anchor_id, citation in (d.get("pageCitations", {}) or {}).items()
        }
        return cls(markdown=d.get("markdown", ""), anchor_index=decoded, page_citations=citations)


# ---------------------------------------------------------------------------
# Annotations
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class DocumentAnnotation:
    """One stored ``DocumentAnnotation`` returned by ``list_annotations``."""

    id: str
    label_id: str
    label: str
    color: str
    bookmark_name: str
    author: str | None = None
    created: str | None = None  # ISO-8601, kept as a string in v1.
    annotated_text: str | None = None
    metadata: Mapping[str, str] = field(default_factory=dict)

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DocumentAnnotation":
        return cls(
            id=d["id"],
            label_id=d.get("labelId", ""),
            label=d.get("label", ""),
            color=d.get("color", ""),
            bookmark_name=d.get("bookmarkName", ""),
            author=d.get("author"),
            created=d.get("created"),
            annotated_text=d.get("annotatedText"),
            metadata=dict(d.get("metadata", {}) or {}),
        )

    def to_wire(self) -> dict[str, Any]:
        wire: dict[str, Any] = {
            "id": self.id,
            "labelId": self.label_id,
            "label": self.label,
            "color": self.color,
            "bookmarkName": self.bookmark_name,
        }
        if self.author is not None:
            wire["author"] = self.author
        if self.created is not None:
            wire["created"] = self.created
        if self.annotated_text is not None:
            wire["annotatedText"] = self.annotated_text
        if self.metadata:
            wire["metadata"] = dict(self.metadata)
        return wire


@dataclass(frozen=True, slots=True)
class AnnotationUpdate:
    """Partial-update payload for :meth:`DocxSession.update_annotation`.

    ``None`` / missing fields leave the existing value unchanged.
    ``metadata_patch`` is a per-key merge: a non-``None`` value sets the
    key, an explicit ``None`` removes it, a missing key leaves it
    unchanged.
    """

    label_id: str | None = None
    label: str | None = None
    color: str | None = None
    author: str | None = None
    metadata_patch: Mapping[str, str | None] | None = None

    def to_wire(self) -> dict[str, Any]:
        wire: dict[str, Any] = {}
        if self.label_id is not None:
            wire["labelId"] = self.label_id
        if self.label is not None:
            wire["label"] = self.label
        if self.color is not None:
            wire["color"] = self.color
        if self.author is not None:
            wire["author"] = self.author
        if self.metadata_patch is not None:
            # Preserve explicit None values — they mean "remove this key".
            wire["metadataPatch"] = dict(self.metadata_patch)
        return wire


# ---------------------------------------------------------------------------
# Edit summary
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class EditSummary:
    """Aggregate snapshot returned by ``get_edit_summary``.

    The shape of the embedded ``remaining_placeholders`` list mirrors
    ``find_placeholders``; ``bare_underscore_runs`` is a list of
    ``{anchor_id, run_unid, text}`` triples.
    """

    total_anchors: int
    remaining_placeholders: tuple[TemplatePlaceholder, ...] = ()
    bare_underscore_runs: tuple[Mapping[str, Any], ...] = ()
    footnote_count: int = 0
    inline_footnote_ref_count: int = 0
    comment_count: int = 0

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "EditSummary":
        return cls(
            total_anchors=int(d.get("totalAnchors", 0)),
            remaining_placeholders=tuple(
                TemplatePlaceholder._from_wire(p) for p in d.get("remainingPlaceholders", ())
            ),
            bare_underscore_runs=tuple(d.get("bareUnderscoreRuns", ())),
            footnote_count=int(d.get("footnoteCount", 0)),
            inline_footnote_ref_count=int(d.get("inlineFootnoteRefCount", 0)),
            comment_count=int(d.get("commentCount", 0)),
        )


# ---------------------------------------------------------------------------
# Option bundles for find / replace
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class FindOptions:
    """Optional filters for ``find_by_text`` / ``find_all_by_text`` / ``find_by_regex``.

    Mirrors the .NET ``FindOptions`` record's two distinct scope controls:

    * ``scopes`` — a :class:`ProjectionScopes` flag set (coarse "which categories
      of part" filter; wire key ``scopes``). Compose with ``|`` to widen, e.g.
      ``ProjectionScopes.HEADERS | ProjectionScopes.FOOTERS``. Defaults to all
      scopes when unset.
    * ``scope_filter`` — a string naming one specific part such as ``"hdr1"``
      (wire key ``scopeFilter``), applied as a finer post-filter on top of
      ``scopes``. Prefer ``scopes`` for whole-category filtering.
    """

    ignore_case: bool = False
    ignore_whitespace: bool = False
    kind_filter: str | None = None
    scopes: ProjectionScopes | None = None
    scope_filter: str | None = None
    citation: PageCitationRequest | None = None

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {}
        if self.ignore_case: out["ignoreCase"] = True
        if self.ignore_whitespace: out["ignoreWhitespace"] = True
        if self.kind_filter is not None: out["kindFilter"] = self.kind_filter
        if self.scopes is not None: out["scopes"] = int(self.scopes)
        if self.scope_filter is not None: out["scopeFilter"] = self.scope_filter
        if self.citation is not None:
            out["citation"] = self.citation.to_wire()
        return out


@dataclass(frozen=True, slots=True)
class ReplaceOptions:
    """Options for ``replace_text_range``."""

    ignore_case: bool = False
    max_replacements: int | None = None
    expected_match_count: int | None = None
    preconditions: MutationPreconditions | None = None

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {}
        if self.ignore_case: out["ignoreCase"] = True
        if self.max_replacements is not None: out["maxReplacements"] = self.max_replacements
        if self.expected_match_count is not None: out["expectedMatchCount"] = self.expected_match_count
        if self.preconditions is not None: out["preconditions"] = self.preconditions.to_wire()
        return out


# ---------------------------------------------------------------------------
# DOCX → HTML conversion options (see convert_docx_to_html / DocxSession.to_html)
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class HtmlOptions:
    """Options for DOCX→HTML conversion. Mirrors the .NET ``HtmlConversionOptions``.

    Integer-coded modes match the wire contract:
    ``comment_render_mode`` -1=disabled,0=endnote,1=inline,2=margin;
    ``pagination_mode`` 0=none,1=paginated;
    ``annotation_label_mode`` 0=above,1=inline,2=tooltip,3=none.
    """

    page_title: str = "Document"
    css_class_prefix: str = "docx-"
    fabricate_css_classes: bool = True
    additional_css: str = ""
    comment_render_mode: int = -1
    comment_css_class_prefix: str = "comment-"
    pagination_mode: int = 0
    pagination_scale: float = 1.0
    pagination_css_class_prefix: str = "page-"
    render_annotations: bool = False
    annotation_label_mode: int = 0
    annotation_css_class_prefix: str = "annot-"
    render_footnotes_and_endnotes: bool = False
    render_headers_and_footers: bool = False
    render_tracked_changes: bool = False
    show_deleted_content: bool = True
    render_move_operations: bool = True
    render_unsupported_content_placeholders: bool = False
    document_language: str | None = None

    def to_wire(self) -> dict[str, Any]:
        """camelCase keys the host dispatcher's ``ParseHtmlOptions`` reads."""
        wire: dict[str, Any] = {
            "pageTitle": self.page_title,
            "cssClassPrefix": self.css_class_prefix,
            "fabricateCssClasses": self.fabricate_css_classes,
            "additionalCss": self.additional_css,
            "commentRenderMode": self.comment_render_mode,
            "commentCssClassPrefix": self.comment_css_class_prefix,
            "paginationMode": self.pagination_mode,
            "paginationScale": self.pagination_scale,
            "paginationCssClassPrefix": self.pagination_css_class_prefix,
            "renderAnnotations": self.render_annotations,
            "annotationLabelMode": self.annotation_label_mode,
            "annotationCssClassPrefix": self.annotation_css_class_prefix,
            "renderFootnotesAndEndnotes": self.render_footnotes_and_endnotes,
            "renderHeadersAndFooters": self.render_headers_and_footers,
            "renderTrackedChanges": self.render_tracked_changes,
            "showDeletedContent": self.show_deleted_content,
            "renderMoveOperations": self.render_move_operations,
            "renderUnsupportedContentPlaceholders": self.render_unsupported_content_placeholders,
        }
        if self.document_language is not None:
            wire["documentLanguage"] = self.document_language
        return wire


# ---------------------------------------------------------------------------
# DocxDiff — IR diff engine (stateless two-document compare)
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class DocxDiffSettings:
    """Settings for the ``docx_diff_*`` module functions. Mirrors the .NET
    ``DocxDiffSettings``; defaults match it field-for-field, so an unset
    ``DocxDiffSettings()`` reproduces the engine's out-of-the-box behavior.

    Every field is sent only when it differs from the default, so the wire
    object stays minimal and the host applies its own defaults for omitted keys.
    """

    author_for_revisions: str = "Open-Xml-PowerTools"
    deterministic: bool = True
    date_time_for_revisions: str | None = None
    #: Accept every tracked revision already present in both inputs before diffing (default False).
    #: ``preserve_input_revisions`` wins when both policies are enabled.
    pre_accept_input_revisions: bool = False
    #: Preserve the inputs' existing tracked revisions Word-style (default False).
    #: This takes precedence over ``pre_accept_input_revisions``.
    preserve_input_revisions: bool = False
    #: Collapse every tracked-revision author in the output to ``author_for_revisions`` (default False).
    #: Matches how author-coloring renderers (LibreOffice) display Word's single-author compare output
    #: when input revisions preserved under their original author would otherwise render in a second color.
    normalize_revision_authors: bool = False
    case_insensitive: bool = False
    culture: str | None = None
    conflate_breaking_and_nonbreaking_spaces: bool = True
    word_separators: str | None = None
    detect_moves: bool = True
    move_similarity_threshold: float = 0.8
    move_minimum_word_count: int = 3
    revision_granularity: DocxDiffRevisionGranularity = DocxDiffRevisionGranularity.FINE
    format_comparison: DocxDiffFormatComparison = DocxDiffFormatComparison.MODELED_ONLY
    #: Compare header/footer stories (default True — Word Compare's own default).
    #: Changed stories get native tracked-changes markup inside their parts; FINE
    #: revisions carry hdr/ftr-scoped anchors; the edit script carries
    #: ``headerFooterOps``. False ignores header/footer scopes entirely.
    compare_headers_footers: bool = True
    #: Track paragraph-and-above property changes (pPr/tcPr/trPr/tblPr/tblGrid/tblPrEx/sectPr) as native
    #: Word markup. Default True; False restores the untracked-right-apply behavior. Consolidate ignores
    #: block-format changes regardless.
    track_block_format_changes: bool = True
    #: Render a run of >=2 adjacent word-matched modified paragraph pairs via a single cross-paragraph
    #: word+pilcrow token-stream diff (the within-run flat-stream shape decoded from Word's compare
    #: output: retained words may cross the pilcrow, paragraph marks are ins/del stream tokens, and the
    #: output paragraph count follows the token-level interleave). Markup (``docx_diff_compare``) only —
    #: the revision list and edit-script JSON are unaffected. Default False.
    cross_paragraph_token_diff: bool = True

    def to_wire(self) -> dict[str, Any]:
        """camelCase keys the host's ``DocxDiffOps.ParseSettings`` reads. Only
        non-default fields are emitted."""
        wire: dict[str, Any] = {}
        if self.author_for_revisions != "Open-Xml-PowerTools":
            wire["authorForRevisions"] = self.author_for_revisions
        if not self.deterministic:
            wire["deterministic"] = False
        if self.date_time_for_revisions is not None:
            wire["dateTimeForRevisions"] = self.date_time_for_revisions
        if self.pre_accept_input_revisions:
            wire["preAcceptInputRevisions"] = True
        if self.preserve_input_revisions:
            wire["preserveInputRevisions"] = True
        if self.normalize_revision_authors:
            wire["normalizeRevisionAuthors"] = True
        if self.case_insensitive:
            wire["caseInsensitive"] = True
        if self.culture is not None:
            wire["culture"] = self.culture
        if not self.conflate_breaking_and_nonbreaking_spaces:
            wire["conflateBreakingAndNonbreakingSpaces"] = False
        if self.word_separators is not None:
            wire["wordSeparators"] = self.word_separators
        if not self.detect_moves:
            wire["detectMoves"] = False
        if self.move_similarity_threshold != 0.8:
            wire["moveSimilarityThreshold"] = self.move_similarity_threshold
        if self.move_minimum_word_count != 3:
            wire["moveMinimumWordCount"] = self.move_minimum_word_count
        if self.revision_granularity != DocxDiffRevisionGranularity.FINE:
            wire["revisionGranularity"] = int(self.revision_granularity)
        if self.format_comparison != DocxDiffFormatComparison.MODELED_ONLY:
            wire["formatComparison"] = int(self.format_comparison)
        if not self.compare_headers_footers:
            wire["compareHeadersFooters"] = False
        if not self.track_block_format_changes:
            wire["trackBlockFormatChanges"] = False
        if not self.cross_paragraph_token_diff:
            wire["crossParagraphTokenDiff"] = False
        return wire


@dataclass(frozen=True, slots=True)
class DocxDiffFormatChange:
    """Details of a ``FORMAT_CHANGED`` revision — the modeled format fields
    before/after plus the names that differ. Mirrors .NET ``DocxDiffFormatChange``.

    ``scope`` names the property container the change describes: ``"run"`` (the default,
    an rPr-grade report) or one of the block-format-change family scopes ``"paragraph"``
    (pPr), ``"tableCell"``/``"tableRow"``/``"table"`` (tcPr/trPr/tblPr+tblGrid), ``"section"``
    (sectPr). Non-run scopes are reported only under Fine revision granularity."""

    old_properties: Mapping[str, str]
    new_properties: Mapping[str, str]
    changed_property_names: Sequence[str]
    scope: str = "run"

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DocxDiffFormatChange":
        return cls(
            old_properties=dict(d.get("oldProperties") or {}),
            new_properties=dict(d.get("newProperties") or {}),
            changed_property_names=tuple(d.get("changedPropertyNames") or ()),
            scope=str(d.get("scope") or "run"),
        )


@dataclass(frozen=True, slots=True)
class DocxDiffRevision:
    """One consumer revision from ``docx_diff_get_revisions``. Mirrors .NET
    ``DocxDiffRevision`` and carries the IR engine's differentiator: the
    left/right block anchors the revision derives from.

    Anchor presence by ``type`` — each type's PRIMARY anchor is ALWAYS
    present; the opposite anchor MAY also be present for a token-level
    revision. ``INSERTED`` → ``right_anchor`` always (plus ``left_anchor``
    when it is a token-level insert inside a modified block); ``DELETED`` →
    ``left_anchor`` always (plus ``right_anchor`` when token-level);
    ``FORMAT_CHANGED`` → both; ``MOVED`` is EXCLUSIVE: source → ``left_anchor``
    only, destination → ``right_anchor`` only. A token-level revision (an
    insert/delete WITHIN a modified paragraph that exists on both sides)
    carries both enclosing-block anchors; a whole-block insert/delete carries
    only its primary anchor.
    """

    type: DocxDiffRevisionType
    text: str
    author: str
    date: str
    move_group_id: int | None = None
    is_move_source: bool | None = None
    format_change: DocxDiffFormatChange | None = None
    left_anchor: str | None = None
    right_anchor: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DocxDiffRevision":
        fc = d.get("formatChange")
        return cls(
            type=DocxDiffRevisionType._from_wire(d["revisionType"]),
            text=d.get("text", ""),
            author=d.get("author", ""),
            date=d.get("date", ""),
            move_group_id=d.get("moveGroupId"),
            is_move_source=d.get("isMoveSource"),
            format_change=DocxDiffFormatChange._from_wire(fc) if fc else None,
            left_anchor=d.get("leftAnchor"),
            right_anchor=d.get("rightAnchor"),
        )


class SemanticChangeOperation(str, Enum):
    """Stable v1 operation names in a :class:`SemanticChangeSet`."""

    INSERT = "insert"
    DELETE = "delete"
    MOVE = "move"
    MODIFY = "modify"


class SemanticChangeFamily(str, Enum):
    """Stable v1 semantic-change families."""

    TEXT = "text"
    BLOCK_STRUCTURE = "block_structure"
    RUN_FORMATTING = "run_formatting"
    PARAGRAPH_FORMATTING = "paragraph_formatting"
    STYLE = "style"
    NUMBERING = "numbering"
    LIST = "list"
    TABLE = "table"
    TABLE_ROW = "table_row"
    TABLE_CELL = "table_cell"
    TABLE_SPAN = "table_span"
    TABLE_WIDTH = "table_width"
    TABLE_STYLE = "table_style"
    SECTION = "section"
    PAGE_SETUP = "page_setup"
    HEADER = "header"
    FOOTER = "footer"
    FIELD = "field"
    FOOTNOTE = "footnote"
    ENDNOTE = "endnote"
    COMMENT = "comment"
    HYPERLINK = "hyperlink"
    BOOKMARK = "bookmark"
    CONTENT_CONTROL = "content_control"
    IMAGE = "image"
    MEDIA = "media"
    RELATIONSHIP = "relationship"
    REVISION = "revision"
    ANNOTATION = "annotation"
    OPAQUE_PACKAGE_PART = "opaque_package_part"


class SemanticValueKind(str, Enum):
    """Discriminator for the closed semantic value union."""

    ABSENT = "absent"
    STRING = "string"
    BOOLEAN = "boolean"
    INTEGER = "integer"
    DIGEST = "digest"
    OBJECT = "object"
    ARRAY = "array"


class FrozenSemanticObject(Mapping[str, "SemanticValue"]):
    """Immutable, hashable member mapping of an object-kind :class:`SemanticValue`.

    Canonical JSON fixes the member order, so ordered equality and the hash both
    follow it. Reads like a ``dict`` (``value["styleId"]``, iteration, ``len``);
    writes raise ``TypeError`` like every other frozen wire mirror.
    """

    __slots__ = ("_items", "_lookup")

    def __init__(self, items: Sequence[tuple[str, "SemanticValue"]]):
        object.__setattr__(self, "_items", tuple(items))
        object.__setattr__(self, "_lookup", dict(self._items))

    def __getitem__(self, key: str) -> "SemanticValue":
        return self._lookup[key]

    def __iter__(self):
        return iter(name for name, _ in self._items)

    def __len__(self) -> int:
        return len(self._items)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, FrozenSemanticObject):
            return self._items == other._items
        return NotImplemented

    def __hash__(self) -> int:
        return hash(self._items)

    def __repr__(self) -> str:
        return f"FrozenSemanticObject({dict(self._items)!r})"


@dataclass(frozen=True, slots=True)
class SemanticValue:
    """Closed typed value used for one semantic change's before/after state.

    Version 1 integer values stay within JavaScript's exactly representable
    ``[-(2**53-1), 2**53-1]`` range so all supported clients retain identity.
    """

    kind: SemanticValueKind
    value: str | bool | int | FrozenSemanticObject | tuple["SemanticValue", ...] | None = None
    algorithm: str | None = None
    profile: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "SemanticValue":
        kind = SemanticValueKind(str(d["kind"]))
        raw = d.get("value")
        # Schema v1 requires the "value" member for every kind except absent (and
        # "algorithm"/"value" for digests); a payload missing them is corrupt, and
        # decoding it as empty content would be indistinguishable from genuine data.
        if kind is SemanticValueKind.OBJECT:
            if not isinstance(raw, Mapping):
                raise ValueError(f"semantic object requires a members object, got {raw!r}")
            value: Any = FrozenSemanticObject(
                [(str(name), cls._from_wire(member)) for name, member in raw.items()]
            )
        elif kind is SemanticValueKind.ARRAY:
            if raw is None or isinstance(raw, (str, bytes)) or not isinstance(raw, Sequence):
                raise ValueError(f"semantic array requires an items array, got {raw!r}")
            value = tuple(cls._from_wire(member) for member in raw)
        elif kind is SemanticValueKind.ABSENT:
            if raw is not None:
                raise ValueError(f"absent semantic value carries no value, got {raw!r}")
            value = None
        elif kind is SemanticValueKind.INTEGER:
            if isinstance(raw, bool) or not isinstance(raw, int):
                raise ValueError(f"semantic integer must be an integer, got {raw!r}")
            if not -(2**53 - 1) <= raw <= 2**53 - 1:
                raise ValueError(f"semantic integer is outside the v1 safe range: {raw}")
            value = raw
        elif kind is SemanticValueKind.STRING:
            if not isinstance(raw, str):
                raise ValueError(f"semantic string must be a string, got {raw!r}")
            value = raw
        elif kind is SemanticValueKind.BOOLEAN:
            if not isinstance(raw, bool):
                raise ValueError(f"semantic boolean must be a boolean, got {raw!r}")
            value = raw
        else:  # SemanticValueKind.DIGEST
            if not isinstance(raw, str) or not raw:
                raise ValueError(f"semantic digest requires a non-empty value, got {raw!r}")
            algorithm = d.get("algorithm")
            if not isinstance(algorithm, str) or not algorithm:
                raise ValueError(
                    f"semantic digest requires a non-empty algorithm, got {algorithm!r}"
                )
            value = raw
        return cls(
            kind=kind,
            value=value,
            algorithm=d.get("algorithm"),
            profile=d.get("profile"),
        )


@dataclass(frozen=True, slots=True)
class SemanticChange:
    """One deterministic, part- and anchor-addressed semantic change."""

    id: str
    operation: SemanticChangeOperation
    family: SemanticChangeFamily
    part_uri: str
    path: str
    before: SemanticValue
    after: SemanticValue
    left_anchor: str | None = None
    right_anchor: str | None = None
    left_scope: str | None = None
    right_scope: str | None = None
    move_id: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "SemanticChange":
        return cls(
            id=str(d["id"]),
            operation=SemanticChangeOperation(str(d["operation"])),
            family=SemanticChangeFamily(str(d["family"])),
            part_uri=str(d["partUri"]),
            path=str(d["path"]),
            left_anchor=d.get("leftAnchor"),
            right_anchor=d.get("rightAnchor"),
            left_scope=d.get("leftScope"),
            right_scope=d.get("rightScope"),
            move_id=d.get("moveId"),
            before=SemanticValue._from_wire(d["before"]),
            after=SemanticValue._from_wire(d["after"]),
        )


@dataclass(frozen=True, slots=True)
class SemanticChangeSet:
    """Version 1 of the public ``docxodus.semantic-changes`` schema."""

    schema: str
    schema_version: int
    change_count: int
    changes: tuple[SemanticChange, ...]

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "SemanticChangeSet":
        schema = str(d.get("schema", ""))
        schema_version = int(d.get("schemaVersion", 0))
        if schema != "docxodus.semantic-changes" or schema_version != 1:
            raise ValueError(
                f"unsupported semantic-change schema {schema!r} version {schema_version}"
            )
        changes = tuple(SemanticChange._from_wire(change) for change in d.get("changes", ()))
        change_count = int(d.get("changeCount", len(changes)))
        if change_count != len(changes):
            raise ValueError(
                f"semantic-change count {change_count} does not match {len(changes)} entries"
            )
        return cls(schema, schema_version, change_count, changes)


# ---------------------------------------------------------------------------
# DocxDiff consolidate — multi-reviewer composite diff
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class DocxDiffReviewer:
    """One reviewer's document for ``docx_diff_consolidate`` and friends.

    ``document`` is the reviewer's DOCX bytes (their redlined or edited version
    of the base); ``author`` is the display name used for their revisions.
    """

    document: bytes
    author: str


@dataclass(frozen=True, slots=True)
class DocxDiffConsolidateSettings:
    """Settings for the ``docx_diff_consolidate_*`` module functions.

    Wraps a :class:`DocxDiffSettings` for the per-reviewer diffs plus
    ``conflict_resolution`` for how to handle competing edits.

    ``to_wire()`` starts from the embedded diff settings wire dict and adds
    ``conflictResolution`` only when it is non-default (i.e. not ``BASE_WINS``),
    matching the sparse-emit convention used by :meth:`DocxDiffSettings.to_wire`.
    """

    diff: DocxDiffSettings = field(default_factory=DocxDiffSettings)
    conflict_resolution: ConflictResolution = ConflictResolution.BASE_WINS

    def to_wire(self) -> dict[str, Any]:
        wire = self.diff.to_wire()
        if self.conflict_resolution != ConflictResolution.BASE_WINS:
            wire["conflictResolution"] = int(self.conflict_resolution)
        return wire


@dataclass(frozen=True, slots=True)
class DocxDiffConflictCompetitor:
    """One reviewer's competing edit within a :class:`DocxDiffConflict`.

    ``author`` is the reviewer's display name; ``result_text`` is the text
    their edit produces at the conflicting span.
    """

    author: str
    result_text: str

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DocxDiffConflictCompetitor":
        return cls(
            author=d.get("author", ""),
            result_text=d.get("resultText", ""),
        )


@dataclass(frozen=True, slots=True)
class DocxDiffConflict:
    """A conflict produced by ``docx_diff_get_conflicts`` — a base span where two
    or more reviewers made incompatible edits.

    ``id`` is a stable integer index for the conflict within this consolidation.
    ``base_anchor`` is the anchor id in the base document. ``token_start`` /
    ``token_end`` delimit the conflicting token range within that block.
    ``policy`` is the :class:`ConflictResolution` that would be applied by
    ``docx_diff_consolidate`` under the current settings. ``competitors`` lists
    each reviewer's competing edit.
    """

    id: int
    base_anchor: str
    token_start: int
    token_end: int
    policy: ConflictResolution
    competitors: tuple[DocxDiffConflictCompetitor, ...]

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DocxDiffConflict":
        return cls(
            id=int(d["id"]),
            base_anchor=d.get("baseAnchor", ""),
            token_start=int(d.get("tokenStart", 0)),
            token_end=int(d.get("tokenEnd", 0)),
            policy=ConflictResolution(int(d.get("policy", 0))),
            competitors=tuple(
                DocxDiffConflictCompetitor._from_wire(c)
                for c in d.get("competitors", ())
            ),
        )


@dataclass(frozen=True, slots=True)
class DocxDiffConsolidatedRevision:
    """One consumer revision from ``docx_diff_get_consolidated_revisions``.

    Mirrors :class:`DocxDiffRevision` with all of its fields plus an optional
    ``conflict_id`` that links this revision to a :class:`DocxDiffConflict`
    when the revision arose from a conflict resolution decision.
    """

    type: DocxDiffRevisionType
    text: str
    author: str
    date: str
    move_group_id: int | None = None
    is_move_source: bool | None = None
    format_change: DocxDiffFormatChange | None = None
    left_anchor: str | None = None
    right_anchor: str | None = None
    conflict_id: int | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DocxDiffConsolidatedRevision":
        fc = d.get("formatChange")
        return cls(
            type=DocxDiffRevisionType._from_wire(d["revisionType"]),
            text=d.get("text", ""),
            author=d.get("author", ""),
            date=d.get("date", ""),
            move_group_id=d.get("moveGroupId"),
            is_move_source=d.get("isMoveSource"),
            format_change=DocxDiffFormatChange._from_wire(fc) if fc else None,
            left_anchor=d.get("leftAnchor"),
            right_anchor=d.get("rightAnchor"),
            conflict_id=d.get("conflictId"),
        )


# ---------------------------------------------------------------------------
# Fill placeholders (client-side multi-pass loop; see DocxSession.fill_placeholders)
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class FillOptions:
    """Options for :meth:`DocxSession.fill_placeholders`.

    Mirrors the C# ``FillOptions`` record; the multi-pass loop runs client-side
    in Python (no new wire op), so these are consumed locally rather than
    serialized.
    """

    kinds: PlaceholderKinds = PlaceholderKinds.ALL
    scope: ProjectionScopes = ProjectionScopes.BODY
    max_passes: int = 8
    preserve_dollar_prefix: bool = True
    context_chars: int = 80
    boundary: ContextBoundary = ContextBoundary.CHAR


@dataclass(frozen=True, slots=True)
class BulkEditResult:
    """Aggregate result returned by :meth:`DocxSession.fill_placeholders`.

    ``filled`` is the count of picker-returned replacements applied;
    ``skipped`` counts placeholders the picker returned ``None`` for (deduped
    across passes — counted once per placeholder lifetime). ``passes`` is the
    highest iteration pass that actually filled something (``0`` = nothing
    filled, ``1`` = one-shot convergence, higher = multi-pass nested-bracket
    convergence). ``still_present`` is a post-loop ``find_placeholders`` count
    — the trustworthy "is the template done?" check (``skipped > 0 &&
    still_present == 0`` means "picker said no on first sight but later passes
    resolved it"; the canonical case from the NVCA Model COI). Mirrors the
    C# ``BulkEditResult.StillPresent`` added in #191. ``unfilled`` and
    ``errors`` mirror the C# shape.
    """

    filled: int
    skipped: int
    passes: int
    still_present: int
    unfilled: tuple["TemplatePlaceholder", ...] = ()
    errors: tuple["EditError", ...] = ()


# ---------------------------------------------------------------------------
# Session settings (nested projection settings are exposed but not yet
# round-tripped by the bridges per the design doc)
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class WmlToMarkdownConverterSettings:
    """Mirror of the C# converter settings, included on ``DocxSessionSettings``."""

    scopes: ProjectionScopes = ProjectionScopes.ALL
    heading_level_offset: int = 0
    anchor_mode: AnchorRenderMode = AnchorRenderMode.BLOCK
    table_mode: TableRenderMode = TableRenderMode.GFM_WITH_OPAQUE_FALLBACK
    table_inline_cell_max: int = 80
    tracked_changes: TrackedChangeMode = TrackedChangeMode.ACCEPT
    resolve_numbering: bool = True
    empty_paragraphs: EmptyParagraphMode = EmptyParagraphMode.ANCHOR_ONLY
    anchor_id_rendering: AnchorIdRendering = AnchorIdRendering.FULL_UNID

    def to_wire(self) -> dict[str, Any]:
        return {
            "scopes": int(self.scopes),
            "headingLevelOffset": self.heading_level_offset,
            "anchorMode": int(self.anchor_mode),
            "tableMode": int(self.table_mode),
            "tableInlineCellMax": self.table_inline_cell_max,
            "trackedChanges": self.tracked_changes.value,
            "resolveNumbering": self.resolve_numbering,
            "emptyParagraphs": int(self.empty_paragraphs),
            "anchorIdRendering": int(self.anchor_id_rendering),
        }


@dataclass(frozen=True, slots=True)
class DocxSessionSettings:
    """Constructor settings passed to ``open_session(bytes, settings=...)``."""

    #: Maximum undo steps retained (default 20, was 50). Each step is a full document
    #: snapshot, so this is a step count and not a memory bound -- see
    #: ``undo_memory_budget_bytes``.
    undo_depth: int = 20
    #: Approximate ceiling in bytes on memory held by undo/redo snapshots (default 128 MiB).
    #: Oldest history is discarded when exceeded, so on a large document undo may not reach
    #: the full ``undo_depth``; one step is always retained. 0 bounds by depth alone.
    undo_memory_budget_bytes: int = 128 * 1024 * 1024
    validate_raw_ops: bool = False
    tracked_changes: TrackedChangeMode = TrackedChangeMode.ACCEPT
    revision_author: str | None = None
    persist_anchor_ids: bool = False
    smart_quotes: bool = False
    #: Capture the initial projection for ``get_diff`` and retain exact opening
    #: package bytes for ``get_semantic_changes``. Disable to avoid both costs.
    capture_initial_projection: bool = True
    projection_settings: WmlToMarkdownConverterSettings | None = None

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {
            "undoDepth": self.undo_depth,
            "undoMemoryBudgetBytes": self.undo_memory_budget_bytes,
            "validateRawOps": self.validate_raw_ops,
            "trackedChanges": self.tracked_changes.value,
            "persistAnchorIds": self.persist_anchor_ids,
            "smartQuotes": self.smart_quotes,
            "captureInitialProjection": self.capture_initial_projection,
        }
        if self.revision_author is not None:
            out["revisionAuthor"] = self.revision_author
        if self.projection_settings is not None:
            out["projectionSettings"] = self.projection_settings.to_wire()
        return out
