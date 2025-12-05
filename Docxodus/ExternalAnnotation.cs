#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Docxodus;

/// <summary>
/// External annotation set that extends OpenContractDocExport with document binding.
/// Uses the SAME annotation structure as OpenContracts for interoperability.
///
/// This enables storing annotations externally (in JSON/database) without modifying
/// the source DOCX file. The documentHash allows validation that annotations
/// still match the document content.
/// </summary>
public class ExternalAnnotationSet : OpenContractDocExport
{
    /// <summary>
    /// Unique identifier for the source document (filename, UUID, or external reference).
    /// </summary>
    public string DocumentId { get; set; } = "";

    /// <summary>
    /// SHA256 hash of the source document for integrity validation.
    /// Required - used to detect if annotations are stale.
    /// </summary>
    public string DocumentHash { get; set; } = "";

    /// <summary>
    /// ISO 8601 timestamp when this annotation set was created.
    /// </summary>
    public string CreatedAt { get; set; } = "";

    /// <summary>
    /// ISO 8601 timestamp when this annotation set was last modified.
    /// </summary>
    public string UpdatedAt { get; set; } = "";

    /// <summary>
    /// Version of the external annotation format (for future migrations).
    /// </summary>
    public string Version { get; set; } = "1.0";

    /// <summary>
    /// Text label definitions keyed by label ID.
    /// These define the annotation categories available for text spans.
    /// </summary>
    public Dictionary<string, AnnotationLabel> TextLabels { get; set; } = new();

    /// <summary>
    /// Document label definitions keyed by label ID.
    /// Note: DocLabels (List&lt;string&gt;) is inherited from OpenContractDocExport
    /// and contains the applied document-level labels. This dictionary provides
    /// the full label definitions.
    /// </summary>
    public Dictionary<string, AnnotationLabel> DocLabelDefinitions { get; set; } = new();
}

/// <summary>
/// Result of validating an external annotation set against a document.
/// </summary>
public class ExternalAnnotationValidationResult
{
    /// <summary>
    /// True if the annotation set is valid for the document.
    /// False if there are any errors (hash mismatch or annotation issues).
    /// </summary>
    public bool IsValid { get; set; }

    /// <summary>
    /// True if the document hash doesn't match, indicating the document
    /// may have been modified since the annotations were created.
    /// </summary>
    public bool HashMismatch { get; set; }

    /// <summary>
    /// List of specific issues found during validation.
    /// </summary>
    public List<ExternalAnnotationValidationIssue> Issues { get; set; } = new();
}

/// <summary>
/// A single validation issue found when validating an external annotation set.
/// </summary>
public class ExternalAnnotationValidationIssue
{
    /// <summary>
    /// ID of the annotation with the issue.
    /// </summary>
    public string AnnotationId { get; set; } = "";

    /// <summary>
    /// Type of issue: "TextMismatch", "OutOfBounds", or "MissingLabel".
    /// </summary>
    public string IssueType { get; set; } = "";

    /// <summary>
    /// Human-readable description of the issue.
    /// </summary>
    public string Description { get; set; } = "";

    /// <summary>
    /// For TextMismatch: the text that was expected (stored in annotation).
    /// </summary>
    public string? ExpectedText { get; set; }

    /// <summary>
    /// For TextMismatch: the actual text found at the annotation's offsets.
    /// </summary>
    public string? ActualText { get; set; }
}

/// <summary>
/// Settings for projecting external annotations onto HTML.
/// </summary>
public class ExternalAnnotationProjectionSettings
{
    /// <summary>
    /// CSS class prefix for annotation elements (default: "ext-annot-").
    /// </summary>
    public string CssClassPrefix { get; set; } = "ext-annot-";

    /// <summary>
    /// How to display annotation labels.
    /// </summary>
    public AnnotationLabelMode LabelMode { get; set; } = AnnotationLabelMode.Above;

    /// <summary>
    /// Whether to include annotation metadata as data attributes.
    /// </summary>
    public bool IncludeMetadata { get; set; } = true;

    /// <summary>
    /// Whether to validate annotations before projection.
    /// If true, invalid annotations will be skipped.
    /// </summary>
    public bool ValidateBeforeProjection { get; set; } = true;
}
