#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

/// <summary>
/// Manages external annotations that don't modify the source document.
/// Provides hash computation, validation, and serialization.
/// </summary>
public static class ExternalAnnotationManager
{
    private static readonly JsonSerializerOptions s_jsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    #region Document Hash

    /// <summary>
    /// Compute the SHA256 hash of a document for integrity validation.
    /// </summary>
    /// <param name="doc">The document to hash.</param>
    /// <returns>Lowercase hex string of the SHA256 hash.</returns>
    public static string ComputeDocumentHash(WmlDocument doc)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        return ComputeDocumentHash(doc.DocumentByteArray);
    }

    /// <summary>
    /// Compute the SHA256 hash of document bytes for integrity validation.
    /// </summary>
    /// <param name="documentBytes">The document bytes to hash.</param>
    /// <returns>Lowercase hex string of the SHA256 hash.</returns>
    public static string ComputeDocumentHash(byte[] documentBytes)
    {
        if (documentBytes == null) throw new ArgumentNullException(nameof(documentBytes));

        var hashBytes = SHA256.HashData(documentBytes);
        return Convert.ToHexString(hashBytes).ToLowerInvariant();
    }

    #endregion

    #region Annotation Set Creation

    /// <summary>
    /// Create an ExternalAnnotationSet from a document.
    /// This extracts the document structure and computes the hash.
    /// </summary>
    /// <param name="doc">The source document.</param>
    /// <param name="documentId">Identifier for the document (filename, UUID, etc.).</param>
    /// <returns>A new ExternalAnnotationSet with document content and hash.</returns>
    public static ExternalAnnotationSet CreateAnnotationSet(WmlDocument doc, string documentId)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (string.IsNullOrEmpty(documentId)) throw new ArgumentException("Document ID is required", nameof(documentId));

        // Use OpenContractExporter to get the document structure
        var export = OpenContractExporter.Export(doc);

        var now = DateTime.UtcNow.ToString("o");
        var hash = ComputeDocumentHash(doc);

        // Create the external annotation set by copying export data
        return new ExternalAnnotationSet
        {
            // Document binding
            DocumentId = documentId,
            DocumentHash = hash,
            CreatedAt = now,
            UpdatedAt = now,
            Version = "1.0",

            // Inherited from OpenContractDocExport
            Title = export.Title,
            Content = export.Content,
            Description = export.Description,
            PageCount = export.PageCount,
            PawlsFileContent = export.PawlsFileContent,
            DocLabels = export.DocLabels,
            LabelledText = export.LabelledText,
            Relationships = export.Relationships,

            // New fields for external annotations
            TextLabels = new Dictionary<string, AnnotationLabel>(),
            DocLabelDefinitions = new Dictionary<string, AnnotationLabel>()
        };
    }

    #endregion

    #region Annotation Creation Helpers

    /// <summary>
    /// Create an annotation from character offsets.
    /// </summary>
    /// <param name="id">Unique identifier for the annotation.</param>
    /// <param name="labelId">Label/category ID for the annotation.</param>
    /// <param name="documentText">Full document text (for extracting rawText).</param>
    /// <param name="startOffset">Start character offset (0-indexed, inclusive).</param>
    /// <param name="endOffset">End character offset (exclusive).</param>
    /// <returns>A new OpenContractsAnnotation with TextSpan targeting.</returns>
    public static OpenContractsAnnotation CreateAnnotation(
        string id,
        string labelId,
        string documentText,
        int startOffset,
        int endOffset)
    {
        if (string.IsNullOrEmpty(id)) throw new ArgumentException("ID is required", nameof(id));
        if (string.IsNullOrEmpty(labelId)) throw new ArgumentException("Label ID is required", nameof(labelId));
        if (documentText == null) throw new ArgumentNullException(nameof(documentText));
        if (startOffset < 0) throw new ArgumentOutOfRangeException(nameof(startOffset), "Start offset must be non-negative");
        if (endOffset < startOffset) throw new ArgumentOutOfRangeException(nameof(endOffset), "End offset must be >= start offset");
        if (endOffset > documentText.Length) throw new ArgumentOutOfRangeException(nameof(endOffset), "End offset exceeds document length");

        var rawText = documentText.Substring(startOffset, endOffset - startOffset);

        return new OpenContractsAnnotation
        {
            Id = id,
            AnnotationLabel = labelId,
            RawText = rawText,
            Page = 0, // Page will be computed during projection if needed
            AnnotationJson = new TextSpan
            {
                Id = id,
                Start = startOffset,
                End = endOffset,
                Text = rawText
            },
            AnnotationType = "text",
            Structural = false
        };
    }

    /// <summary>
    /// Create an annotation by searching for text in the document.
    /// </summary>
    /// <param name="id">Unique identifier for the annotation.</param>
    /// <param name="labelId">Label/category ID for the annotation.</param>
    /// <param name="documentText">Full document text to search in.</param>
    /// <param name="searchText">Text to search for.</param>
    /// <param name="occurrence">Which occurrence to use (1-based, default: 1).</param>
    /// <returns>A new OpenContractsAnnotation, or null if text not found.</returns>
    public static OpenContractsAnnotation? CreateAnnotationFromSearch(
        string id,
        string labelId,
        string documentText,
        string searchText,
        int occurrence = 1)
    {
        if (string.IsNullOrEmpty(id)) throw new ArgumentException("ID is required", nameof(id));
        if (string.IsNullOrEmpty(labelId)) throw new ArgumentException("Label ID is required", nameof(labelId));
        if (documentText == null) throw new ArgumentNullException(nameof(documentText));
        if (string.IsNullOrEmpty(searchText)) throw new ArgumentException("Search text is required", nameof(searchText));
        if (occurrence < 1) throw new ArgumentOutOfRangeException(nameof(occurrence), "Occurrence must be >= 1");

        var offsets = FindTextOccurrences(documentText, searchText);

        if (occurrence > offsets.Count)
        {
            return null; // Occurrence not found
        }

        var (start, end) = offsets[occurrence - 1];
        return CreateAnnotation(id, labelId, documentText, start, end);
    }

    /// <summary>
    /// Find all occurrences of a text string in the document.
    /// </summary>
    /// <param name="documentText">Full document text to search in.</param>
    /// <param name="searchText">Text to search for.</param>
    /// <param name="maxResults">Maximum number of results (default: 100).</param>
    /// <returns>List of (startOffset, endOffset) tuples.</returns>
    public static List<(int start, int end)> FindTextOccurrences(
        string documentText,
        string searchText,
        int maxResults = 100)
    {
        if (documentText == null) throw new ArgumentNullException(nameof(documentText));
        if (string.IsNullOrEmpty(searchText)) return new List<(int, int)>();

        var results = new List<(int start, int end)>();
        var index = 0;

        while (results.Count < maxResults)
        {
            index = documentText.IndexOf(searchText, index, StringComparison.Ordinal);
            if (index < 0) break;

            results.Add((index, index + searchText.Length));
            index += 1; // Move past start of this match to find overlapping matches
        }

        return results;
    }

    #endregion

    #region Validation

    /// <summary>
    /// Validate an external annotation set against a document.
    /// Checks hash match and verifies each annotation's text still matches.
    /// </summary>
    /// <param name="doc">The document to validate against.</param>
    /// <param name="annotationSet">The annotation set to validate.</param>
    /// <returns>Validation result with any issues found.</returns>
    public static ExternalAnnotationValidationResult Validate(
        WmlDocument doc,
        ExternalAnnotationSet annotationSet)
    {
        if (doc == null) throw new ArgumentNullException(nameof(doc));
        if (annotationSet == null) throw new ArgumentNullException(nameof(annotationSet));

        var result = new ExternalAnnotationValidationResult
        {
            IsValid = true,
            HashMismatch = false,
            Issues = new List<ExternalAnnotationValidationIssue>()
        };

        // Check document hash
        var currentHash = ComputeDocumentHash(doc);
        if (!string.Equals(currentHash, annotationSet.DocumentHash, StringComparison.OrdinalIgnoreCase))
        {
            result.HashMismatch = true;
            result.IsValid = false;
        }

        // Get current document text
        var export = OpenContractExporter.Export(doc);
        var documentText = export.Content;

        // Validate each annotation
        foreach (var annotation in annotationSet.LabelledText)
        {
            // Skip structural annotations (they don't have user-specified text)
            if (annotation.Structural) continue;

            var issue = ValidateAnnotation(annotation, documentText, annotationSet.TextLabels);
            if (issue != null)
            {
                result.Issues.Add(issue);
                result.IsValid = false;
            }
        }

        return result;
    }

    private static ExternalAnnotationValidationIssue? ValidateAnnotation(
        OpenContractsAnnotation annotation,
        string documentText,
        Dictionary<string, AnnotationLabel> labels)
    {
        var annotationId = annotation.Id ?? "(unnamed)";

        // Check if label exists
        if (!string.IsNullOrEmpty(annotation.AnnotationLabel) &&
            labels.Count > 0 &&
            !labels.ContainsKey(annotation.AnnotationLabel))
        {
            return new ExternalAnnotationValidationIssue
            {
                AnnotationId = annotationId,
                IssueType = "MissingLabel",
                Description = $"Label '{annotation.AnnotationLabel}' is not defined in TextLabels"
            };
        }

        // Check text span targeting
        if (annotation.AnnotationJson is TextSpan textSpan)
        {
            // Check bounds
            if (textSpan.Start < 0 || textSpan.End < textSpan.Start)
            {
                return new ExternalAnnotationValidationIssue
                {
                    AnnotationId = annotationId,
                    IssueType = "OutOfBounds",
                    Description = $"Invalid offsets: start={textSpan.Start}, end={textSpan.End}"
                };
            }

            if (textSpan.End > documentText.Length)
            {
                return new ExternalAnnotationValidationIssue
                {
                    AnnotationId = annotationId,
                    IssueType = "OutOfBounds",
                    Description = $"End offset {textSpan.End} exceeds document length {documentText.Length}"
                };
            }

            // Check text match
            var actualText = documentText.Substring(textSpan.Start, textSpan.End - textSpan.Start);
            var expectedText = textSpan.Text ?? annotation.RawText;

            if (!string.Equals(actualText, expectedText, StringComparison.Ordinal))
            {
                return new ExternalAnnotationValidationIssue
                {
                    AnnotationId = annotationId,
                    IssueType = "TextMismatch",
                    Description = "Text at annotation offsets does not match expected text",
                    ExpectedText = expectedText,
                    ActualText = actualText
                };
            }
        }

        return null;
    }

    #endregion

    #region Serialization

    /// <summary>
    /// Serialize an annotation set to JSON.
    /// </summary>
    /// <param name="set">The annotation set to serialize.</param>
    /// <returns>JSON string representation.</returns>
    public static string SerializeToJson(ExternalAnnotationSet set)
    {
        if (set == null) throw new ArgumentNullException(nameof(set));
        return JsonSerializer.Serialize(set, s_jsonOptions);
    }

    /// <summary>
    /// Deserialize an annotation set from JSON.
    /// </summary>
    /// <param name="json">JSON string to deserialize.</param>
    /// <returns>The deserialized annotation set, or null if invalid.</returns>
    public static ExternalAnnotationSet? DeserializeFromJson(string json)
    {
        if (string.IsNullOrEmpty(json)) return null;

        try
        {
            return JsonSerializer.Deserialize<ExternalAnnotationSet>(json, s_jsonOptions);
        }
        catch (JsonException)
        {
            return null;
        }
    }

    #endregion
}
