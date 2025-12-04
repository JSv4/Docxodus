// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    /// <summary>
    /// Tests for the External Annotation system (Issue #57).
    /// Tests cover hash computation, annotation creation, validation, and projection.
    /// </summary>
    public class ExternalAnnotationTests
    {
        private static readonly DirectoryInfo TestFilesDir = new DirectoryInfo("../../../../TestFiles/");

        #region Helper Methods

        /// <summary>
        /// Creates a properly structured test document with all required parts.
        /// </summary>
        private static WmlDocument CreateTestDocument(Action<Body> configureBody)
        {
            using var ms = new MemoryStream();
            using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = wDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = new Body();
                mainPart.Document.Body = body;

                // Add StyleDefinitionsPart (required for many operations)
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles();

                // Add DocumentSettingsPart
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings();

                // Configure body content
                configureBody(body);

                mainPart.Document.Save();
            }

            return new WmlDocument("test.docx", ms.ToArray());
        }

        private static WmlDocument CreateSimpleTestDocument(string text)
        {
            return CreateTestDocument(body =>
            {
                body.AppendChild(new Paragraph(new Run(new Text(text))));
            });
        }

        #endregion

        #region Document Hash Tests

        [Fact]
        public void EA001_ComputeDocumentHash_SameDocument_ReturnsSameHash()
        {
            // Arrange
            var doc = CreateSimpleTestDocument("Hello, world!");

            // Act
            var hash1 = ExternalAnnotationManager.ComputeDocumentHash(doc);
            var hash2 = ExternalAnnotationManager.ComputeDocumentHash(doc);

            // Assert
            Assert.NotNull(hash1);
            Assert.NotEmpty(hash1);
            Assert.Equal(hash1, hash2);
            Assert.Equal(64, hash1.Length); // SHA256 produces 64 hex characters
        }

        [Fact]
        public void EA002_ComputeDocumentHash_ModifiedDocument_ReturnsDifferentHash()
        {
            // Arrange
            var doc1 = CreateSimpleTestDocument("Hello, world!");
            var doc2 = CreateSimpleTestDocument("Hello, world!!");

            // Act
            var hash1 = ExternalAnnotationManager.ComputeDocumentHash(doc1);
            var hash2 = ExternalAnnotationManager.ComputeDocumentHash(doc2);

            // Assert
            Assert.NotEqual(hash1, hash2);
        }

        [Fact]
        public void EA003_ComputeDocumentHash_ByteArray_MatchesWmlDocument()
        {
            // Arrange
            var doc = CreateSimpleTestDocument("Test content");

            // Act
            var hashFromDoc = ExternalAnnotationManager.ComputeDocumentHash(doc);
            var hashFromBytes = ExternalAnnotationManager.ComputeDocumentHash(doc.DocumentByteArray);

            // Assert
            Assert.Equal(hashFromDoc, hashFromBytes);
        }

        #endregion

        #region Annotation Set Creation Tests

        [Fact]
        public void EA010_CreateAnnotationSet_ReturnsValidSet()
        {
            // Arrange
            var doc = CreateSimpleTestDocument("Hello, world!");
            var documentId = "test-doc-001";

            // Act
            var set = ExternalAnnotationManager.CreateAnnotationSet(doc, documentId);

            // Assert
            Assert.NotNull(set);
            Assert.Equal(documentId, set.DocumentId);
            Assert.NotEmpty(set.DocumentHash);
            Assert.NotEmpty(set.CreatedAt);
            Assert.NotEmpty(set.UpdatedAt);
            Assert.Equal("1.0", set.Version);
            Assert.NotNull(set.Content);
            Assert.Contains("Hello, world!", set.Content);
        }

        [Fact]
        public void EA011_CreateAnnotationSet_InheritsFromOpenContractDocExport()
        {
            // Arrange
            var doc = CreateTestDocument(body =>
            {
                body.AppendChild(new Paragraph(new Run(new Text("First paragraph"))));
                body.AppendChild(new Paragraph(new Run(new Text("Second paragraph"))));
            });

            // Act
            var set = ExternalAnnotationManager.CreateAnnotationSet(doc, "test");

            // Assert
            Assert.True(set.PageCount >= 1);
            Assert.NotNull(set.PawlsFileContent);
            Assert.NotNull(set.LabelledText);
            Assert.NotNull(set.TextLabels);
            Assert.NotNull(set.DocLabelDefinitions);
        }

        #endregion

        #region Annotation Creation Tests

        [Fact]
        public void EA020_CreateAnnotation_ValidOffsets_CreatesAnnotation()
        {
            // Arrange
            var documentText = "Hello, world! This is a test.";

            // Act
            var annotation = ExternalAnnotationManager.CreateAnnotation(
                "ann-001", "IMPORTANT", documentText, 0, 5);

            // Assert
            Assert.NotNull(annotation);
            Assert.Equal("ann-001", annotation.Id);
            Assert.Equal("IMPORTANT", annotation.AnnotationLabel);
            Assert.Equal("Hello", annotation.RawText);
            Assert.False(annotation.Structural);

            var span = annotation.AnnotationJson as TextSpan;
            Assert.NotNull(span);
            Assert.Equal(0, span.Start);
            Assert.Equal(5, span.End);
            Assert.Equal("Hello", span.Text);
        }

        [Fact]
        public void EA021_CreateAnnotation_InvalidOffsets_ThrowsException()
        {
            // Arrange
            var documentText = "Hello";

            // Assert - Start > End
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                ExternalAnnotationManager.CreateAnnotation("ann", "label", documentText, 3, 1));

            // Assert - Negative start
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                ExternalAnnotationManager.CreateAnnotation("ann", "label", documentText, -1, 3));

            // Assert - End > Length
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                ExternalAnnotationManager.CreateAnnotation("ann", "label", documentText, 0, 10));
        }

        [Fact]
        public void EA022_CreateAnnotationFromSearch_ExistingText_FindsOffsets()
        {
            // Arrange
            var documentText = "Hello, world! This is a test.";

            // Act
            var annotation = ExternalAnnotationManager.CreateAnnotationFromSearch(
                "ann-001", "GREETING", documentText, "world");

            // Assert
            Assert.NotNull(annotation);
            Assert.Equal("world", annotation.RawText);

            var span = annotation.AnnotationJson as TextSpan;
            Assert.NotNull(span);
            Assert.Equal(7, span.Start);
            Assert.Equal(12, span.End);
        }

        [Fact]
        public void EA023_CreateAnnotationFromSearch_MultipleOccurrences_FindsCorrectOne()
        {
            // Arrange - use lowercase consistently
            var documentText = "hello world, hello world, hello world!";

            // Act
            var ann1 = ExternalAnnotationManager.CreateAnnotationFromSearch(
                "ann-1", "WORD", documentText, "hello", 1);
            var ann2 = ExternalAnnotationManager.CreateAnnotationFromSearch(
                "ann-2", "WORD", documentText, "hello", 2);
            var ann3 = ExternalAnnotationManager.CreateAnnotationFromSearch(
                "ann-3", "WORD", documentText, "hello", 3);

            // Assert
            Assert.NotNull(ann1);
            Assert.NotNull(ann2);
            Assert.NotNull(ann3);

            var span1 = ann1.AnnotationJson as TextSpan;
            var span2 = ann2.AnnotationJson as TextSpan;
            var span3 = ann3.AnnotationJson as TextSpan;

            Assert.NotNull(span1);
            Assert.NotNull(span2);
            Assert.NotNull(span3);

            // First occurrence at 0, second at 13, third at 26
            Assert.Equal(0, span1.Start);
            Assert.Equal(13, span2.Start);
            Assert.Equal(26, span3.Start);
        }

        [Fact]
        public void EA024_CreateAnnotationFromSearch_NotFound_ReturnsNull()
        {
            // Arrange
            var documentText = "Hello, world!";

            // Act
            var annotation = ExternalAnnotationManager.CreateAnnotationFromSearch(
                "ann-001", "LABEL", documentText, "goodbye");

            // Assert
            Assert.Null(annotation);
        }

        [Fact]
        public void EA025_FindTextOccurrences_ReturnsAllOccurrences()
        {
            // Arrange
            var documentText = "the quick brown fox jumps over the lazy dog";

            // Act
            var occurrences = ExternalAnnotationManager.FindTextOccurrences(documentText, "the");

            // Assert
            Assert.Equal(2, occurrences.Count);
            Assert.Equal((0, 3), occurrences[0]);
            Assert.Equal((31, 34), occurrences[1]);
        }

        #endregion

        #region Validation Tests

        [Fact]
        public void EA030_Validate_MatchingHash_ReturnsValid()
        {
            // Arrange
            var doc = CreateSimpleTestDocument("Hello, world!");
            var set = ExternalAnnotationManager.CreateAnnotationSet(doc, "test");

            // Act
            var result = ExternalAnnotationManager.Validate(doc, set);

            // Assert
            Assert.True(result.IsValid);
            Assert.False(result.HashMismatch);
            Assert.Empty(result.Issues);
        }

        [Fact]
        public void EA031_Validate_MismatchedHash_ReturnsHashMismatch()
        {
            // Arrange
            var doc1 = CreateSimpleTestDocument("Hello, world!");
            var doc2 = CreateSimpleTestDocument("Hello, world!!");
            var set = ExternalAnnotationManager.CreateAnnotationSet(doc1, "test");

            // Act
            var result = ExternalAnnotationManager.Validate(doc2, set);

            // Assert
            Assert.False(result.IsValid);
            Assert.True(result.HashMismatch);
        }

        [Fact]
        public void EA032_Validate_StaleAnnotation_ReportsTextMismatch()
        {
            // Arrange
            var doc = CreateSimpleTestDocument("Hello, world!");
            var set = ExternalAnnotationManager.CreateAnnotationSet(doc, "test");

            // Add an annotation with wrong text
            var badAnnotation = new OpenContractsAnnotation
            {
                Id = "bad-ann",
                AnnotationLabel = "TEST",
                RawText = "wrong text",
                Page = 0,
                AnnotationJson = new TextSpan
                {
                    Id = "bad-ann",
                    Start = 0,
                    End = 5,
                    Text = "wrong text" // Doesn't match "Hello"
                },
                Structural = false
            };
            set.LabelledText.Add(badAnnotation);

            // Act
            var result = ExternalAnnotationManager.Validate(doc, set);

            // Assert
            Assert.False(result.IsValid);
            Assert.Contains(result.Issues, i => i.IssueType == "TextMismatch");
        }

        [Fact]
        public void EA033_Validate_OutOfBoundsAnnotation_ReportsOutOfBounds()
        {
            // Arrange
            var doc = CreateSimpleTestDocument("Hello");
            var set = ExternalAnnotationManager.CreateAnnotationSet(doc, "test");

            // Add an annotation with invalid offsets
            var badAnnotation = new OpenContractsAnnotation
            {
                Id = "bad-ann",
                AnnotationLabel = "TEST",
                RawText = "out of bounds",
                Page = 0,
                AnnotationJson = new TextSpan
                {
                    Id = "bad-ann",
                    Start = 0,
                    End = 1000, // Way beyond document length
                    Text = "out of bounds"
                },
                Structural = false
            };
            set.LabelledText.Add(badAnnotation);

            // Act
            var result = ExternalAnnotationManager.Validate(doc, set);

            // Assert
            Assert.False(result.IsValid);
            Assert.Contains(result.Issues, i => i.IssueType == "OutOfBounds");
        }

        #endregion

        #region Serialization Tests

        [Fact]
        public void EA040_SerializeDeserialize_RoundTrip_PreservesAllData()
        {
            // Arrange
            var doc = CreateSimpleTestDocument("Hello, world!");
            var set = ExternalAnnotationManager.CreateAnnotationSet(doc, "test-doc");

            // Add some labels
            set.TextLabels["IMPORTANT"] = new AnnotationLabel
            {
                Id = "IMPORTANT",
                Text = "Important",
                Color = "#FF0000",
                Description = "Important text",
                LabelType = "text"
            };

            // Add an annotation
            var annotation = ExternalAnnotationManager.CreateAnnotation(
                "ann-001", "IMPORTANT", set.Content, 0, 5);
            set.LabelledText.Add(annotation);

            // Act
            var json = ExternalAnnotationManager.SerializeToJson(set);
            var deserialized = ExternalAnnotationManager.DeserializeFromJson(json);

            // Assert
            Assert.NotNull(deserialized);
            Assert.Equal(set.DocumentId, deserialized.DocumentId);
            Assert.Equal(set.DocumentHash, deserialized.DocumentHash);
            Assert.Equal(set.Version, deserialized.Version);
            Assert.Contains("Hello, world!", deserialized.Content);
            Assert.Single(deserialized.TextLabels);
            Assert.True(deserialized.TextLabels.ContainsKey("IMPORTANT"));
        }

        [Fact]
        public void EA041_DeserializeFromJson_InvalidJson_ReturnsNull()
        {
            // Act
            var result = ExternalAnnotationManager.DeserializeFromJson("not valid json");

            // Assert
            Assert.Null(result);
        }

        [Fact]
        public void EA042_DeserializeFromJson_EmptyString_ReturnsNull()
        {
            // Act
            var result = ExternalAnnotationManager.DeserializeFromJson("");

            // Assert
            Assert.Null(result);
        }

        #endregion

        #region Projection Tests

        [Fact]
        public void EA050_ProjectAnnotations_SingleAnnotation_AddsWrapperSpan()
        {
            // Arrange
            var doc = CreateSimpleTestDocument("Hello, world!");
            var set = ExternalAnnotationManager.CreateAnnotationSet(doc, "test");

            // Add label
            set.TextLabels["GREETING"] = new AnnotationLabel
            {
                Id = "GREETING",
                Text = "Greeting",
                Color = "#FFEB3B"
            };

            // Add annotation
            var annotation = ExternalAnnotationManager.CreateAnnotation(
                "ann-001", "GREETING", set.Content, 0, 5);
            set.LabelledText.Add(annotation);

            var settings = new WmlToHtmlConverterSettings { PageTitle = "Test" };
            var html = WmlToHtmlConverter.ConvertToHtml(doc, settings);

            var projectionSettings = new ExternalAnnotationProjectionSettings
            {
                CssClassPrefix = "ext-annot-",
                LabelMode = AnnotationLabelMode.Inline
            };

            // Act
            var result = ExternalAnnotationProjector.ProjectAnnotations(html, set, projectionSettings);

            // Assert
            var resultStr = result.ToString();
            Assert.Contains("ext-annot-highlight", resultStr);
            Assert.Contains("data-annotation-id", resultStr);
            Assert.Contains("ann-001", resultStr);
        }

        [Fact]
        public void EA051_ConvertWithAnnotations_IntegratedFlow()
        {
            // Arrange
            var doc = CreateSimpleTestDocument("Hello, world!");
            var set = ExternalAnnotationManager.CreateAnnotationSet(doc, "test");

            set.TextLabels["TEST"] = new AnnotationLabel
            {
                Id = "TEST",
                Text = "Test Label",
                Color = "#00FF00"
            };

            var annotation = ExternalAnnotationManager.CreateAnnotation(
                "ann-001", "TEST", set.Content, 0, 5);
            set.LabelledText.Add(annotation);

            var htmlSettings = new WmlToHtmlConverterSettings { PageTitle = "Test" };
            var projectionSettings = new ExternalAnnotationProjectionSettings();

            // Act
            var html = ExternalAnnotationProjector.ConvertWithAnnotations(
                doc, set, htmlSettings, projectionSettings);

            // Assert
            Assert.NotNull(html);
            Assert.Contains("ext-annot-highlight", html);
        }

        #endregion

        #region Integration Tests with Real Documents

        [Fact]
        public void EA060_Integration_RealDocument_CreatesAndValidatesSet()
        {
            // Arrange
            var sourceDocx = new FileInfo(Path.Combine(TestFilesDir.FullName, "HC001-5DayTourPlanTemplate.docx"));
            if (!sourceDocx.Exists)
            {
                // Skip if test file doesn't exist
                return;
            }

            var wmlDoc = new WmlDocument(sourceDocx.FullName);

            // Act
            var set = ExternalAnnotationManager.CreateAnnotationSet(wmlDoc, "tour-template");

            // Assert
            Assert.NotEmpty(set.DocumentHash);
            Assert.NotEmpty(set.Content);
            Assert.True(set.PageCount >= 1);

            // Validate should pass
            var validation = ExternalAnnotationManager.Validate(wmlDoc, set);
            Assert.True(validation.IsValid);
            Assert.False(validation.HashMismatch);
        }

        #endregion
    }
}

#endif
