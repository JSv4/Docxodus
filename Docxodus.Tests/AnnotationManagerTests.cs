// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests
{
    public class AnnotationManagerTests
    {
        /// <summary>
        /// Creates a simple test document with paragraphs.
        /// </summary>
        private WmlDocument CreateTestDocument(params string[] paragraphTexts)
        {
            using (var ms = new MemoryStream())
            {
                using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = doc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    var body = new Body();
                    mainPart.Document.Body = body;

                    // Add StyleDefinitionsPart (required for many operations)
                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();

                    // Add DocumentSettingsPart
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();

                    // Add paragraphs
                    foreach (var text in paragraphTexts)
                    {
                        var para = new Paragraph(
                            new Run(
                                new Text(text) { Space = SpaceProcessingModeValues.Preserve }
                            )
                        );
                        body.Append(para);
                    }

                    mainPart.Document.Save();
                }

                return new WmlDocument("TestDocument.docx", ms.ToArray());
            }
        }

        [Fact]
        public void AM001_AddAnnotation_WithSearchText_CreatesBookmarkAndCustomXml()
        {
            // Arrange
            var doc = CreateTestDocument("This is a test paragraph.", "Another paragraph here.");
            var annotation = new DocumentAnnotation("ann1", "IMPORTANT", "Important", "#FF5722")
            {
                Author = "Test User"
            };
            var range = AnnotationRange.FromSearch("test paragraph");

            // Act
            var result = AnnotationManager.AddAnnotation(doc, annotation, range);

            // Assert
            Assert.True(AnnotationManager.HasAnnotations(result));
            var annotations = AnnotationManager.GetAnnotations(result);
            Assert.Single(annotations);
            Assert.Equal("ann1", annotations[0].Id);
            Assert.Equal("IMPORTANT", annotations[0].LabelId);
            Assert.Equal("Important", annotations[0].Label);
            Assert.Equal("#FF5722", annotations[0].Color);
            Assert.Equal("Test User", annotations[0].Author);
            Assert.StartsWith(AnnotationManager.BookmarkPrefix, annotations[0].BookmarkName);
        }

        [Fact]
        public void AM002_AddAnnotation_WithParagraphIndex_CreatesBookmark()
        {
            // Arrange
            var doc = CreateTestDocument("First paragraph.", "Second paragraph.", "Third paragraph.");
            var annotation = new DocumentAnnotation("ann2", "SECTION", "Section Header", "#4CAF50");
            var range = AnnotationRange.FromParagraphs(1, 1); // Second paragraph

            // Act
            var result = AnnotationManager.AddAnnotation(doc, annotation, range);

            // Assert
            var annotations = AnnotationManager.GetAnnotations(result);
            Assert.Single(annotations);
            Assert.Equal("ann2", annotations[0].Id);

            // Verify the annotated text
            var annotatedText = AnnotationManager.GetAnnotatedText(result, "ann2");
            Assert.Equal("Second paragraph.", annotatedText);
        }

        [Fact]
        public void AM003_GetAnnotation_ReturnsSpecificAnnotation()
        {
            // Arrange
            var doc = CreateTestDocument("Test content here.");
            var annotation = new DocumentAnnotation("specific-id", "TYPE_A", "Type A", "#2196F3");
            var range = AnnotationRange.FromSearch("Test");
            doc = AnnotationManager.AddAnnotation(doc, annotation, range);

            // Act
            var retrieved = AnnotationManager.GetAnnotation(doc, "specific-id");

            // Assert
            Assert.NotNull(retrieved);
            Assert.Equal("specific-id", retrieved.Id);
            Assert.Equal("TYPE_A", retrieved.LabelId);
            Assert.Equal("Type A", retrieved.Label);
        }

        [Fact]
        public void AM004_GetAnnotation_ReturnsNullForNonExistent()
        {
            // Arrange
            var doc = CreateTestDocument("Test content.");

            // Act
            var result = AnnotationManager.GetAnnotation(doc, "non-existent-id");

            // Assert
            Assert.Null(result);
        }

        [Fact]
        public void AM005_RemoveAnnotation_RemovesAnnotationAndBookmark()
        {
            // Arrange
            var doc = CreateTestDocument("Test content for removal.");
            var annotation = new DocumentAnnotation("to-remove", "TEMP", "Temporary", "#9E9E9E");
            var range = AnnotationRange.FromSearch("content for removal");
            doc = AnnotationManager.AddAnnotation(doc, annotation, range);
            Assert.True(AnnotationManager.HasAnnotations(doc));

            // Act
            var result = AnnotationManager.RemoveAnnotation(doc, "to-remove");

            // Assert
            Assert.False(AnnotationManager.HasAnnotations(result));
            Assert.Empty(AnnotationManager.GetAnnotations(result));
        }

        [Fact]
        public void AM006_UpdateAnnotation_UpdatesMetadata()
        {
            // Arrange
            var doc = CreateTestDocument("Content to annotate.");
            var annotation = new DocumentAnnotation("update-test", "OLD_TYPE", "Old Label", "#FF0000");
            var range = AnnotationRange.FromSearch("Content");
            doc = AnnotationManager.AddAnnotation(doc, annotation, range);

            // Modify the annotation
            var updatedAnnotation = new DocumentAnnotation
            {
                Id = "update-test",
                LabelId = "NEW_TYPE",
                Label = "New Label",
                Color = "#00FF00",
                BookmarkName = AnnotationManager.BookmarkPrefix + "update-test"
            };

            // Act
            var result = AnnotationManager.UpdateAnnotation(doc, updatedAnnotation);

            // Assert
            var retrieved = AnnotationManager.GetAnnotation(result, "update-test");
            Assert.NotNull(retrieved);
            Assert.Equal("NEW_TYPE", retrieved.LabelId);
            Assert.Equal("New Label", retrieved.Label);
            Assert.Equal("#00FF00", retrieved.Color);
        }

        [Fact]
        public void AM007_MultipleAnnotations_CanCoexist()
        {
            // Arrange
            var doc = CreateTestDocument("First section content.", "Second section content.", "Third section content.");

            // Act - Add multiple annotations
            doc = AnnotationManager.AddAnnotation(doc,
                new DocumentAnnotation("ann-1", "TYPE_A", "Type A", "#FF5722"),
                AnnotationRange.FromParagraphs(0, 0));

            doc = AnnotationManager.AddAnnotation(doc,
                new DocumentAnnotation("ann-2", "TYPE_B", "Type B", "#4CAF50"),
                AnnotationRange.FromParagraphs(1, 1));

            doc = AnnotationManager.AddAnnotation(doc,
                new DocumentAnnotation("ann-3", "TYPE_C", "Type C", "#2196F3"),
                AnnotationRange.FromParagraphs(2, 2));

            // Assert
            var annotations = AnnotationManager.GetAnnotations(doc);
            Assert.Equal(3, annotations.Count);
            Assert.Contains(annotations, a => a.Id == "ann-1");
            Assert.Contains(annotations, a => a.Id == "ann-2");
            Assert.Contains(annotations, a => a.Id == "ann-3");
        }

        [Fact]
        public void AM008_AnnotationWithMetadata_PersistsCorrectly()
        {
            // Arrange
            var doc = CreateTestDocument("Content with metadata.");
            var annotation = new DocumentAnnotation("meta-test", "CUSTOM", "Custom Label", "#9C27B0")
            {
                Author = "Jane Doe",
                Metadata = new Dictionary<string, string>
                {
                    { "confidence", "0.95" },
                    { "source", "automated" },
                    { "category", "legal-clause" }
                }
            };
            var range = AnnotationRange.FromSearch("Content");

            // Act
            var result = AnnotationManager.AddAnnotation(doc, annotation, range);

            // Assert
            var retrieved = AnnotationManager.GetAnnotation(result, "meta-test");
            Assert.NotNull(retrieved);
            Assert.Equal("Jane Doe", retrieved.Author);
            Assert.Equal(3, retrieved.Metadata.Count);
            Assert.Equal("0.95", retrieved.Metadata["confidence"]);
            Assert.Equal("automated", retrieved.Metadata["source"]);
            Assert.Equal("legal-clause", retrieved.Metadata["category"]);
        }

        [Fact]
        public void AM009_HasAnnotations_ReturnsFalseForCleanDocument()
        {
            // Arrange
            var doc = CreateTestDocument("Clean document without annotations.");

            // Act & Assert
            Assert.False(AnnotationManager.HasAnnotations(doc));
        }

        [Fact]
        public void AM010_GetAnnotatedText_ReturnsCorrectText()
        {
            // Arrange
            var doc = CreateTestDocument("The quick brown fox jumps over the lazy dog.");
            var annotation = new DocumentAnnotation("text-test", "ANIMAL", "Animal", "#795548");
            var range = AnnotationRange.FromSearch("brown fox");
            doc = AnnotationManager.AddAnnotation(doc, annotation, range);

            // Act
            var annotatedText = AnnotationManager.GetAnnotatedText(doc, "text-test");

            // Assert
            Assert.Equal("brown fox", annotatedText);
        }

        [Fact]
        public void AM011_SearchTextWithOccurrence_FindsCorrectInstance()
        {
            // Arrange
            var doc = CreateTestDocument("The word test appears here. Another test appears here too.");
            var annotation = new DocumentAnnotation("second-test", "KEYWORD", "Keyword", "#E91E63");
            var range = AnnotationRange.FromSearch("test", occurrence: 2);

            // Act
            var result = AnnotationManager.AddAnnotation(doc, annotation, range);

            // Assert
            // The annotation should be on the second occurrence of "test"
            var annotations = AnnotationManager.GetAnnotations(result);
            Assert.Single(annotations);
        }

        [Fact]
        public void AM012_UpdatePageSpans_UpdatesAnnotationPageInfo()
        {
            // Arrange
            var doc = CreateTestDocument("Multi-page content simulation.");
            var annotation = new DocumentAnnotation("page-test", "SECTION", "Section", "#607D8B");
            var range = AnnotationRange.FromSearch("Multi-page");
            doc = AnnotationManager.AddAnnotation(doc, annotation, range);

            var pageSpans = new Dictionary<string, (int startPage, int endPage)>
            {
                { "page-test", (1, 3) }
            };

            // Act
            var result = AnnotationManager.UpdateAnnotationPageSpans(doc, pageSpans);

            // Assert
            var retrieved = AnnotationManager.GetAnnotation(result, "page-test");
            Assert.NotNull(retrieved);
            Assert.Equal(1, retrieved.StartPage);
            Assert.Equal(3, retrieved.EndPage);
        }

        [Fact]
        public void AM013_RenderAnnotationsInHtml_ProducesHighlightSpans()
        {
            // Arrange
            var doc = CreateTestDocument("This document has annotated content.");
            var annotation = new DocumentAnnotation("render-test", "HIGHLIGHT", "Highlight", "#FFEB3B");
            var range = AnnotationRange.FromSearch("annotated content");
            doc = AnnotationManager.AddAnnotation(doc, annotation, range);

            // Verify annotation was added with correct bookmark
            var retrievedAnnotation = AnnotationManager.GetAnnotation(doc, "render-test");
            Assert.NotNull(retrievedAnnotation);
            Assert.StartsWith(AnnotationManager.BookmarkPrefix, retrievedAnnotation.BookmarkName);

            var settings = new WmlToHtmlConverterSettings
            {
                RenderAnnotations = true,
                AnnotationLabelMode = AnnotationLabelMode.Above
            };

            // Act
            var html = WmlToHtmlConverter.ConvertToHtml(doc, settings);

            // Assert
            var htmlString = html.ToString();

            // Debug: Check if the bookmark anchor is in HTML
            Assert.Contains(retrievedAnnotation.BookmarkName, htmlString);

            Assert.Contains("annot-highlight", htmlString);
            Assert.Contains("data-annotation-id", htmlString);
            Assert.Contains("render-test", htmlString);
        }

        [Fact]
        public void AM014_RenderAnnotationsDisabled_NoHighlightSpans()
        {
            // Arrange
            var doc = CreateTestDocument("This document has annotated content.");
            var annotation = new DocumentAnnotation("no-render-test", "HIGHLIGHT", "Highlight", "#FFEB3B");
            var range = AnnotationRange.FromSearch("annotated");
            doc = AnnotationManager.AddAnnotation(doc, annotation, range);

            var settings = new WmlToHtmlConverterSettings
            {
                RenderAnnotations = false // Disabled
            };

            // Act
            var html = WmlToHtmlConverter.ConvertToHtml(doc, settings);

            // Assert
            var htmlString = html.ToString();
            Assert.DoesNotContain("annot-highlight", htmlString);
            Assert.DoesNotContain("data-annotation-id", htmlString);
        }

        [Fact]
        public void AM015_AnnotationLabelModes_ApplyCorrectAttributes()
        {
            // Arrange
            var doc = CreateTestDocument("Content for label mode testing.");
            var annotation = new DocumentAnnotation("mode-test", "TEST", "Test Label", "#3F51B5");
            var range = AnnotationRange.FromSearch("label mode");
            doc = AnnotationManager.AddAnnotation(doc, annotation, range);

            // Test different label modes
            var modes = new[] { AnnotationLabelMode.Above, AnnotationLabelMode.Inline, AnnotationLabelMode.Tooltip, AnnotationLabelMode.None };

            foreach (var mode in modes)
            {
                var settings = new WmlToHtmlConverterSettings
                {
                    RenderAnnotations = true,
                    AnnotationLabelMode = mode
                };

                // Act
                var html = WmlToHtmlConverter.ConvertToHtml(doc, settings);

                // Assert
                var htmlString = html.ToString();
                Assert.Contains($"data-label-mode=\"{mode.ToString().ToLowerInvariant()}\"", htmlString);
            }
        }

        [Fact]
        public void AM016_AnnotationCssGenerated_WhenRenderingEnabled()
        {
            // Arrange
            var doc = CreateTestDocument("Document for CSS test.");
            var annotation = new DocumentAnnotation("css-test", "STYLE", "Style", "#00BCD4");
            var range = AnnotationRange.FromSearch("Document");
            doc = AnnotationManager.AddAnnotation(doc, annotation, range);

            var settings = new WmlToHtmlConverterSettings
            {
                RenderAnnotations = true,
                AnnotationCssClassPrefix = "annot-"
            };

            // Act
            var html = WmlToHtmlConverter.ConvertToHtml(doc, settings);

            // Assert
            var htmlString = html.ToString();
            // CSS should contain annotation styles
            Assert.Contains(".annot-highlight", htmlString);
            Assert.Contains(".annot-label", htmlString);
            Assert.Contains("--annot-color", htmlString);
        }

        [Fact]
        public void AM017_MultiParagraphAnnotation_SpansCorrectly()
        {
            // Arrange
            var doc = CreateTestDocument("First paragraph.", "Second paragraph.", "Third paragraph.");
            var annotation = new DocumentAnnotation("multi-para", "SECTION", "Section", "#8BC34A");
            var range = AnnotationRange.FromParagraphs(0, 2); // All three paragraphs

            // Act
            var result = AnnotationManager.AddAnnotation(doc, annotation, range);

            // Assert
            var annotations = AnnotationManager.GetAnnotations(result);
            Assert.Single(annotations);

            // The annotation should span all paragraphs
            var annotatedText = AnnotationManager.GetAnnotatedText(result, "multi-para");
            Assert.Contains("First paragraph", annotatedText);
            Assert.Contains("Second paragraph", annotatedText);
            Assert.Contains("Third paragraph", annotatedText);
        }

        [Fact]
        public void AM018_DocumentAnnotation_Constructor_SetsDefaults()
        {
            // Arrange & Act
            var annotation = new DocumentAnnotation("test-id", "TYPE", "Label", "#123456");

            // Assert
            Assert.Equal("test-id", annotation.Id);
            Assert.Equal("TYPE", annotation.LabelId);
            Assert.Equal("Label", annotation.Label);
            Assert.Equal("#123456", annotation.Color);
            Assert.Equal(AnnotationManager.BookmarkPrefix + "test-id", annotation.BookmarkName);
            Assert.NotNull(annotation.Created);
            Assert.True(annotation.PageInfoStale);
        }

        [Fact]
        public void AM019_AnnotationRange_StaticMethods_CreateCorrectRanges()
        {
            // Test FromSearch
            var searchRange = AnnotationRange.FromSearch("test text", 2);
            Assert.Equal("test text", searchRange.SearchText);
            Assert.Equal(2, searchRange.Occurrence);

            // Test FromBookmark
            var bookmarkRange = AnnotationRange.FromBookmark("MyBookmark");
            Assert.Equal("MyBookmark", bookmarkRange.ExistingBookmarkName);

            // Test FromParagraphs
            var paraRange = AnnotationRange.FromParagraphs(1, 5);
            Assert.Equal(1, paraRange.StartParagraphIndex);
            Assert.Equal(5, paraRange.EndParagraphIndex);

            // Test FromRuns
            var runRange = AnnotationRange.FromRuns(0, 1, 2, 3);
            Assert.Equal(0, runRange.StartParagraphIndex);
            Assert.Equal(1, runRange.StartRunIndex);
            Assert.Equal(2, runRange.EndParagraphIndex);
            Assert.Equal(3, runRange.EndRunIndex);
        }

        [Fact]
        public void AM020_RemoveNonExistentAnnotation_ReturnsUnchangedDocument()
        {
            // Arrange
            var doc = CreateTestDocument("Test content.");
            var annotation = new DocumentAnnotation("existing", "TYPE", "Label", "#000000");
            var range = AnnotationRange.FromSearch("Test");
            doc = AnnotationManager.AddAnnotation(doc, annotation, range);

            // Act - Try to remove non-existent annotation
            var result = AnnotationManager.RemoveAnnotation(doc, "non-existent");

            // Assert - Existing annotation should still be there
            Assert.True(AnnotationManager.HasAnnotations(result));
            var annotations = AnnotationManager.GetAnnotations(result);
            Assert.Single(annotations);
            Assert.Equal("existing", annotations[0].Id);
        }
    }
}
