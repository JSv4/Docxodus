// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;
using static Docxodus.WmlComparer;

namespace OxPt
{
    /// <summary>
    /// Tests for move detection in WmlComparer.GetRevisions().
    /// </summary>
    public class WmlComparerMoveDetectionTests
    {
        #region Helper Methods

        /// <summary>
        /// Creates a minimal valid DOCX document with the specified paragraphs.
        /// </summary>
        private static WmlDocument CreateDocumentWithParagraphs(params string[] paragraphs)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(
                    new Body(
                        paragraphs.Select(text =>
                            new Paragraph(
                                new Run(
                                    new Text(text)
                                )
                            )
                        )
                    )
                );

                // Add required styles part
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new Styles(
                    new DocDefaults(
                        new RunPropertiesDefault(
                            new RunPropertiesBaseStyle(
                                new RunFonts { Ascii = "Calibri" },
                                new FontSize { Val = "22" }
                            )
                        ),
                        new ParagraphPropertiesDefault()
                    )
                );

                // Add required settings part
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings();

                doc.Save();
            }

            stream.Position = 0;
            return new WmlDocument("test.docx", stream.ToArray());
        }

        #endregion

        #region Basic Move Detection Tests

        [Fact]
        public void MoveDetection_IdenticalText_ShouldMarkAsMove()
        {
            // Arrange: Create two documents where a paragraph is moved
            // Doc1: A, B, C
            // Doc2: B, A, C  (A moved after B)
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here.",
                "This is paragraph C with more words added."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection.",
                "This is paragraph C with more words added."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert
            var movedRevisions = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.Moved).ToList();
            Assert.True(movedRevisions.Count >= 2, $"Expected at least 2 moved revisions, got {movedRevisions.Count}");

            // Verify move pairs are properly linked
            var moveGroups = movedRevisions.GroupBy(r => r.MoveGroupId).ToList();
            foreach (var group in moveGroups)
            {
                Assert.NotNull(group.Key);
                var items = group.ToList();
                Assert.Equal(2, items.Count);
                Assert.True(items.Any(r => r.IsMoveSource == true), "Should have a move source");
                Assert.True(items.Any(r => r.IsMoveSource == false), "Should have a move destination");
            }
        }

        [Fact]
        public void MoveDetection_SimilarText_AboveThreshold_ShouldMarkAsMove()
        {
            // Arrange: Text is 90% similar (above 80% threshold)
            var doc1 = CreateDocumentWithParagraphs(
                "The quick brown fox jumps over the lazy dog today.",
                "Another paragraph here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "Another paragraph here.",
                "The quick brown fox jumps over the lazy dog now."  // "today" -> "now" (90% similar)
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert
            var movedRevisions = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.Moved).ToList();
            Assert.True(movedRevisions.Count >= 2, $"Expected at least 2 moved revisions for similar text, got {movedRevisions.Count}");
        }

        [Fact]
        public void MoveDetection_DissimilarText_BelowThreshold_ShouldRemainInsertedDeleted()
        {
            // Arrange: Text is very different (below 80% threshold)
            var doc1 = CreateDocumentWithParagraphs(
                "The quick brown fox jumps over the lazy dog.",
                "Another paragraph here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "Another paragraph here.",
                "A completely different sentence with new words entirely."  // Very different
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert
            var movedRevisions = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.Moved).ToList();
            Assert.Empty(movedRevisions);

            // Should have separate deletions and insertions
            var deletions = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.Deleted).ToList();
            var insertions = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.Inserted).ToList();
            Assert.True(deletions.Count > 0 || insertions.Count > 0, "Should have deletions or insertions");
        }

        [Fact]
        public void MoveDetection_ShortText_BelowMinimum_ShouldRemainInsertedDeleted()
        {
            // Arrange: Very short text (below 3 word minimum)
            var doc1 = CreateDocumentWithParagraphs(
                "Hello world",  // Only 2 words
                "Another paragraph here with more content."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "Another paragraph here with more content.",
                "Hello world"  // Moved but too short
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert: Short text should not be detected as a move
            var movedRevisions = revisions
                .Where(r => r.RevisionType == WmlComparerRevisionType.Moved)
                .Where(r => r.Text?.Contains("Hello") == true || r.Text?.Contains("world") == true)
                .ToList();
            Assert.Empty(movedRevisions);
        }

        #endregion

        #region Settings Tests

        [Fact]
        public void MoveDetection_Disabled_ShouldNotDetectMoves()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = false  // Disabled
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert
            var movedRevisions = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.Moved).ToList();
            Assert.Empty(movedRevisions);
        }

        [Fact]
        public void MoveDetection_CustomThreshold_ShouldRespectSetting()
        {
            // Arrange: Text is 70% similar
            var doc1 = CreateDocumentWithParagraphs(
                "The quick brown fox jumps over the lazy dog in the park.",
                "Another paragraph here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "Another paragraph here.",
                "The quick brown cat runs under the sleepy dog in the yard."  // ~60% similar
            );

            // With low threshold, should detect as move
            var lowThresholdSettings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.5,  // 50% threshold
                MoveMinimumWordCount = 3
            };

            // With high threshold, should not detect as move
            var highThresholdSettings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.9,  // 90% threshold
                MoveMinimumWordCount = 3
            };

            // Act
            var comparedLow = WmlComparer.Compare(doc1, doc2, lowThresholdSettings);
            var revisionsLow = WmlComparer.GetRevisions(comparedLow, lowThresholdSettings);

            var comparedHigh = WmlComparer.Compare(doc1, doc2, highThresholdSettings);
            var revisionsHigh = WmlComparer.GetRevisions(comparedHigh, highThresholdSettings);

            // Assert
            var movesLow = revisionsLow.Where(r => r.RevisionType == WmlComparerRevisionType.Moved).Count();
            var movesHigh = revisionsHigh.Where(r => r.RevisionType == WmlComparerRevisionType.Moved).Count();

            Assert.True(movesLow >= movesHigh,
                $"Lower threshold should detect more or equal moves: low={movesLow}, high={movesHigh}");
        }

        [Fact]
        public void MoveDetection_CustomMinWordCount_ShouldRespectSetting()
        {
            // Arrange: 4-word sentence
            var doc1 = CreateDocumentWithParagraphs(
                "Four word sentence here.",
                "Another paragraph with more content for testing purposes."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "Another paragraph with more content for testing purposes.",
                "Four word sentence here."
            );

            // With min 3 words, should detect
            var minThreeSettings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // With min 5 words, should not detect (sentence is only 4 words)
            var minFiveSettings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 5
            };

            // Act
            var comparedThree = WmlComparer.Compare(doc1, doc2, minThreeSettings);
            var revisionsThree = WmlComparer.GetRevisions(comparedThree, minThreeSettings);

            var comparedFive = WmlComparer.Compare(doc1, doc2, minFiveSettings);
            var revisionsFive = WmlComparer.GetRevisions(comparedFive, minFiveSettings);

            // Assert
            var movesWithMinThree = revisionsThree
                .Where(r => r.RevisionType == WmlComparerRevisionType.Moved)
                .Where(r => r.Text?.Contains("Four") == true)
                .Count();

            var movesWithMinFive = revisionsFive
                .Where(r => r.RevisionType == WmlComparerRevisionType.Moved)
                .Where(r => r.Text?.Contains("Four") == true)
                .Count();

            Assert.True(movesWithMinThree >= movesWithMinFive,
                $"Lower min word count should detect more moves: min3={movesWithMinThree}, min5={movesWithMinFive}");
        }

        [Fact]
        public void MoveDetection_CaseInsensitive_ShouldMatchIgnoringCase()
        {
            // Arrange: Same text with different case
            var doc1 = CreateDocumentWithParagraphs(
                "THE QUICK BROWN FOX JUMPS OVER THE LAZY DOG.",
                "Another paragraph here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "Another paragraph here.",
                "the quick brown fox jumps over the lazy dog."  // Same words, different case
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3,
                CaseInsensitive = true
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert: Should detect as move when case-insensitive
            var movedRevisions = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.Moved).ToList();
            Assert.True(movedRevisions.Count >= 2,
                $"Case-insensitive should detect moves: got {movedRevisions.Count}");
        }

        #endregion

        #region Multiple Moves Tests

        [Fact]
        public void MoveDetection_MultipleMoves_ShouldMatchCorrectly()
        {
            // Arrange: Multiple paragraphs swapped
            // Doc1: A, B, C, D
            // Doc2: C, D, A, B  (A,B moved to end; C,D moved to start)
            var doc1 = CreateDocumentWithParagraphs(
                "First paragraph with content alpha beta gamma.",
                "Second paragraph with content delta epsilon zeta.",
                "Third paragraph with content eta theta iota.",
                "Fourth paragraph with content kappa lambda mu."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "Third paragraph with content eta theta iota.",
                "Fourth paragraph with content kappa lambda mu.",
                "First paragraph with content alpha beta gamma.",
                "Second paragraph with content delta epsilon zeta."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert
            var movedRevisions = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.Moved).ToList();

            // Each move pair should have unique MoveGroupId
            var moveGroupIds = movedRevisions
                .Where(r => r.MoveGroupId.HasValue)
                .Select(r => r.MoveGroupId.Value)
                .Distinct()
                .ToList();

            // Verify each group has at least one source and one destination
            // (may have multiple revisions per group due to paragraph marks, etc.)
            foreach (var groupId in moveGroupIds)
            {
                var groupRevisions = movedRevisions.Where(r => r.MoveGroupId == groupId).ToList();
                Assert.True(groupRevisions.Count >= 2, $"Group {groupId} should have at least 2 revisions");
                Assert.True(groupRevisions.Any(r => r.IsMoveSource == true), $"Group {groupId} should have at least one source");
                Assert.True(groupRevisions.Any(r => r.IsMoveSource == false), $"Group {groupId} should have at least one destination");
            }
        }

        #endregion

        #region Edge Cases

        [Fact]
        public void MoveDetection_EmptyDocument_ShouldNotThrow()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs();
            var doc2 = CreateDocumentWithParagraphs("New content added here with several words.");

            var settings = new WmlComparerSettings
            {
                DetectMoves = true
            };

            // Act & Assert: Should not throw
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // No moves expected (nothing to move from empty doc)
            var moves = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.Moved).ToList();
            Assert.Empty(moves);
        }

        [Fact]
        public void MoveDetection_IdenticalDocuments_ShouldHaveNoRevisions()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs(
                "Same content in both documents with enough words."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "Same content in both documents with enough words."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert
            Assert.Empty(revisions);
        }

        [Fact]
        public void MoveDetection_OnlyDeletions_ShouldNotCreateMoves()
        {
            // Arrange: Content removed, nothing added
            var doc1 = CreateDocumentWithParagraphs(
                "First paragraph that will be deleted.",
                "Second paragraph that stays here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "Second paragraph that stays here."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert: Should be deletion, not move
            var moves = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.Moved).ToList();
            Assert.Empty(moves);

            var deletions = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.Deleted).ToList();
            Assert.True(deletions.Count > 0, "Should have deletions");
        }

        [Fact]
        public void MoveDetection_OnlyInsertions_ShouldNotCreateMoves()
        {
            // Arrange: Content added, nothing removed
            var doc1 = CreateDocumentWithParagraphs(
                "First paragraph that stays here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "First paragraph that stays here.",
                "Second paragraph that is newly added."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert: Should be insertion, not move
            var moves = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.Moved).ToList();
            Assert.Empty(moves);

            var insertions = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.Inserted).ToList();
            Assert.True(insertions.Count > 0, "Should have insertions");
        }

        #endregion

        #region Revision Properties Tests

        [Fact]
        public void MoveDetection_RevisionProperties_ShouldBeCorrect()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs(
                "Paragraph to be moved with enough words for detection.",
                "Static paragraph that does not change here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "Static paragraph that does not change here.",
                "Paragraph to be moved with enough words for detection."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert
            var movedRevisions = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.Moved).ToList();

            foreach (var rev in movedRevisions)
            {
                // All moved revisions should have MoveGroupId
                Assert.NotNull(rev.MoveGroupId);
                Assert.True(rev.MoveGroupId > 0);

                // All moved revisions should have IsMoveSource set
                Assert.NotNull(rev.IsMoveSource);
            }

            // Find a pair and verify they reference each other
            if (movedRevisions.Count >= 2)
            {
                var source = movedRevisions.FirstOrDefault(r => r.IsMoveSource == true);
                var dest = movedRevisions.FirstOrDefault(r => r.IsMoveSource == false && r.MoveGroupId == source?.MoveGroupId);

                Assert.NotNull(source);
                Assert.NotNull(dest);
                Assert.Equal(source.MoveGroupId, dest.MoveGroupId);
            }
        }

        #endregion

        #region Native Move Markup Tests

        /// <summary>
        /// Verifies that the compared document contains native w:moveFrom elements
        /// instead of just w:del elements for moved content.
        /// </summary>
        [Fact]
        public void NativeMoveMarkup_ShouldContainMoveFromElement()
        {
            // Arrange: Create two documents where a paragraph is moved
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here.",
                "This is paragraph C with more words added."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection.",
                "This is paragraph C with more words added."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract the document XML
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);
            var bodyXml = doc.MainDocumentPart.Document.Body.OuterXml;
            var bodyElement = XElement.Parse(bodyXml);

            // Assert: Should contain w:moveFrom elements
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var moveFromElements = bodyElement.Descendants(w + "moveFrom").ToList();

            Assert.True(moveFromElements.Count > 0,
                $"Expected w:moveFrom elements in the document, but found none. Body XML contains: {(bodyXml.Length > 500 ? bodyXml.Substring(0, 500) + "..." : bodyXml)}");
        }

        /// <summary>
        /// Verifies that the compared document contains native w:moveTo elements
        /// instead of just w:ins elements for moved content.
        /// </summary>
        [Fact]
        public void NativeMoveMarkup_ShouldContainMoveToElement()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here.",
                "This is paragraph C with more words added."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection.",
                "This is paragraph C with more words added."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract the document XML
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);
            var bodyXml = doc.MainDocumentPart.Document.Body.OuterXml;
            var bodyElement = XElement.Parse(bodyXml);

            // Assert: Should contain w:moveTo elements
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var moveToElements = bodyElement.Descendants(w + "moveTo").ToList();

            Assert.True(moveToElements.Count > 0,
                $"Expected w:moveTo elements in the document, but found none.");
        }

        /// <summary>
        /// Verifies that move range markers (moveFromRangeStart/End, moveToRangeStart/End)
        /// are present and properly paired.
        /// </summary>
        [Fact]
        public void NativeMoveMarkup_ShouldContainRangeMarkers()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract the document XML
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);
            var bodyXml = doc.MainDocumentPart.Document.Body.OuterXml;
            var bodyElement = XElement.Parse(bodyXml);

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Assert: Should contain range start/end elements
            var moveFromRangeStart = bodyElement.Descendants(w + "moveFromRangeStart").ToList();
            var moveFromRangeEnd = bodyElement.Descendants(w + "moveFromRangeEnd").ToList();
            var moveToRangeStart = bodyElement.Descendants(w + "moveToRangeStart").ToList();
            var moveToRangeEnd = bodyElement.Descendants(w + "moveToRangeEnd").ToList();

            Assert.True(moveFromRangeStart.Count > 0, "Expected w:moveFromRangeStart elements");
            Assert.True(moveFromRangeEnd.Count > 0, "Expected w:moveFromRangeEnd elements");
            Assert.True(moveToRangeStart.Count > 0, "Expected w:moveToRangeStart elements");
            Assert.True(moveToRangeEnd.Count > 0, "Expected w:moveToRangeEnd elements");

            // Verify range start and end counts match
            Assert.Equal(moveFromRangeStart.Count, moveFromRangeEnd.Count);
            Assert.Equal(moveToRangeStart.Count, moveToRangeEnd.Count);
        }

        /// <summary>
        /// Verifies that move pairs are linked via the w:name attribute.
        /// </summary>
        [Fact]
        public void NativeMoveMarkup_ShouldLinkPairsViaNameAttribute()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract the document XML
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);
            var bodyXml = doc.MainDocumentPart.Document.Body.OuterXml;
            var bodyElement = XElement.Parse(bodyXml);

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Get the w:name values from moveFromRangeStart and moveToRangeStart
            var moveFromNames = bodyElement.Descendants(w + "moveFromRangeStart")
                .Select(e => e.Attribute(w + "name")?.Value)
                .Where(n => n != null)
                .ToHashSet();

            var moveToNames = bodyElement.Descendants(w + "moveToRangeStart")
                .Select(e => e.Attribute(w + "name")?.Value)
                .Where(n => n != null)
                .ToHashSet();

            // Assert: Each moveFrom name should have a matching moveTo name
            Assert.True(moveFromNames.Count > 0, "Expected moveFromRangeStart elements with w:name attribute");
            Assert.True(moveToNames.Count > 0, "Expected moveToRangeStart elements with w:name attribute");

            // All names should match between source and destination
            Assert.True(moveFromNames.SetEquals(moveToNames),
                $"Move names should match. From: [{string.Join(", ", moveFromNames)}], To: [{string.Join(", ", moveToNames)}]");
        }

        /// <summary>
        /// Verifies that when move detection is disabled, no move markup is generated.
        /// </summary>
        [Fact]
        public void NativeMoveMarkup_WhenDisabled_ShouldNotContainMoveElements()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = false  // Disabled
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract the document XML
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);
            var bodyXml = doc.MainDocumentPart.Document.Body.OuterXml;
            var bodyElement = XElement.Parse(bodyXml);

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Assert: Should NOT contain move elements
            var moveFromElements = bodyElement.Descendants(w + "moveFrom").ToList();
            var moveToElements = bodyElement.Descendants(w + "moveTo").ToList();

            Assert.Empty(moveFromElements);
            Assert.Empty(moveToElements);

            // Should have regular ins/del instead
            var insElements = bodyElement.Descendants(w + "ins").ToList();
            var delElements = bodyElement.Descendants(w + "del").ToList();

            Assert.True(insElements.Count > 0 || delElements.Count > 0,
                "Should have regular ins/del elements when move detection is disabled");
        }

        /// <summary>
        /// Verifies that move elements have required attributes (w:id, w:author, w:date).
        /// </summary>
        [Fact]
        public void NativeMoveMarkup_ShouldHaveRequiredAttributes()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3,
                AuthorForRevisions = "TestAuthor"
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract the document XML
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);
            var bodyXml = doc.MainDocumentPart.Document.Body.OuterXml;
            var bodyElement = XElement.Parse(bodyXml);

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Check moveFrom elements
            var moveFromElements = bodyElement.Descendants(w + "moveFrom").ToList();
            foreach (var elem in moveFromElements)
            {
                Assert.NotNull(elem.Attribute(w + "id"));
                Assert.NotNull(elem.Attribute(w + "author"));
                Assert.Equal("TestAuthor", elem.Attribute(w + "author")?.Value);
                Assert.NotNull(elem.Attribute(w + "date"));
            }

            // Check moveTo elements
            var moveToElements = bodyElement.Descendants(w + "moveTo").ToList();
            foreach (var elem in moveToElements)
            {
                Assert.NotNull(elem.Attribute(w + "id"));
                Assert.NotNull(elem.Attribute(w + "author"));
                Assert.Equal("TestAuthor", elem.Attribute(w + "author")?.Value);
                Assert.NotNull(elem.Attribute(w + "date"));
            }
        }

        /// <summary>
        /// Verifies that range IDs are properly paired between start and end elements.
        /// </summary>
        [Fact]
        public void NativeMoveMarkup_RangeIdsShouldBeProperlyPaired()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract the document XML
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);
            var bodyXml = doc.MainDocumentPart.Document.Body.OuterXml;
            var bodyElement = XElement.Parse(bodyXml);

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Get range IDs from moveFrom range markers
            var moveFromStartIds = bodyElement.Descendants(w + "moveFromRangeStart")
                .Select(e => e.Attribute(w + "id")?.Value)
                .Where(id => id != null)
                .ToHashSet();

            var moveFromEndIds = bodyElement.Descendants(w + "moveFromRangeEnd")
                .Select(e => e.Attribute(w + "id")?.Value)
                .Where(id => id != null)
                .ToHashSet();

            // Get range IDs from moveTo range markers
            var moveToStartIds = bodyElement.Descendants(w + "moveToRangeStart")
                .Select(e => e.Attribute(w + "id")?.Value)
                .Where(id => id != null)
                .ToHashSet();

            var moveToEndIds = bodyElement.Descendants(w + "moveToRangeEnd")
                .Select(e => e.Attribute(w + "id")?.Value)
                .Where(id => id != null)
                .ToHashSet();

            // Assert: Start and end IDs should match
            Assert.True(moveFromStartIds.SetEquals(moveFromEndIds),
                $"moveFrom range IDs should match. Start: [{string.Join(", ", moveFromStartIds)}], End: [{string.Join(", ", moveFromEndIds)}]");

            Assert.True(moveToStartIds.SetEquals(moveToEndIds),
                $"moveTo range IDs should match. Start: [{string.Join(", ", moveToStartIds)}], End: [{string.Join(", ", moveToEndIds)}]");
        }

        #endregion

        #region SimplifyMoveMarkup Tests

        /// <summary>
        /// Verifies that SimplifyMoveMarkup converts moveFrom to del elements.
        /// </summary>
        [Fact]
        public void SimplifyMoveMarkup_ShouldConvertMoveFromToDel()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                SimplifyMoveMarkup = true,  // Enable simplification
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract the document XML
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);
            var bodyXml = doc.MainDocumentPart.Document.Body.OuterXml;
            var bodyElement = XElement.Parse(bodyXml);

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Assert: Should NOT contain move elements
            var moveFromElements = bodyElement.Descendants(w + "moveFrom").ToList();
            var moveToElements = bodyElement.Descendants(w + "moveTo").ToList();

            Assert.Empty(moveFromElements);
            Assert.Empty(moveToElements);

            // Should have del elements instead
            var delElements = bodyElement.Descendants(w + "del").ToList();
            Assert.True(delElements.Count > 0, "Should have w:del elements after simplification");
        }

        /// <summary>
        /// Verifies that SimplifyMoveMarkup converts moveTo to ins elements.
        /// </summary>
        [Fact]
        public void SimplifyMoveMarkup_ShouldConvertMoveToToIns()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                SimplifyMoveMarkup = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract the document XML
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);
            var bodyXml = doc.MainDocumentPart.Document.Body.OuterXml;
            var bodyElement = XElement.Parse(bodyXml);

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Should have ins elements
            var insElements = bodyElement.Descendants(w + "ins").ToList();
            Assert.True(insElements.Count > 0, "Should have w:ins elements after simplification");
        }

        /// <summary>
        /// Verifies that SimplifyMoveMarkup removes all move range markers.
        /// </summary>
        [Fact]
        public void SimplifyMoveMarkup_ShouldRemoveRangeMarkers()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                SimplifyMoveMarkup = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract the document XML
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);
            var bodyXml = doc.MainDocumentPart.Document.Body.OuterXml;
            var bodyElement = XElement.Parse(bodyXml);

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Assert: Should NOT contain any range markers
            Assert.Empty(bodyElement.Descendants(w + "moveFromRangeStart").ToList());
            Assert.Empty(bodyElement.Descendants(w + "moveFromRangeEnd").ToList());
            Assert.Empty(bodyElement.Descendants(w + "moveToRangeStart").ToList());
            Assert.Empty(bodyElement.Descendants(w + "moveToRangeEnd").ToList());
        }

        /// <summary>
        /// Verifies that SimplifyMoveMarkup preserves author and date attributes.
        /// </summary>
        [Fact]
        public void SimplifyMoveMarkup_ShouldPreserveAttributes()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                SimplifyMoveMarkup = true,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3,
                AuthorForRevisions = "TestAuthor"
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract the document XML
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);
            var bodyXml = doc.MainDocumentPart.Document.Body.OuterXml;
            var bodyElement = XElement.Parse(bodyXml);

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Check that del/ins elements have the expected attributes
            var delElements = bodyElement.Descendants(w + "del").ToList();
            var insElements = bodyElement.Descendants(w + "ins").ToList();

            foreach (var del in delElements)
            {
                Assert.NotNull(del.Attribute(w + "author"));
                Assert.NotNull(del.Attribute(w + "id"));
            }

            foreach (var ins in insElements)
            {
                Assert.NotNull(ins.Attribute(w + "author"));
                Assert.NotNull(ins.Attribute(w + "id"));
            }
        }

        /// <summary>
        /// Verifies that SimplifyMoveMarkup = false (default) preserves move elements.
        /// </summary>
        [Fact]
        public void SimplifyMoveMarkup_WhenFalse_ShouldPreserveMoveElements()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                SimplifyMoveMarkup = false,  // Explicitly false (default)
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract the document XML
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);
            var bodyXml = doc.MainDocumentPart.Document.Body.OuterXml;
            var bodyElement = XElement.Parse(bodyXml);

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Assert: Should contain move elements
            var moveFromElements = bodyElement.Descendants(w + "moveFrom").ToList();
            var moveToElements = bodyElement.Descendants(w + "moveTo").ToList();

            Assert.True(moveFromElements.Count > 0, "Should have w:moveFrom elements when SimplifyMoveMarkup is false");
            Assert.True(moveToElements.Count > 0, "Should have w:moveTo elements when SimplifyMoveMarkup is false");
        }

        /// <summary>
        /// Verifies that DetectMoves defaults to true.
        /// </summary>
        [Fact]
        public void DetectMoves_ShouldDefaultToTrue()
        {
            var settings = new WmlComparerSettings();
            Assert.True(settings.DetectMoves, "DetectMoves should default to true");
        }

        /// <summary>
        /// Verifies that SimplifyMoveMarkup defaults to false.
        /// </summary>
        [Fact]
        public void SimplifyMoveMarkup_ShouldDefaultToFalse()
        {
            var settings = new WmlComparerSettings();
            Assert.False(settings.SimplifyMoveMarkup, "SimplifyMoveMarkup should default to false");
        }

        #endregion

        #region ID Uniqueness Tests (Issue #96 Phase II)

        /// <summary>
        /// Verifies that all revision IDs are unique across the document when moves are present.
        /// This is the core test for Issue #96 - duplicate IDs cause Word "unreadable content" warnings.
        /// </summary>
        [Fact]
        public void MoveMarkup_AllRevisionIdsShouldBeUnique()
        {
            // Arrange: Create documents with moved content
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here.",
                "This is paragraph C that stays in place.",
                "This is paragraph D with additional content."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection.",
                "This is paragraph C that stays in place but modified slightly.",
                "This is paragraph D with additional content."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                SimplifyMoveMarkup = false,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract all revision IDs from all content parts
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);

            var allIds = new List<string>();
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var revisionElements = new[] { "ins", "del", "moveFrom", "moveTo",
                "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd", "rPrChange" };

            // Check main document
            var mainXDoc = doc.MainDocumentPart.GetXDocument();
            foreach (var elemName in revisionElements)
            {
                allIds.AddRange(mainXDoc.Descendants(w + elemName)
                    .Select(e => e.Attribute(w + "id")?.Value)
                    .Where(id => id != null));
            }

            // Check footnotes if present
            if (doc.MainDocumentPart.FootnotesPart != null)
            {
                var fnXDoc = doc.MainDocumentPart.FootnotesPart.GetXDocument();
                foreach (var elemName in revisionElements)
                {
                    allIds.AddRange(fnXDoc.Descendants(w + elemName)
                        .Select(e => e.Attribute(w + "id")?.Value)
                        .Where(id => id != null));
                }
            }

            // Check endnotes if present
            if (doc.MainDocumentPart.EndnotesPart != null)
            {
                var enXDoc = doc.MainDocumentPart.EndnotesPart.GetXDocument();
                foreach (var elemName in revisionElements)
                {
                    allIds.AddRange(enXDoc.Descendants(w + elemName)
                        .Select(e => e.Attribute(w + "id")?.Value)
                        .Where(id => id != null));
                }
            }

            // Assert: No duplicate IDs (excluding range start/end pairs which intentionally share IDs)
            // For range elements, start and end share the same ID by design
            // But NO other element should share an ID with any other element
            var duplicates = allIds.GroupBy(x => x)
                .Where(g => g.Count() > 2)  // Allow pairs (start/end) but not more
                .Select(g => new { Id = g.Key, Count = g.Count() })
                .ToList();

            Assert.True(duplicates.Count == 0,
                $"Found revision IDs used more than twice (only range pairs should share IDs): " +
                $"{string.Join(", ", duplicates.Select(d => $"id={d.Id} count={d.Count}"))}");
        }

        /// <summary>
        /// Verifies that move names properly pair moveFrom and moveTo elements.
        /// Each move name should appear exactly once in moveFromRangeStart and once in moveToRangeStart.
        /// Note: Consecutive paragraphs may be grouped as a single move block.
        /// </summary>
        [Fact]
        public void MoveMarkup_MoveNamesShouldProperlyPairSourceAndDestination()
        {
            // Arrange: Create documents with moved content
            var doc1 = CreateDocumentWithParagraphs(
                "This is paragraph A with enough words for move detection.",
                "This is paragraph B with sufficient content here."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This is paragraph B with sufficient content here.",
                "This is paragraph A with enough words for move detection."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                SimplifyMoveMarkup = false,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract move names
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var mainXDoc = doc.MainDocumentPart.GetXDocument();

            var moveFromNames = mainXDoc.Descendants(w + "moveFromRangeStart")
                .Select(e => e.Attribute(w + "name")?.Value)
                .Where(n => n != null)
                .ToList();

            var moveToNames = mainXDoc.Descendants(w + "moveToRangeStart")
                .Select(e => e.Attribute(w + "name")?.Value)
                .Where(n => n != null)
                .ToList();

            // Assert: Should have at least one move detected
            Assert.True(moveFromNames.Count > 0, "Expected at least one moveFromRangeStart with w:name");
            Assert.True(moveToNames.Count > 0, "Expected at least one moveToRangeStart with w:name");

            // Assert: moveFrom and moveTo names should match (same names, same count)
            Assert.True(moveFromNames.OrderBy(x => x).SequenceEqual(moveToNames.OrderBy(x => x)),
                $"Move names should match between moveFrom and moveTo. " +
                $"From: [{string.Join(", ", moveFromNames)}], To: [{string.Join(", ", moveToNames)}]");

            // Assert: No empty or null move names
            Assert.DoesNotContain("", moveFromNames);
            Assert.DoesNotContain("", moveToNames);
            Assert.True(moveFromNames.All(n => n.StartsWith("move")),
                "All move names should follow the 'moveN' pattern");
        }

        /// <summary>
        /// Verifies that a document with moves and other changes has unique IDs.
        /// This specifically tests the scenario that caused Issue #96.
        /// </summary>
        [Fact]
        public void MoveMarkup_WithMixedChanges_ShouldHaveUniqueIds()
        {
            // Arrange: Create documents with moves AND other ins/del changes
            var doc1 = CreateDocumentWithParagraphs(
                "This paragraph will be moved to a new location.",
                "This paragraph stays but will be modified here.",
                "This paragraph will be deleted entirely from doc.",
                "This is static content that does not change."
            );
            var doc2 = CreateDocumentWithParagraphs(
                "This paragraph stays but has been changed now.",
                "This is static content that does not change.",
                "This paragraph will be moved to a new location.",
                "This is a completely new paragraph inserted."
            );

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                SimplifyMoveMarkup = false,
                MoveSimilarityThreshold = 0.8,
                MoveMinimumWordCount = 3
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Extract all revision IDs
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var doc = WordprocessingDocument.Open(stream, false);

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var mainXDoc = doc.MainDocumentPart.GetXDocument();

            // Get IDs from different element types
            var insIds = mainXDoc.Descendants(w + "ins")
                .Select(e => e.Attribute(w + "id")?.Value).Where(id => id != null).ToList();
            var delIds = mainXDoc.Descendants(w + "del")
                .Select(e => e.Attribute(w + "id")?.Value).Where(id => id != null).ToList();
            var moveFromIds = mainXDoc.Descendants(w + "moveFrom")
                .Select(e => e.Attribute(w + "id")?.Value).Where(id => id != null).ToList();
            var moveToIds = mainXDoc.Descendants(w + "moveTo")
                .Select(e => e.Attribute(w + "id")?.Value).Where(id => id != null).ToList();

            // Combine non-range IDs (these should all be unique)
            var nonRangeIds = insIds.Concat(delIds).Concat(moveFromIds).Concat(moveToIds).ToList();

            // Check for duplicates
            var duplicates = nonRangeIds.GroupBy(x => x)
                .Where(g => g.Count() > 1)
                .ToList();

            Assert.True(duplicates.Count == 0,
                $"Found duplicate IDs among ins/del/moveFrom/moveTo elements: " +
                $"{string.Join(", ", duplicates.Select(g => $"id={g.Key}"))}. " +
                $"This is the Issue #96 bug - FixUpRevMarkIds was overwriting IDs.");
        }

        #endregion

        #region Stress Tests (Issue #96)

        /// <summary>
        /// Generates a paragraph with unique content for stress testing.
        /// </summary>
        private static string GenerateStressTestParagraph(int index)
        {
            var templates = new[]
            {
                "Paragraph {0}: This document section contains important information about the project requirements and specifications. Reference ID: {1}",
                "Section {0}: The following content describes the technical implementation details for the proposed system architecture. Doc: {1}",
                "Item {0}: According to the agreement dated herein, the parties shall comply with all terms and conditions specified. Contract: {1}",
                "Clause {0}: The licensee agrees to use the software only for purposes permitted under this license agreement. License: {1}",
                "Article {0}: This paragraph establishes the fundamental principles governing the relationship between the entities. Ref: {1}",
                "Point {0}: The data processing activities shall be conducted in accordance with applicable privacy regulations. GDPR: {1}",
                "Note {0}: All modifications to this document must be tracked and approved by the designated review committee. Rev: {1}",
                "Entry {0}: The financial statements have been prepared in accordance with generally accepted accounting principles. GAAP: {1}",
                "Record {0}: This memorandum summarizes the key decisions made during the executive committee meeting. Minutes: {1}",
                "Statement {0}: The undersigned hereby certifies that all information provided is true and accurate. Cert: {1}",
                "Provision {0}: Notwithstanding the foregoing, the obligations set forth herein shall survive termination. Legal: {1}",
                "Stipulation {0}: The contractor shall deliver all work products by the specified deadline. Deadline: {1}",
                "Requirement {0}: The system shall support concurrent users and maintain response times under load. Perf: {1}",
                "Specification {0}: All API endpoints must implement proper authentication and authorization. Security: {1}",
                "Definition {0}: For purposes of this agreement, the following terms shall have the meanings ascribed. Terms: {1}",
            };

            var template = templates[index % templates.Length];
            return string.Format(template, index, $"DOC-{index:D4}-{Guid.NewGuid().ToString().Substring(0, 8).ToUpper()}");
        }

        /// <summary>
        /// Stress test for Issue #96: Validates that revision IDs remain unique even with
        /// dozens of moves and hundreds of other changes. Uses fixed seed for reproducibility.
        /// </summary>
        [Theory]
        [InlineData(50, 15, 30, "Small")]      // 50 paragraphs, ~15 moves, ~30 other changes
        [InlineData(100, 25, 50, "Medium")]    // 100 paragraphs, ~25 moves, ~50 other changes
        [InlineData(200, 40, 100, "Large")]    // 200 paragraphs, ~40 moves, ~100 other changes
        public void StressTest_ManyMovesAndChanges_ShouldHaveUniqueIds(
            int paragraphCount, int moveCount, int changeCount, string testName)
        {
            // Arrange: Use fixed seed for reproducibility
            var rng = new Random(42);

            // Generate original document with numbered paragraphs
            var originalParagraphs = Enumerable.Range(1, paragraphCount)
                .Select(i => GenerateStressTestParagraph(i))
                .ToList();

            // Create modified version with moves and changes
            var modifiedParagraphs = new List<string>(originalParagraphs);

            // Apply moves: pick random paragraphs and move them to new positions
            var availableForMove = Enumerable.Range(0, modifiedParagraphs.Count).ToList();
            for (int i = 0; i < moveCount && availableForMove.Count > 2; i++)
            {
                int fromIdx = availableForMove[rng.Next(availableForMove.Count)];
                availableForMove.Remove(fromIdx);

                var para = modifiedParagraphs[fromIdx];
                modifiedParagraphs.RemoveAt(fromIdx);

                // Adjust available indices after removal
                availableForMove = availableForMove.Select(x => x > fromIdx ? x - 1 : x).ToList();

                int toIdx = rng.Next(modifiedParagraphs.Count + 1);
                modifiedParagraphs.Insert(toIdx, para);

                // Adjust available indices after insertion
                availableForMove = availableForMove.Select(x => x >= toIdx ? x + 1 : x).ToList();
            }

            // Apply deletions
            int deleteCount = changeCount / 3;
            for (int i = 0; i < deleteCount && modifiedParagraphs.Count > paragraphCount / 2; i++)
            {
                int idx = rng.Next(modifiedParagraphs.Count);
                modifiedParagraphs.RemoveAt(idx);
            }

            // Apply insertions
            int insertCount = changeCount / 3;
            for (int i = 0; i < insertCount; i++)
            {
                int idx = rng.Next(modifiedParagraphs.Count + 1);
                modifiedParagraphs.Insert(idx, $"[NEW-{i + 1}] This is a newly inserted paragraph with enough words to be meaningful. " +
                    $"It contains various content including technical terms, legal jargon, and general prose. " +
                    $"The purpose is to test the comparison engine with substantial insertions. Reference: INS-{Guid.NewGuid():N}");
            }

            // Apply modifications (change words in existing paragraphs)
            int modifyCount = changeCount / 3;
            for (int i = 0; i < modifyCount && modifiedParagraphs.Count > 0; i++)
            {
                int idx = rng.Next(modifiedParagraphs.Count);
                var para = modifiedParagraphs[idx];
                para = para.Replace("paragraph", "section")
                           .Replace("content", "material")
                           .Replace("document", "file");
                if (!para.Contains("[MODIFIED]"))
                {
                    para = "[MODIFIED] " + para;
                }
                modifiedParagraphs[idx] = para;
            }

            // Create documents
            var doc1 = CreateDocumentWithParagraphs(originalParagraphs.ToArray());
            var doc2 = CreateDocumentWithParagraphs(modifiedParagraphs.ToArray());

            var settings = new WmlComparerSettings
            {
                DetectMoves = true,
                SimplifyMoveMarkup = false,
                MoveSimilarityThreshold = 0.75,
                MoveMinimumWordCount = 5,
                AuthorForRevisions = "StressTest"
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Assert: Analyze results
            using var stream = new MemoryStream(compared.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(stream, false);
            var mainXDoc = wDoc.MainDocumentPart.GetXDocument();

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            // Collect all revision IDs
            var revisionElements = new[] { "ins", "del", "moveFrom", "moveTo", "rPrChange" };
            var allIds = new List<(string Id, string Type)>();

            foreach (var elemName in revisionElements)
            {
                foreach (var elem in mainXDoc.Descendants(w + elemName))
                {
                    var id = elem.Attribute(w + "id")?.Value;
                    if (id != null)
                    {
                        allIds.Add((id, elemName));
                    }
                }
            }

            // Check for duplicates - THE CRITICAL TEST
            var duplicates = allIds.GroupBy(x => x.Id)
                .Where(g => g.Count() > 1)
                .ToList();

            Assert.True(duplicates.Count == 0,
                $"StressTest {testName}: Found {duplicates.Count} duplicate revision IDs. " +
                $"First duplicates: {string.Join(", ", duplicates.Take(5).Select(d => $"id={d.Key}:[{string.Join(",", d.Select(x => x.Type))}]"))}. " +
                $"Total elements: {allIds.Count}");

            // Check move name pairing
            var moveFromNames = mainXDoc.Descendants(w + "moveFromRangeStart")
                .Select(e => e.Attribute(w + "name")?.Value)
                .Where(n => !string.IsNullOrEmpty(n))
                .ToList();

            var moveToNames = mainXDoc.Descendants(w + "moveToRangeStart")
                .Select(e => e.Attribute(w + "name")?.Value)
                .Where(n => !string.IsNullOrEmpty(n))
                .ToList();

            // Validate all names are paired
            var unpairedFrom = moveFromNames.Except(moveToNames).ToList();
            var unpairedTo = moveToNames.Except(moveFromNames).ToList();

            Assert.True(!unpairedFrom.Any() && !unpairedTo.Any(),
                $"StressTest {testName}: Unpaired move names found. " +
                $"From without To: [{string.Join(", ", unpairedFrom)}], " +
                $"To without From: [{string.Join(", ", unpairedTo)}]");

            // OpenXML validation
            var validator = new OpenXmlValidator(FileFormatVersions.Office2019);
            var errors = validator.Validate(wDoc).ToList();

            // Note: Some validation errors may be acceptable (e.g., missing optional parts)
            // We focus on ensuring no critical structural errors
            var criticalErrors = errors
                .Where(e => e.ErrorType == DocumentFormat.OpenXml.Validation.ValidationErrorType.Schema)
                .ToList();

            Assert.True(criticalErrors.Count == 0,
                $"StressTest {testName}: OpenXML schema validation failed with {criticalErrors.Count} errors. " +
                $"First errors: {string.Join("; ", criticalErrors.Take(3).Select(e => e.Description))}");
        }

        #endregion
    }
}
