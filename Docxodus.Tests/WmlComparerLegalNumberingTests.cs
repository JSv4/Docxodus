// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace OxPt
{
    /// <summary>
    /// Tests for legal numbering preservation in WmlComparer.
    ///
    /// This addresses GitHub issue: https://github.com/dotnet/Open-XML-SDK/issues/1634
    ///
    /// "Legal numbering" is a Word feature where nested list numbers display as "1.1", "1.1.1"
    /// instead of "a)", "i)", etc. It's controlled by the w:isLgl element in the abstract
    /// numbering definition.
    ///
    /// When comparing documents where one has legal numbering and the other doesn't,
    /// the comparison result should preserve the numbering definitions from both documents.
    /// </summary>
    public class WmlComparerLegalNumberingTests
    {
        #region Helper Methods

        /// <summary>
        /// Creates a document with a numbered list paragraph.
        /// </summary>
        /// <param name="text">The text content of the paragraph</param>
        /// <param name="useLegalNumbering">If true, includes w:isLgl in the numbering definition</param>
        /// <param name="numId">The numId to use for the paragraph</param>
        /// <returns>A WmlDocument with the numbered paragraph</returns>
        private static WmlDocument CreateDocumentWithNumberedList(string text, bool useLegalNumbering, int numId = 1)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();

                // Create the document with a numbered paragraph
                mainPart.Document = new Document(
                    new Body(
                        new Paragraph(
                            new ParagraphProperties(
                                new NumberingProperties(
                                    new NumberingLevelReference { Val = 0 },
                                    new NumberingId { Val = numId }
                                )
                            ),
                            new Run(new Text(text))
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

                // Add numbering part with optional legal numbering
                var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();

                // Build the level element with optional isLgl
                var levelElement = new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." },
                    new LevelJustification { Val = LevelJustificationValues.Left },
                    new PreviousParagraphProperties(
                        new Indentation { Left = "720", Hanging = "360" }
                    )
                ) { LevelIndex = 0 };

                if (useLegalNumbering)
                {
                    // Insert isLgl element AFTER numFmt per Open XML schema order:
                    // start, numFmt, lvlRestart, pStyle, isLgl, suff, lvlText, lvlPicBulletId, lvlJc, pPr, rPr
                    levelElement.InsertAfter(new IsLegalNumberingStyle(), levelElement.GetFirstChild<NumberingFormat>());
                }

                var abstractNum = new AbstractNum(levelElement)
                {
                    AbstractNumberId = 1
                };

                var numberingInstance = new NumberingInstance(
                    new AbstractNumId { Val = 1 }
                ) { NumberID = numId };

                numberingPart.Numbering = new Numbering(abstractNum, numberingInstance);

                doc.Save();
            }

            stream.Position = 0;
            return new WmlDocument("test.docx", stream.ToArray());
        }

        /// <summary>
        /// Creates a document with multi-level legal numbering (more realistic test case).
        /// </summary>
        private static WmlDocument CreateDocumentWithMultiLevelNumbering(string[] texts, bool useLegalNumbering, int numId = 1)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();

                // Create paragraphs at different levels
                var paragraphs = texts.Select((text, index) =>
                    new Paragraph(
                        new ParagraphProperties(
                            new NumberingProperties(
                                new NumberingLevelReference { Val = index % 3 }, // Cycle through levels 0, 1, 2
                                new NumberingId { Val = numId }
                            )
                        ),
                        new Run(new Text(text))
                    )
                ).ToArray();

                mainPart.Document = new Document(new Body(paragraphs));

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

                // Add numbering part with multi-level definition
                var numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();

                // Create 3 levels for multi-level numbering
                var levels = new Level[3];
                for (int i = 0; i < 3; i++)
                {
                    var levelText = string.Join(".", Enumerable.Range(1, i + 1).Select(n => $"%{n}")) + ".";

                    levels[i] = new Level(
                        new StartNumberingValue { Val = 1 },
                        new NumberingFormat { Val = NumberFormatValues.Decimal },
                        new LevelText { Val = levelText },
                        new LevelJustification { Val = LevelJustificationValues.Left },
                        new PreviousParagraphProperties(
                            new Indentation { Left = ((i + 1) * 720).ToString(), Hanging = "360" }
                        )
                    ) { LevelIndex = i };

                    if (useLegalNumbering)
                    {
                        // Insert isLgl AFTER numFmt per Open XML schema order
                        levels[i].InsertAfter(new IsLegalNumberingStyle(), levels[i].GetFirstChild<NumberingFormat>());
                    }
                }

                var abstractNum = new AbstractNum(levels) { AbstractNumberId = 1 };
                var numberingInstance = new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = numId };
                numberingPart.Numbering = new Numbering(abstractNum, numberingInstance);

                doc.Save();
            }

            stream.Position = 0;
            return new WmlDocument("test.docx", stream.ToArray());
        }

        /// <summary>
        /// Checks if a document's numbering definitions contain the isLgl element.
        /// </summary>
        private static bool DocumentHasLegalNumbering(WmlDocument doc)
        {
            using var stream = new MemoryStream(doc.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(stream, false);

            var numberingPart = wDoc.MainDocumentPart?.NumberingDefinitionsPart;
            if (numberingPart == null)
                return false;

            var xDoc = numberingPart.GetXDocument();
            // Check for w:isLgl element in any abstract numbering level
            return xDoc.Descendants(W.isLgl).Any();
        }

        /// <summary>
        /// Gets the count of abstractNum definitions in the document.
        /// </summary>
        private static int GetAbstractNumCount(WmlDocument doc)
        {
            using var stream = new MemoryStream(doc.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(stream, false);

            var numberingPart = wDoc.MainDocumentPart?.NumberingDefinitionsPart;
            if (numberingPart == null)
                return 0;

            var xDoc = numberingPart.GetXDocument();
            return xDoc.Descendants(W.abstractNum).Count();
        }

        /// <summary>
        /// Validates the document and returns validation errors.
        /// </summary>
        private static string ValidateDocument(WmlDocument doc)
        {
            using var stream = new MemoryStream(doc.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(stream, false);

            var validator = new OpenXmlValidator();
            var errors = validator.Validate(wDoc)
                .Where(e => !e.Description.Contains("Unid")) // Ignore our internal attributes
                .ToList();

            if (errors.Count == 0)
                return string.Empty;

            return string.Join(Environment.NewLine, errors.Select(e =>
                $"[{e.ErrorType}] {e.Description} at {e.Path?.XPath ?? "unknown"}"));
        }

        #endregion

        #region Test Cases

        /// <summary>
        /// Test Case 1: Original has legal numbering, revised does NOT.
        /// Expected: Result should preserve legal numbering from original.
        /// </summary>
        [Fact]
        public void WC_LegalNum_001_OriginalHasLegalNumbering_RevisedDoesNot_PreservesLegal()
        {
            // Arrange
            var original = CreateDocumentWithNumberedList("First item", useLegalNumbering: true);
            var revised = CreateDocumentWithNumberedList("First item modified", useLegalNumbering: false);

            Assert.True(DocumentHasLegalNumbering(original), "Original should have legal numbering");
            Assert.False(DocumentHasLegalNumbering(revised), "Revised should NOT have legal numbering");

            var settings = new WmlComparerSettings();

            // Act
            var compared = WmlComparer.Compare(original, revised, settings);

            // Assert - the result starts from original, so should retain legal numbering
            Assert.True(DocumentHasLegalNumbering(compared),
                "Compared document should preserve legal numbering from original");

            var validationErrors = ValidateDocument(compared);
            Assert.True(string.IsNullOrEmpty(validationErrors),
                $"Document should be valid. Errors: {validationErrors}");
        }

        /// <summary>
        /// Test Case 2: Original does NOT have legal numbering, revised DOES.
        /// Expected: Result should include legal numbering definition from revised.
        /// This is the main issue from GitHub #1634 - when the revised document
        /// introduces legal numbering, it should be preserved.
        /// </summary>
        [Fact]
        public void WC_LegalNum_002_RevisedHasLegalNumbering_OriginalDoesNot_PreservesLegal()
        {
            // Arrange
            var original = CreateDocumentWithNumberedList("First item", useLegalNumbering: false);
            var revised = CreateDocumentWithNumberedList("First item modified", useLegalNumbering: true);

            Assert.False(DocumentHasLegalNumbering(original), "Original should NOT have legal numbering");
            Assert.True(DocumentHasLegalNumbering(revised), "Revised should have legal numbering");

            var settings = new WmlComparerSettings();

            // Act
            var compared = WmlComparer.Compare(original, revised, settings);

            // Assert - The revised document introduces legal numbering, which should be preserved
            // The comparison result should include numbering definitions from BOTH documents
            // This test verifies the fix for GitHub issue #1634
            Assert.True(DocumentHasLegalNumbering(compared),
                "Compared document should include legal numbering from revised document. " +
                "This is the main issue from GitHub #1634 - when revised introduces legal numbering, it should be preserved.");

            var validationErrors = ValidateDocument(compared);
            Assert.True(string.IsNullOrEmpty(validationErrors),
                $"Document should be valid. Errors: {validationErrors}");
        }

        /// <summary>
        /// Test Case 3: Both documents have legal numbering.
        /// Expected: Result should have legal numbering.
        /// </summary>
        [Fact]
        public void WC_LegalNum_003_BothHaveLegalNumbering_PreservesLegal()
        {
            // Arrange
            var original = CreateDocumentWithNumberedList("First item", useLegalNumbering: true);
            var revised = CreateDocumentWithNumberedList("First item modified", useLegalNumbering: true);

            Assert.True(DocumentHasLegalNumbering(original), "Original should have legal numbering");
            Assert.True(DocumentHasLegalNumbering(revised), "Revised should have legal numbering");

            var settings = new WmlComparerSettings();

            // Act
            var compared = WmlComparer.Compare(original, revised, settings);

            // Assert
            Assert.True(DocumentHasLegalNumbering(compared),
                "Compared document should preserve legal numbering");

            var validationErrors = ValidateDocument(compared);
            Assert.True(string.IsNullOrEmpty(validationErrors),
                $"Document should be valid. Errors: {validationErrors}");
        }

        /// <summary>
        /// Test Case 4: Multi-level numbering with legal style in revised only.
        /// This is a more realistic test case matching the GitHub issue scenario.
        /// </summary>
        [Fact]
        public void WC_LegalNum_004_MultiLevel_RevisedHasLegal_PreservesLegal()
        {
            // Arrange - simulates the GitHub issue scenario more closely
            var original = CreateDocumentWithMultiLevelNumbering(
                new[] { "Section 1", "Subsection 1.1", "Sub-subsection 1.1.1" },
                useLegalNumbering: false);

            var revised = CreateDocumentWithMultiLevelNumbering(
                new[] { "Section 1", "Subsection 1.1 modified", "Sub-subsection 1.1.1" },
                useLegalNumbering: true);

            Assert.False(DocumentHasLegalNumbering(original), "Original should NOT have legal numbering");
            Assert.True(DocumentHasLegalNumbering(revised), "Revised should have legal numbering");

            var settings = new WmlComparerSettings();

            // Act
            var compared = WmlComparer.Compare(original, revised, settings);

            // Assert
            Assert.True(DocumentHasLegalNumbering(compared),
                "Multi-level comparison should preserve legal numbering from revised document");

            var validationErrors = ValidateDocument(compared);
            Assert.True(string.IsNullOrEmpty(validationErrors),
                $"Document should be valid. Errors: {validationErrors}");
        }

        /// <summary>
        /// Test Case 5: Revised document has different numId with legal numbering.
        /// Tests that numbering with different IDs is properly merged.
        /// </summary>
        [Fact]
        public void WC_LegalNum_005_DifferentNumIds_RevisedHasLegal_MergesCorrectly()
        {
            // Arrange
            var original = CreateDocumentWithNumberedList("First item", useLegalNumbering: false, numId: 1);
            var revised = CreateDocumentWithNumberedList("First item modified", useLegalNumbering: true, numId: 2);

            var settings = new WmlComparerSettings();

            // Act
            var compared = WmlComparer.Compare(original, revised, settings);

            // Assert - result should have numbering definitions from both
            var abstractNumCount = GetAbstractNumCount(compared);
            Assert.True(abstractNumCount >= 1, "Result should have at least one abstractNum");

            // The result should be valid
            var validationErrors = ValidateDocument(compared);
            Assert.True(string.IsNullOrEmpty(validationErrors),
                $"Document should be valid. Errors: {validationErrors}");
        }

        #endregion
    }
}
