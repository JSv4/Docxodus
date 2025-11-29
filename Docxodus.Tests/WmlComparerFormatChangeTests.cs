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
using static Docxodus.WmlComparer;

namespace OxPt
{
    /// <summary>
    /// Tests for format change detection in WmlComparer.
    /// Tests verify that when text content is identical but formatting changes,
    /// the comparison produces native w:rPrChange markup.
    /// </summary>
    public class WmlComparerFormatChangeTests
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

        /// <summary>
        /// Creates a DOCX document with a paragraph that has bold text.
        /// </summary>
        private static WmlDocument CreateDocumentWithBoldParagraph(string text)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(
                    new Body(
                        new Paragraph(
                            new Run(
                                new RunProperties(new Bold()),
                                new Text(text)
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

        /// <summary>
        /// Creates a DOCX document with a paragraph that has italic text.
        /// </summary>
        private static WmlDocument CreateDocumentWithItalicParagraph(string text)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(
                    new Body(
                        new Paragraph(
                            new Run(
                                new RunProperties(new Italic()),
                                new Text(text)
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

        /// <summary>
        /// Creates a DOCX document with a paragraph that has both bold and italic text.
        /// </summary>
        private static WmlDocument CreateDocumentWithBoldItalicParagraph(string text)
        {
            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(
                    new Body(
                        new Paragraph(
                            new Run(
                                new RunProperties(new Bold(), new Italic()),
                                new Text(text)
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

        private static XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        #endregion

        #region Basic Format Change Detection Tests

        [Fact]
        public void FormatChange_AddBold_ShouldContainRPrChangeElement()
        {
            // Arrange: Create two documents - same text, doc2 has bold
            var doc1 = CreateDocumentWithParagraphs("This is some sample text.");
            var doc2 = CreateDocumentWithBoldParagraph("This is some sample text.");

            var settings = new WmlComparerSettings
            {
                DetectFormatChanges = true
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Assert: Check that the compared document contains w:rPrChange
            using var ms = new MemoryStream(compared.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var mainDoc = wDoc.MainDocumentPart.GetXDocument();
            var rPrChanges = mainDoc.Descendants(W + "rPrChange").ToList();

            Assert.NotEmpty(rPrChanges);
        }

        [Fact]
        public void FormatChange_RemoveBold_ShouldContainRPrChangeElement()
        {
            // Arrange: Create two documents - doc1 has bold, doc2 doesn't
            var doc1 = CreateDocumentWithBoldParagraph("This is some sample text.");
            var doc2 = CreateDocumentWithParagraphs("This is some sample text.");

            var settings = new WmlComparerSettings
            {
                DetectFormatChanges = true
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Assert: Check that the compared document contains w:rPrChange
            using var ms = new MemoryStream(compared.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var mainDoc = wDoc.MainDocumentPart.GetXDocument();
            var rPrChanges = mainDoc.Descendants(W + "rPrChange").ToList();

            Assert.NotEmpty(rPrChanges);
        }

        [Fact]
        public void FormatChange_BoldToItalic_ShouldContainRPrChangeElement()
        {
            // Arrange: Create two documents - doc1 has bold, doc2 has italic
            var doc1 = CreateDocumentWithBoldParagraph("This is some sample text.");
            var doc2 = CreateDocumentWithItalicParagraph("This is some sample text.");

            var settings = new WmlComparerSettings
            {
                DetectFormatChanges = true
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Assert: Check that the compared document contains w:rPrChange
            using var ms = new MemoryStream(compared.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var mainDoc = wDoc.MainDocumentPart.GetXDocument();
            var rPrChanges = mainDoc.Descendants(W + "rPrChange").ToList();

            Assert.NotEmpty(rPrChanges);
        }

        [Fact]
        public void FormatChange_AddMultipleFormats_ShouldContainRPrChangeElement()
        {
            // Arrange: Create two documents - plain to bold+italic
            var doc1 = CreateDocumentWithParagraphs("This is some sample text.");
            var doc2 = CreateDocumentWithBoldItalicParagraph("This is some sample text.");

            var settings = new WmlComparerSettings
            {
                DetectFormatChanges = true
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Assert: Check that the compared document contains w:rPrChange
            using var ms = new MemoryStream(compared.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var mainDoc = wDoc.MainDocumentPart.GetXDocument();
            var rPrChanges = mainDoc.Descendants(W + "rPrChange").ToList();

            Assert.NotEmpty(rPrChanges);
        }

        #endregion

        #region RprChange Attributes Tests

        [Fact]
        public void FormatChange_ShouldHaveRequiredAttributes()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs("This is some sample text.");
            var doc2 = CreateDocumentWithBoldParagraph("This is some sample text.");

            var settings = new WmlComparerSettings
            {
                DetectFormatChanges = true,
                AuthorForRevisions = "Test Author"
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Assert: Check that rPrChange has required attributes
            using var ms = new MemoryStream(compared.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var mainDoc = wDoc.MainDocumentPart.GetXDocument();
            var rPrChanges = mainDoc.Descendants(W + "rPrChange").ToList();

            Assert.NotEmpty(rPrChanges);
            foreach (var rPrChange in rPrChanges)
            {
                Assert.NotNull(rPrChange.Attribute(W + "id"));
                Assert.NotNull(rPrChange.Attribute(W + "author"));
                Assert.NotNull(rPrChange.Attribute(W + "date"));
                Assert.Equal("Test Author", (string)rPrChange.Attribute(W + "author"));
            }
        }

        [Fact]
        public void FormatChange_ShouldContainOldProperties()
        {
            // Arrange: bold to plain - the old properties should be bold
            var doc1 = CreateDocumentWithBoldParagraph("This is some sample text.");
            var doc2 = CreateDocumentWithParagraphs("This is some sample text.");

            var settings = new WmlComparerSettings
            {
                DetectFormatChanges = true
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Assert: Check that rPrChange contains rPr with the old properties
            using var ms = new MemoryStream(compared.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var mainDoc = wDoc.MainDocumentPart.GetXDocument();
            var rPrChanges = mainDoc.Descendants(W + "rPrChange").ToList();

            Assert.NotEmpty(rPrChanges);
            var rPrChange = rPrChanges.First();
            var oldRPr = rPrChange.Element(W + "rPr");
            Assert.NotNull(oldRPr);
            // The old properties should contain bold
            Assert.NotNull(oldRPr.Element(W + "b"));
        }

        #endregion

        #region Settings Tests

        [Fact]
        public void FormatChange_WhenDisabled_ShouldNotContainRPrChangeElement()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs("This is some sample text.");
            var doc2 = CreateDocumentWithBoldParagraph("This is some sample text.");

            var settings = new WmlComparerSettings
            {
                DetectFormatChanges = false  // Disabled
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Assert: Should NOT contain rPrChange
            using var ms = new MemoryStream(compared.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var mainDoc = wDoc.MainDocumentPart.GetXDocument();
            var rPrChanges = mainDoc.Descendants(W + "rPrChange").ToList();

            Assert.Empty(rPrChanges);
        }

        [Fact]
        public void FormatChange_NoFormattingChange_ShouldNotContainRPrChangeElement()
        {
            // Arrange: Same text, same formatting
            var doc1 = CreateDocumentWithBoldParagraph("This is some sample text.");
            var doc2 = CreateDocumentWithBoldParagraph("This is some sample text.");

            var settings = new WmlComparerSettings
            {
                DetectFormatChanges = true
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);

            // Assert: Should NOT contain rPrChange since formatting is the same
            using var ms = new MemoryStream(compared.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var mainDoc = wDoc.MainDocumentPart.GetXDocument();
            var rPrChanges = mainDoc.Descendants(W + "rPrChange").ToList();

            Assert.Empty(rPrChanges);
        }

        #endregion

        #region GetRevisions Tests

        [Fact]
        public void GetRevisions_FormatChange_ShouldReturnFormatChangedType()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs("This is some sample text.");
            var doc2 = CreateDocumentWithBoldParagraph("This is some sample text.");

            var settings = new WmlComparerSettings
            {
                DetectFormatChanges = true
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert
            var formatChangedRevisions = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.FormatChanged).ToList();
            Assert.NotEmpty(formatChangedRevisions);
        }

        [Fact]
        public void GetRevisions_FormatChange_ShouldHaveFormatChangeDetails()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs("This is some sample text.");
            var doc2 = CreateDocumentWithBoldParagraph("This is some sample text.");

            var settings = new WmlComparerSettings
            {
                DetectFormatChanges = true
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert
            var formatChangedRevisions = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.FormatChanged).ToList();
            Assert.NotEmpty(formatChangedRevisions);

            var revision = formatChangedRevisions.First();
            Assert.NotNull(revision.FormatChange);
            Assert.NotNull(revision.FormatChange.ChangedPropertyNames);
            Assert.Contains("bold", revision.FormatChange.ChangedPropertyNames);
        }

        [Fact]
        public void GetRevisions_FormatChange_ShouldHaveCorrectText()
        {
            // Arrange
            var doc1 = CreateDocumentWithParagraphs("This is some sample text.");
            var doc2 = CreateDocumentWithBoldParagraph("This is some sample text.");

            var settings = new WmlComparerSettings
            {
                DetectFormatChanges = true
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert
            var formatChangedRevisions = revisions.Where(r => r.RevisionType == WmlComparerRevisionType.FormatChanged).ToList();
            Assert.NotEmpty(formatChangedRevisions);

            var revision = formatChangedRevisions.First();
            Assert.NotNull(revision.Text);
            Assert.Contains("sample", revision.Text);
        }

        #endregion

        #region Integration with Text Changes Tests

        [Fact]
        public void FormatChange_WithTextChange_ShouldTrackBoth()
        {
            // Arrange: Both text and format change
            // Doc1: plain "Hello world"
            // Doc2: bold "Hello there" (both text and format change)
            var doc1 = CreateDocumentWithParagraphs("Hello world");

            using var stream = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(
                    new Body(
                        new Paragraph(
                            new Run(
                                new RunProperties(new Bold()),
                                new Text("Hello there")
                            )
                        )
                    )
                );

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

                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings();

                doc.Save();
            }

            stream.Position = 0;
            var doc2 = new WmlDocument("test.docx", stream.ToArray());

            var settings = new WmlComparerSettings
            {
                DetectFormatChanges = true
            };

            // Act
            var compared = WmlComparer.Compare(doc1, doc2, settings);
            var revisions = WmlComparer.GetRevisions(compared, settings);

            // Assert: Should have both text changes and potentially format changes
            Assert.True(revisions.Count > 0, "Should have revisions");

            // Text changes should result in ins/del
            var textRevisions = revisions.Where(r =>
                r.RevisionType == WmlComparerRevisionType.Inserted ||
                r.RevisionType == WmlComparerRevisionType.Deleted).ToList();
            Assert.NotEmpty(textRevisions);
        }

        #endregion
    }
}
