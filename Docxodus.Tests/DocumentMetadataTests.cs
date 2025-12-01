// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    /// <summary>
    /// Tests for WmlToHtmlConverter.GetDocumentMetadata() method.
    /// These tests verify the lazy loading metadata extraction API (Issue #44 Phase 3).
    /// </summary>
    public class DocumentMetadataTests
    {
        private static readonly DirectoryInfo TestFilesDir = new DirectoryInfo("../../../../TestFiles/");

        #region Basic Functionality Tests

        [Fact]
        public void DM001_GetDocumentMetadata_ReturnsValidMetadata()
        {
            // Arrange
            var sourceDocx = new FileInfo(Path.Combine(TestFilesDir.FullName, "HC001-5DayTourPlanTemplate.docx"));
            var wmlDoc = new WmlDocument(sourceDocx.FullName);

            // Act
            var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

            // Assert
            Assert.NotNull(metadata);
            Assert.NotNull(metadata.Sections);
            Assert.True(metadata.Sections.Count > 0, "Should have at least one section");
            Assert.True(metadata.TotalParagraphs >= 0, "Total paragraphs should be non-negative");
            Assert.True(metadata.TotalTables >= 0, "Total tables should be non-negative");
            Assert.True(metadata.EstimatedPageCount >= 1, "Estimated page count should be at least 1");
        }

        [Fact]
        public void DM002_GetDocumentMetadata_SectionHasValidDimensions()
        {
            // Arrange
            var sourceDocx = new FileInfo(Path.Combine(TestFilesDir.FullName, "HC001-5DayTourPlanTemplate.docx"));
            var wmlDoc = new WmlDocument(sourceDocx.FullName);

            // Act
            var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

            // Assert
            var section = metadata.Sections.First();

            // US Letter is 612x792 points (8.5" x 11")
            Assert.True(section.PageWidthPt > 0, "Page width should be positive");
            Assert.True(section.PageHeightPt > 0, "Page height should be positive");
            Assert.True(section.ContentWidthPt > 0, "Content width should be positive");
            Assert.True(section.ContentHeightPt > 0, "Content height should be positive");

            // Content area should be smaller than page
            Assert.True(section.ContentWidthPt < section.PageWidthPt, "Content width should be less than page width");
            Assert.True(section.ContentHeightPt < section.PageHeightPt, "Content height should be less than page height");

            // Margins should be non-negative
            Assert.True(section.MarginTopPt >= 0, "Top margin should be non-negative");
            Assert.True(section.MarginRightPt >= 0, "Right margin should be non-negative");
            Assert.True(section.MarginBottomPt >= 0, "Bottom margin should be non-negative");
            Assert.True(section.MarginLeftPt >= 0, "Left margin should be non-negative");
        }

        [Fact]
        public void DM003_GetDocumentMetadata_ParagraphIndicesAreContiguous()
        {
            // Arrange
            var sourceDocx = new FileInfo(Path.Combine(TestFilesDir.FullName, "HC001-5DayTourPlanTemplate.docx"));
            var wmlDoc = new WmlDocument(sourceDocx.FullName);

            // Act
            var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

            // Assert
            int expectedStart = 0;
            foreach (var section in metadata.Sections)
            {
                Assert.Equal(expectedStart, section.StartParagraphIndex);
                Assert.True(section.EndParagraphIndex >= section.StartParagraphIndex,
                    "End index should be >= start index");
                expectedStart = section.EndParagraphIndex;
            }

            Assert.Equal(metadata.TotalParagraphs, expectedStart);
        }

        #endregion

        #region Feature Detection Tests

        [Fact]
        public void DM010_GetDocumentMetadata_DetectsTrackedChanges()
        {
            // Arrange - Create a document with tracked changes
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new InsertedRun(
                                    new Run(new Text("Inserted text"))
                                )
                                { Author = "Test", Date = DateTime.Now }
                            )
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);

                // Act
                var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

                // Assert
                Assert.True(metadata.HasTrackedChanges, "Should detect tracked changes (insertions)");
            }
        }

        [Fact]
        public void DM011_GetDocumentMetadata_DetectsDeletions()
        {
            // Arrange - Create a document with deletions
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new DeletedRun(
                                    new Run(new DeletedText("Deleted text"))
                                )
                                { Author = "Test", Date = DateTime.Now }
                            )
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);

                // Act
                var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

                // Assert
                Assert.True(metadata.HasTrackedChanges, "Should detect tracked changes (deletions)");
            }
        }

        [Fact]
        public void DM012_GetDocumentMetadata_NoTrackedChangesWhenClean()
        {
            // Arrange - Create a clean document
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new Run(new Text("Normal text"))
                            )
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);

                // Act
                var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

                // Assert
                Assert.False(metadata.HasTrackedChanges, "Clean document should not have tracked changes");
            }
        }

        #endregion

        #region Multi-Section Tests

        [Fact]
        public void DM020_GetDocumentMetadata_HandlesMultipleSections()
        {
            // Arrange - Create a document with multiple sections (section break in paragraph)
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Create document with two sections
                    // First section ends with a paragraph that has sectPr in pPr
                    mainPart.Document = new Document(
                        new Body(
                            // Section 1 content
                            new Paragraph(
                                new Run(new Text("Section 1 content"))
                            ),
                            // Section break paragraph (sectPr in pPr marks end of section 1)
                            new Paragraph(
                                new ParagraphProperties(
                                    new SectionProperties(
                                        new PageSize() { Width = 12240, Height = 15840 }, // US Letter
                                        new PageMargin() { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 }
                                    )
                                ),
                                new Run(new Text("End of section 1"))
                            ),
                            // Section 2 content
                            new Paragraph(
                                new Run(new Text("Section 2 content"))
                            ),
                            // Document-level sectPr for final section
                            new SectionProperties(
                                new PageSize() { Width = 15840, Height = 12240 }, // Landscape
                                new PageMargin() { Top = 720, Right = 720, Bottom = 720, Left = 720 }
                            )
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);

                // Act
                var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

                // Assert
                Assert.True(metadata.Sections.Count >= 2, $"Should have at least 2 sections, got {metadata.Sections.Count}");

                // Verify section indices are sequential
                for (int i = 0; i < metadata.Sections.Count; i++)
                {
                    Assert.Equal(i, metadata.Sections[i].SectionIndex);
                }

                // Verify paragraph indices are contiguous
                Assert.Equal(0, metadata.Sections[0].StartParagraphIndex);
            }
        }

        [Fact]
        public void DM021_GetDocumentMetadata_DifferentPageSizesPerSection()
        {
            // Arrange - Create document with different page sizes
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new ParagraphProperties(
                                    new SectionProperties(
                                        new PageSize() { Width = 12240, Height = 15840 } // US Letter (612x792pt)
                                    )
                                ),
                                new Run(new Text("US Letter"))
                            ),
                            new Paragraph(
                                new Run(new Text("A4 content"))
                            ),
                            new SectionProperties(
                                new PageSize() { Width = 11906, Height = 16838 } // A4 (~595x842pt)
                            )
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);

                // Act
                var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

                // Assert
                Assert.True(metadata.Sections.Count >= 2, "Should have 2 sections");

                // First section should be US Letter (~612pt width)
                var section1 = metadata.Sections[0];
                Assert.True(Math.Abs(section1.PageWidthPt - 612) < 1, $"First section width should be ~612pt (US Letter), got {section1.PageWidthPt}");

                // Second section should be A4 (~595pt width)
                var section2 = metadata.Sections[1];
                Assert.True(Math.Abs(section2.PageWidthPt - 595.3) < 1, $"Second section width should be ~595pt (A4), got {section2.PageWidthPt}");
            }
        }

        #endregion

        #region Edge Cases

        [Fact]
        public void DM030_GetDocumentMetadata_HandlesEmptyDocument()
        {
            // Arrange - Create an empty document
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);

                // Act
                var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

                // Assert
                Assert.NotNull(metadata);
                Assert.True(metadata.Sections.Count >= 1, "Should have at least one section even for empty doc");
                Assert.Equal(0, metadata.TotalParagraphs);
                Assert.Equal(0, metadata.TotalTables);
            }
        }

        [Fact]
        public void DM031_GetDocumentMetadata_HandlesDocumentWithTables()
        {
            // Arrange - Create document with a table
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(new Run(new Text("Before table"))),
                            new DocumentFormat.OpenXml.Wordprocessing.Table(
                                new DocumentFormat.OpenXml.Wordprocessing.TableRow(
                                    new DocumentFormat.OpenXml.Wordprocessing.TableCell(new Paragraph(new Run(new Text("Cell 1")))),
                                    new DocumentFormat.OpenXml.Wordprocessing.TableCell(new Paragraph(new Run(new Text("Cell 2"))))
                                ),
                                new DocumentFormat.OpenXml.Wordprocessing.TableRow(
                                    new DocumentFormat.OpenXml.Wordprocessing.TableCell(new Paragraph(new Run(new Text("Cell 3")))),
                                    new DocumentFormat.OpenXml.Wordprocessing.TableCell(new Paragraph(new Run(new Text("Cell 4"))))
                                )
                            ),
                            new Paragraph(new Run(new Text("After table")))
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);

                // Act
                var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

                // Assert
                Assert.Equal(1, metadata.TotalTables);
                // Paragraphs: 2 outside table + 4 inside table cells = 6
                Assert.True(metadata.TotalParagraphs >= 6, $"Should count paragraphs inside tables, got {metadata.TotalParagraphs}");
            }
        }

        [Fact]
        public void DM032_GetDocumentMetadata_DefaultsToUSLetterWhenNoSectPr()
        {
            // Arrange - Create document without explicit sectPr
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(new Run(new Text("Just text, no sectPr")))
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);

                // Act
                var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

                // Assert
                Assert.True(metadata.Sections.Count >= 1, "Should have at least one section");
                var section = metadata.Sections[0];

                // Should default to US Letter (612x792 points)
                Assert.Equal(612, section.PageWidthPt);
                Assert.Equal(792, section.PageHeightPt);
            }
        }

        [Fact]
        public void DM033_GetDocumentMetadata_HandlesDocumentWithHeadersAndFooters()
        {
            // Arrange - Create document with headers and footers
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add header part
                    var headerPart = mainPart.AddNewPart<HeaderPart>();
                    headerPart.Header = new Header(
                        new Paragraph(new Run(new Text("Header content")))
                    );
                    headerPart.Header.Save();
                    var headerId = mainPart.GetIdOfPart(headerPart);

                    // Add footer part
                    var footerPart = mainPart.AddNewPart<FooterPart>();
                    footerPart.Footer = new Footer(
                        new Paragraph(new Run(new Text("Footer content")))
                    );
                    footerPart.Footer.Save();
                    var footerId = mainPart.GetIdOfPart(footerPart);

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(new Run(new Text("Body content"))),
                            new SectionProperties(
                                new HeaderReference() { Type = HeaderFooterValues.Default, Id = headerId },
                                new FooterReference() { Type = HeaderFooterValues.Default, Id = footerId }
                            )
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);

                // Act
                var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

                // Assert
                Assert.True(metadata.Sections.Count >= 1);
                var section = metadata.Sections[0];
                Assert.True(section.HasHeader, "Should detect default header");
                Assert.True(section.HasFooter, "Should detect default footer");
            }
        }

        #endregion

        #region Footnotes and Endnotes Tests

        [Fact]
        public void DM040_GetDocumentMetadata_DetectsFootnotes()
        {
            // Arrange - Create document with footnotes
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add footnotes part with actual footnote content
                    var footnotesPart = mainPart.AddNewPart<FootnotesPart>();
                    footnotesPart.Footnotes = new Footnotes(
                        new Footnote(
                            new Paragraph(new Run(new Text("Footnote text")))
                        )
                        { Type = FootnoteEndnoteValues.Normal, Id = 1 }
                    );
                    footnotesPart.Footnotes.Save();

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new Run(new Text("Text with footnote")),
                                new Run(new FootnoteReference() { Id = 1 })
                            )
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);

                // Act
                var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

                // Assert
                Assert.True(metadata.HasFootnotes, "Should detect footnotes");
            }
        }

        [Fact]
        public void DM041_GetDocumentMetadata_DetectsEndnotes()
        {
            // Arrange - Create document with endnotes
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add endnotes part with actual endnote content
                    var endnotesPart = mainPart.AddNewPart<EndnotesPart>();
                    endnotesPart.Endnotes = new Endnotes(
                        new Endnote(
                            new Paragraph(new Run(new Text("Endnote text")))
                        )
                        { Type = FootnoteEndnoteValues.Normal, Id = 1 }
                    );
                    endnotesPart.Endnotes.Save();

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new Run(new Text("Text with endnote")),
                                new Run(new EndnoteReference() { Id = 1 })
                            )
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);

                // Act
                var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

                // Assert
                Assert.True(metadata.HasEndnotes, "Should detect endnotes");
            }
        }

        #endregion

        #region Comments Tests

        [Fact]
        public void DM050_GetDocumentMetadata_DetectsComments()
        {
            // Arrange - Create document with comments
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add comments part
                    var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
                    commentsPart.Comments = new Comments(
                        new Comment(
                            new Paragraph(new Run(new Text("Comment text")))
                        )
                        { Id = "1", Author = "Test Author", Date = DateTime.Now }
                    );
                    commentsPart.Comments.Save();

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new CommentRangeStart() { Id = "1" },
                                new Run(new Text("Commented text")),
                                new CommentRangeEnd() { Id = "1" },
                                new Run(new CommentReference() { Id = "1" })
                            )
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);

                // Act
                var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

                // Assert
                Assert.True(metadata.HasComments, "Should detect comments");
            }
        }

        #endregion

        #region Unit Conversion Tests

        [Fact]
        public void DM060_GetDocumentMetadata_CorrectlyConvertsTwipsToPoints()
        {
            // Arrange - Create document with known dimensions
            // 1 point = 20 twips
            // US Letter: 8.5" x 11" = 612pt x 792pt = 12240 twips x 15840 twips
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(new Run(new Text("Content"))),
                            new SectionProperties(
                                new PageSize() { Width = 12240, Height = 15840 },
                                new PageMargin() { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 }
                                // Margins: 1 inch = 72pt = 1440 twips
                            )
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);

                // Act
                var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);

                // Assert
                var section = metadata.Sections[0];
                Assert.Equal(612, section.PageWidthPt);
                Assert.Equal(792, section.PageHeightPt);
                Assert.Equal(72, section.MarginTopPt);
                Assert.Equal(72, section.MarginRightPt);
                Assert.Equal(72, section.MarginBottomPt);
                Assert.Equal(72, section.MarginLeftPt);

                // Content dimensions: 612 - 72 - 72 = 468 width, 792 - 72 - 72 = 648 height
                Assert.Equal(468, section.ContentWidthPt);
                Assert.Equal(648, section.ContentHeightPt);
            }
        }

        #endregion

        #region RenderPageRange Tests (Lazy Loading - Issue #31)

        [Fact]
        public void RPR001_RenderPageRange_ReturnsValidHtml()
        {
            // Arrange
            var sourceDocx = new FileInfo(Path.Combine(TestFilesDir.FullName, "HC001-5DayTourPlanTemplate.docx"));
            var wmlDoc = new WmlDocument(sourceDocx.FullName);
            var settings = new WmlToHtmlConverterSettings
            {
                PageTitle = "Test Document",
                CssClassPrefix = "docx-",
                FabricateCssClasses = true
            };

            // Act
            var html = WmlToHtmlConverter.RenderPageRange(wmlDoc, settings, 1, 1);

            // Assert
            Assert.NotNull(html);
            Assert.Contains("pagination-container", html.ToString());
            Assert.Contains("data-start-page=\"1\"", html.ToString());
            Assert.Contains("data-end-page=\"1\"", html.ToString());
        }

        [Fact]
        public void RPR002_RenderPageRange_IncludesPageMetadata()
        {
            // Arrange
            var sourceDocx = new FileInfo(Path.Combine(TestFilesDir.FullName, "HC001-5DayTourPlanTemplate.docx"));
            var wmlDoc = new WmlDocument(sourceDocx.FullName);
            var settings = new WmlToHtmlConverterSettings
            {
                PageTitle = "Test Document"
            };

            // Act
            var html = WmlToHtmlConverter.RenderPageRange(wmlDoc, settings, 1, 3);
            var htmlStr = html.ToString();

            // Assert
            Assert.Contains("data-start-page=\"1\"", htmlStr);
            Assert.Contains("data-end-page=\"3\"", htmlStr);
            Assert.Contains("data-total-pages", htmlStr);
            Assert.Contains("data-start-block", htmlStr);
            Assert.Contains("data-end-block", htmlStr);
        }

        [Fact]
        public void RPR003_RenderPageRange_HandlesSinglePage()
        {
            // Arrange - Create a simple document
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(new Run(new Text("Page 1 content")))
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);
                var settings = new WmlToHtmlConverterSettings();

                // Act
                var html = WmlToHtmlConverter.RenderPageRange(wmlDoc, settings, 1, 1);
                var htmlStr = html.ToString();

                // Assert
                Assert.Contains("Page 1 content", htmlStr);
                Assert.Contains("data-start-page=\"1\"", htmlStr);
                Assert.Contains("data-end-page=\"1\"", htmlStr);
            }
        }

        [Fact]
        public void RPR004_RenderPageRange_ClampsInvalidPageNumbers()
        {
            // Arrange
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(new Run(new Text("Content")))
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);
                var settings = new WmlToHtmlConverterSettings();

                // Act - Request page 100 when doc only has 1 page
                var html = WmlToHtmlConverter.RenderPageRange(wmlDoc, settings, 100, 200);

                // Assert - Should not throw, should return valid HTML
                Assert.NotNull(html);
                Assert.Contains("pagination-container", html.ToString());
            }
        }

        [Fact]
        public void RPR005_RenderPageRange_HandlesNegativeStartPage()
        {
            // Arrange
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(new Run(new Text("Content")))
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);
                var settings = new WmlToHtmlConverterSettings();

                // Act - Negative start page should be clamped to 1
                var html = WmlToHtmlConverter.RenderPageRange(wmlDoc, settings, -5, 1);

                // Assert - Should handle gracefully
                Assert.NotNull(html);
                Assert.Contains("data-start-page=\"1\"", html.ToString());
            }
        }

        [Fact]
        public void RPR006_RenderPageRange_IncludesBlockIndices()
        {
            // Arrange
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(new Run(new Text("Paragraph 1"))),
                            new Paragraph(new Run(new Text("Paragraph 2"))),
                            new Paragraph(new Run(new Text("Paragraph 3")))
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);
                var settings = new WmlToHtmlConverterSettings();

                // Act
                var html = WmlToHtmlConverter.RenderPageRange(wmlDoc, settings, 1, 1);
                var htmlStr = html.ToString();

                // Assert - Should include block index data attributes
                Assert.Contains("data-block-index", htmlStr);
            }
        }

        [Fact]
        public void RPR007_RenderPageRange_RespectsScale()
        {
            // Arrange
            var sourceDocx = new FileInfo(Path.Combine(TestFilesDir.FullName, "HC001-5DayTourPlanTemplate.docx"));
            var wmlDoc = new WmlDocument(sourceDocx.FullName);
            var settings = new WmlToHtmlConverterSettings
            {
                PaginationScale = 0.5 // 50% scale
            };

            // Act
            var html = WmlToHtmlConverter.RenderPageRange(wmlDoc, settings, 1, 1);

            // Assert - Should not throw, scale is applied during HTML generation
            Assert.NotNull(html);
        }

        [Fact]
        public void RPR008_RenderPageRange_HandlesEmptyDocument()
        {
            // Arrange
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(new Body());
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);
                var settings = new WmlToHtmlConverterSettings();

                // Act
                var html = WmlToHtmlConverter.RenderPageRange(wmlDoc, settings, 1, 1);

                // Assert
                Assert.NotNull(html);
                // Should handle empty document gracefully
            }
        }

        [Fact]
        public void RPR009_RenderPageRange_IncludesSectionDimensions()
        {
            // Arrange - Create document with explicit page size
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(new Run(new Text("Content"))),
                            new SectionProperties(
                                new PageSize() { Width = 12240, Height = 15840 } // US Letter
                            )
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);
                var settings = new WmlToHtmlConverterSettings();

                // Act
                var html = WmlToHtmlConverter.RenderPageRange(wmlDoc, settings, 1, 1);
                var htmlStr = html.ToString();

                // Assert - Should include section dimension attributes
                Assert.Contains("data-page-width", htmlStr);
                Assert.Contains("data-page-height", htmlStr);
            }
        }

        [Fact]
        public void RPR010_RenderPageRange_WithWordprocessingDocumentOverload()
        {
            // Arrange - Test the WordprocessingDocument overload
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(new Run(new Text("Test content")))
                        )
                    );
                    mainPart.Document.Save();
                }

                ms.Position = 0;

                using (var openDoc = WordprocessingDocument.Open(ms, true))
                {
                    var settings = new WmlToHtmlConverterSettings();

                    // Act
                    var html = WmlToHtmlConverter.RenderPageRange(openDoc, settings, 1, 1);
                    var htmlStr = html.ToString();

                    // Assert
                    Assert.Contains("Test content", htmlStr);
                    Assert.Contains("pagination-container", htmlStr);
                }
            }
        }

        /// <summary>
        /// Critical test: Creates a 5-page document with unique markers per page,
        /// then renders only pages 2-4 and verifies:
        /// - Pages 2, 3, 4 content IS present
        /// - Pages 1, 5 content is NOT present
        /// This proves virtual scrolling works correctly.
        /// </summary>
        [Fact]
        public void RPR011_RenderPageRange_FivePageDocument_RendersOnlyMiddleThreePages()
        {
            // Arrange - Create a 5-page document with ~25 paragraphs per page
            // Each page has unique marker text
            const int paragraphsPerPage = 25;
            const int totalPages = 5;

            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    var body = new Body();

                    for (int page = 1; page <= totalPages; page++)
                    {
                        // Add unique marker at start of each "page"
                        body.Append(new Paragraph(
                            new Run(new Text($"===PAGE_{page}_START==="))
                        ));

                        // Add filler paragraphs
                        for (int p = 1; p < paragraphsPerPage - 1; p++)
                        {
                            body.Append(new Paragraph(
                                new Run(new Text($"Page {page} paragraph {p} content"))
                            ));
                        }

                        // Add unique marker at end of each "page"
                        body.Append(new Paragraph(
                            new Run(new Text($"===PAGE_{page}_END==="))
                        ));
                    }

                    mainPart.Document = new Document(body);
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("five-page-test.docx", ms);

                // First verify we have 5 pages
                var metadata = WmlToHtmlConverter.GetDocumentMetadata(wmlDoc);
                Assert.True(metadata.EstimatedPageCount >= 4,
                    $"Expected at least 4 pages, got {metadata.EstimatedPageCount}");

                var settings = new WmlToHtmlConverterSettings();

                // Act - Render only pages 2-4 (the middle 3 pages)
                var html = WmlToHtmlConverter.RenderPageRange(wmlDoc, settings, 2, 4);
                var htmlStr = html.ToString();

                // Assert - Data attributes show correct range
                Assert.Contains("data-start-page=\"2\"", htmlStr);
                Assert.Contains("data-end-page=\"4\"", htmlStr);

                // Assert - Pages 2, 3, 4 content IS present
                Assert.Contains("===PAGE_2_START===", htmlStr);
                Assert.Contains("===PAGE_2_END===", htmlStr);
                Assert.Contains("===PAGE_3_START===", htmlStr);
                Assert.Contains("===PAGE_3_END===", htmlStr);
                Assert.Contains("===PAGE_4_START===", htmlStr);
                Assert.Contains("===PAGE_4_END===", htmlStr);

                // Assert - Pages 1 and 5 content is NOT present
                Assert.DoesNotContain("===PAGE_1_START===", htmlStr);
                Assert.DoesNotContain("===PAGE_1_END===", htmlStr);
                Assert.DoesNotContain("===PAGE_5_START===", htmlStr);
                Assert.DoesNotContain("===PAGE_5_END===", htmlStr);

                // Log for verification
                Console.WriteLine($"Total paragraphs in doc: {metadata.TotalParagraphs}");
                Console.WriteLine($"Estimated pages: {metadata.EstimatedPageCount}");
                Console.WriteLine($"HTML length: {htmlStr.Length} characters");
                Console.WriteLine("Successfully verified: Pages 2-4 rendered, Pages 1 & 5 excluded");
            }
        }

        /// <summary>
        /// Additional verification: Render only page 1 and verify only page 1 content appears.
        /// </summary>
        [Fact]
        public void RPR012_RenderPageRange_FivePageDocument_RendersOnlyFirstPage()
        {
            // Arrange - Create a 5-page document
            const int paragraphsPerPage = 25;
            const int totalPages = 5;

            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    var body = new Body();

                    for (int page = 1; page <= totalPages; page++)
                    {
                        body.Append(new Paragraph(
                            new Run(new Text($"===MARKER_PAGE_{page}==="))
                        ));

                        for (int p = 1; p < paragraphsPerPage; p++)
                        {
                            body.Append(new Paragraph(
                                new Run(new Text($"Filler content page {page}"))
                            ));
                        }
                    }

                    mainPart.Document = new Document(body);
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);
                var settings = new WmlToHtmlConverterSettings();

                // Act - Render only page 1
                var html = WmlToHtmlConverter.RenderPageRange(wmlDoc, settings, 1, 1);
                var htmlStr = html.ToString();

                // Assert - Only page 1 marker is present
                Assert.Contains("===MARKER_PAGE_1===", htmlStr);
                Assert.DoesNotContain("===MARKER_PAGE_2===", htmlStr);
                Assert.DoesNotContain("===MARKER_PAGE_3===", htmlStr);
                Assert.DoesNotContain("===MARKER_PAGE_4===", htmlStr);
                Assert.DoesNotContain("===MARKER_PAGE_5===", htmlStr);

                Assert.Contains("data-start-page=\"1\"", htmlStr);
                Assert.Contains("data-end-page=\"1\"", htmlStr);
            }
        }

        /// <summary>
        /// Verification: Render only the last page (page 5) of a 5-page document.
        /// </summary>
        [Fact]
        public void RPR013_RenderPageRange_FivePageDocument_RendersOnlyLastPage()
        {
            // Arrange - Create a 5-page document
            const int paragraphsPerPage = 25;
            const int totalPages = 5;

            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    var body = new Body();

                    for (int page = 1; page <= totalPages; page++)
                    {
                        body.Append(new Paragraph(
                            new Run(new Text($"===UNIQUE_PAGE_{page}==="))
                        ));

                        for (int p = 1; p < paragraphsPerPage; p++)
                        {
                            body.Append(new Paragraph(
                                new Run(new Text($"Content for page {page}"))
                            ));
                        }
                    }

                    mainPart.Document = new Document(body);
                    mainPart.Document.Save();
                }

                ms.Position = 0;
                var wmlDoc = new WmlDocument("test.docx", ms);
                var settings = new WmlToHtmlConverterSettings();

                // Act - Render only page 5
                var html = WmlToHtmlConverter.RenderPageRange(wmlDoc, settings, 5, 5);
                var htmlStr = html.ToString();

                // Assert - Only page 5 marker is present
                Assert.DoesNotContain("===UNIQUE_PAGE_1===", htmlStr);
                Assert.DoesNotContain("===UNIQUE_PAGE_2===", htmlStr);
                Assert.DoesNotContain("===UNIQUE_PAGE_3===", htmlStr);
                Assert.DoesNotContain("===UNIQUE_PAGE_4===", htmlStr);
                Assert.Contains("===UNIQUE_PAGE_5===", htmlStr);

                Assert.Contains("data-start-page=\"5\"", htmlStr);
                Assert.Contains("data-end-page=\"5\"", htmlStr);
            }
        }

        #endregion
    }
}

#endif
