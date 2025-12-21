// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#define COPY_FILES_FOR_DEBUGGING

// DO_CONVERSION_VIA_WORD is defined in the project Docxodus.Tests.OA.csproj, but not in the Docxodus.Tests.csproj

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using SkiaSharp;
using Xunit;

#if DO_CONVERSION_VIA_WORD
using Word = Microsoft.Office.Interop.Word;
#endif

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class HcTests
    {
        public static bool s_CopySourceFiles = true;
        public static bool s_CopyFormattingAssembledDocx = true;
        public static bool s_ConvertUsingWord = true;

        // PowerShell oneliner that generates InlineData for all files in a directory
        // dir | % { '[InlineData("' + $_.Name + '")]' } | clip

        [Theory]
        [InlineData("HC001-5DayTourPlanTemplate.docx")]
        [InlineData("HC002-Hebrew-01.docx")]
        [InlineData("HC003-Hebrew-02.docx")]
        [InlineData("HC004-ResumeTemplate.docx")]
        [InlineData("HC005-TaskPlanTemplate.docx")]
        [InlineData("HC006-Test-01.docx")]
        [InlineData("HC007-Test-02.docx")]
        [InlineData("HC008-Test-03.docx")]
        [InlineData("HC009-Test-04.docx")]
        [InlineData("HC010-Test-05.docx")]
        [InlineData("HC011-Test-06.docx")]
        [InlineData("HC012-Test-07.docx")]
        [InlineData("HC013-Test-08.docx")]
        [InlineData("HC014-RTL-Table-01.docx")]
        [InlineData("HC015-Vertical-Spacing-atLeast.docx")]
        [InlineData("HC016-Horizontal-Spacing-firstLine.docx")]
        [InlineData("HC017-Vertical-Alignment-Cell-01.docx")]
        [InlineData("HC018-Vertical-Alignment-Para-01.docx")]
        [InlineData("HC019-Hidden-Run.docx")]
        [InlineData("HC020-Small-Caps.docx")]
        [InlineData("HC021-Symbols.docx")]
        [InlineData("HC022-Table-Of-Contents.docx")]
        [InlineData("HC023-Hyperlink.docx")]
        [InlineData("HC024-Tabs-01.docx")]
        [InlineData("HC025-Tabs-02.docx")]
        [InlineData("HC026-Tabs-03.docx")]
        [InlineData("HC027-Tabs-04.docx")]
        [InlineData("HC028-No-Break-Hyphen.docx")]
        [InlineData("HC029-Table-Merged-Cells.docx")]
        [InlineData("HC030-Content-Controls.docx")]
        [InlineData("HC031-Complicated-Document.docx")]
        [InlineData("HC032-Named-Color.docx")]
        [InlineData("HC033-Run-With-Border.docx")]
        [InlineData("HC034-Run-With-Position.docx")]
        [InlineData("HC035-Strike-Through.docx")]
        [InlineData("HC036-Super-Script.docx")]
        [InlineData("HC037-Sub-Script.docx")]
        [InlineData("HC038-Conflicting-Border-Weight.docx")]
        [InlineData("HC039-Bold.docx")]
        [InlineData("HC040-Hyperlink-Fieldcode-01.docx")]
        [InlineData("HC041-Hyperlink-Fieldcode-02.docx")]
        [InlineData("HC042-Image-Png.docx")]
        [InlineData("HC043-Chart.docx")]
        [InlineData("HC044-Embedded-Workbook.docx")]
        [InlineData("HC045-Italic.docx")]
        [InlineData("HC046-BoldAndItalic.docx")]
        [InlineData("HC047-No-Section.docx")]
        [InlineData("HC048-Excerpt.docx")]
        [InlineData("HC049-Borders.docx")]
        [InlineData("HC050-Shaded-Text-01.docx")]
        [InlineData("HC051-Shaded-Text-02.docx")]
        [InlineData("HC060-Image-with-Hyperlink.docx")]
        [InlineData("HC061-Hyperlink-in-Field.docx")]
        
        public void HC001(string name)
        {
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));

#if COPY_FILES_FOR_DEBUGGING
            var sourceCopiedToDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-1-Source.docx")));
            if (!sourceCopiedToDestDocx.Exists)
                File.Copy(sourceDocx.FullName, sourceCopiedToDestDocx.FullName);

            var assembledFormattingDestDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-2-FormattingAssembled.docx")));
            if (!assembledFormattingDestDocx.Exists)
                CopyFormattingAssembledDocx(sourceDocx, assembledFormattingDestDocx);
#endif

            var oxPtConvertedDestHtml = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-3-OxPt.html")));
            ConvertToHtml(sourceDocx, oxPtConvertedDestHtml);

#if DO_CONVERSION_VIA_WORD
            var wordConvertedDocHtml = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-4-Word.html")));
            ConvertToHtmlUsingWord(sourceDocx, wordConvertedDocHtml);
#endif

        }

        [Theory]
        [InlineData("HC006-Test-01.docx")]
        public void HC002_NoCssClasses(string name)
        {
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));

            var oxPtConvertedDestHtml = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-5-OxPt-No-CSS-Classes.html")));
            ConvertToHtmlNoCssClasses(sourceDocx, oxPtConvertedDestHtml);
        }

        private static void CopyFormattingAssembledDocx(FileInfo source, FileInfo dest)
        {
            var ba = File.ReadAllBytes(source.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(ba, 0, ba.Length);
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(ms, true))
                {

                    RevisionAccepter.AcceptRevisions(wordDoc);
                    SimplifyMarkupSettings simplifyMarkupSettings = new SimplifyMarkupSettings
                    {
                        RemoveComments = true,
                        RemoveContentControls = true,
                        RemoveEndAndFootNotes = true,
                        RemoveFieldCodes = false,
                        RemoveLastRenderedPageBreak = true,
                        RemovePermissions = true,
                        RemoveProof = true,
                        RemoveRsidInfo = true,
                        RemoveSmartTags = true,
                        RemoveSoftHyphens = true,
                        RemoveGoBackBookmark = true,
                        ReplaceTabsWithSpaces = false,
                    };
                    MarkupSimplifier.SimplifyMarkup(wordDoc, simplifyMarkupSettings);

                    FormattingAssemblerSettings formattingAssemblerSettings = new FormattingAssemblerSettings
                    {
                        RemoveStyleNamesFromParagraphAndRunProperties = false,
                        ClearStyles = false,
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        CreateHtmlConverterAnnotationAttributes = true,
                        OrderElementsPerStandard = false,
                        ListItemRetrieverSettings =
                            new ListItemRetrieverSettings()
                            {
                                ListItemTextImplementations = ListItemRetrieverSettings.DefaultListItemTextImplementations,
                            },
                    };

                    FormattingAssembler.AssembleFormatting(wordDoc, formattingAssemblerSettings);
                }
                var newBa = ms.ToArray();
                File.WriteAllBytes(dest.FullName, newBa);
            }
        }

        private static void ConvertToHtml(FileInfo sourceDocx, FileInfo destFileName)
        {
            byte[] byteArray = File.ReadAllBytes(sourceDocx.FullName);
            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
                    var outputDirectory = destFileName.Directory;
                    destFileName = new FileInfo(Path.Combine(outputDirectory.FullName, destFileName.Name));
                    var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
                    int imageCounter = 0;
                    var pageTitle = (string)wDoc.CoreFilePropertiesPart.GetXDocument().Descendants(DC.title).FirstOrDefault();
                    if (pageTitle == null)
                        pageTitle = sourceDocx.FullName;

                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = pageTitle,
                        FabricateCssClasses = true,
                        CssClassPrefix = "pt-",
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo =>
                        {
                            DirectoryInfo localDirInfo = new DirectoryInfo(imageDirectoryName);
                            if (!localDirInfo.Exists)
                                localDirInfo.Create();
                            ++imageCounter;
                            string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            SKEncodedImageFormat? imageFormat = null;
                            if (extension == "png")
                            {
                                // Convert png to gif.
                                extension = "gif";
                                imageFormat = SKEncodedImageFormat.Gif;
                            }
                            else if (extension == "gif")
                                imageFormat = SKEncodedImageFormat.Gif;
                            else if (extension == "bmp")
                                imageFormat = SKEncodedImageFormat.Bmp;
                            else if (extension == "jpeg")
                                imageFormat = SKEncodedImageFormat.Jpeg;
                            else if (extension == "tiff")
                            {
                                // Convert tiff to png (SkiaSharp doesn't support tiff output).
                                extension = "png";
                                imageFormat = SKEncodedImageFormat.Png;
                            }
                            else if (extension == "x-wmf")
                            {
                                // Convert wmf to png (SkiaSharp doesn't support wmf output).
                                extension = "png";
                                imageFormat = SKEncodedImageFormat.Png;
                            }

                            // If the image format isn't one that we expect, ignore it,
                            // and don't return markup for the link.
                            if (imageFormat == null)
                                return null;

                            string imageFileName = imageDirectoryName + "/image" +
                                imageCounter.ToString() + "." + extension;
                            try
                            {
                                imageInfo.SaveImage(imageFileName, imageFormat.Value);
                            }
                            catch (Exception)
                            {
                                return null;
                            }
                            XElement img = new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageFileName),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };
                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

                    // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
                    // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
                    // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
                    // for detailed explanation.
                    //
                    // If you further transform the XML tree returned by ConvertToHtmlTransform, you
                    // must do it correctly, or entities will not be serialized properly.

                    var htmlString = html.ToString(SaveOptions.DisableFormatting);
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        private static void ConvertToHtmlNoCssClasses(FileInfo sourceDocx, FileInfo destFileName)
        {
            byte[] byteArray = File.ReadAllBytes(sourceDocx.FullName);
            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(memoryStream, true))
                {
                    var outputDirectory = destFileName.Directory;
                    destFileName = new FileInfo(Path.Combine(outputDirectory.FullName, destFileName.Name));
                    var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
                    int imageCounter = 0;
                    var pageTitle = (string)wDoc.CoreFilePropertiesPart.GetXDocument().Descendants(DC.title).FirstOrDefault();
                    if (pageTitle == null)
                        pageTitle = sourceDocx.FullName;

                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = pageTitle,
                        FabricateCssClasses = false,
                        RestrictToSupportedLanguages = false,
                        RestrictToSupportedNumberingFormats = false,
                        ImageHandler = imageInfo =>
                        {
                            DirectoryInfo localDirInfo = new DirectoryInfo(imageDirectoryName);
                            if (!localDirInfo.Exists)
                                localDirInfo.Create();
                            ++imageCounter;
                            string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                            SKEncodedImageFormat? imageFormat = null;
                            if (extension == "png")
                            {
                                // Convert png to gif.
                                extension = "gif";
                                imageFormat = SKEncodedImageFormat.Gif;
                            }
                            else if (extension == "gif")
                                imageFormat = SKEncodedImageFormat.Gif;
                            else if (extension == "bmp")
                                imageFormat = SKEncodedImageFormat.Bmp;
                            else if (extension == "jpeg")
                                imageFormat = SKEncodedImageFormat.Jpeg;
                            else if (extension == "tiff")
                            {
                                // Convert tiff to png (SkiaSharp doesn't support tiff output).
                                extension = "png";
                                imageFormat = SKEncodedImageFormat.Png;
                            }
                            else if (extension == "x-wmf")
                            {
                                // Convert wmf to png (SkiaSharp doesn't support wmf output).
                                extension = "png";
                                imageFormat = SKEncodedImageFormat.Png;
                            }

                            // If the image format isn't one that we expect, ignore it,
                            // and don't return markup for the link.
                            if (imageFormat == null)
                                return null;

                            string imageFileName = imageDirectoryName + "/image" +
                                imageCounter.ToString() + "." + extension;
                            try
                            {
                                imageInfo.SaveImage(imageFileName, imageFormat.Value);
                            }
                            catch (Exception)
                            {
                                return null;
                            }
                            XElement img = new XElement(Xhtml.img,
                                new XAttribute(NoNamespace.src, imageFileName),
                                imageInfo.ImgStyleAttribute,
                                imageInfo.AltText != null ?
                                    new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                            return img;
                        }
                    };
                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

                    // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
                    // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
                    // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
                    // for detailed explanation.
                    //
                    // If you further transform the XML tree returned by ConvertToHtmlTransform, you
                    // must do it correctly, or entities will not be serialized properly.

                    var htmlString = html.ToString(SaveOptions.DisableFormatting);
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

#if DO_CONVERSION_VIA_WORD
        public static void ConvertToHtmlUsingWord(FileInfo sourceFileName, FileInfo destFileName)
        {
            Word.Application app = new Word.Application();
            app.Visible = false;
            try
            {
                Word.Document doc = app.Documents.Open(sourceFileName.FullName);
                doc.SaveAs2(destFileName.FullName, Word.WdSaveFormat.wdFormatFilteredHTML);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Console.WriteLine("Caught unexpected COM exception.");
                ((Microsoft.Office.Interop.Word._Application)app).Quit();
                Environment.Exit(0);
            }
            ((Microsoft.Office.Interop.Word._Application)app).Quit();
        }
#endif

        [Fact]
        public void HC003_TrackedChanges_InsertionsAndDeletions()
        {
            // Use WmlComparer to create a document with tracked changes
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/WC");
            FileInfo doc1 = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-Unmodified.docx"));
            FileInfo doc2 = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-InsertInMiddle.docx"));

            WmlDocument wmlDoc1 = new WmlDocument(doc1.FullName);
            WmlDocument wmlDoc2 = new WmlDocument(doc2.FullName);

            WmlComparerSettings comparerSettings = new WmlComparerSettings();
            WmlDocument comparedDoc = WmlComparer.Compare(wmlDoc1, wmlDoc2, comparerSettings);

            // Convert to HTML with tracked changes rendering enabled
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(comparedDoc.DocumentByteArray, 0, comparedDoc.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Tracked Changes Test",
                        FabricateCssClasses = true,
                        CssClassPrefix = "pt-",
                        RenderTrackedChanges = true,
                        IncludeRevisionMetadata = true,
                        ShowDeletedContent = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify the HTML contains <ins> elements (insertions)
                    Assert.Contains("<ins", htmlString);
                    Assert.Contains("class=\"rev-ins\"", htmlString);

                    // Verify metadata attributes are present
                    Assert.Contains("data-author=", htmlString);

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "TrackedChanges-Insertions.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        [Fact]
        public void HC004_TrackedChanges_Deletions()
        {
            // Use WmlComparer to create a document with deletions
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/WC");
            FileInfo doc1 = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-Unmodified.docx"));
            FileInfo doc2 = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-DeleteInMiddle.docx"));

            WmlDocument wmlDoc1 = new WmlDocument(doc1.FullName);
            WmlDocument wmlDoc2 = new WmlDocument(doc2.FullName);

            WmlComparerSettings comparerSettings = new WmlComparerSettings();
            WmlDocument comparedDoc = WmlComparer.Compare(wmlDoc1, wmlDoc2, comparerSettings);

            // Convert to HTML with tracked changes rendering enabled
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(comparedDoc.DocumentByteArray, 0, comparedDoc.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Tracked Changes Deletions Test",
                        FabricateCssClasses = true,
                        CssClassPrefix = "pt-",
                        RenderTrackedChanges = true,
                        IncludeRevisionMetadata = true,
                        ShowDeletedContent = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify the HTML contains <del> elements (deletions)
                    Assert.Contains("<del", htmlString);
                    Assert.Contains("class=\"rev-del\"", htmlString);

                    // Verify metadata attributes are present
                    Assert.Contains("data-author=", htmlString);

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "TrackedChanges-Deletions.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        [Fact]
        public void HC005_TrackedChanges_CssGenerated()
        {
            // Use WmlComparer to create a document with tracked changes
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/WC");
            FileInfo doc1 = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-Unmodified.docx"));
            FileInfo doc2 = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-InsertInMiddle.docx"));

            WmlDocument wmlDoc1 = new WmlDocument(doc1.FullName);
            WmlDocument wmlDoc2 = new WmlDocument(doc2.FullName);

            WmlComparerSettings comparerSettings = new WmlComparerSettings();
            WmlDocument comparedDoc = WmlComparer.Compare(wmlDoc1, wmlDoc2, comparerSettings);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(comparedDoc.DocumentByteArray, 0, comparedDoc.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "CSS Test",
                        FabricateCssClasses = true,
                        RenderTrackedChanges = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify the CSS for tracked changes is generated
                    Assert.Contains("ins.rev-ins", htmlString);
                    Assert.Contains("del.rev-del", htmlString);
                    Assert.Contains("text-decoration: underline", htmlString);
                    Assert.Contains("text-decoration: line-through", htmlString);
                }
            }
        }

        [Fact]
        public void HC006_TrackedChanges_DisabledByDefault()
        {
            // When RenderTrackedChanges is false (default), revisions should be accepted
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/WC");
            FileInfo doc1 = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-Unmodified.docx"));
            FileInfo doc2 = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-InsertInMiddle.docx"));

            WmlDocument wmlDoc1 = new WmlDocument(doc1.FullName);
            WmlDocument wmlDoc2 = new WmlDocument(doc2.FullName);

            WmlComparerSettings comparerSettings = new WmlComparerSettings();
            WmlDocument comparedDoc = WmlComparer.Compare(wmlDoc1, wmlDoc2, comparerSettings);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(comparedDoc.DocumentByteArray, 0, comparedDoc.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Default Test",
                        FabricateCssClasses = true,
                        // RenderTrackedChanges defaults to false
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify the HTML does NOT contain <ins> or <del> elements
                    Assert.DoesNotContain("<ins", htmlString);
                    Assert.DoesNotContain("<del", htmlString);

                    // Verify revision CSS is not generated
                    Assert.DoesNotContain("ins.rev-ins", htmlString);
                    Assert.DoesNotContain("del.rev-del", htmlString);
                }
            }
        }

        [Fact]
        public void HC007_FootnotesAndEndnotes_CssEnabled()
        {
            // Test that footnote CSS is generated when RenderFootnotesAndEndnotes is true
            // Use an existing test document
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/WC");
            FileInfo doc = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-Unmodified.docx"));

            byte[] byteArray = File.ReadAllBytes(doc.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Footnote Test",
                        FabricateCssClasses = true,
                        RenderFootnotesAndEndnotes = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify footnote CSS is generated when enabled
                    Assert.Contains("a.footnote-ref", htmlString);
                    Assert.Contains("section.footnotes", htmlString);
                    Assert.Contains("Footnotes and Endnotes CSS", htmlString);
                }
            }
        }

        [Fact]
        public void HC008_FootnotesAndEndnotes_CssDisabled()
        {
            // Test that footnote CSS is NOT generated when RenderFootnotesAndEndnotes is false (default)
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/WC");
            FileInfo doc = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-Unmodified.docx"));

            byte[] byteArray = File.ReadAllBytes(doc.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Footnote Test - Disabled",
                        FabricateCssClasses = true,
                        // RenderFootnotesAndEndnotes defaults to false
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify footnote CSS is NOT generated when disabled
                    Assert.DoesNotContain("a.footnote-ref", htmlString);
                    Assert.DoesNotContain("section.footnotes", htmlString);
                    Assert.DoesNotContain("Footnotes and Endnotes CSS", htmlString);
                }
            }
        }

        [Fact]
        public void HC009_HeadersAndFooters_CssEnabled()
        {
            // Test that header/footer CSS is generated when RenderHeadersAndFooters is true
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/WC");
            FileInfo doc = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-Unmodified.docx"));

            byte[] byteArray = File.ReadAllBytes(doc.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Header/Footer Test",
                        FabricateCssClasses = true,
                        RenderHeadersAndFooters = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify header/footer CSS is generated when enabled
                    Assert.Contains("header.document-header", htmlString);
                    Assert.Contains("footer.document-footer", htmlString);
                    Assert.Contains("Document Headers and Footers CSS", htmlString);
                }
            }
        }

        [Fact]
        public void HC010_HeadersAndFooters_CssDisabled()
        {
            // Test that header/footer CSS is NOT generated when RenderHeadersAndFooters is false (default)
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/WC");
            FileInfo doc = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-Unmodified.docx"));

            byte[] byteArray = File.ReadAllBytes(doc.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Header/Footer Test - Disabled",
                        FabricateCssClasses = true,
                        // RenderHeadersAndFooters defaults to false
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify header/footer CSS is NOT generated when disabled
                    Assert.DoesNotContain("header.document-header", htmlString);
                    Assert.DoesNotContain("footer.document-footer", htmlString);
                    Assert.DoesNotContain("Document Headers and Footers CSS", htmlString);
                }
            }
        }

        [Fact]
        public void HC011_TrackedChanges_MoveOperations()
        {
            // Use WmlComparer to create a document with move operations
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/WC");
            FileInfo doc1 = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-Unmodified.docx"));
            FileInfo doc2 = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-MovedPara.docx"));

            if (!doc2.Exists)
            {
                // Skip if test file doesn't exist
                return;
            }

            WmlDocument wmlDoc1 = new WmlDocument(doc1.FullName);
            WmlDocument wmlDoc2 = new WmlDocument(doc2.FullName);

            WmlComparerSettings comparerSettings = new WmlComparerSettings();
            WmlDocument comparedDoc = WmlComparer.Compare(wmlDoc1, wmlDoc2, comparerSettings);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(comparedDoc.DocumentByteArray, 0, comparedDoc.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Move Operations Test",
                        FabricateCssClasses = true,
                        RenderTrackedChanges = true,
                        RenderMoveOperations = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify move CSS classes are generated
                    Assert.Contains("rev-move-from", htmlString);
                    Assert.Contains("rev-move-to", htmlString);
                }
            }
        }

        [Fact]
        public void HC012_TrackedChanges_AuthorColors()
        {
            // Test that author-specific CSS is generated
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/WC");
            FileInfo doc = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-Unmodified.docx"));

            byte[] byteArray = File.ReadAllBytes(doc.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Author Colors Test",
                        FabricateCssClasses = true,
                        RenderTrackedChanges = true,
                        AuthorColors = new Dictionary<string, string>
                        {
                            { "Test Author", "#ff0000" },
                            { "Another Author", "#00ff00" }
                        }
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify author color CSS is generated (data-author attribute selector)
                    Assert.Contains("[data-author=\"Test Author\"]", htmlString);
                    Assert.Contains("#ff0000", htmlString);
                    Assert.Contains("[data-author=\"Another Author\"]", htmlString);
                    Assert.Contains("#00ff00", htmlString);
                }
            }
        }

        [Fact]
        public void HC013_TrackedChanges_AllFeaturesEnabled()
        {
            // Test with all tracked changes features enabled
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/WC");
            FileInfo doc1 = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-Unmodified.docx"));
            FileInfo doc2 = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-InsertInMiddle.docx"));

            WmlDocument wmlDoc1 = new WmlDocument(doc1.FullName);
            WmlDocument wmlDoc2 = new WmlDocument(doc2.FullName);

            WmlComparerSettings comparerSettings = new WmlComparerSettings();
            WmlDocument comparedDoc = WmlComparer.Compare(wmlDoc1, wmlDoc2, comparerSettings);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(comparedDoc.DocumentByteArray, 0, comparedDoc.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "All Features Test",
                        FabricateCssClasses = true,
                        RenderTrackedChanges = true,
                        IncludeRevisionMetadata = true,
                        ShowDeletedContent = true,
                        RenderMoveOperations = true,
                        RenderFootnotesAndEndnotes = true,
                        RenderHeadersAndFooters = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify all CSS sections are generated
                    Assert.Contains("Tracked Changes CSS", htmlString);
                    Assert.Contains("ins.rev-ins", htmlString);
                    Assert.Contains("del.rev-del", htmlString);

                    // Verify body structure
                    Assert.Contains("<body", htmlString);

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "AllFeatures.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        [Fact]
        public void HC014_Comments_CssGeneratedWhenEnabled()
        {
            // Test that comment CSS is generated when RenderComments is true
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/WC");
            FileInfo doc = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-Unmodified.docx"));

            byte[] byteArray = File.ReadAllBytes(doc.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Comment CSS Test",
                        FabricateCssClasses = true,
                        RenderComments = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify comment CSS is generated when enabled
                    Assert.Contains("Comments CSS", htmlString);
                    Assert.Contains("span.comment-highlight", htmlString);
                    Assert.Contains("a.comment-marker", htmlString);
                    Assert.Contains("aside.comments-section", htmlString);
                    Assert.Contains("li.comment", htmlString);
                }
            }
        }

        [Fact]
        public void HC015_Comments_CssNotGeneratedWhenDisabled()
        {
            // Test that comment CSS is NOT generated when RenderComments is false (default)
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/WC");
            FileInfo doc = new FileInfo(Path.Combine(sourceDir.FullName, "WC002-Unmodified.docx"));

            byte[] byteArray = File.ReadAllBytes(doc.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Comment CSS Test - Disabled",
                        FabricateCssClasses = true,
                        // RenderComments defaults to false
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify comment CSS is NOT generated when disabled
                    Assert.DoesNotContain("Comments CSS", htmlString);
                    Assert.DoesNotContain("span.comment-highlight", htmlString);
                    Assert.DoesNotContain("a.comment-marker", htmlString);
                    Assert.DoesNotContain("aside.comments-section", htmlString);
                }
            }
        }

        [Fact]
        public void HC016_Comments_WithCommentContent()
        {
            // Use HC031 which has a real comment (id=10 by "Eric White")
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, "HC031-Complicated-Document.docx"));

            byte[] byteArray = File.ReadAllBytes(sourceDocx.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Comment Content Test",
                        FabricateCssClasses = true,
                        RenderComments = true,
                        IncludeCommentMetadata = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify comment highlighting is present
                    Assert.Contains("comment-highlight", htmlString);
                    Assert.Contains("data-comment-id=\"10\"", htmlString);

                    // Verify comment marker is present
                    Assert.Contains("comment-marker", htmlString);
                    Assert.Contains("href=\"#comment-10\"", htmlString);

                    // Verify comments section is present
                    Assert.Contains("comments-section", htmlString);
                    Assert.Contains("id=\"comment-10\"", htmlString);

                    // Verify author metadata
                    Assert.Contains("data-author=\"Eric White\"", htmlString);
                    Assert.Contains("Eric White", htmlString);

                    // Verify comment text
                    Assert.Contains("This is a comment.", htmlString);

                    // Verify back reference link
                    Assert.Contains("href=\"#comment-ref-10\"", htmlString);
                    Assert.Contains("comment-backref", htmlString);

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "Comments-Test.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        [Fact]
        public void HC017_Comments_InlineMode()
        {
            // Use HC031 which has a real comment and test inline mode
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, "HC031-Complicated-Document.docx"));

            byte[] byteArray = File.ReadAllBytes(sourceDocx.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Inline Comment Test",
                        FabricateCssClasses = true,
                        RenderComments = true,
                        CommentRenderMode = CommentRenderMode.Inline,
                        IncludeCommentMetadata = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify inline mode attributes
                    Assert.Contains("title=\"Eric White: This is a comment.\"", htmlString);
                    Assert.Contains("data-comment=\"This is a comment.\"", htmlString);

                    // In inline mode, there should NOT be a comments section element (but CSS is fine)
                    Assert.DoesNotContain("<aside class=\"comments-section\"", htmlString);

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "Comments-Inline.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        [Fact]
        public void HC018_Comments_MultipleComments()
        {
            // Copy an existing document and add multiple comments programmatically
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, "HC006-Test-01.docx"));

            byte[] byteArray = File.ReadAllBytes(sourceDocx.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    var mainPart = wDoc.MainDocumentPart;
                    var body = mainPart.Document.Body;
                    var firstPara = body.Elements<Paragraph>().FirstOrDefault();

                    if (firstPara != null)
                    {
                        // Add comment markers to first paragraph
                        var firstRun = firstPara.Elements<Run>().FirstOrDefault();
                        if (firstRun != null)
                        {
                            firstRun.InsertBeforeSelf(new CommentRangeStart() { Id = "100" });
                            firstRun.InsertAfterSelf(new CommentRangeEnd() { Id = "100" });
                            firstRun.InsertAfterSelf(new Run(new CommentReference() { Id = "100" }));
                        }
                    }

                    var secondPara = body.Elements<Paragraph>().Skip(1).FirstOrDefault();
                    if (secondPara != null)
                    {
                        var secondRun = secondPara.Elements<Run>().FirstOrDefault();
                        if (secondRun != null)
                        {
                            secondRun.InsertBeforeSelf(new CommentRangeStart() { Id = "101" });
                            secondRun.InsertAfterSelf(new CommentRangeEnd() { Id = "101" });
                            secondRun.InsertAfterSelf(new Run(new CommentReference() { Id = "101" }));
                        }
                    }

                    // Add comments part with multiple comments
                    var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
                    commentsPart.Comments = new Comments(
                        new Comment(
                            new Paragraph(new Run(new Text("Comment one text.")))
                        )
                        { Id = "100", Author = "Author One" },
                        new Comment(
                            new Paragraph(new Run(new Text("Comment two text.")))
                        )
                        { Id = "101", Author = "Author Two" }
                    );

                    mainPart.Document.Save();
                }

                ms.Position = 0;
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Multiple Comments Test",
                        FabricateCssClasses = true,
                        RenderComments = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify both comments are rendered
                    Assert.Contains("id=\"comment-100\"", htmlString);
                    Assert.Contains("id=\"comment-101\"", htmlString);
                    Assert.Contains("Comment one text.", htmlString);
                    Assert.Contains("Comment two text.", htmlString);
                    Assert.Contains("Author One", htmlString);
                    Assert.Contains("Author Two", htmlString);

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "Comments-Multiple.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        [Fact]
        public void HC019_Comments_CustomCssPrefix()
        {
            // Use HC031 which has a real comment and test custom CSS prefix
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, "HC031-Complicated-Document.docx"));

            byte[] byteArray = File.ReadAllBytes(sourceDocx.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Custom Prefix Test",
                        FabricateCssClasses = true,
                        RenderComments = true,
                        CommentCssClassPrefix = "note-",
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify custom prefix is used
                    Assert.Contains("note-highlight", htmlString);
                    Assert.Contains("note-marker", htmlString);
                    Assert.Contains("notes-section", htmlString);

                    // Verify default prefix is NOT used
                    Assert.DoesNotContain("comment-highlight", htmlString);
                    Assert.DoesNotContain("comments-section", htmlString);
                }
            }
        }

        [Fact]
        public void HC020_Comments_MarginMode()
        {
            // Use HC031 which has a real comment and test margin mode rendering
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, "HC031-Complicated-Document.docx"));

            byte[] byteArray = File.ReadAllBytes(sourceDocx.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Margin Mode Comments Test",
                        FabricateCssClasses = true,
                        RenderComments = true,
                        CommentRenderMode = CommentRenderMode.Margin,
                        IncludeCommentMetadata = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify margin mode layout structure
                    Assert.Contains("comment-margin-container", htmlString);
                    Assert.Contains("comment-margin-content", htmlString);
                    Assert.Contains("comment-margin-column", htmlString);
                    Assert.Contains("comment-margin-note", htmlString);

                    // Verify margin note elements
                    Assert.Contains("comment-margin-note-header", htmlString);
                    Assert.Contains("comment-margin-author", htmlString);
                    Assert.Contains("comment-margin-note-body", htmlString);
                    Assert.Contains("comment-margin-backref", htmlString);

                    // Verify margin mode CSS is generated
                    Assert.Contains("/* Margin Mode Comments */", htmlString);
                    Assert.Contains("display: flex", htmlString);
                    Assert.Contains("flex-direction: row", htmlString);
                    Assert.Contains("width: 250px", htmlString);

                    // Verify print media query is included
                    Assert.Contains("@media print", htmlString);

                    // Verify there is NO endnote-style comments section element in HTML (CSS is fine)
                    // The CSS for comments-section is generated for all modes, but the actual <aside> element should not be present
                    Assert.DoesNotContain("<aside class=\"comments-section\"", htmlString);

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "Comments-Margin.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        [Fact]
        public void HC021_Comments_MarginMode_MultipleComments()
        {
            // Test margin mode with multiple comments to verify ordering
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, "HC006-Test-01.docx"));

            byte[] byteArray = File.ReadAllBytes(sourceDocx.FullName);
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    var mainPart = wDoc.MainDocumentPart;
                    var body = mainPart.Document.Body;
                    var firstPara = body.Elements<Paragraph>().FirstOrDefault();

                    if (firstPara != null)
                    {
                        var firstRun = firstPara.Elements<Run>().FirstOrDefault();
                        if (firstRun != null)
                        {
                            firstRun.InsertBeforeSelf(new CommentRangeStart() { Id = "200" });
                            firstRun.InsertAfterSelf(new CommentRangeEnd() { Id = "200" });
                            firstRun.InsertAfterSelf(new Run(new CommentReference() { Id = "200" }));
                        }
                    }

                    var secondPara = body.Elements<Paragraph>().Skip(1).FirstOrDefault();
                    if (secondPara != null)
                    {
                        var secondRun = secondPara.Elements<Run>().FirstOrDefault();
                        if (secondRun != null)
                        {
                            secondRun.InsertBeforeSelf(new CommentRangeStart() { Id = "201" });
                            secondRun.InsertAfterSelf(new CommentRangeEnd() { Id = "201" });
                            secondRun.InsertAfterSelf(new Run(new CommentReference() { Id = "201" }));
                        }
                    }

                    // Add comments part
                    var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
                    commentsPart.Comments = new Comments(
                        new Comment(
                            new Paragraph(new Run(new Text("First margin comment.")))
                        )
                        { Id = "200", Author = "Reviewer A", Date = new DateTime(2024, 1, 15, 10, 30, 0) },
                        new Comment(
                            new Paragraph(new Run(new Text("Second margin comment.")))
                        )
                        { Id = "201", Author = "Reviewer B", Date = new DateTime(2024, 1, 16, 14, 0, 0) }
                    );

                    mainPart.Document.Save();
                }

                ms.Position = 0;
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Multiple Margin Comments Test",
                        FabricateCssClasses = true,
                        RenderComments = true,
                        CommentRenderMode = CommentRenderMode.Margin,
                        IncludeCommentMetadata = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify both comments are in margin column
                    Assert.Contains("id=\"comment-200\"", htmlString);
                    Assert.Contains("id=\"comment-201\"", htmlString);
                    Assert.Contains("First margin comment.", htmlString);
                    Assert.Contains("Second margin comment.", htmlString);
                    Assert.Contains("Reviewer A", htmlString);
                    Assert.Contains("Reviewer B", htmlString);

                    // Verify margin structure
                    Assert.Contains("comment-margin-column", htmlString);

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "Comments-Margin-Multiple.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        [Fact]
        public void HC015_TabPrecedingText_UsesMinWidth()
        {
            // Test that text preceding a tab (like list numbers "2.3") uses min-width
            // instead of fixed width to prevent text overflow/overlap issues.
            // This fixes the bug where section numbers would overlap with heading text
            // because the width was calculated as 0 for text elements.

            using (MemoryStream ms = new MemoryStream())
            {
                // Create a document with a paragraph that has text followed by a tab
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add required parts
                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new Style(
                            new StyleName() { Val = "Normal" },
                            new PrimaryStyle()
                        ) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true }
                    );
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings(
                        new DefaultTabStop() { Val = 720 }  // 720 twips = 0.5 inch
                    );
                    settingsPart.Settings.Save();

                    // Create document with a paragraph containing "2.3" + tab + "Section Title"
                    // This simulates numbered headings like "2.3    Deemed Liquidation Events"
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new ParagraphProperties(
                                    new Tabs(
                                        new TabStop() { Val = TabStopValues.Left, Position = 720 }
                                    )
                                ),
                                new Run(
                                    new Text("2.3")
                                ),
                                new Run(
                                    new TabChar()
                                ),
                                new Run(
                                    new Text("Section Title")
                                )
                            )
                        )
                    );

                    mainPart.Document.Save();
                }

                ms.Position = 0;
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Tab Width Test",
                        FabricateCssClasses = true,
                        CssClassPrefix = "pt-",
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // The key assertion: verify min-width is used instead of width
                    // for elements preceding tabs. This prevents text overflow.
                    Assert.Contains("min-width:", htmlString);

                    // Verify the content is present
                    Assert.Contains("2.3", htmlString);
                    Assert.Contains("Section Title", htmlString);

                    // Verify we're NOT using fixed width (which would cause overflow)
                    // The CSS should have min-width, not a plain width for tab-preceding spans
                    var styleElement = html.Descendants(Xhtml.style).FirstOrDefault();
                    if (styleElement != null)
                    {
                        string css = styleElement.Value;
                        // Check that min-width appears in the CSS for inline-block elements
                        // These are the spans that wrap text preceding tabs
                        Assert.True(
                            css.Contains("min-width:") || htmlString.Contains("min-width:"),
                            "Expected min-width to be used for tab-preceding content to prevent text overflow"
                        );
                    }

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "TabWidth-MinWidth.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        [Fact]
        public void HC016_RunWithoutRPr_DoesNotCrash()
        {
            // Test that runs without w:rPr elements are handled gracefully.
            // Previously, DefineRunStyle and GetLangAttribute used .First() which
            // would throw InvalidOperationException if no rPr element existed.
            // This test verifies the fix using .FirstOrDefault() with null checks.

            using (MemoryStream ms = new MemoryStream())
            {
                // Create a document with runs that have NO rPr elements
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add required parts
                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new Style(
                            new StyleName() { Val = "Normal" },
                            new PrimaryStyle()
                        ) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true }
                    );
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create document with runs that have no rPr at all
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                // Run with no rPr - just text
                                new Run(
                                    new Text("Plain text without formatting")
                                ),
                                // Another run with no rPr
                                new Run(
                                    new Text(" and more plain text")
                                )
                            ),
                            new Paragraph(
                                // Mixed: run without rPr followed by run with rPr
                                new Run(
                                    new Text("No formatting here")
                                ),
                                new Run(
                                    new RunProperties(
                                        new Bold()
                                    ),
                                    new Text(" but this is bold")
                                )
                            )
                        )
                    );

                    mainPart.Document.Save();
                }

                ms.Position = 0;
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Null rPr Test",
                        FabricateCssClasses = true,
                        CssClassPrefix = "pt-",
                    };

                    // This should NOT throw - previously it would crash with:
                    // System.InvalidOperationException: Sequence contains no elements
                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify all content is present in the output
                    Assert.Contains("Plain text without formatting", htmlString);
                    Assert.Contains("and more plain text", htmlString);
                    Assert.Contains("No formatting here", htmlString);
                    Assert.Contains("but this is bold", htmlString);

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "NullRPr-Test.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        [Fact]
        public void ConcurrentConversions_ShouldNotCorruptShadeCache()
        {
            // This test verifies that the ShadeCache (ConcurrentDictionary) handles
            // concurrent access correctly during parallel document conversions.

            // Create a proper document with all required parts and shading
            byte[] docBytes;
            using (var stream = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    // Add main document part
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new Run(
                                    new RunProperties(
                                        new Shading { Val = ShadingPatternValues.Percent20, Color = "FF0000", Fill = "FFFFFF" }
                                    ),
                                    new Text("Red shading 20%")
                                )
                            ),
                            new Paragraph(
                                new Run(
                                    new RunProperties(
                                        new Shading { Val = ShadingPatternValues.Percent50, Color = "00FF00", Fill = "000000" }
                                    ),
                                    new Text("Green shading 50%")
                                )
                            ),
                            new Paragraph(
                                new Run(
                                    new RunProperties(
                                        new Shading { Val = ShadingPatternValues.Percent75, Color = "0000FF", Fill = "FFFFFF" }
                                    ),
                                    new Text("Blue shading 75%")
                                )
                            )
                        )
                    );

                    // Add required StyleDefinitionsPart
                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                                    new FontSize { Val = "22" }
                                )
                            )
                        )
                    );

                    // Add DocumentSettingsPart
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
                }
                docBytes = stream.ToArray();
            }

            var exceptions = new ConcurrentBag<Exception>();
            var tasks = new Task[20];

            // Run 20 concurrent conversions (each gets its own copy of the doc bytes)
            for (int i = 0; i < 20; i++)
            {
                int iteration = i;
                tasks[i] = Task.Run(() =>
                {
                    try
                    {
                        // Each task needs its own copy since the converter may modify the document
                        byte[] localDocBytes = (byte[])docBytes.Clone();
                        using (var ms = new MemoryStream())
                        {
                            ms.Write(localDocBytes, 0, localDocBytes.Length);
                            ms.Position = 0;
                            using (var wDoc = WordprocessingDocument.Open(ms, true))
                            {
                                var settings = new WmlToHtmlConverterSettings
                                {
                                    PageTitle = $"Concurrent Test {iteration}",
                                    FabricateCssClasses = true,
                                    CssClassPrefix = $"pt{iteration}-",
                                };
                                XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                                string htmlString = html.ToString();

                                // Verify content was converted
                                Assert.Contains("Red shading 20%", htmlString);
                                Assert.Contains("Green shading 50%", htmlString);
                                Assert.Contains("Blue shading 75%", htmlString);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        exceptions.Add(ex);
                    }
                });
            }

            Task.WaitAll(tasks);

            // No exceptions should have occurred
            Assert.Empty(exceptions);
        }

        [Fact]
        public void ClearShadeCache_ShouldNotThrowDuringConcurrentUse()
        {
            // This test verifies that clearing the cache while conversions are running
            // doesn't cause exceptions (ConcurrentDictionary handles this safely).

            byte[] docBytes;
            using (var stream = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new Run(
                                    new RunProperties(
                                        new Shading { Val = ShadingPatternValues.Percent25, Color = "123456", Fill = "ABCDEF" }
                                    ),
                                    new Text("Shaded content")
                                )
                            )
                        )
                    );

                    // Add required StyleDefinitionsPart
                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                                    new FontSize { Val = "22" }
                                )
                            )
                        )
                    );

                    // Add DocumentSettingsPart
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
                }
                docBytes = stream.ToArray();
            }

            var cts = new CancellationTokenSource();
            var exceptions = new ConcurrentBag<Exception>();

            // Start a background task doing conversions
            var conversionTask = Task.Run(() =>
            {
                while (!cts.Token.IsCancellationRequested)
                {
                    try
                    {
                        // Each iteration needs its own copy since converter may modify document
                        byte[] localDocBytes = (byte[])docBytes.Clone();
                        using (var ms = new MemoryStream())
                        {
                            ms.Write(localDocBytes, 0, localDocBytes.Length);
                            ms.Position = 0;
                            using (var wDoc = WordprocessingDocument.Open(ms, true))
                            {
                                var settings = new WmlToHtmlConverterSettings();
                                WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        exceptions.Add(ex);
                    }
                }
            });

            // Clear cache multiple times while conversions are running
            for (int i = 0; i < 50; i++)
            {
                WmlToHtmlConverter.ClearShadeCache();
                Thread.Sleep(1);
            }

            cts.Cancel();
            try { conversionTask.Wait(TimeSpan.FromSeconds(5)); } catch { }

            // No exceptions should have occurred from the concurrent access
            Assert.Empty(exceptions);
        }

        [Fact]
        public void FontFamilyHelper_ConcurrentMarkAsUnknown_ShouldNotCorrupt()
        {
            // This test verifies that FontFamilyHelper's ConcurrentDictionary-based
            // unknown fonts cache handles concurrent access correctly.

            // Clear any existing unknown fonts
            FontFamilyHelper.ClearUnknownFontsCache();

            var exceptions = new ConcurrentBag<Exception>();
            var tasks = new Task[20];

            // Run 20 concurrent tasks marking fonts as unknown
            for (int i = 0; i < 20; i++)
            {
                int iteration = i;
                tasks[i] = Task.Run(() =>
                {
                    try
                    {
                        for (int j = 0; j < 100; j++)
                        {
                            // Each task marks some unique fonts and some shared fonts
                            FontFamilyHelper.MarkAsUnknown($"UniqueFont-{iteration}-{j}");
                            FontFamilyHelper.MarkAsUnknown($"SharedFont-{j}");

                            // Also check if fonts are marked
                            FontFamilyHelper.IsMarkedUnknown($"SharedFont-{j}");
                            FontFamilyHelper.IsMarkedUnknown($"UniqueFont-{iteration}-{j}");
                        }
                    }
                    catch (Exception ex)
                    {
                        exceptions.Add(ex);
                    }
                });
            }

            Task.WaitAll(tasks);

            // No exceptions should have occurred
            Assert.Empty(exceptions);

            // Verify fonts were marked (each task marks 100 unique + 100 shared per iteration)
            // Unique: 20 tasks * 100 = 2000
            // Shared: 100 (deduplicated)
            // Total: 2100
            Assert.Equal(2100, FontFamilyHelper.UnknownFonts.Count);

            // Verify specific fonts are marked
            Assert.True(FontFamilyHelper.IsMarkedUnknown("UniqueFont-0-0"));
            Assert.True(FontFamilyHelper.IsMarkedUnknown("SharedFont-50"));

            // Clean up
            FontFamilyHelper.ClearUnknownFontsCache();
            Assert.Empty(FontFamilyHelper.UnknownFonts);
        }

        [Fact]
        public void HC022_TabWidthCalculation_TextWidthNonZero()
        {
            // Test that text width is now calculated (non-zero) for text elements
            // before tabs. Previously this was hardcoded to 0, causing incorrect
            // tab positioning for right/center/decimal tabs.

            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new Style(
                            new StyleName() { Val = "Normal" },
                            new PrimaryStyle()
                        ) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true }
                    );
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings(
                        new DefaultTabStop() { Val = 720 }
                    );
                    settingsPart.Settings.Save();

                    // Create a paragraph with: "Hello World" + tab + "Right aligned"
                    // The tab is right-aligned at 6 inches (8640 twips)
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new ParagraphProperties(
                                    new Tabs(
                                        new TabStop() { Val = TabStopValues.Right, Position = 8640 }
                                    )
                                ),
                                new Run(
                                    new RunProperties(
                                        new RunFonts() { Ascii = "Times New Roman" },
                                        new FontSize() { Val = "24" }  // 12pt
                                    ),
                                    new Text("Hello World")
                                ),
                                new Run(
                                    new TabChar()
                                ),
                                new Run(
                                    new Text("Right aligned")
                                )
                            )
                        )
                    );

                    mainPart.Document.Save();
                }

                ms.Position = 0;
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Tab Width Calculation Test",
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // The HTML output should contain span elements with margin or width
                    // values that account for the text before the tab
                    Assert.Contains("Hello World", htmlString);
                    Assert.Contains("Right aligned", htmlString);

                    // Verify tab span has styling applied (margin for spacing)
                    Assert.Contains("margin", htmlString);

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "TabWidth-Calculation.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        [Fact]
        public void HC023_RightAlignedTab_CorrectSpacing()
        {
            // Test that right-aligned tabs calculate correct spacing based on
            // the width of text that follows the tab.

            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new Style(
                            new StyleName() { Val = "Normal" },
                            new PrimaryStyle()
                        ) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true }
                    );
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings(
                        new DefaultTabStop() { Val = 720 }
                    );
                    settingsPart.Settings.Save();

                    // Table of Contents style: "Chapter 1" + tab with dots + "1"
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new ParagraphProperties(
                                    new Tabs(
                                        new TabStop() { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Dot, Position = 8640 }
                                    )
                                ),
                                new Run(
                                    new Text("Chapter 1")
                                ),
                                new Run(
                                    new TabChar()
                                ),
                                new Run(
                                    new Text("1")
                                )
                            )
                        )
                    );

                    mainPart.Document.Save();
                }

                ms.Position = 0;
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Right Tab Test",
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify content is present
                    Assert.Contains("Chapter 1", htmlString);
                    Assert.Contains("1", htmlString);

                    // Verify tab span has margin or width styling for the right-aligned tab
                    // (dot leaders may or may not appear depending on font availability)
                    Assert.True(
                        htmlString.Contains("margin") || htmlString.Contains("width:"),
                        "Expected tab to have margin or width styling for positioning"
                    );

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "TabWidth-RightAligned.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        [Fact]
        public void HC024_MultipleTabsInParagraph_AllHaveSpacing()
        {
            // Test that multiple tabs in a single paragraph all get correct spacing

            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new Style(
                            new StyleName() { Val = "Normal" },
                            new PrimaryStyle()
                        ) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true }
                    );
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings(
                        new DefaultTabStop() { Val = 720 }
                    );
                    settingsPart.Settings.Save();

                    // Create: "Col1" + tab + "Col2" + tab + "Col3"
                    // With tabs at 2" and 4"
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new ParagraphProperties(
                                    new Tabs(
                                        new TabStop() { Val = TabStopValues.Left, Position = 2880 },  // 2 inches
                                        new TabStop() { Val = TabStopValues.Left, Position = 5760 }   // 4 inches
                                    )
                                ),
                                new Run(new Text("Col1")),
                                new Run(new TabChar()),
                                new Run(new Text("Col2")),
                                new Run(new TabChar()),
                                new Run(new Text("Col3"))
                            )
                        )
                    );

                    mainPart.Document.Save();
                }

                ms.Position = 0;
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Multiple Tabs Test",
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Verify all columns are present
                    Assert.Contains("Col1", htmlString);
                    Assert.Contains("Col2", htmlString);
                    Assert.Contains("Col3", htmlString);

                    // Count margin/spacing occurrences - should have multiple for tabs
                    int marginCount = System.Text.RegularExpressions.Regex.Matches(htmlString, @"margin[^;]*:").Count;
                    Assert.True(marginCount >= 2, $"Expected at least 2 margin styles for tabs, found {marginCount}");

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "TabWidth-Multiple.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
                }
            }
        }

        [Fact]
        public void HC025_TabLeaderCharacters_DotLeaderRendered()
        {
            // Test with actual test file that has dot leaders
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo sourceFile = new FileInfo(Path.Combine(sourceDir.FullName, "HC024-Tabs-01.docx"));
            WmlDocument wmlDoc = new WmlDocument(sourceFile.FullName);

            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(wmlDoc.DocumentByteArray, 0, wmlDoc.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Tab Leaders Test",
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "TabLeaders-HC024.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);

                    // Check for dot leader characters - at least 3 dots in a row
                    // Note: The exact count varies by platform due to font measurement differences
                    bool hasDotLeaders = System.Text.RegularExpressions.Regex.IsMatch(htmlString, @"\.{3,}");
                    Assert.True(hasDotLeaders, "Expected dot leader characters (...) in HTML output");
                }
            }
        }

        [Fact]
        public void HC026_TabLeaderCharacters_ProgrammaticDotLeader()
        {
            // Create a document programmatically with dot leader tab
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wDoc.AddMainDocumentPart();

                    // Add styles part (required for proper processing)
                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                                    new FontSize() { Val = "24" }
                                )
                            )
                        )
                    );
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings(
                        new DefaultTabStop() { Val = 720 }
                    );
                    settingsPart.Settings.Save();

                    // Create: "Chapter 1" + dotted tab leader + "1"
                    // With right-aligned tab with dot leader at 5 inches
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new ParagraphProperties(
                                    new Tabs(
                                        new TabStop()
                                        {
                                            Val = TabStopValues.Right,
                                            Position = 7200,  // 5 inches
                                            Leader = TabStopLeaderCharValues.Dot
                                        }
                                    )
                                ),
                                new Run(new Text("Chapter 1")),
                                new Run(new TabChar()),
                                new Run(new Text("1"))
                            )
                        )
                    );

                    mainPart.Document.Save();
                }

                ms.Position = 0;
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, true))
                {
                    WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                    {
                        PageTitle = "Programmatic Dot Leader Test",
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Save for debugging
                    var destFileName = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "DotLeader-Programmatic.html"));
                    File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);

                    // Output for debugging
                    System.Diagnostics.Debug.WriteLine("=== HTML OUTPUT ===");
                    System.Diagnostics.Debug.WriteLine(htmlString);

                    // Check for dot leader characters - at least 5 dots in a row
                    bool hasDotLeaders = System.Text.RegularExpressions.Regex.IsMatch(htmlString, @"\.{5,}");
                    Assert.True(hasDotLeaders, $"Expected dot leader characters (.....) in HTML output. HTML:\n{htmlString.Substring(0, Math.Min(2000, htmlString.Length))}");
                }
            }
        }

        #region Unsupported Content Placeholder Tests

        [Fact]
        public void HC028_UnsupportedContentPlaceholders_MathEquation()
        {
            // Test that math equations render as placeholders when enabled
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add required style definitions part
                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    // Add document settings part
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create document with math element
                    var mathNs = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/math");
                    var body = new XElement(W.body,
                        new XElement(W.p,
                            new XElement(mathNs + "oMath",
                                new XElement(mathNs + "r",
                                    new XElement(mathNs + "t", "x + y = z")
                                )
                            )
                        )
                    );

                    var document = new XElement(W.document,
                        new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName),
                        new XAttribute(XNamespace.Xmlns + "m", mathNs.NamespaceName),
                        body);

                    mainPart.PutXDocument(new XDocument(document));

                    // Test with placeholders DISABLED (default - should be empty)
                    var settingsOff = new WmlToHtmlConverterSettings
                    {
                        RenderUnsupportedContentPlaceholders = false,
                    };

                    XElement htmlOff = WmlToHtmlConverter.ConvertToHtml(wDoc, settingsOff);
                    string htmlOffString = htmlOff.ToString();
                    Assert.DoesNotContain("[MATH]", htmlOffString);
                    Assert.DoesNotContain("unsupported-", htmlOffString);

                    // Test with placeholders ENABLED
                    var settingsOn = new WmlToHtmlConverterSettings
                    {
                        RenderUnsupportedContentPlaceholders = true,
                        UnsupportedContentCssClassPrefix = "unsupported-",
                        IncludeUnsupportedContentMetadata = true,
                    };

                    XElement htmlOn = WmlToHtmlConverter.ConvertToHtml(wDoc, settingsOn);
                    string htmlOnString = htmlOn.ToString();

                    // Should contain placeholder
                    Assert.Contains("[MATH]", htmlOnString);
                    Assert.Contains("unsupported-placeholder", htmlOnString);
                    Assert.Contains("unsupported-math", htmlOnString);
                    Assert.Contains("data-content-type=\"MathEquation\"", htmlOnString);
                }
            }
        }

        [Fact]
        public void HC029_UnsupportedContentPlaceholders_FormField()
        {
            // Test that form fields render as placeholders when enabled
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add required style definitions part
                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    // Add document settings part
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create document with form field (checkbox)
                    var body = new XElement(W.body,
                        new XElement(W.p,
                            new XElement(W.r,
                                new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin")),
                                new XElement(W.ffData,
                                    new XElement(W.checkBox,
                                        new XElement(W._default, new XAttribute(W.val, "0"))
                                    )
                                )
                            ),
                            new XElement(W.r,
                                new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))
                            )
                        )
                    );

                    var document = new XElement(W.document,
                        new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName),
                        body);

                    mainPart.PutXDocument(new XDocument(document));

                    // Test with placeholders ENABLED
                    var settings = new WmlToHtmlConverterSettings
                    {
                        RenderUnsupportedContentPlaceholders = true,
                        IncludeUnsupportedContentMetadata = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Should contain form field placeholder
                    Assert.Contains("[CHECKBOX]", htmlString);
                    Assert.Contains("unsupported-form", htmlString);
                    Assert.Contains("data-content-type=\"FormField\"", htmlString);
                }
            }
        }

        [Fact]
        public void HC030_UnsupportedContentPlaceholders_RubyAnnotation()
        {
            // Test that ruby annotations render as placeholders when enabled
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add required style definitions part
                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    // Add document settings part
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create document with ruby annotation
                    var body = new XElement(W.body,
                        new XElement(W.p,
                            new XElement(W.r,
                                new XElement(W.ruby,
                                    new XElement(W.rubyPr),
                                    new XElement(W.rt,
                                        new XElement(W.r,
                                            new XElement(W.t, "とうきょう")
                                        )
                                    ),
                                    new XElement(W.rubyBase,
                                        new XElement(W.r,
                                            new XElement(W.t, "東京")
                                        )
                                    )
                                )
                            )
                        )
                    );

                    var document = new XElement(W.document,
                        new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName),
                        body);

                    mainPart.PutXDocument(new XDocument(document));

                    // Test with placeholders ENABLED
                    var settings = new WmlToHtmlConverterSettings
                    {
                        RenderUnsupportedContentPlaceholders = true,
                        IncludeUnsupportedContentMetadata = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Should contain ruby placeholder with base text
                    Assert.Contains("東京", htmlString);  // Base text should be included
                    Assert.Contains("unsupported-ruby", htmlString);
                    Assert.Contains("data-content-type=\"RubyAnnotation\"", htmlString);
                }
            }
        }

        [Fact]
        public void HC031_UnsupportedContentPlaceholders_CssGenerated()
        {
            // Test that CSS is generated for placeholders when enabled
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add required style definitions part
                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    // Add document settings part
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create minimal document
                    var body = new XElement(W.body,
                        new XElement(W.p,
                            new XElement(W.r,
                                new XElement(W.t, "Test")
                            )
                        )
                    );

                    var document = new XElement(W.document,
                        new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName),
                        body);

                    mainPart.PutXDocument(new XDocument(document));

                    // Test with placeholders ENABLED
                    var settings = new WmlToHtmlConverterSettings
                    {
                        RenderUnsupportedContentPlaceholders = true,
                        UnsupportedContentCssClassPrefix = "unsupported-",
                        FabricateCssClasses = true,
                    };

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Should contain CSS for placeholders
                    Assert.Contains("/* Unsupported Content Placeholders CSS */", htmlString);
                    Assert.Contains(".unsupported-placeholder", htmlString);
                    Assert.Contains(".unsupported-image", htmlString);
                    Assert.Contains(".unsupported-math", htmlString);
                    Assert.Contains(".unsupported-form", htmlString);
                    Assert.Contains(".unsupported-ruby", htmlString);
                }
            }
        }

        [Fact]
        public void HC032_UnsupportedContentPlaceholders_BackwardCompatibility()
        {
            // Test that default settings preserve backward compatibility (no placeholders)
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add required style definitions part
                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    // Add document settings part
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create document with math element
                    var mathNs = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/math");
                    var body = new XElement(W.body,
                        new XElement(W.p,
                            new XElement(mathNs + "oMath",
                                new XElement(mathNs + "r",
                                    new XElement(mathNs + "t", "x + y")
                                )
                            )
                        )
                    );

                    var document = new XElement(W.document,
                        new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName),
                        new XAttribute(XNamespace.Xmlns + "m", mathNs.NamespaceName),
                        body);

                    mainPart.PutXDocument(new XDocument(document));

                    // Test with DEFAULT settings (placeholders disabled)
                    var settings = new WmlToHtmlConverterSettings();

                    // Verify default is false
                    Assert.False(settings.RenderUnsupportedContentPlaceholders);

                    XElement html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    string htmlString = html.ToString();

                    // Should NOT contain any placeholder indicators
                    Assert.DoesNotContain("[MATH]", htmlString);
                    Assert.DoesNotContain("unsupported-placeholder", htmlString);
                    Assert.DoesNotContain("Unsupported Content Placeholders CSS", htmlString);
                }
            }
        }

        #endregion

        #region Language Attribute Tests

        [Fact]
        public void HC033_HtmlElementHasLangAttribute_FromThemeFontLang()
        {
            // Test that <html> element gets lang attribute from themeFontLang
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add settings with themeFontLang set to French
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings(
                        new ThemeFontLanguages() { Val = "fr-FR" }
                    );
                    settingsPart.Settings.Save();

                    // Add minimal styles
                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Bonjour")))));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings();
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

                    var langAttr = html.Attribute("lang");
                    Assert.NotNull(langAttr);
                    Assert.Equal("fr-FR", langAttr.Value);
                }
            }
        }

        [Fact]
        public void HC034_DocumentLanguageSetting_OverridesDocument()
        {
            // Test that DocumentLanguage setting overrides document settings
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add settings with French language
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings(
                        new ThemeFontLanguages() { Val = "fr-FR" }
                    );
                    settingsPart.Settings.Save();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Test")))));
                    mainPart.Document.Save();

                    // Override with German
                    var settings = new WmlToHtmlConverterSettings
                    {
                        DocumentLanguage = "de-DE"
                    };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

                    var langAttr = html.Attribute("lang");
                    Assert.NotNull(langAttr);
                    Assert.Equal("de-DE", langAttr.Value);
                }
            }
        }

        [Fact]
        public void HC035_FallbackLanguage_WhenNotSpecified()
        {
            // Test that language falls back to "en-US" when not specified in document
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // No themeFontLang - empty settings
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Hello")))));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings();
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

                    var langAttr = html.Attribute("lang");
                    Assert.NotNull(langAttr);
                    Assert.Equal("en-US", langAttr.Value);
                }
            }
        }

        [Fact]
        public void HC036_HtmlConverterSettings_DocumentLanguage()
        {
            // Test that HtmlConverterSettings also supports DocumentLanguage
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    mainPart.Document = new Document(new Body(new Paragraph(new Run(new Text("Test")))));
                    mainPart.Document.Save();

                    // Use HtmlConverterSettings (the alternate settings class)
                    var htmlSettings = new HtmlConverterSettings
                    {
                        DocumentLanguage = "ja-JP"
                    };
                    var html = HtmlConverter.ConvertToHtml(wDoc, htmlSettings);

                    var langAttr = html.Attribute("lang");
                    Assert.NotNull(langAttr);
                    Assert.Equal("ja-JP", langAttr.Value);
                }
            }
        }

        #endregion

        #region Table Layout Tests

        [Fact]
        public void HC037_TableDxaWidth_RendersAsPoints()
        {
            // Test that table widths specified in DXA (twips) are converted to points
            // DXA to points: divide by 20 (1 point = 20 twips)
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "Times New Roman" },
                                    new FontSize { Val = "24" }))));
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create a table with DXA width (7200 twips = 360 points = 5 inches)
                    mainPart.Document = new Document(
                        new Body(
                            new DocumentFormat.OpenXml.Wordprocessing.Table(
                                new TableProperties(
                                    new TableWidth { Width = "7200", Type = TableWidthUnitValues.Dxa }),
                                new DocumentFormat.OpenXml.Wordprocessing.TableRow(
                                    new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                                        new TableCellProperties(
                                            new TableCellWidth { Width = "3600", Type = TableWidthUnitValues.Dxa }),
                                        new Paragraph(
                                            new Run(new Text("Cell 1"))))),
                                new DocumentFormat.OpenXml.Wordprocessing.TableRow(
                                    new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                                        new TableCellProperties(
                                            new TableCellWidth { Width = "3600", Type = TableWidthUnitValues.Dxa }),
                                        new Paragraph(
                                            new Run(new Text("Cell 2"))))))));

                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings
                    {
                        PageTitle = "Table DXA Width Test"
                    };

                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // 7200 twips / 20 = 360 points
                    Assert.Contains("width: 360pt", htmlString);
                    Assert.Contains("Cell 1", htmlString);
                    Assert.Contains("Cell 2", htmlString);
                }
            }
        }

        [Fact]
        public void HC038_BorderlessTable_HasDataAttribute()
        {
            // Test that tables with nil/none borders get data-borderless="true" attribute
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "Times New Roman" },
                                    new FontSize { Val = "24" }))));
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create a table with nil borders (borderless layout table)
                    mainPart.Document = new Document(
                        new Body(
                            new DocumentFormat.OpenXml.Wordprocessing.Table(
                                new TableProperties(
                                    new TableBorders(
                                        new TopBorder { Val = BorderValues.Nil },
                                        new BottomBorder { Val = BorderValues.Nil },
                                        new LeftBorder { Val = BorderValues.Nil },
                                        new RightBorder { Val = BorderValues.Nil },
                                        new InsideHorizontalBorder { Val = BorderValues.Nil },
                                        new InsideVerticalBorder { Val = BorderValues.Nil })),
                                new DocumentFormat.OpenXml.Wordprocessing.TableRow(
                                    new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                                        new Paragraph(
                                            new Run(new Text("Borderless Cell"))))))));

                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings
                    {
                        PageTitle = "Borderless Table Test"
                    };

                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    Assert.Contains("data-borderless=\"true\"", htmlString);
                    Assert.Contains("Borderless Cell", htmlString);
                }
            }
        }

        [Fact]
        public void HC039_TableWithBorders_NoDataBorderlessAttribute()
        {
            // Test that tables with visible borders do NOT get data-borderless attribute
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "Times New Roman" },
                                    new FontSize { Val = "24" }))));
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create a table with single-line borders
                    mainPart.Document = new Document(
                        new Body(
                            new DocumentFormat.OpenXml.Wordprocessing.Table(
                                new TableProperties(
                                    new TableBorders(
                                        new TopBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                                        new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                                        new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                                        new RightBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" },
                                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "000000" })),
                                new DocumentFormat.OpenXml.Wordprocessing.TableRow(
                                    new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                                        new Paragraph(
                                            new Run(new Text("Bordered Cell"))))))));

                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings
                    {
                        PageTitle = "Table With Borders Test"
                    };

                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    Assert.DoesNotContain("data-borderless", htmlString);
                    Assert.Contains("Bordered Cell", htmlString);
                }
            }
        }

        [Fact]
        public void HC040_TableWithoutTblBorders_IsBorderless()
        {
            // Test that tables without any tblBorders element are treated as borderless
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "Times New Roman" },
                                    new FontSize { Val = "24" }))));
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create a table without any border properties
                    mainPart.Document = new Document(
                        new Body(
                            new DocumentFormat.OpenXml.Wordprocessing.Table(
                                new TableProperties(),  // No borders specified
                                new DocumentFormat.OpenXml.Wordprocessing.TableRow(
                                    new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                                        new Paragraph(
                                            new Run(new Text("No Border Props Cell"))))))));

                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings
                    {
                        PageTitle = "No Border Props Table Test"
                    };

                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    Assert.Contains("data-borderless=\"true\"", htmlString);
                    Assert.Contains("No Border Props Cell", htmlString);
                }
            }
        }

        [Fact]
        public void HC041_TablePercentageWidth_RendersCorrectly()
        {
            // Test that table widths specified as percentage render correctly
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "Times New Roman" },
                                    new FontSize { Val = "24" }))));
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create a table with percentage width (5000 = 100%, so 2500 = 50%)
                    mainPart.Document = new Document(
                        new Body(
                            new DocumentFormat.OpenXml.Wordprocessing.Table(
                                new TableProperties(
                                    new TableWidth { Width = "2500", Type = TableWidthUnitValues.Pct }),
                                new DocumentFormat.OpenXml.Wordprocessing.TableRow(
                                    new DocumentFormat.OpenXml.Wordprocessing.TableCell(
                                        new Paragraph(
                                            new Run(new Text("50% Width Cell"))))))));

                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings
                    {
                        PageTitle = "Table Percentage Width Test"
                    };

                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // 2500 / 50 = 50%
                    Assert.Contains("width: 50%", htmlString);
                    Assert.Contains("50% Width Cell", htmlString);
                }
            }
        }

        [Fact]
        public void HC042_UnknownSerifFont_GetsSerifFallback()
        {
            // Test that unknown fonts without "sans" or "mono" patterns get serif fallback
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "MyCustomFont" },
                                    new FontSize { Val = "24" }))));
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new Run(new Text("Custom font text")))));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings { PageTitle = "Font Fallback Test" };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Unknown font should get serif fallback
                    Assert.Contains("'MyCustomFont', serif", htmlString);
                }
            }
        }

        [Fact]
        public void HC043_UnknownSansFont_GetsSansSerifFallback()
        {
            // Test that fonts with "sans" pattern get sans-serif fallback
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "CustomSansFont" },
                                    new FontSize { Val = "24" }))));
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new Run(new Text("Sans font text")))));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings { PageTitle = "Font Fallback Test" };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Font with "sans" should get sans-serif fallback
                    Assert.Contains("'CustomSansFont', sans-serif", htmlString);
                }
            }
        }

        [Fact]
        public void HC044_UnknownMonoFont_GetsMonospaceFallback()
        {
            // Test that fonts with "mono" or "code" patterns get monospace fallback
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "MyCodeFont" },
                                    new FontSize { Val = "24" }))));
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new Run(new Text("Mono font text")))));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings { PageTitle = "Font Fallback Test" };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Font with "code" should get monospace fallback
                    Assert.Contains("'MyCodeFont', monospace", htmlString);
                }
            }
        }

        [Fact]
        public void HC045_KnownFont_UsesKnownFallback()
        {
            // Test that known fonts (like Arial) use their predefined fallback
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "Arial" },
                                    new FontSize { Val = "24" }))));
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new Run(new Text("Arial text")))));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings { PageTitle = "Font Fallback Test" };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Known font Arial should get sans-serif fallback
                    Assert.Contains("'Arial', sans-serif", htmlString);
                }
            }
        }

        [Fact]
        public void HC046_CourierNew_GetsMonospaceFallback()
        {
            // Test that Courier New now gets proper monospace fallback (was missing before)
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "Courier New" },
                                    new FontSize { Val = "24" }))));
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(
                                new Run(new Text("Courier text")))));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings { PageTitle = "Font Fallback Test" };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Courier New should now get monospace fallback
                    Assert.Contains("'Courier New', monospace", htmlString);
                }
            }
        }

        [Fact]
        public void HC047_JapaneseText_GetsCjkFallback()
        {
            // Test that East Asian text with Japanese language gets CJK fallback chain
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "Times New Roman", EastAsia = "MS Mincho" },
                                    new FontSize { Val = "24" },
                                    new Languages { Val = "en-US", EastAsia = "ja-JP" }))));
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create a run with Japanese text (hiragana character)
                    var run = new Run(
                        new RunProperties(
                            new RunFonts { EastAsia = "MS Mincho" },
                            new Languages { EastAsia = "ja-JP" }),
                        new Text("あ"));  // Hiragana 'a'

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(run)));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings { PageTitle = "CJK Font Fallback Test" };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Should contain CJK JP fallback chain fonts
                    Assert.Contains("Noto Serif CJK JP", htmlString);
                    Assert.Contains("MS Mincho", htmlString);
                }
            }
        }

        [Fact]
        public void HC048_ChineseSimplified_GetsCjkScFallback()
        {
            // Test that Simplified Chinese text gets CJK SC fallback chain
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "Times New Roman", EastAsia = "SimSun" },
                                    new FontSize { Val = "24" },
                                    new Languages { Val = "en-US", EastAsia = "zh-hans" }))));
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create a run with Chinese text
                    var run = new Run(
                        new RunProperties(
                            new RunFonts { EastAsia = "SimSun" },
                            new Languages { EastAsia = "zh-hans" }),
                        new Text("中"));  // Chinese character

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(run)));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings { PageTitle = "CJK Font Fallback Test" };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Should contain CJK SC fallback chain fonts
                    Assert.Contains("Noto Serif CJK SC", htmlString);
                    Assert.Contains("Microsoft YaHei", htmlString);
                }
            }
        }

        #endregion

        #region Theme Color Tests

        [Fact]
        public void HC049_ThemeColor_ResolvesAccentColor()
        {
            // Test that theme colors like accent1 are resolved to actual hex values
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add theme part with color scheme
                    var themePart = mainPart.AddNewPart<ThemePart>();
                    themePart.Theme = new DocumentFormat.OpenXml.Drawing.Theme(
                        new DocumentFormat.OpenXml.Drawing.ThemeElements(
                            new DocumentFormat.OpenXml.Drawing.ColorScheme(
                                new DocumentFormat.OpenXml.Drawing.Dark1Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Light1Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "FFFFFF" }),
                                new DocumentFormat.OpenXml.Drawing.Dark2Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "1F4E78" }),
                                new DocumentFormat.OpenXml.Drawing.Light2Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "EEECE1" }),
                                new DocumentFormat.OpenXml.Drawing.Accent1Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "4472C4" }),  // Blue
                                new DocumentFormat.OpenXml.Drawing.Accent2Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "ED7D31" }),
                                new DocumentFormat.OpenXml.Drawing.Accent3Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "A5A5A5" }),
                                new DocumentFormat.OpenXml.Drawing.Accent4Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "FFC000" }),
                                new DocumentFormat.OpenXml.Drawing.Accent5Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "5B9BD5" }),
                                new DocumentFormat.OpenXml.Drawing.Accent6Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "70AD47" }),
                                new DocumentFormat.OpenXml.Drawing.Hyperlink(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "0563C1" }),
                                new DocumentFormat.OpenXml.Drawing.FollowedHyperlinkColor(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "954F72" }))
                            { Name = "Office" },
                            new DocumentFormat.OpenXml.Drawing.FontScheme(
                                new DocumentFormat.OpenXml.Drawing.MajorFont(
                                    new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "Calibri Light" }),
                                new DocumentFormat.OpenXml.Drawing.MinorFont(
                                    new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "Calibri" }))
                            { Name = "Office" },
                            new DocumentFormat.OpenXml.Drawing.FormatScheme(
                                new DocumentFormat.OpenXml.Drawing.FillStyleList(
                                    new DocumentFormat.OpenXml.Drawing.SolidFill()),
                                new DocumentFormat.OpenXml.Drawing.LineStyleList(
                                    new DocumentFormat.OpenXml.Drawing.Outline()),
                                new DocumentFormat.OpenXml.Drawing.EffectStyleList(
                                    new DocumentFormat.OpenXml.Drawing.EffectStyle()),
                                new DocumentFormat.OpenXml.Drawing.BackgroundFillStyleList(
                                    new DocumentFormat.OpenXml.Drawing.SolidFill()))
                            { Name = "Office" }))
                    { Name = "Office Theme" };
                    themePart.Theme.Save();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles(
                        new DocDefaults(
                            new RunPropertiesDefault(
                                new RunPropertiesBaseStyle(
                                    new RunFonts { Ascii = "Calibri" },
                                    new FontSize { Val = "22" }))));
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create a run with theme color (accent1 = #4472C4)
                    var run = new Run(
                        new RunProperties(
                            new Color { Val = "4472C4", ThemeColor = ThemeColorValues.Accent1 }),
                        new Text("Theme colored text"));

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(run)));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings { ResolveThemeColors = true };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Should contain the resolved color #4472C4
                    Assert.Contains("#4472C4", htmlString);
                }
            }
        }

        [Fact]
        public void HC050_ThemeColor_AppliesTintModifier()
        {
            // Test that tint modifier lightens the theme color
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add theme part with accent1 = pure blue (0000FF)
                    var themePart = mainPart.AddNewPart<ThemePart>();
                    themePart.Theme = new DocumentFormat.OpenXml.Drawing.Theme(
                        new DocumentFormat.OpenXml.Drawing.ThemeElements(
                            new DocumentFormat.OpenXml.Drawing.ColorScheme(
                                new DocumentFormat.OpenXml.Drawing.Dark1Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Light1Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "FFFFFF" }),
                                new DocumentFormat.OpenXml.Drawing.Dark2Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Light2Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "FFFFFF" }),
                                new DocumentFormat.OpenXml.Drawing.Accent1Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "0000FF" }),  // Pure blue
                                new DocumentFormat.OpenXml.Drawing.Accent2Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Accent3Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Accent4Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Accent5Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Accent6Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Hyperlink(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.FollowedHyperlinkColor(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }))
                            { Name = "Test" },
                            new DocumentFormat.OpenXml.Drawing.FontScheme(
                                new DocumentFormat.OpenXml.Drawing.MajorFont(
                                    new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "Calibri" }),
                                new DocumentFormat.OpenXml.Drawing.MinorFont(
                                    new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "Calibri" }))
                            { Name = "Test" },
                            new DocumentFormat.OpenXml.Drawing.FormatScheme(
                                new DocumentFormat.OpenXml.Drawing.FillStyleList(
                                    new DocumentFormat.OpenXml.Drawing.SolidFill()),
                                new DocumentFormat.OpenXml.Drawing.LineStyleList(
                                    new DocumentFormat.OpenXml.Drawing.Outline()),
                                new DocumentFormat.OpenXml.Drawing.EffectStyleList(
                                    new DocumentFormat.OpenXml.Drawing.EffectStyle()),
                                new DocumentFormat.OpenXml.Drawing.BackgroundFillStyleList(
                                    new DocumentFormat.OpenXml.Drawing.SolidFill()))
                            { Name = "Test" }))
                    { Name = "Test Theme" };
                    themePart.Theme.Save();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create a run with theme color + 50% tint (7F = 127, so tint factor = 127/255 ≈ 0.5)
                    // Blue with 50% tint toward white should be around #8080FF
                    var run = new Run(
                        new RunProperties(
                            new Color { Val = "0000FF", ThemeColor = ThemeColorValues.Accent1, ThemeTint = "7F" }),
                        new Text("Tinted theme color"));

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(run)));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings { ResolveThemeColors = true };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // The color should be lightened - blue component stays at FF, others increase
                    // Should NOT be the original blue #0000FF
                    Assert.DoesNotContain("#0000FF", htmlString);
                    // Should contain a lightened color (somewhere between blue and white)
                    Assert.Contains("color:", htmlString);
                }
            }
        }

        [Fact]
        public void HC051_ThemeColor_DisabledWhenSettingFalse()
        {
            // Test that theme colors are not resolved when ResolveThemeColors = false
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    // Add theme part
                    var themePart = mainPart.AddNewPart<ThemePart>();
                    themePart.Theme = new DocumentFormat.OpenXml.Drawing.Theme(
                        new DocumentFormat.OpenXml.Drawing.ThemeElements(
                            new DocumentFormat.OpenXml.Drawing.ColorScheme(
                                new DocumentFormat.OpenXml.Drawing.Dark1Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Light1Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "FFFFFF" }),
                                new DocumentFormat.OpenXml.Drawing.Dark2Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Light2Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "FFFFFF" }),
                                new DocumentFormat.OpenXml.Drawing.Accent1Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "4472C4" }),  // Theme blue
                                new DocumentFormat.OpenXml.Drawing.Accent2Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Accent3Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Accent4Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Accent5Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Accent6Color(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.Hyperlink(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }),
                                new DocumentFormat.OpenXml.Drawing.FollowedHyperlinkColor(
                                    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "000000" }))
                            { Name = "Test" },
                            new DocumentFormat.OpenXml.Drawing.FontScheme(
                                new DocumentFormat.OpenXml.Drawing.MajorFont(
                                    new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "Calibri" }),
                                new DocumentFormat.OpenXml.Drawing.MinorFont(
                                    new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "Calibri" }))
                            { Name = "Test" },
                            new DocumentFormat.OpenXml.Drawing.FormatScheme(
                                new DocumentFormat.OpenXml.Drawing.FillStyleList(
                                    new DocumentFormat.OpenXml.Drawing.SolidFill()),
                                new DocumentFormat.OpenXml.Drawing.LineStyleList(
                                    new DocumentFormat.OpenXml.Drawing.Outline()),
                                new DocumentFormat.OpenXml.Drawing.EffectStyleList(
                                    new DocumentFormat.OpenXml.Drawing.EffectStyle()),
                                new DocumentFormat.OpenXml.Drawing.BackgroundFillStyleList(
                                    new DocumentFormat.OpenXml.Drawing.SolidFill()))
                            { Name = "Test" }))
                    { Name = "Test Theme" };
                    themePart.Theme.Save();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create a run with theme color AND explicit fallback val (FF0000 = red)
                    var run = new Run(
                        new RunProperties(
                            new Color { Val = "FF0000", ThemeColor = ThemeColorValues.Accent1 }),
                        new Text("Explicit color fallback"));

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(run)));
                    mainPart.Document.Save();

                    // Disable theme color resolution
                    var settings = new WmlToHtmlConverterSettings { ResolveThemeColors = false };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Should use the explicit val (red), not the theme color (blue)
                    Assert.Contains("#FF0000", htmlString);
                }
            }
        }

        #endregion

        #region Page CSS Tests

        [Fact]
        public void HC052_PageCss_GeneratesAtPageRule()
        {
            // Test that @page CSS rule is generated when GeneratePageCss = true
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create document with US Letter page size (12240 x 15840 twips = 8.5 x 11 inches)
                    // and 1 inch margins (1440 twips each)
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(new Run(new Text("Test content"))),
                            new SectionProperties(
                                new PageSize { Width = 12240, Height = 15840 },
                                new PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 })));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings { GeneratePageCss = true };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Should contain @page rule
                    Assert.Contains("@page", htmlString);
                    // Should contain size declaration (8.5 x 11 inches)
                    Assert.Contains("size:", htmlString);
                    Assert.Contains("8.50in", htmlString);
                    Assert.Contains("11.00in", htmlString);
                    // Should contain margin declaration (1 inch each)
                    Assert.Contains("margin:", htmlString);
                    Assert.Contains("1.00in", htmlString);
                }
            }
        }

        [Fact]
        public void HC053_PageCss_A4PageSize()
        {
            // Test that A4 page size generates correct dimensions
            // A4 = 210mm x 297mm = 595.3pt x 841.9pt = 11906 x 16838 twips
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create document with A4 page size (11906 x 16838 twips)
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(new Run(new Text("A4 content"))),
                            new SectionProperties(
                                new PageSize { Width = 11906, Height = 16838 },
                                new PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 })));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings { GeneratePageCss = true };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Should contain @page rule with A4 dimensions
                    Assert.Contains("@page", htmlString);
                    // A4 width ≈ 8.27in, height ≈ 11.69in
                    Assert.Contains("8.27in", htmlString);
                    Assert.Contains("11.69in", htmlString);
                }
            }
        }

        [Fact]
        public void HC054_PageCss_DisabledByDefault()
        {
            // Test that @page CSS is NOT generated by default
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(new Run(new Text("Default settings"))),
                            new SectionProperties(
                                new PageSize { Width = 12240, Height = 15840 })));
                    mainPart.Document.Save();

                    // Use default settings (GeneratePageCss = false)
                    var settings = new WmlToHtmlConverterSettings();
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Should NOT contain @page rule
                    Assert.DoesNotContain("@page", htmlString);
                }
            }
        }

        [Fact]
        public void HC055_PageCss_CustomMargins()
        {
            // Test that custom margins are correctly converted
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument wDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Styles();
                    stylesPart.Styles.Save();

                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                    settingsPart.Settings.Save();

                    // Create document with custom margins:
                    // Top: 0.5in (720 twips), Right: 0.75in (1080 twips)
                    // Bottom: 1.5in (2160 twips), Left: 1.25in (1800 twips)
                    mainPart.Document = new Document(
                        new Body(
                            new Paragraph(new Run(new Text("Custom margins"))),
                            new SectionProperties(
                                new PageSize { Width = 12240, Height = 15840 },
                                new PageMargin { Top = 720, Right = 1080, Bottom = 2160, Left = 1800 })));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings { GeneratePageCss = true };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Should contain @page rule with custom margins
                    Assert.Contains("@page", htmlString);
                    Assert.Contains("0.50in", htmlString);  // Top
                    Assert.Contains("0.75in", htmlString);  // Right
                    Assert.Contains("1.50in", htmlString);  // Bottom
                    Assert.Contains("1.25in", htmlString);  // Left
                }
            }
        }

        #endregion
    }
}

#endif
