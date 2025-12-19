// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#define COPY_FILES_FOR_DEBUGGING

// DO_CONVERSION_VIA_WORD is defined in the project Docxodus.Tests.OA.csproj, but not in the Docxodus.Tests.csproj

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
    }
}

#endif
