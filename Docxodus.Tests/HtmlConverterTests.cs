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
    }
}

#endif
