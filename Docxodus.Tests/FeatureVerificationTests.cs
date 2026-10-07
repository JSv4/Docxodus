#nullable enable

// Feature verification tests for resolved WmlToHtmlConverter gaps
// Tests all features marked as RESOLVED in docs/architecture/wml_to_html_converter_gaps.md

using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;
using Drw = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests
{
    /// <summary>
    /// Comprehensive feature verification tests for resolved WmlToHtmlConverter gaps.
    /// Each test verifies one of the features marked as RESOLVED in the gaps document.
    /// </summary>
    public class FeatureVerificationTests
    {
        #region 1. @page CSS Rule Tests

        [Fact]
        public void FV001_PageCss_GeneratesAtPageRule_USLetter()
        {
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    AddBasicStyles(mainPart);
                    AddBasicSettings(mainPart);

                    mainPart.Document = new Wp.Document(
                        new Wp.Body(
                            new Wp.Paragraph(new Wp.Run(new Wp.Text("Test content"))),
                            new Wp.SectionProperties(
                                new Wp.PageSize { Width = 12240, Height = 15840 }, // US Letter in twips
                                new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 }
                            )
                        )
                    );
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings { GeneratePageCss = true };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    Assert.Contains("@page", htmlString);
                    Assert.Contains("size:", htmlString);
                    Assert.Contains("8.50in", htmlString);
                    Assert.Contains("11.00in", htmlString);
                    Assert.Contains("margin:", htmlString);
                    Assert.Contains("1.00in", htmlString);
                }
            }
        }

        #endregion

        #region 2. Table Width Calculation (DXA to Points)

        [Fact]
        public void FV002_TableDxaWidth_ConvertsToPoints()
        {
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    AddBasicStyles(mainPart);
                    AddBasicSettings(mainPart);

                    // Table with 6 inch DXA width = 8640 twips = 432pt
                    var table = new Wp.Table(
                        new Wp.TableProperties(
                            new Wp.TableWidth { Width = "8640", Type = Wp.TableWidthUnitValues.Dxa },
                            new Wp.TableBorders(
                                new Wp.TopBorder { Val = Wp.BorderValues.Single, Size = 4 },
                                new Wp.BottomBorder { Val = Wp.BorderValues.Single, Size = 4 }
                            )
                        )
                    );
                    var row = new Wp.TableRow();
                    var cell = new Wp.TableCell(new Wp.Paragraph(new Wp.Run(new Wp.Text("Cell content"))));
                    row.Append(cell);
                    table.Append(row);

                    mainPart.Document = new Wp.Document(new Wp.Body(table));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings();
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // 8640 twips / 20 = 432pt
                    Assert.Contains("432pt", htmlString);
                }
            }
        }

        #endregion

        #region 3. Borderless Table Detection

        [Fact]
        public void FV003_BorderlessTable_HasDataAttribute()
        {
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    AddBasicStyles(mainPart);
                    AddBasicSettings(mainPart);

                    var table = new Wp.Table(
                        new Wp.TableProperties(
                            new Wp.TableBorders(
                                new Wp.TopBorder { Val = Wp.BorderValues.Nil },
                                new Wp.LeftBorder { Val = Wp.BorderValues.Nil },
                                new Wp.BottomBorder { Val = Wp.BorderValues.Nil },
                                new Wp.RightBorder { Val = Wp.BorderValues.Nil },
                                new Wp.InsideHorizontalBorder { Val = Wp.BorderValues.Nil },
                                new Wp.InsideVerticalBorder { Val = Wp.BorderValues.Nil }
                            )
                        )
                    );
                    var row = new Wp.TableRow();
                    var cell = new Wp.TableCell(new Wp.Paragraph(new Wp.Run(new Wp.Text("Borderless cell"))));
                    row.Append(cell);
                    table.Append(row);

                    mainPart.Document = new Wp.Document(new Wp.Body(table));
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings();
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    Assert.Contains("data-borderless=\"true\"", htmlString);
                }
            }
        }

        #endregion

        #region 4. Theme Color Resolution

        [Fact]
        public void FV004_ThemeColor_ResolvesAccent1()
        {
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    AddThemePart(mainPart, accent1Color: "4472C4"); // Blue
                    AddBasicStyles(mainPart);
                    AddBasicSettings(mainPart);

                    mainPart.Document = new Wp.Document(
                        new Wp.Body(
                            new Wp.Paragraph(
                                new Wp.Run(
                                    new Wp.RunProperties(new Wp.Color { ThemeColor = Wp.ThemeColorValues.Accent1 }),
                                    new Wp.Text("Theme colored text")
                                )
                            )
                        )
                    );
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings { ResolveThemeColors = true };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // accent1 = #4472C4 from our theme
                    Assert.Contains("#4472C4", htmlString);
                }
            }
        }

        [Fact]
        public void FV005_ThemeColor_DisabledWhenSettingFalse()
        {
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    AddThemePart(mainPart, accent1Color: "4472C4"); // Blue
                    AddBasicStyles(mainPart);
                    AddBasicSettings(mainPart);

                    // Use theme color but also provide explicit Val (fallback)
                    mainPart.Document = new Wp.Document(
                        new Wp.Body(
                            new Wp.Paragraph(
                                new Wp.Run(
                                    new Wp.RunProperties(
                                        new Wp.Color { Val = "FF0000", ThemeColor = Wp.ThemeColorValues.Accent1 }
                                    ),
                                    new Wp.Text("Text with fallback color")
                                )
                            )
                        )
                    );
                    mainPart.Document.Save();

                    // Disable theme color resolution
                    var settings = new WmlToHtmlConverterSettings { ResolveThemeColors = false };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Should use the fallback red color, not theme blue
                    Assert.Contains("#FF0000", htmlString);
                    Assert.DoesNotContain("#4472C4", htmlString);
                }
            }
        }

        #endregion

        #region 5. Document Language on <html>

        [Fact]
        public void FV006_DocumentLanguage_FromThemeFontLang()
        {
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    AddBasicStyles(mainPart);

                    // Set document language to French
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Wp.Settings(
                        new Wp.ThemeFontLanguages { Val = "fr-FR" }
                    );
                    settingsPart.Settings.Save();

                    mainPart.Document = new Wp.Document(
                        new Wp.Body(new Wp.Paragraph(new Wp.Run(new Wp.Text("Bonjour"))))
                    );
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
        public void FV007_DocumentLanguage_SettingOverridesDocument()
        {
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    AddBasicStyles(mainPart);

                    // Document is French
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Wp.Settings(
                        new Wp.ThemeFontLanguages { Val = "fr-FR" }
                    );
                    settingsPart.Settings.Save();

                    mainPart.Document = new Wp.Document(
                        new Wp.Body(new Wp.Paragraph(new Wp.Run(new Wp.Text("Test"))))
                    );
                    mainPart.Document.Save();

                    // Override to German
                    var settings = new WmlToHtmlConverterSettings { DocumentLanguage = "de-DE" };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

                    var langAttr = html.Attribute("lang");
                    Assert.NotNull(langAttr);
                    Assert.Equal("de-DE", langAttr.Value);
                }
            }
        }

        #endregion

        #region 6. Foreign Language Span Attributes

        [Fact]
        public void FV008_ForeignTextSpan_HasLangAttribute()
        {
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    AddBasicStyles(mainPart);
                    AddBasicSettings(mainPart); // Default is en-US

                    mainPart.Document = new Wp.Document(
                        new Wp.Body(
                            new Wp.Paragraph(
                                new Wp.Run(new Wp.Text("English text ") { Space = SpaceProcessingModeValues.Preserve }),
                                new Wp.Run(
                                    new Wp.RunProperties(new Wp.Languages { Val = "es" }),
                                    new Wp.Text("Texto en español")
                                )
                            )
                        )
                    );
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings();
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Spanish text should have lang="es" since document default is en-US
                    Assert.Contains("lang=\"es\"", htmlString);
                    Assert.Contains("Texto en español", htmlString);
                }
            }
        }

        [Fact]
        public void FV009_ForeignTextSpan_Japanese_HasLangAttribute()
        {
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    AddBasicStyles(mainPart);
                    AddBasicSettings(mainPart); // Default is en-US

                    mainPart.Document = new Wp.Document(
                        new Wp.Body(
                            new Wp.Paragraph(
                                new Wp.Run(
                                    new Wp.RunProperties(
                                        new Wp.RunFonts { EastAsia = "MS Mincho" },
                                        new Wp.Languages { EastAsia = "ja-JP" }
                                    ),
                                    new Wp.Text("日本語テスト")
                                )
                            )
                        )
                    );
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings();
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Japanese text should have lang attribute
                    Assert.Contains("日本語テスト", htmlString);
                    // Check for Japanese language marker
                    Assert.True(
                        htmlString.Contains("lang=\"ja\"") || htmlString.Contains("lang=\"ja-JP\""),
                        "Expected Japanese lang attribute (ja or ja-JP)"
                    );
                }
            }
        }

        #endregion

        #region 7. Font Fallback Improvements

        [Fact]
        public void FV010_UnknownFont_GetsSerifFallback()
        {
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Wp.Styles(
                        new Wp.DocDefaults(
                            new Wp.RunPropertiesDefault(
                                new Wp.RunPropertiesBaseStyle(
                                    new Wp.RunFonts { Ascii = "MyUnknownProprietaryFont" },
                                    new Wp.FontSize { Val = "24" }
                                )
                            )
                        )
                    );
                    stylesPart.Styles.Save();
                    AddBasicSettings(mainPart);

                    mainPart.Document = new Wp.Document(
                        new Wp.Body(new Wp.Paragraph(new Wp.Run(new Wp.Text("Test with unknown font"))))
                    );
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings();
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Unknown font should get generic serif fallback
                    Assert.Contains("MyUnknownProprietaryFont", htmlString);
                    Assert.Contains("serif", htmlString);
                }
            }
        }

        [Fact]
        public void FV011_UnknownSansFont_GetsSansSerifFallback()
        {
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Wp.Styles(
                        new Wp.DocDefaults(
                            new Wp.RunPropertiesDefault(
                                new Wp.RunPropertiesBaseStyle(
                                    new Wp.RunFonts { Ascii = "CustomSansFont" },
                                    new Wp.FontSize { Val = "24" }
                                )
                            )
                        )
                    );
                    stylesPart.Styles.Save();
                    AddBasicSettings(mainPart);

                    mainPart.Document = new Wp.Document(
                        new Wp.Body(new Wp.Paragraph(new Wp.Run(new Wp.Text("Test with sans font"))))
                    );
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings();
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Font with "sans" in name should get sans-serif fallback
                    Assert.Contains("CustomSansFont", htmlString);
                    Assert.Contains("sans-serif", htmlString);
                }
            }
        }

        [Fact]
        public void FV012_CourierNew_GetsMonospaceFallback()
        {
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Wp.Styles(
                        new Wp.DocDefaults(
                            new Wp.RunPropertiesDefault(
                                new Wp.RunPropertiesBaseStyle(
                                    new Wp.RunFonts { Ascii = "Courier New" },
                                    new Wp.FontSize { Val = "24" }
                                )
                            )
                        )
                    );
                    stylesPart.Styles.Save();
                    AddBasicSettings(mainPart);

                    mainPart.Document = new Wp.Document(
                        new Wp.Body(new Wp.Paragraph(new Wp.Run(new Wp.Text("Code sample"))))
                    );
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings();
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Courier New should get monospace fallback
                    Assert.Contains("Courier New", htmlString);
                    Assert.Contains("monospace", htmlString);
                }
            }
        }

        #endregion

        #region 8. CJK Font-Family Fallback Chain

        [Fact]
        public void FV013_JapaneseText_GetsCjkFallbackChain()
        {
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Wp.Styles(
                        new Wp.DocDefaults(
                            new Wp.RunPropertiesDefault(
                                new Wp.RunPropertiesBaseStyle(
                                    new Wp.RunFonts { Ascii = "Times New Roman", EastAsia = "MS Mincho" },
                                    new Wp.FontSize { Val = "24" }
                                )
                            )
                        )
                    );
                    stylesPart.Styles.Save();
                    AddBasicSettings(mainPart);

                    mainPart.Document = new Wp.Document(
                        new Wp.Body(
                            new Wp.Paragraph(
                                new Wp.Run(
                                    new Wp.RunProperties(
                                        new Wp.RunFonts { EastAsia = "MS Mincho" },
                                        new Wp.Languages { EastAsia = "ja-JP" }
                                    ),
                                    new Wp.Text("日本語テスト")
                                )
                            )
                        )
                    );
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings();
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Should include Japanese CJK fallback fonts like Noto Serif CJK JP
                    Assert.Contains("Noto Serif CJK JP", htmlString);
                }
            }
        }

        [Fact]
        public void FV014_SimplifiedChinese_GetsCjkScFallbackChain()
        {
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();

                    var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                    stylesPart.Styles = new Wp.Styles(
                        new Wp.DocDefaults(
                            new Wp.RunPropertiesDefault(
                                new Wp.RunPropertiesBaseStyle(
                                    new Wp.RunFonts { Ascii = "Times New Roman", EastAsia = "SimSun" },
                                    new Wp.FontSize { Val = "24" }
                                )
                            )
                        )
                    );
                    stylesPart.Styles.Save();
                    AddBasicSettings(mainPart);

                    mainPart.Document = new Wp.Document(
                        new Wp.Body(
                            new Wp.Paragraph(
                                new Wp.Run(
                                    new Wp.RunProperties(
                                        new Wp.RunFonts { EastAsia = "SimSun" },
                                        new Wp.Languages { EastAsia = "zh-CN" }
                                    ),
                                    new Wp.Text("简体中文测试")
                                )
                            )
                        )
                    );
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings();
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Should include Simplified Chinese CJK fallback fonts
                    Assert.Contains("Noto Serif CJK SC", htmlString);
                }
            }
        }

        #endregion

        #region Comprehensive Test - All Features Together

        [Fact]
        public void FV099_ComprehensiveTest_AllResolvedFeatures()
        {
            // This test verifies ALL resolved features work together in a single document
            using (var ms = new MemoryStream())
            {
                using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    var mainPart = wDoc.AddMainDocumentPart();
                    AddThemePart(mainPart, accent1Color: "4472C4");
                    AddBasicStyles(mainPart);

                    // Set document language to en-US
                    var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Wp.Settings(
                        new Wp.ThemeFontLanguages { Val = "en-US" }
                    );
                    settingsPart.Settings.Save();

                    var body = new Wp.Body();

                    // 1. Add themed paragraph
                    body.Append(new Wp.Paragraph(
                        new Wp.Run(
                            new Wp.RunProperties(new Wp.Color { ThemeColor = Wp.ThemeColorValues.Accent1 }),
                            new Wp.Text("AI Report - Theme colored heading")
                        )
                    ));

                    // 2. Add foreign language text
                    body.Append(new Wp.Paragraph(
                        new Wp.Run(new Wp.Text("Markets: ") { Space = SpaceProcessingModeValues.Preserve }),
                        new Wp.Run(
                            new Wp.RunProperties(new Wp.Languages { Val = "fr" }),
                            new Wp.Text("France ")
                        ),
                        new Wp.Run(
                            new Wp.RunProperties(new Wp.Languages { Val = "es" }),
                            new Wp.Text("España")
                        )
                    ));

                    // 3. Add table with DXA width
                    var table = new Wp.Table(
                        new Wp.TableProperties(
                            new Wp.TableWidth { Width = "4320", Type = Wp.TableWidthUnitValues.Dxa }, // 3 inches = 216pt
                            new Wp.TableBorders(
                                new Wp.TopBorder { Val = Wp.BorderValues.Single, Size = 4 },
                                new Wp.BottomBorder { Val = Wp.BorderValues.Single, Size = 4 }
                            )
                        )
                    );
                    var tRow = new Wp.TableRow();
                    var tCell = new Wp.TableCell(new Wp.Paragraph(new Wp.Run(new Wp.Text("Data cell"))));
                    tRow.Append(tCell);
                    table.Append(tRow);
                    body.Append(table);

                    // 4. Add borderless table
                    var borderlessTable = new Wp.Table(
                        new Wp.TableProperties(
                            new Wp.TableBorders(
                                new Wp.TopBorder { Val = Wp.BorderValues.Nil },
                                new Wp.LeftBorder { Val = Wp.BorderValues.Nil },
                                new Wp.BottomBorder { Val = Wp.BorderValues.Nil },
                                new Wp.RightBorder { Val = Wp.BorderValues.Nil }
                            )
                        )
                    );
                    var bRow = new Wp.TableRow();
                    var bCell = new Wp.TableCell(new Wp.Paragraph(new Wp.Run(new Wp.Text("Signature: _______________"))));
                    bRow.Append(bCell);
                    borderlessTable.Append(bRow);
                    body.Append(borderlessTable);

                    // 5. Add Japanese text
                    body.Append(new Wp.Paragraph(
                        new Wp.Run(new Wp.Text("Japan: ") { Space = SpaceProcessingModeValues.Preserve }),
                        new Wp.Run(
                            new Wp.RunProperties(
                                new Wp.RunFonts { EastAsia = "MS Mincho" },
                                new Wp.Languages { EastAsia = "ja-JP" }
                            ),
                            new Wp.Text("人工知能")
                        )
                    ));

                    // 6. Add code sample with monospace font
                    body.Append(new Wp.Paragraph(
                        new Wp.Run(
                            new Wp.RunProperties(new Wp.RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" }),
                            new Wp.Text("console.log('AI')")
                        )
                    ));

                    // 7. Add unknown font text
                    body.Append(new Wp.Paragraph(
                        new Wp.Run(
                            new Wp.RunProperties(new Wp.RunFonts { Ascii = "ProprietaryBrandFont" }),
                            new Wp.Text("Custom branded content")
                        )
                    ));

                    // Add page settings
                    body.Append(new Wp.SectionProperties(
                        new Wp.PageSize { Width = 12240, Height = 15840 },
                        new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 }
                    ));

                    mainPart.Document = new Wp.Document(body);
                    mainPart.Document.Save();

                    var settings = new WmlToHtmlConverterSettings
                    {
                        GeneratePageCss = true,
                        ResolveThemeColors = true
                    };
                    var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                    var htmlString = html.ToString();

                    // Write to file for manual inspection
                    var outputPath = Path.Combine(Directory.GetCurrentDirectory(), "comprehensive_test_output.html");
                    File.WriteAllText(outputPath, htmlString);

                    // Verify all features
                    var failures = new System.Collections.Generic.List<string>();

                    // 1. @page CSS
                    if (!htmlString.Contains("@page")) failures.Add("FAIL: No @page CSS rule");
                    if (!htmlString.Contains("8.50in")) failures.Add("FAIL: No 8.50in page width");

                    // 2. Document language
                    var htmlElement = html;
                    var langAttr = htmlElement.Attribute("lang");
                    if (langAttr == null || langAttr.Value != "en-US") failures.Add("FAIL: lang attribute not en-US");

                    // 3. Table DXA width (4320 twips = 216pt)
                    if (!htmlString.Contains("216pt")) failures.Add("FAIL: No 216pt table width");

                    // 4. Borderless table
                    if (!htmlString.Contains("data-borderless=\"true\"")) failures.Add("FAIL: No data-borderless attribute");

                    // 5. Theme color resolution
                    if (!htmlString.Contains("#4472C4")) failures.Add("FAIL: Theme color #4472C4 not resolved");

                    // 6. Foreign language spans
                    if (!htmlString.Contains("lang=\"fr\"")) failures.Add("FAIL: No French lang attribute");
                    if (!htmlString.Contains("lang=\"es\"")) failures.Add("FAIL: No Spanish lang attribute");

                    // 7. CJK fallback
                    if (!htmlString.Contains("Noto Serif CJK JP")) failures.Add("FAIL: No Japanese CJK fallback chain");

                    // 8. Monospace fallback
                    if (!htmlString.Contains("monospace")) failures.Add("FAIL: No monospace fallback for Courier New");

                    // 9. Unknown font gets serif fallback
                    if (!htmlString.Contains("ProprietaryBrandFont")) failures.Add("FAIL: Unknown font not preserved");
                    if (!htmlString.Contains("serif")) failures.Add("FAIL: No serif fallback for unknown font");

                    if (failures.Count > 0)
                    {
                        Assert.Fail("Feature verification failures:\n" + string.Join("\n", failures));
                    }
                }
            }
        }

        #endregion

        #region Helper Methods

        private void AddBasicStyles(MainDocumentPart mainPart)
        {
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Wp.Styles(
                new Wp.DocDefaults(
                    new Wp.RunPropertiesDefault(
                        new Wp.RunPropertiesBaseStyle(
                            new Wp.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                            new Wp.FontSize { Val = "24" }
                        )
                    )
                )
            );
            stylesPart.Styles.Save();
        }

        private void AddBasicSettings(MainDocumentPart mainPart)
        {
            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Wp.Settings(
                new Wp.ThemeFontLanguages { Val = "en-US" }
            );
            settingsPart.Settings.Save();
        }

        private void AddThemePart(MainDocumentPart mainPart, string accent1Color)
        {
            var themePart = mainPart.AddNewPart<ThemePart>();
            themePart.Theme = new Drw.Theme(
                new Drw.ThemeElements(
                    new Drw.ColorScheme(
                        new Drw.Dark1Color(new Drw.RgbColorModelHex { Val = "000000" }),
                        new Drw.Light1Color(new Drw.RgbColorModelHex { Val = "FFFFFF" }),
                        new Drw.Dark2Color(new Drw.RgbColorModelHex { Val = "44546A" }),
                        new Drw.Light2Color(new Drw.RgbColorModelHex { Val = "E7E6E6" }),
                        new Drw.Accent1Color(new Drw.RgbColorModelHex { Val = accent1Color }),
                        new Drw.Accent2Color(new Drw.RgbColorModelHex { Val = "ED7D31" }),
                        new Drw.Accent3Color(new Drw.RgbColorModelHex { Val = "A5A5A5" }),
                        new Drw.Accent4Color(new Drw.RgbColorModelHex { Val = "FFC000" }),
                        new Drw.Accent5Color(new Drw.RgbColorModelHex { Val = "5B9BD5" }),
                        new Drw.Accent6Color(new Drw.RgbColorModelHex { Val = "70AD47" }),
                        new Drw.Hyperlink(new Drw.RgbColorModelHex { Val = "0563C1" }),
                        new Drw.FollowedHyperlinkColor(new Drw.RgbColorModelHex { Val = "954F72" })
                    )
                    { Name = "Office" },
                    new Drw.FontScheme(
                        new Drw.MajorFont(new Drw.LatinFont { Typeface = "Calibri Light" }),
                        new Drw.MinorFont(new Drw.LatinFont { Typeface = "Calibri" })
                    )
                    { Name = "Office" },
                    new Drw.FormatScheme(
                        new Drw.FillStyleList(
                            new Drw.SolidFill(new Drw.SchemeColor { Val = Drw.SchemeColorValues.PhColor })),
                        new Drw.LineStyleList(
                            new Drw.Outline(new Drw.SolidFill(new Drw.SchemeColor { Val = Drw.SchemeColorValues.PhColor }))
                            { Width = 6350 }),
                        new Drw.EffectStyleList(new Drw.EffectStyle(new Drw.EffectList())),
                        new Drw.BackgroundFillStyleList(
                            new Drw.SolidFill(new Drw.SchemeColor { Val = Drw.SchemeColorValues.PhColor }))
                    )
                    { Name = "Office" }
                )
            )
            { Name = "Office Theme" };
            themePart.Theme.Save();
        }

        #endregion
    }
}
