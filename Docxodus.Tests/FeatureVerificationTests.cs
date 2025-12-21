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
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OxPt
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

                    mainPart.Document = new W.Document(
                        new W.Body(
                            new W.Paragraph(new W.Run(new W.Text("Test content"))),
                            new W.SectionProperties(
                                new W.PageSize { Width = 12240, Height = 15840 }, // US Letter in twips
                                new W.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 }
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
                    var table = new W.Table(
                        new W.TableProperties(
                            new W.TableWidth { Width = "8640", Type = W.TableWidthUnitValues.Dxa },
                            new W.TableBorders(
                                new W.TopBorder { Val = W.BorderValues.Single, Size = 4 },
                                new W.BottomBorder { Val = W.BorderValues.Single, Size = 4 }
                            )
                        )
                    );
                    var row = new W.TableRow();
                    var cell = new W.TableCell(new W.Paragraph(new W.Run(new W.Text("Cell content"))));
                    row.Append(cell);
                    table.Append(row);

                    mainPart.Document = new W.Document(new W.Body(table));
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

                    var table = new W.Table(
                        new W.TableProperties(
                            new W.TableBorders(
                                new W.TopBorder { Val = W.BorderValues.Nil },
                                new W.LeftBorder { Val = W.BorderValues.Nil },
                                new W.BottomBorder { Val = W.BorderValues.Nil },
                                new W.RightBorder { Val = W.BorderValues.Nil },
                                new W.InsideHorizontalBorder { Val = W.BorderValues.Nil },
                                new W.InsideVerticalBorder { Val = W.BorderValues.Nil }
                            )
                        )
                    );
                    var row = new W.TableRow();
                    var cell = new W.TableCell(new W.Paragraph(new W.Run(new W.Text("Borderless cell"))));
                    row.Append(cell);
                    table.Append(row);

                    mainPart.Document = new W.Document(new W.Body(table));
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

                    mainPart.Document = new W.Document(
                        new W.Body(
                            new W.Paragraph(
                                new W.Run(
                                    new W.RunProperties(new W.Color { ThemeColor = W.ThemeColorValues.Accent1 }),
                                    new W.Text("Theme colored text")
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
                    mainPart.Document = new W.Document(
                        new W.Body(
                            new W.Paragraph(
                                new W.Run(
                                    new W.RunProperties(
                                        new W.Color { Val = "FF0000", ThemeColor = W.ThemeColorValues.Accent1 }
                                    ),
                                    new W.Text("Text with fallback color")
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
                    settingsPart.Settings = new W.Settings(
                        new W.ThemeFontLanguages { Val = "fr-FR" }
                    );
                    settingsPart.Settings.Save();

                    mainPart.Document = new W.Document(
                        new W.Body(new W.Paragraph(new W.Run(new W.Text("Bonjour"))))
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
                    settingsPart.Settings = new W.Settings(
                        new W.ThemeFontLanguages { Val = "fr-FR" }
                    );
                    settingsPart.Settings.Save();

                    mainPart.Document = new W.Document(
                        new W.Body(new W.Paragraph(new W.Run(new W.Text("Test"))))
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

                    mainPart.Document = new W.Document(
                        new W.Body(
                            new W.Paragraph(
                                new W.Run(new W.Text("English text ") { Space = SpaceProcessingModeValues.Preserve }),
                                new W.Run(
                                    new W.RunProperties(new W.Languages { Val = "es" }),
                                    new W.Text("Texto en español")
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

                    mainPart.Document = new W.Document(
                        new W.Body(
                            new W.Paragraph(
                                new W.Run(
                                    new W.RunProperties(
                                        new W.RunFonts { EastAsia = "MS Mincho" },
                                        new W.Languages { EastAsia = "ja-JP" }
                                    ),
                                    new W.Text("日本語テスト")
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
                    stylesPart.Styles = new W.Styles(
                        new W.DocDefaults(
                            new W.RunPropertiesDefault(
                                new W.RunPropertiesBaseStyle(
                                    new W.RunFonts { Ascii = "MyUnknownProprietaryFont" },
                                    new W.FontSize { Val = "24" }
                                )
                            )
                        )
                    );
                    stylesPart.Styles.Save();
                    AddBasicSettings(mainPart);

                    mainPart.Document = new W.Document(
                        new W.Body(new W.Paragraph(new W.Run(new W.Text("Test with unknown font"))))
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
                    stylesPart.Styles = new W.Styles(
                        new W.DocDefaults(
                            new W.RunPropertiesDefault(
                                new W.RunPropertiesBaseStyle(
                                    new W.RunFonts { Ascii = "CustomSansFont" },
                                    new W.FontSize { Val = "24" }
                                )
                            )
                        )
                    );
                    stylesPart.Styles.Save();
                    AddBasicSettings(mainPart);

                    mainPart.Document = new W.Document(
                        new W.Body(new W.Paragraph(new W.Run(new W.Text("Test with sans font"))))
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
                    stylesPart.Styles = new W.Styles(
                        new W.DocDefaults(
                            new W.RunPropertiesDefault(
                                new W.RunPropertiesBaseStyle(
                                    new W.RunFonts { Ascii = "Courier New" },
                                    new W.FontSize { Val = "24" }
                                )
                            )
                        )
                    );
                    stylesPart.Styles.Save();
                    AddBasicSettings(mainPart);

                    mainPart.Document = new W.Document(
                        new W.Body(new W.Paragraph(new W.Run(new W.Text("Code sample"))))
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
                    stylesPart.Styles = new W.Styles(
                        new W.DocDefaults(
                            new W.RunPropertiesDefault(
                                new W.RunPropertiesBaseStyle(
                                    new W.RunFonts { Ascii = "Times New Roman", EastAsia = "MS Mincho" },
                                    new W.FontSize { Val = "24" }
                                )
                            )
                        )
                    );
                    stylesPart.Styles.Save();
                    AddBasicSettings(mainPart);

                    mainPart.Document = new W.Document(
                        new W.Body(
                            new W.Paragraph(
                                new W.Run(
                                    new W.RunProperties(
                                        new W.RunFonts { EastAsia = "MS Mincho" },
                                        new W.Languages { EastAsia = "ja-JP" }
                                    ),
                                    new W.Text("日本語テスト")
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
                    stylesPart.Styles = new W.Styles(
                        new W.DocDefaults(
                            new W.RunPropertiesDefault(
                                new W.RunPropertiesBaseStyle(
                                    new W.RunFonts { Ascii = "Times New Roman", EastAsia = "SimSun" },
                                    new W.FontSize { Val = "24" }
                                )
                            )
                        )
                    );
                    stylesPart.Styles.Save();
                    AddBasicSettings(mainPart);

                    mainPart.Document = new W.Document(
                        new W.Body(
                            new W.Paragraph(
                                new W.Run(
                                    new W.RunProperties(
                                        new W.RunFonts { EastAsia = "SimSun" },
                                        new W.Languages { EastAsia = "zh-CN" }
                                    ),
                                    new W.Text("简体中文测试")
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
                    settingsPart.Settings = new W.Settings(
                        new W.ThemeFontLanguages { Val = "en-US" }
                    );
                    settingsPart.Settings.Save();

                    var body = new W.Body();

                    // 1. Add themed paragraph
                    body.Append(new W.Paragraph(
                        new W.Run(
                            new W.RunProperties(new W.Color { ThemeColor = W.ThemeColorValues.Accent1 }),
                            new W.Text("AI Report - Theme colored heading")
                        )
                    ));

                    // 2. Add foreign language text
                    body.Append(new W.Paragraph(
                        new W.Run(new W.Text("Markets: ") { Space = SpaceProcessingModeValues.Preserve }),
                        new W.Run(
                            new W.RunProperties(new W.Languages { Val = "fr" }),
                            new W.Text("France ")
                        ),
                        new W.Run(
                            new W.RunProperties(new W.Languages { Val = "es" }),
                            new W.Text("España")
                        )
                    ));

                    // 3. Add table with DXA width
                    var table = new W.Table(
                        new W.TableProperties(
                            new W.TableWidth { Width = "4320", Type = W.TableWidthUnitValues.Dxa }, // 3 inches = 216pt
                            new W.TableBorders(
                                new W.TopBorder { Val = W.BorderValues.Single, Size = 4 },
                                new W.BottomBorder { Val = W.BorderValues.Single, Size = 4 }
                            )
                        )
                    );
                    var tRow = new W.TableRow();
                    var tCell = new W.TableCell(new W.Paragraph(new W.Run(new W.Text("Data cell"))));
                    tRow.Append(tCell);
                    table.Append(tRow);
                    body.Append(table);

                    // 4. Add borderless table
                    var borderlessTable = new W.Table(
                        new W.TableProperties(
                            new W.TableBorders(
                                new W.TopBorder { Val = W.BorderValues.Nil },
                                new W.LeftBorder { Val = W.BorderValues.Nil },
                                new W.BottomBorder { Val = W.BorderValues.Nil },
                                new W.RightBorder { Val = W.BorderValues.Nil }
                            )
                        )
                    );
                    var bRow = new W.TableRow();
                    var bCell = new W.TableCell(new W.Paragraph(new W.Run(new W.Text("Signature: _______________"))));
                    bRow.Append(bCell);
                    borderlessTable.Append(bRow);
                    body.Append(borderlessTable);

                    // 5. Add Japanese text
                    body.Append(new W.Paragraph(
                        new W.Run(new W.Text("Japan: ") { Space = SpaceProcessingModeValues.Preserve }),
                        new W.Run(
                            new W.RunProperties(
                                new W.RunFonts { EastAsia = "MS Mincho" },
                                new W.Languages { EastAsia = "ja-JP" }
                            ),
                            new W.Text("人工知能")
                        )
                    ));

                    // 6. Add code sample with monospace font
                    body.Append(new W.Paragraph(
                        new W.Run(
                            new W.RunProperties(new W.RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" }),
                            new W.Text("console.log('AI')")
                        )
                    ));

                    // 7. Add unknown font text
                    body.Append(new W.Paragraph(
                        new W.Run(
                            new W.RunProperties(new W.RunFonts { Ascii = "ProprietaryBrandFont" }),
                            new W.Text("Custom branded content")
                        )
                    ));

                    // Add page settings
                    body.Append(new W.SectionProperties(
                        new W.PageSize { Width = 12240, Height = 15840 },
                        new W.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 }
                    ));

                    mainPart.Document = new W.Document(body);
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
            stylesPart.Styles = new W.Styles(
                new W.DocDefaults(
                    new W.RunPropertiesDefault(
                        new W.RunPropertiesBaseStyle(
                            new W.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                            new W.FontSize { Val = "24" }
                        )
                    )
                )
            );
            stylesPart.Styles.Save();
        }

        private void AddBasicSettings(MainDocumentPart mainPart)
        {
            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new W.Settings(
                new W.ThemeFontLanguages { Val = "en-US" }
            );
            settingsPart.Settings.Save();
        }

        private void AddThemePart(MainDocumentPart mainPart, string accent1Color)
        {
            var themePart = mainPart.AddNewPart<ThemePart>();
            themePart.Theme = new A.Theme(
                new A.ThemeElements(
                    new A.ColorScheme(
                        new A.Dark1Color(new A.RgbColorModelHex { Val = "000000" }),
                        new A.Light1Color(new A.RgbColorModelHex { Val = "FFFFFF" }),
                        new A.Dark2Color(new A.RgbColorModelHex { Val = "44546A" }),
                        new A.Light2Color(new A.RgbColorModelHex { Val = "E7E6E6" }),
                        new A.Accent1Color(new A.RgbColorModelHex { Val = accent1Color }),
                        new A.Accent2Color(new A.RgbColorModelHex { Val = "ED7D31" }),
                        new A.Accent3Color(new A.RgbColorModelHex { Val = "A5A5A5" }),
                        new A.Accent4Color(new A.RgbColorModelHex { Val = "FFC000" }),
                        new A.Accent5Color(new A.RgbColorModelHex { Val = "5B9BD5" }),
                        new A.Accent6Color(new A.RgbColorModelHex { Val = "70AD47" }),
                        new A.Hyperlink(new A.RgbColorModelHex { Val = "0563C1" }),
                        new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "954F72" })
                    )
                    { Name = "Office" },
                    new A.FontScheme(
                        new A.MajorFont(new A.LatinFont { Typeface = "Calibri Light" }),
                        new A.MinorFont(new A.LatinFont { Typeface = "Calibri" })
                    )
                    { Name = "Office" },
                    new A.FormatScheme(
                        new A.FillStyleList(
                            new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor })),
                        new A.LineStyleList(
                            new A.Outline(new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }))
                            { Width = 6350 }),
                        new A.EffectStyleList(new A.EffectStyle(new A.EffectList())),
                        new A.BackgroundFillStyleList(
                            new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }))
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
