// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Docxodus;
using W = DocumentFormat.OpenXml.Wordprocessing;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    // Regression coverage for issue #210: WmlToHtmlConverter threw
    // FormatException ("Format_InvalidStringWithValue, 100%") when a table or
    // cell width used w:type="pct" with a percent-suffixed value (e.g.
    // w:w="100%"), the form emitted by the `docx` JS library. Per the OOXML
    // ST_TblWidth / ST_MeasurementOrPercent schema the value may be either a
    // plain integer (fiftieths of a percent) OR a "<number>%" string.
    public class HcTablePercentageWidthTests
    {
        // Builds a minimal in-memory .docx with a single 2x1 table.
        // tblWidth/tblType set the table-level w:tblW; cellWidth/cellType set
        // the per-cell w:tcW. Widths are passed through verbatim as the raw
        // OOXML attribute string so we can exercise the "100%" / "50%" forms.
        private static byte[] CreateDocxWithTableWidth(
            string tblWidth, W.TableWidthUnitValues tblType,
            string cellWidth, W.TableWidthUnitValues cellType)
        {
            using var ms = new MemoryStream();
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();

                // WmlToHtmlConverter routes through FormattingAssembler, which
                // dereferences MainDocumentPart.StyleDefinitionsPart. A minimal
                // in-memory document must supply one (plus default run props) or
                // conversion throws ArgumentNullException before any width parsing.
                var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                stylesPart.Styles = new W.Styles(
                    new W.DocDefaults(
                        new W.RunPropertiesDefault(
                            new W.RunPropertiesBaseStyle(
                                new W.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                                new W.FontSize { Val = "24" }))));
                stylesPart.Styles.Save();

                // ConvertToHtml also dereferences MainDocumentPart.DocumentSettingsPart
                // (CalculateSpanWidthForTabs reads w:defaultTabStop).
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new W.Settings();
                settingsPart.Settings.Save();

                W.TableCell MakeCell(string text) => new W.TableCell(
                    new W.TableCellProperties(
                        new W.TableCellWidth { Width = cellWidth, Type = cellType }),
                    new W.Paragraph(
                        new W.Run(
                            new W.Text(text) { Space = SpaceProcessingModeValues.Preserve })));

                var table = new W.Table(
                    new W.TableProperties(
                        new W.TableWidth { Width = tblWidth, Type = tblType }),
                    new W.TableGrid(
                        new W.GridColumn { Width = "4500" },
                        new W.GridColumn { Width = "4500" }),
                    new W.TableRow(MakeCell("Item"), MakeCell("Amount")),
                    new W.TableRow(MakeCell("Fee"), MakeCell("1000")));

                mainPart.Document = new W.Document(
                    new W.Body(
                        new W.Paragraph(
                            new W.Run(
                                new W.Text("Header") { Space = SpaceProcessingModeValues.Preserve })),
                        table,
                        new W.SectionProperties(
                            new W.PageSize { Width = 12240, Height = 15840 },
                            new W.PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440 })));
                mainPart.Document.Save();
            }
            return ms.ToArray();
        }

        private static string RenderToHtml(byte[] docxBytes)
        {
            var wmlDoc = new WmlDocument("test.docx", docxBytes);
            var settings = new WmlToHtmlConverterSettings();
            var html = WmlToHtmlConverter.ConvertToHtml(wmlDoc, settings);
            return html.ToString(SaveOptions.DisableFormatting);
        }

        // Issue #210 core repro: percent-suffixed string widths (what `docx`
        // emits for WidthType.PERCENTAGE). Previously threw FormatException.
        [Fact]
        public void HC_Pct_PercentSuffixString_DoesNotThrow_AndEmitsPercentWidths()
        {
            var bytes = CreateDocxWithTableWidth(
                "100%", W.TableWidthUnitValues.Pct,
                "50%", W.TableWidthUnitValues.Pct);

            string html = null;
            var ex = Record.Exception(() => html = RenderToHtml(bytes));

            Assert.Null(ex);                       // #210: must not throw
            Assert.NotNull(html);

            var normalized = html.Replace(" ", "");
            // Explicit "100%" is already a percentage (not fiftieths).
            Assert.Contains("width:100%", normalized);
            // Explicit "50%" cell width -> rendered with one decimal place.
            Assert.Contains("width:50.0%", normalized);
        }

        // The integer fiftieths-of-a-percent form must keep working: 5000 -> 100%,
        // 2500 -> 50.0%.
        [Fact]
        public void HC_Pct_IntegerFiftieths_StillYieldsPercentWidths()
        {
            var bytes = CreateDocxWithTableWidth(
                "5000", W.TableWidthUnitValues.Pct,
                "2500", W.TableWidthUnitValues.Pct);

            var html = RenderToHtml(bytes);
            var normalized = html.Replace(" ", "");

            Assert.Contains("width:100%", normalized);
            Assert.Contains("width:50.0%", normalized);
        }

        // DXA (twips) widths are unaffected by the fix: 9000 twips -> 450pt,
        // 4500 twips -> 225pt.
        [Fact]
        public void HC_Dxa_TwipsWidths_StillYieldPointWidths()
        {
            var bytes = CreateDocxWithTableWidth(
                "9000", W.TableWidthUnitValues.Dxa,
                "4500", W.TableWidthUnitValues.Dxa);

            var html = RenderToHtml(bytes);
            var normalized = html.Replace(" ", "");

            Assert.Contains("width:450pt", normalized);
            Assert.Contains("width:225pt", normalized);
        }

        // A malformed / non-numeric width must be ignored gracefully rather than
        // throwing — the helper returns null and no width is emitted.
        [Fact]
        public void HC_Pct_GarbageWidth_IsIgnored_DoesNotThrow()
        {
            var bytes = CreateDocxWithTableWidth(
                "not-a-number", W.TableWidthUnitValues.Pct,
                "2500", W.TableWidthUnitValues.Pct);

            string html = null;
            var ex = Record.Exception(() => html = RenderToHtml(bytes));

            Assert.Null(ex);
            Assert.NotNull(html);
            // Cell width still parses (2500 -> 50.0%); table width silently dropped.
            Assert.Contains("width:50.0%", html.Replace(" ", ""));
        }
    }
}

#endif
