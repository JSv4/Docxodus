// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Docxodus;
using Wp = DocumentFormat.OpenXml.Wordprocessing;


namespace Docxodus.Tests
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
            string tblWidth, Wp.TableWidthUnitValues tblType,
            string cellWidth, Wp.TableWidthUnitValues cellType)
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
                stylesPart.Styles = new Wp.Styles(
                    new Wp.DocDefaults(
                        new Wp.RunPropertiesDefault(
                            new Wp.RunPropertiesBaseStyle(
                                new Wp.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                                new Wp.FontSize { Val = "24" }))));
                stylesPart.Styles.Save();

                // ConvertToHtml also dereferences MainDocumentPart.DocumentSettingsPart
                // (CalculateSpanWidthForTabs reads w:defaultTabStop).
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Wp.Settings();
                settingsPart.Settings.Save();

                Wp.TableCell MakeCell(string text) => new Wp.TableCell(
                    new Wp.TableCellProperties(
                        new Wp.TableCellWidth { Width = cellWidth, Type = cellType }),
                    new Wp.Paragraph(
                        new Wp.Run(
                            new Wp.Text(text) { Space = SpaceProcessingModeValues.Preserve })));

                var table = new Wp.Table(
                    new Wp.TableProperties(
                        new Wp.TableWidth { Width = tblWidth, Type = tblType }),
                    new Wp.TableGrid(
                        new Wp.GridColumn { Width = "4500" },
                        new Wp.GridColumn { Width = "4500" }),
                    new Wp.TableRow(MakeCell("Item"), MakeCell("Amount")),
                    new Wp.TableRow(MakeCell("Fee"), MakeCell("1000")));

                mainPart.Document = new Wp.Document(
                    new Wp.Body(
                        new Wp.Paragraph(
                            new Wp.Run(
                                new Wp.Text("Header") { Space = SpaceProcessingModeValues.Preserve })),
                        table,
                        new Wp.SectionProperties(
                            new Wp.PageSize { Width = 12240, Height = 15840 },
                            new Wp.PageMargin { Top = 1440, Bottom = 1440, Left = 1440, Right = 1440 })));
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
                "100%", Wp.TableWidthUnitValues.Pct,
                "50%", Wp.TableWidthUnitValues.Pct);

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
                "5000", Wp.TableWidthUnitValues.Pct,
                "2500", Wp.TableWidthUnitValues.Pct);

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
                "9000", Wp.TableWidthUnitValues.Dxa,
                "4500", Wp.TableWidthUnitValues.Dxa);

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
                "not-a-number", Wp.TableWidthUnitValues.Pct,
                "2500", Wp.TableWidthUnitValues.Pct);

            string html = null;
            var ex = Record.Exception(() => html = RenderToHtml(bytes));

            Assert.Null(ex);
            Assert.NotNull(html);
            // Cell width still parses (2500 -> 50.0%); table width silently dropped.
            Assert.Contains("width:50.0%", html.Replace(" ", ""));
        }
    }
}

