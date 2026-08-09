// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus
{
    public static partial class WmlToHtmlConverter
    {
        private static readonly XNamespace Svg = "http://www.w3.org/2000/svg";

        private sealed class CachedChartSeries
        {
            public string Name { get; init; }

            public string Color { get; init; }

            public Dictionary<int, double> Values { get; init; }
        }

        /// <summary>
        /// Projects the cached data of a standard clustered Word bar/column chart into inline SVG.
        /// Word stores those values in the chart part specifically so consumers can display the
        /// chart without recalculating its embedded workbook. Keeping this projection in the DOCX
        /// converter makes charts visible in both standalone and WASM output without a JS runtime
        /// dependency or a server-side Office process.
        /// </summary>
        private static XElement ProcessChart(WordprocessingDocument wordDoc, XElement drawing)
        {
            var chartReference = drawing.Descendants(C.chart).FirstOrDefault();
            var relationshipId = (string)chartReference?.Attribute(R.id);
            if (string.IsNullOrWhiteSpace(relationshipId))
                return null;

            ChartPart chartPart;
            try
            {
                chartPart = wordDoc.MainDocumentPart?.GetPartById(relationshipId) as ChartPart;
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
            catch (InvalidOperationException)
            {
                return null;
            }

            var chartSpace = chartPart?.GetXDocument().Root;
            var chart = chartSpace?.Element(C.chart);
            var plotArea = chart?.Element(C.plotArea);
            var barChart = plotArea?.Element(C.barChart);
            if (barChart == null)
                return null;

            var grouping = (string)barChart.Element(C.grouping)?.Attribute("val");
            if (!string.IsNullOrEmpty(grouping) && grouping != "clustered")
                return null;

            var container = drawing.Elements().FirstOrDefault(e => e.Name == WP.inline || e.Name == WP.anchor);
            var extent = container?.Element(WP.extent);
            var widthEmu = (long?)extent?.Attribute("cx");
            var heightEmu = (long?)extent?.Attribute("cy");
            if (widthEmu is null or <= 0 || heightEmu is null or <= 0)
                return null;

            var width = widthEmu.Value / 9525.0;   // EMUs -> CSS px at the converter's 96-DPI contract
            var height = heightEmu.Value / 9525.0;
            if (width < 80 || height < 60)
                return null;

            var theme = LoadThemeColorScheme(wordDoc);
            var series = barChart.Elements(C.ser)
                .Select((element, index) => ReadCachedChartSeries(element, index, theme))
                .Where(item => item != null && item.Values.Count > 0)
                .ToList();
            if (series.Count == 0)
                return null;

            var categoryCount = series.SelectMany(item => item.Values.Keys).DefaultIfEmpty(-1).Max() + 1;
            if (categoryCount <= 0)
                return null;

            var categories = ReadCachedCategories(barChart.Elements(C.ser).First(), categoryCount);
            var title = ReadChartTitle(chart);
            var titleFontSize = ReadChartFontSize(chart.Element(C.title), 14 * 96 / 72.0);
            var legend = chart.Element(C.legend);
            var showLegend = legend != null && !IsDeleted(legend);
            var legendFontSize = ReadChartFontSize(legend, 9 * 96 / 72.0);
            var horizontal = (string)barChart.Element(C.barDir)?.Attribute("val") == "bar";
            var gapWidth = ReadIntAttribute(barChart.Element(C.gapWidth), "val", 150, 0, 500);
            var overlap = ReadIntAttribute(barChart.Element(C.overlap), "val", 0, -100, 100);

            var categoryAxis = plotArea.Element(C.catAx);
            var valueAxis = plotArea.Element(C.valAx);
            var categoryFontSize = ReadChartFontSize(categoryAxis, 9 * 96 / 72.0);
            var valueFontSize = ReadChartFontSize(valueAxis, 9 * 96 / 72.0);
            var scaling = valueAxis?.Element(C.c + "scaling");
            var explicitMinimum = ReadDoubleAttribute(scaling?.Element(C.min), "val");
            var explicitMaximum = ReadDoubleAttribute(scaling?.Element(C.max), "val");
            var explicitMajorUnit = ReadDoubleAttribute(valueAxis?.Element(C.majorUnit), "val");

            var altText = (string)container?.Element(WP.docPr)?.Attribute("descr")
                ?? (string)container?.Element(WP.docPr)?.Attribute("name")
                ?? title
                ?? "Chart";
            var svg = new XElement(Svg + "svg",
                new XAttribute("viewBox", $"0 0 {Format(width)} {Format(height)}"),
                new XAttribute("role", "img"),
                new XAttribute("aria-label", altText),
                new XAttribute("data-chart-type", horizontal ? "bar" : "column"),
                new XAttribute("style", string.Format(CultureInfo.InvariantCulture,
                    "display: inline-block; width: {0:0.##}pt; height: {1:0.##}pt; vertical-align: top;",
                    widthEmu.Value / 12700.0, heightEmu.Value / 12700.0)),
                new XElement(Svg + "title", altText),
                new XElement(Svg + "rect",
                    new XAttribute("x", "0.5"),
                    new XAttribute("y", "0.5"),
                    new XAttribute("width", Format(width - 1)),
                    new XAttribute("height", Format(height - 1)),
                    new XAttribute("fill", "#FFFFFF"),
                    new XAttribute("stroke", "#D0D0D0"),
                    new XAttribute("stroke-width", "1")));

            if (!string.IsNullOrWhiteSpace(title))
                AddSvgText(svg, width / 2, Math.Max(titleFontSize, height * 0.105), title,
                    "middle", titleFontSize);

            var values = series.SelectMany(item => item.Values.Values).ToList();
            var scale = BuildChartScale(values, explicitMinimum, explicitMaximum, explicitMajorUnit);
            if (horizontal)
                RenderHorizontalBarChart(svg, width, height, title, showLegend, series, categories,
                    categoryCount, gapWidth, overlap, scale, categoryFontSize, valueFontSize);
            else
                RenderColumnChart(svg, width, height, title, showLegend, series, categories,
                    categoryCount, gapWidth, overlap, scale, categoryFontSize, valueFontSize);

            if (showLegend)
                RenderChartLegend(svg, width, height, series, legendFontSize);

            return svg;
        }

        private sealed class ChartScale
        {
            public double Minimum { get; init; }

            public double Maximum { get; init; }

            public double MajorUnit { get; init; }
        }

        private static ChartScale BuildChartScale(IReadOnlyCollection<double> values,
            double? explicitMinimum, double? explicitMaximum, double? explicitMajorUnit)
        {
            var dataMinimum = Math.Min(0, values.Min());
            var dataMaximum = Math.Max(0, values.Max());
            var rawRange = Math.Max(1e-9, dataMaximum - dataMinimum);
            var majorUnit = explicitMajorUnit is > 0
                ? explicitMajorUnit.Value
                : NiceNumber(rawRange / 5, true);
            if (majorUnit <= 0 || double.IsNaN(majorUnit) || double.IsInfinity(majorUnit))
                majorUnit = 1;

            var minimum = explicitMinimum ?? Math.Floor(dataMinimum / majorUnit) * majorUnit;
            // Office adds one tick of headroom when the largest value lands exactly on a major
            // gridline. Besides matching the stock chart, this keeps the tallest bar off the frame.
            var maximum = explicitMaximum ??
                Math.Ceiling((dataMaximum + majorUnit * 0.001) / majorUnit) * majorUnit;
            if (maximum <= minimum)
                maximum = minimum + majorUnit;

            // Protect output size against a malformed axis with a microscopic explicit unit.
            if ((maximum - minimum) / majorUnit > 20)
                majorUnit = NiceNumber((maximum - minimum) / 10, true);

            return new ChartScale { Minimum = minimum, Maximum = maximum, MajorUnit = majorUnit };
        }

        private static double NiceNumber(double value, bool round)
        {
            var exponent = Math.Floor(Math.Log10(Math.Max(value, 1e-9)));
            var fraction = value / Math.Pow(10, exponent);
            double niceFraction;
            if (round)
                niceFraction = fraction < 1.5 ? 1 : fraction < 3 ? 2 : fraction < 7 ? 5 : 10;
            else
                niceFraction = fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10;
            return niceFraction * Math.Pow(10, exponent);
        }

        private static void RenderColumnChart(XElement svg, double width, double height, string title,
            bool showLegend, IReadOnlyList<CachedChartSeries> series, IReadOnlyList<string> categories,
            int categoryCount, int gapWidth, int overlap, ChartScale scale,
            double categoryFontSize, double valueFontSize)
        {
            var plotLeft = Math.Max(22, 10 + FormatAxisValue(scale.Maximum, scale.MajorUnit).Length * 6);
            var plotRight = width - 12;
            var plotTop = string.IsNullOrWhiteSpace(title) ? 14 : 60;
            var plotBottom = height - (showLegend ? 54 : 28);
            var plotWidth = Math.Max(1, plotRight - plotLeft);
            var plotHeight = Math.Max(1, plotBottom - plotTop);
            var range = scale.Maximum - scale.Minimum;
            double Y(double value) => plotBottom - (value - scale.Minimum) / range * plotHeight;

            for (var value = scale.Minimum; value <= scale.Maximum + scale.MajorUnit * 0.0001;
                 value += scale.MajorUnit)
            {
                var y = Y(value);
                svg.Add(new XElement(Svg + "line",
                    new XAttribute("x1", Format(plotLeft)),
                    new XAttribute("x2", Format(plotRight)),
                    new XAttribute("y1", Format(y)),
                    new XAttribute("y2", Format(y)),
                    new XAttribute("stroke", "#D9D9D9"),
                    new XAttribute("stroke-width", "1")));
                AddSvgText(svg, plotLeft - 5, y, FormatAxisValue(value, scale.MajorUnit), "end",
                    valueFontSize, true);
            }

            var slot = plotWidth / categoryCount;
            var negativeOverlapGap = Math.Max(0, -overlap / 100.0);
            var positiveOverlap = Math.Max(0, overlap / 100.0);
            var denominator = series.Count + gapWidth / 100.0 +
                              (series.Count - 1) * (negativeOverlapGap - positiveOverlap);
            var barWidth = Math.Max(1, slot / Math.Max(1, denominator));
            var barStep = barWidth * (1 + negativeOverlapGap - positiveOverlap);
            var groupWidth = barWidth + Math.Max(0, series.Count - 1) * barStep;
            var zeroY = Y(0);

            for (var categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++)
            {
                var groupLeft = plotLeft + categoryIndex * slot + (slot - groupWidth) / 2;
                for (var seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
                {
                    if (!series[seriesIndex].Values.TryGetValue(categoryIndex, out var value))
                        continue;
                    var valueY = Y(value);
                    var top = Math.Min(zeroY, valueY);
                    var barHeight = Math.Max(0.5, Math.Abs(zeroY - valueY));
                    svg.Add(ChartBar(groupLeft + seriesIndex * barStep, top, barWidth, barHeight,
                        series[seriesIndex], seriesIndex, categoryIndex, value));
                }

                AddSvgText(svg, plotLeft + (categoryIndex + 0.5) * slot, plotBottom + 15,
                    categories[categoryIndex], "middle", categoryFontSize);
            }
        }

        private static void RenderHorizontalBarChart(XElement svg, double width, double height, string title,
            bool showLegend, IReadOnlyList<CachedChartSeries> series, IReadOnlyList<string> categories,
            int categoryCount, int gapWidth, int overlap, ChartScale scale,
            double categoryFontSize, double valueFontSize)
        {
            var longestCategory = categories.DefaultIfEmpty(string.Empty).Max(item => item?.Length ?? 0);
            var plotLeft = Math.Min(width * 0.38, Math.Max(40, 12 + longestCategory * 6));
            var plotRight = width - 12;
            var plotTop = string.IsNullOrWhiteSpace(title) ? 14 : 60;
            var plotBottom = height - (showLegend ? 54 : 32);
            var plotWidth = Math.Max(1, plotRight - plotLeft);
            var plotHeight = Math.Max(1, plotBottom - plotTop);
            var range = scale.Maximum - scale.Minimum;
            double X(double value) => plotLeft + (value - scale.Minimum) / range * plotWidth;

            for (var value = scale.Minimum; value <= scale.Maximum + scale.MajorUnit * 0.0001;
                 value += scale.MajorUnit)
            {
                var x = X(value);
                svg.Add(new XElement(Svg + "line",
                    new XAttribute("x1", Format(x)),
                    new XAttribute("x2", Format(x)),
                    new XAttribute("y1", Format(plotTop)),
                    new XAttribute("y2", Format(plotBottom)),
                    new XAttribute("stroke", "#D9D9D9"),
                    new XAttribute("stroke-width", "1")));
                AddSvgText(svg, x, plotBottom + 15, FormatAxisValue(value, scale.MajorUnit), "middle",
                    valueFontSize);
            }

            var slot = plotHeight / categoryCount;
            var negativeOverlapGap = Math.Max(0, -overlap / 100.0);
            var positiveOverlap = Math.Max(0, overlap / 100.0);
            var denominator = series.Count + gapWidth / 100.0 +
                              (series.Count - 1) * (negativeOverlapGap - positiveOverlap);
            var barHeight = Math.Max(1, slot / Math.Max(1, denominator));
            var barStep = barHeight * (1 + negativeOverlapGap - positiveOverlap);
            var groupHeight = barHeight + Math.Max(0, series.Count - 1) * barStep;
            var zeroX = X(0);

            for (var categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++)
            {
                var groupTop = plotTop + categoryIndex * slot + (slot - groupHeight) / 2;
                AddSvgText(svg, plotLeft - 6, plotTop + (categoryIndex + 0.5) * slot,
                    categories[categoryIndex], "end", categoryFontSize, true);
                for (var seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
                {
                    if (!series[seriesIndex].Values.TryGetValue(categoryIndex, out var value))
                        continue;
                    var valueX = X(value);
                    var left = Math.Min(zeroX, valueX);
                    var barWidth = Math.Max(0.5, Math.Abs(zeroX - valueX));
                    svg.Add(ChartBar(left, groupTop + seriesIndex * barStep, barWidth, barHeight,
                        series[seriesIndex], seriesIndex, categoryIndex, value));
                }
            }
        }

        private static XElement ChartBar(double x, double y, double width, double height,
            CachedChartSeries series, int seriesIndex, int categoryIndex, double value) =>
            new XElement(Svg + "rect",
                new XAttribute("class", "docx-chart-bar"),
                new XAttribute("data-chart-series", seriesIndex),
                new XAttribute("data-chart-category", categoryIndex),
                new XAttribute("data-chart-value", value.ToString("R", CultureInfo.InvariantCulture)),
                new XAttribute("x", Format(x)),
                new XAttribute("y", Format(y)),
                new XAttribute("width", Format(width)),
                new XAttribute("height", Format(height)),
                new XAttribute("fill", series.Color));

        private static void RenderChartLegend(XElement svg, double width, double height,
            IReadOnlyList<CachedChartSeries> series, double fontSize)
        {
            const double interItemGap = 7.5;
            var widths = series.Select(item => 20.5 + Math.Max(1, item.Name.Length) * fontSize * 0.375)
                .ToList();
            var total = widths.Sum() - interItemGap;
            var x = Math.Max(4, (width - total) / 2);
            var y = height - 11;
            for (var index = 0; index < series.Count; index++)
            {
                svg.Add(new XElement(Svg + "rect",
                    new XAttribute("x", Format(x)),
                    new XAttribute("y", Format(y - 4)),
                    new XAttribute("width", "8"),
                    new XAttribute("height", "8"),
                    new XAttribute("fill", series[index].Color)));
                AddSvgText(svg, x + 12, y, series[index].Name, "start", fontSize, true);
                x += widths[index];
            }
        }

        private static void AddSvgText(XElement svg, double x, double y, string value,
            string anchor, double fontSize, bool middleBaseline = false)
        {
            if (string.IsNullOrEmpty(value))
                return;
            svg.Add(new XElement(Svg + "text",
                new XAttribute("x", Format(x)),
                new XAttribute("y", Format(y)),
                new XAttribute("text-anchor", anchor),
                middleBaseline ? new XAttribute("dominant-baseline", "middle") : null,
                new XAttribute("fill", "#595959"),
                new XAttribute("font-family", "Calibri, Carlito, sans-serif"),
                new XAttribute("font-size", Format(fontSize)),
                new XText(value)));
        }

        private static CachedChartSeries ReadCachedChartSeries(XElement series, int index,
            ThemeColorScheme theme)
        {
            var values = ReadCachedPoints(series.Element(C.val));
            if (values.Count == 0)
                return null;
            var name = ReadCachedText(series.Element(C.tx));
            if (string.IsNullOrWhiteSpace(name))
                name = $"Series {index + 1}";
            return new CachedChartSeries
            {
                Name = name,
                Color = ReadChartColor(series.Element(C.spPr), index, theme),
                Values = values,
            };
        }

        private static IReadOnlyList<string> ReadCachedCategories(XElement firstSeries, int count)
        {
            var points = firstSeries.Element(C.cat)?.Descendants(C.pt)
                .Select(point => new
                {
                    Index = ReadIntAttribute(point, "idx", -1, -1, int.MaxValue),
                    Value = point.Element(C.v)?.Value,
                })
                .Where(point => point.Index >= 0)
                .GroupBy(point => point.Index)
                .ToDictionary(group => group.Key, group => group.First().Value ?? string.Empty)
                ?? new Dictionary<int, string>();
            return Enumerable.Range(0, count)
                .Select(index => points.TryGetValue(index, out var value) && !string.IsNullOrWhiteSpace(value)
                    ? value
                    : (index + 1).ToString(CultureInfo.InvariantCulture))
                .ToList();
        }

        private static Dictionary<int, double> ReadCachedPoints(XElement valueContainer)
        {
            if (valueContainer == null)
                return new Dictionary<int, double>();
            return valueContainer.Descendants(C.pt)
                .Select(point => new
                {
                    Index = ReadIntAttribute(point, "idx", -1, -1, int.MaxValue),
                    Text = point.Element(C.v)?.Value,
                })
                .Where(point => point.Index >= 0 &&
                                double.TryParse(point.Text, NumberStyles.Float,
                                    CultureInfo.InvariantCulture, out _))
                .GroupBy(point => point.Index)
                .ToDictionary(group => group.Key, group =>
                    double.Parse(group.First().Text, NumberStyles.Float, CultureInfo.InvariantCulture));
        }

        private static string ReadChartTitle(XElement chart)
        {
            var title = chart?.Element(C.title);
            if (title == null || IsDeleted(title))
                return null;
            var explicitText = string.Concat(title.Descendants(A.t).Select(element => element.Value));
            if (string.IsNullOrWhiteSpace(explicitText))
                explicitText = ReadCachedText(title.Element(C.tx));
            return string.IsNullOrWhiteSpace(explicitText) ? "Chart Title" : explicitText;
        }

        private static string ReadCachedText(XElement container) =>
            container?.Descendants(C.v).Select(element => element.Value)
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));

        private static string ReadChartColor(XElement shapeProperties, int seriesIndex,
            ThemeColorScheme theme)
        {
            var fill = shapeProperties?.Element(A.solidFill);
            var rgb = (string)fill?.Element(A.srgbClr)?.Attribute("val");
            var schemeColor = fill?.Element(A.schemeClr);
            if (!IsHexColor(rgb) && schemeColor != null)
            {
                var name = (string)schemeColor.Attribute("val");
                if (!string.IsNullOrEmpty(name))
                    theme?.Colors.TryGetValue(name, out rgb);
                if (IsHexColor(rgb))
                    rgb = ApplyChartLuminosity(rgb, schemeColor);
            }

            if (!IsHexColor(rgb))
            {
                var defaults = new[]
                {
                    "5B9BD5", "ED7D31", "A5A5A5", "FFC000", "4472C4", "70AD47",
                };
                rgb = defaults[seriesIndex % defaults.Length];
            }
            return "#" + rgb.ToUpperInvariant();
        }

        private static string ApplyChartLuminosity(string color, XElement schemeColor)
        {
            var red = int.Parse(color.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            var green = int.Parse(color.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            var blue = int.Parse(color.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            var lumMod = ReadDoubleAttribute(schemeColor.Element(A.lumMod), "val") ?? 100000;
            var lumOff = ReadDoubleAttribute(schemeColor.Element(A.lumOff), "val") ?? 0;
            int Apply(int channel) => (int)Math.Round(Math.Clamp(
                channel * lumMod / 100000 + 255 * lumOff / 100000, 0, 255));
            return $"{Apply(red):X2}{Apply(green):X2}{Apply(blue):X2}";
        }

        private static bool IsDeleted(XElement element) =>
            string.Equals((string)element.Element(C.delete)?.Attribute("val"), "1",
                StringComparison.OrdinalIgnoreCase) ||
            string.Equals((string)element.Element(C.delete)?.Attribute("val"), "true",
                StringComparison.OrdinalIgnoreCase);

        private static int ReadIntAttribute(XElement element, string name, int fallback, int minimum, int maximum)
        {
            return int.TryParse((string)element?.Attribute(name), NumberStyles.Integer,
                       CultureInfo.InvariantCulture, out var value)
                ? Math.Clamp(value, minimum, maximum)
                : fallback;
        }

        private static double? ReadDoubleAttribute(XElement element, string name)
        {
            return double.TryParse((string)element?.Attribute(name), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var value) && double.IsFinite(value)
                ? value
                : null;
        }

        private static double ReadChartFontSize(XElement textContainer, double fallback)
        {
            var size = textContainer?.Descendants(A.defRPr)
                .Select(element => ReadDoubleAttribute(element, "sz"))
                .FirstOrDefault(value => value is > 0);
            // DrawingML stores font size in hundredths of a point. CSS uses 96 px per inch.
            return size is > 0 ? size.Value / 75.0 : fallback;
        }

        private static string FormatAxisValue(double value, double majorUnit)
        {
            if (Math.Abs(value) < majorUnit * 0.0001)
                value = 0;
            var decimals = majorUnit >= 1 ? 0 : Math.Min(4,
                Math.Max(0, (int)Math.Ceiling(-Math.Log10(majorUnit))));
            return value.ToString("F" + decimals, CultureInfo.InvariantCulture);
        }

        private static string Format(double value) =>
            value.ToString("0.###", CultureInfo.InvariantCulture);
    }
}
