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

        // Chart families that can be projected from cached data. 3-D variants carry the same
        // cached series as their 2-D counterparts, so they project to the flat rendering.
        private static readonly Dictionary<XName, string> SupportedChartPlots = new Dictionary<XName, string>
        {
            [C.barChart] = "bar",
            [C.bar3DChart] = "bar",
            [C.lineChart] = "line",
            [C.line3DChart] = "line",
            [C.areaChart] = "area",
            [C.area3DChart] = "area",
            [C.pieChart] = "pie",
            [C.pie3DChart] = "pie",
            [C.doughnutChart] = "doughnut",
        };

        /// <summary>
        /// Projects the cached data of a Word chart into inline SVG. Covers the bar/column
        /// (clustered, stacked, percent-stacked), line, area, pie, and doughnut families, with
        /// 3-D variants rendered as their 2-D projection. Word stores cached values in the chart
        /// part specifically so consumers can display the chart without recalculating its embedded
        /// workbook. Keeping this projection in the DOCX converter makes charts visible in both
        /// standalone and WASM output without a JS runtime dependency or a server-side Office
        /// process.
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
            var plot = plotArea?.Elements()
                .FirstOrDefault(element => SupportedChartPlots.ContainsKey(element.Name));
            if (plot == null)
                return null;

            var family = SupportedChartPlots[plot.Name];
            var grouping = (string)plot.Element(C.grouping)?.Attribute("val")
                ?? (family == "bar" ? "clustered" : "standard");

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
            var series = plot.Elements(C.ser)
                .Select((element, index) => ReadCachedChartSeries(element, index, theme))
                .Where(item => item != null && item.Values.Count > 0)
                .ToList();
            if (series.Count == 0)
                return null;

            var categoryCount = series.SelectMany(item => item.Values.Keys).DefaultIfEmpty(-1).Max() + 1;
            if (categoryCount <= 0)
                return null;

            var categories = ReadCachedCategories(plot.Elements(C.ser).First(), categoryCount);
            var title = ReadChartTitle(chart);
            var titleFontSize = ReadChartFontSize(chart.Element(C.title), 14 * 96 / 72.0);
            var legend = chart.Element(C.legend);
            var showLegend = legend != null && !IsDeleted(legend);
            var legendFontSize = ReadChartFontSize(legend, 9 * 96 / 72.0);
            var horizontal = family == "bar" && (string)plot.Element(C.barDir)?.Attribute("val") == "bar";
            var gapWidth = ReadIntAttribute(plot.Element(C.gapWidth), "val", 150, 0, 500);
            var overlap = ReadIntAttribute(plot.Element(C.overlap), "val", 0, -100, 100);

            var categoryAxis = plotArea.Element(C.catAx) ?? plotArea.Element(C.dateAx);
            var valueAxis = plotArea.Element(C.valAx);
            var categoryFontSize = ReadChartFontSize(categoryAxis, 9 * 96 / 72.0);
            var valueFontSize = ReadChartFontSize(valueAxis, 9 * 96 / 72.0);
            var scaling = valueAxis?.Element(C.c + "scaling");
            var explicitMinimum = ReadDoubleAttribute(scaling?.Element(C.min), "val");
            var explicitMaximum = ReadDoubleAttribute(scaling?.Element(C.max), "val");
            var explicitMajorUnit = ReadDoubleAttribute(valueAxis?.Element(C.majorUnit), "val");

            var groupingSuffix = grouping switch
            {
                "stacked" => "-stacked",
                "percentStacked" => "-percent-stacked",
                _ => string.Empty,
            };
            var chartType = family switch
            {
                "bar" => (horizontal ? "bar" : "column") + groupingSuffix,
                "line" => "line" + groupingSuffix,
                "area" => "area" + groupingSuffix,
                _ => family,
            };

            var altText = (string)container?.Element(WP.docPr)?.Attribute("descr")
                ?? (string)container?.Element(WP.docPr)?.Attribute("name")
                ?? title
                ?? "Chart";
            var svg = new XElement(Svg + "svg",
                new XAttribute("viewBox", $"0 0 {Format(width)} {Format(height)}"),
                new XAttribute("role", "img"),
                new XAttribute("aria-label", altText),
                new XAttribute("data-chart-type", chartType),
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

            if (family is "pie" or "doughnut")
            {
                var sliceColors = RenderPieChart(svg, width, height, title, showLegend, plot,
                    series[0], categoryCount, family == "doughnut", theme);
                if (showLegend)
                    RenderChartLegend(svg, width, height,
                        Enumerable.Range(0, categoryCount)
                            .Select(index => (categories[index], sliceColors[index]))
                            .ToList(),
                        legendFontSize);
                return svg;
            }

            var stacked = grouping is "stacked" or "percentStacked";
            var axisSuffix = string.Empty;
            if (grouping == "percentStacked")
            {
                series = NormalizeSeriesToPercentages(series, categoryCount);
                axisSuffix = "%";
            }

            var scaleValues = stacked
                ? StackedExtents(series, categoryCount)
                : series.SelectMany(item => item.Values.Values).ToList();
            if (grouping == "percentStacked")
            {
                // Word pins a percent-stacked axis to 100% rather than adding tick headroom.
                explicitMinimum ??= Math.Min(0, scaleValues.Min());
                explicitMaximum ??= 100;
            }
            var scale = BuildChartScale(scaleValues, explicitMinimum, explicitMaximum, explicitMajorUnit);

            if (family is "line" or "area")
                RenderLineOrAreaChart(svg, width, height, title, showLegend, series, categories,
                    categoryCount, family == "area", stacked, scale, categoryFontSize, valueFontSize,
                    axisSuffix);
            else if (horizontal)
                RenderHorizontalBarChart(svg, width, height, title, showLegend, series, categories,
                    categoryCount, gapWidth, overlap, stacked, scale, categoryFontSize, valueFontSize,
                    axisSuffix);
            else
                RenderColumnChart(svg, width, height, title, showLegend, series, categories,
                    categoryCount, gapWidth, overlap, stacked, scale, categoryFontSize, valueFontSize,
                    axisSuffix);

            if (showLegend)
                RenderChartLegend(svg, width, height,
                    series.Select(item => (item.Name, item.Color)).ToList(), legendFontSize);

            return svg;
        }

        private sealed class CachedChartSeries
        {
            public string Name { get; init; }

            public string Color { get; init; }

            public Dictionary<int, double> Values { get; init; }
        }

        private sealed class ChartScale
        {
            public double Minimum { get; init; }

            public double Maximum { get; init; }

            public double MajorUnit { get; init; }
        }

        private static List<CachedChartSeries> NormalizeSeriesToPercentages(
            IReadOnlyList<CachedChartSeries> series, int categoryCount)
        {
            var totals = new double[categoryCount];
            foreach (var item in series)
                foreach (var pair in item.Values)
                    totals[pair.Key] += Math.Abs(pair.Value);
            return series
                .Select(item => new CachedChartSeries
                {
                    Name = item.Name,
                    Color = item.Color,
                    Values = item.Values.ToDictionary(pair => pair.Key,
                        pair => totals[pair.Key] > 0 ? pair.Value / totals[pair.Key] * 100 : 0),
                })
                .ToList();
        }

        private static List<double> StackedExtents(IReadOnlyList<CachedChartSeries> series,
            int categoryCount)
        {
            // A stacked axis spans the per-category totals, with positive and negative values
            // accumulating on their own sides of zero.
            var extents = new List<double>();
            for (var categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++)
            {
                double positive = 0, negative = 0;
                foreach (var item in series)
                {
                    if (!item.Values.TryGetValue(categoryIndex, out var value))
                        continue;
                    if (value >= 0)
                        positive += value;
                    else
                        negative += value;
                }
                extents.Add(positive);
                extents.Add(negative);
            }
            return extents;
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
            int categoryCount, int gapWidth, int overlap, bool stacked, ChartScale scale,
            double categoryFontSize, double valueFontSize, string axisSuffix)
        {
            var plotLeft = Math.Max(22,
                10 + (FormatAxisValue(scale.Maximum, scale.MajorUnit) + axisSuffix).Length * 6);
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
                AddSvgText(svg, plotLeft - 5, y, FormatAxisValue(value, scale.MajorUnit) + axisSuffix,
                    "end", valueFontSize, true);
            }

            var slot = plotWidth / categoryCount;
            double barWidth, barStep, groupWidth;
            if (stacked)
            {
                barWidth = Math.Max(1, slot / (1 + gapWidth / 100.0));
                barStep = 0;
                groupWidth = barWidth;
            }
            else
            {
                var negativeOverlapGap = Math.Max(0, -overlap / 100.0);
                var positiveOverlap = Math.Max(0, overlap / 100.0);
                var denominator = series.Count + gapWidth / 100.0 +
                                  (series.Count - 1) * (negativeOverlapGap - positiveOverlap);
                barWidth = Math.Max(1, slot / Math.Max(1, denominator));
                barStep = barWidth * (1 + negativeOverlapGap - positiveOverlap);
                groupWidth = barWidth + Math.Max(0, series.Count - 1) * barStep;
            }
            var zeroY = Y(0);

            for (var categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++)
            {
                var groupLeft = plotLeft + categoryIndex * slot + (slot - groupWidth) / 2;
                double positive = 0, negative = 0;
                for (var seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
                {
                    if (!series[seriesIndex].Values.TryGetValue(categoryIndex, out var value))
                        continue;
                    double top, barHeight;
                    if (stacked)
                    {
                        var from = value >= 0 ? positive : negative;
                        var to = from + value;
                        top = Math.Min(Y(from), Y(to));
                        barHeight = Math.Max(0.5, Math.Abs(Y(from) - Y(to)));
                        if (value >= 0)
                            positive = to;
                        else
                            negative = to;
                    }
                    else
                    {
                        var valueY = Y(value);
                        top = Math.Min(zeroY, valueY);
                        barHeight = Math.Max(0.5, Math.Abs(zeroY - valueY));
                    }
                    svg.Add(ChartBar(groupLeft + seriesIndex * barStep, top, barWidth, barHeight,
                        series[seriesIndex], seriesIndex, categoryIndex, value));
                }

                AddSvgText(svg, plotLeft + (categoryIndex + 0.5) * slot, plotBottom + 15,
                    categories[categoryIndex], "middle", categoryFontSize);
            }
        }

        private static void RenderHorizontalBarChart(XElement svg, double width, double height, string title,
            bool showLegend, IReadOnlyList<CachedChartSeries> series, IReadOnlyList<string> categories,
            int categoryCount, int gapWidth, int overlap, bool stacked, ChartScale scale,
            double categoryFontSize, double valueFontSize, string axisSuffix)
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
                AddSvgText(svg, x, plotBottom + 15, FormatAxisValue(value, scale.MajorUnit) + axisSuffix,
                    "middle", valueFontSize);
            }

            var slot = plotHeight / categoryCount;
            double barHeight, barStep, groupHeight;
            if (stacked)
            {
                barHeight = Math.Max(1, slot / (1 + gapWidth / 100.0));
                barStep = 0;
                groupHeight = barHeight;
            }
            else
            {
                var negativeOverlapGap = Math.Max(0, -overlap / 100.0);
                var positiveOverlap = Math.Max(0, overlap / 100.0);
                var denominator = series.Count + gapWidth / 100.0 +
                                  (series.Count - 1) * (negativeOverlapGap - positiveOverlap);
                barHeight = Math.Max(1, slot / Math.Max(1, denominator));
                barStep = barHeight * (1 + negativeOverlapGap - positiveOverlap);
                groupHeight = barHeight + Math.Max(0, series.Count - 1) * barStep;
            }
            var zeroX = X(0);

            for (var categoryIndex = 0; categoryIndex < categoryCount; categoryIndex++)
            {
                var groupTop = plotTop + categoryIndex * slot + (slot - groupHeight) / 2;
                AddSvgText(svg, plotLeft - 6, plotTop + (categoryIndex + 0.5) * slot,
                    categories[categoryIndex], "end", categoryFontSize, true);
                double positive = 0, negative = 0;
                for (var seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
                {
                    if (!series[seriesIndex].Values.TryGetValue(categoryIndex, out var value))
                        continue;
                    double left, barWidth;
                    if (stacked)
                    {
                        var from = value >= 0 ? positive : negative;
                        var to = from + value;
                        left = Math.Min(X(from), X(to));
                        barWidth = Math.Max(0.5, Math.Abs(X(from) - X(to)));
                        if (value >= 0)
                            positive = to;
                        else
                            negative = to;
                    }
                    else
                    {
                        var valueX = X(value);
                        left = Math.Min(zeroX, valueX);
                        barWidth = Math.Max(0.5, Math.Abs(zeroX - valueX));
                    }
                    svg.Add(ChartBar(left, groupTop + seriesIndex * barStep, barWidth, barHeight,
                        series[seriesIndex], seriesIndex, categoryIndex, value));
                }
            }
        }

        private static void RenderLineOrAreaChart(XElement svg, double width, double height, string title,
            bool showLegend, IReadOnlyList<CachedChartSeries> series, IReadOnlyList<string> categories,
            int categoryCount, bool area, bool stacked, ChartScale scale,
            double categoryFontSize, double valueFontSize, string axisSuffix)
        {
            var plotLeft = Math.Max(22,
                10 + (FormatAxisValue(scale.Maximum, scale.MajorUnit) + axisSuffix).Length * 6);
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
                AddSvgText(svg, plotLeft - 5, y, FormatAxisValue(value, scale.MajorUnit) + axisSuffix,
                    "end", valueFontSize, true);
            }

            var slot = plotWidth / categoryCount;
            double X(int index) => plotLeft + (index + 0.5) * slot;

            // A dense category axis (e.g. daily data) cannot fit every label; thin to what the
            // plot width can hold given the widest label.
            var labelWidth = Math.Max(30,
                categories.DefaultIfEmpty(string.Empty).Max(item => item?.Length ?? 0) * 6 + 12);
            var labelStep = Math.Max(1, (int)Math.Ceiling(categoryCount * labelWidth / plotWidth));
            for (var index = 0; index < categoryCount; index += labelStep)
                AddSvgText(svg, X(index), plotBottom + 15, categories[index], "middle",
                    categoryFontSize);

            var baseline = new double[categoryCount];
            for (var seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
            {
                var item = series[seriesIndex];
                double ValueAt(int index) => item.Values.TryGetValue(index, out var v) ? v : 0;
                if (area)
                {
                    var points = new List<string>();
                    for (var index = 0; index < categoryCount; index++)
                        points.Add(Format(X(index)) + "," + Format(Y(baseline[index] + ValueAt(index))));
                    for (var index = categoryCount - 1; index >= 0; index--)
                        points.Add(Format(X(index)) + "," + Format(Y(stacked ? baseline[index] : 0)));
                    svg.Add(new XElement(Svg + "polygon",
                        new XAttribute("class", "docx-chart-area"),
                        new XAttribute("data-chart-series", seriesIndex),
                        new XAttribute("points", string.Join(" ", points)),
                        new XAttribute("fill", item.Color),
                        stacked ? null : new XAttribute("fill-opacity", "0.75"),
                        new XAttribute("stroke", "#FFFFFF"),
                        new XAttribute("stroke-width", "1")));
                }
                else
                {
                    // Word's default for a blank point is a gap; without stacking, simply skip it.
                    var points = Enumerable.Range(0, categoryCount)
                        .Where(index => stacked || item.Values.ContainsKey(index))
                        .Select(index => Format(X(index)) + "," +
                                         Format(Y(baseline[index] + ValueAt(index))))
                        .ToList();
                    if (points.Count > 0)
                        svg.Add(new XElement(Svg + "polyline",
                            new XAttribute("class", "docx-chart-line"),
                            new XAttribute("data-chart-series", seriesIndex),
                            new XAttribute("points", string.Join(" ", points)),
                            new XAttribute("fill", "none"),
                            new XAttribute("stroke", item.Color),
                            new XAttribute("stroke-width", "2"),
                            new XAttribute("stroke-linejoin", "round"),
                            new XAttribute("stroke-linecap", "round")));
                }

                if (stacked)
                    for (var index = 0; index < categoryCount; index++)
                        baseline[index] += ValueAt(index);
            }
        }

        private static IReadOnlyList<string> RenderPieChart(XElement svg, double width, double height,
            string title, bool showLegend, XElement plot, CachedChartSeries series, int categoryCount,
            bool doughnut, ThemeColorScheme theme)
        {
            var plotTop = string.IsNullOrWhiteSpace(title) ? 14 : 60;
            var plotBottom = height - (showLegend ? 54 : 20);
            var centerX = width / 2;
            var centerY = (plotTop + plotBottom) / 2;
            var radius = Math.Max(8, Math.Min(width - 40, plotBottom - plotTop) / 2 - 4);
            var innerRadius = doughnut
                ? radius * ReadIntAttribute(plot.Element(C.holeSize), "val", 50, 10, 90) / 100.0
                : 0;

            var colors = ReadPieSliceColors(plot.Elements(C.ser).First(), categoryCount, theme);
            var total = 0.0;
            for (var index = 0; index < categoryCount; index++)
                if (series.Values.TryGetValue(index, out var value))
                    total += Math.Abs(value);
            if (total <= 0)
                return colors;

            // Slices run clockwise from 12 o'clock (plus any explicit first-slice rotation),
            // matching Word. Negative values plot at their absolute size, also matching Word.
            var angle = (ReadIntAttribute(plot.Element(C.firstSliceAng), "val", 0, 0, 360) - 90)
                * Math.PI / 180;
            for (var index = 0; index < categoryCount; index++)
            {
                if (!series.Values.TryGetValue(index, out var value) || value == 0)
                    continue;
                var sweep = Math.Min(Math.Abs(value) / total * Math.PI * 2, Math.PI * 2 - 1e-4);
                svg.Add(PieSlice(centerX, centerY, radius, innerRadius, angle, angle + sweep,
                    colors[index], index, value));
                angle += sweep;
            }
            return colors;
        }

        private static XElement PieSlice(double centerX, double centerY, double outerRadius,
            double innerRadius, double start, double end, string color, int categoryIndex, double value)
        {
            var largeArc = end - start > Math.PI ? 1 : 0;
            string Point(double radius, double angle) =>
                Format(centerX + radius * Math.Cos(angle)) + " " +
                Format(centerY + radius * Math.Sin(angle));
            var path = innerRadius > 0
                ? $"M {Point(outerRadius, start)} " +
                  $"A {Format(outerRadius)} {Format(outerRadius)} 0 {largeArc} 1 {Point(outerRadius, end)} " +
                  $"L {Point(innerRadius, end)} " +
                  $"A {Format(innerRadius)} {Format(innerRadius)} 0 {largeArc} 0 {Point(innerRadius, start)} Z"
                : $"M {Format(centerX)} {Format(centerY)} " +
                  $"L {Point(outerRadius, start)} " +
                  $"A {Format(outerRadius)} {Format(outerRadius)} 0 {largeArc} 1 {Point(outerRadius, end)} Z";
            return new XElement(Svg + "path",
                new XAttribute("class", "docx-chart-slice"),
                new XAttribute("data-chart-category", categoryIndex),
                new XAttribute("data-chart-value", value.ToString("R", CultureInfo.InvariantCulture)),
                new XAttribute("d", path),
                new XAttribute("fill", color),
                new XAttribute("stroke", "#FFFFFF"),
                new XAttribute("stroke-width", "1"));
        }

        private static IReadOnlyList<string> ReadPieSliceColors(XElement series, int categoryCount,
            ThemeColorScheme theme)
        {
            var explicitColors = series.Elements(C.dPt)
                .Select(point => new
                {
                    Index = ReadIntAttribute(point.Element(C.idx), "val", -1, -1, int.MaxValue),
                    Properties = point.Element(C.spPr),
                })
                .Where(point => point.Index >= 0 && point.Properties != null)
                .GroupBy(point => point.Index)
                .ToDictionary(group => group.Key, group => group.First().Properties);
            return Enumerable.Range(0, categoryCount)
                .Select(index => ReadChartColor(
                    explicitColors.TryGetValue(index, out var properties) ? properties : null,
                    index, theme))
                .ToList();
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
            IReadOnlyList<(string Name, string Color)> items, double fontSize)
        {
            const double interItemGap = 7.5;
            var widths = items
                .Select(item => 20.5 + Math.Max(1, item.Name?.Length ?? 1) * fontSize * 0.375)
                .ToList();
            var total = widths.Sum() - interItemGap;
            var x = Math.Max(4, (width - total) / 2);
            var y = height - 11;
            for (var index = 0; index < items.Count; index++)
            {
                svg.Add(new XElement(Svg + "rect",
                    new XAttribute("x", Format(x)),
                    new XAttribute("y", Format(y - 4)),
                    new XAttribute("width", "8"),
                    new XAttribute("height", "8"),
                    new XAttribute("fill", items[index].Color)));
                AddSvgText(svg, x + 12, y, items[index].Name, "start", fontSize, true);
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
            var cache = firstSeries.Element(C.cat);
            // A date axis caches serial day numbers; surface them as dates, the way every
            // renderer with the embedded workbook would.
            var formatCode = cache?.Descendants(C.formatCode).FirstOrDefault()?.Value;
            var isDate = formatCode != null && formatCode.Contains('y') &&
                (cache.Descendants(C.numCache).Any() || cache.Descendants(C.numLit).Any());
            var points = cache?.Descendants(C.pt)
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
                .Select(index =>
                {
                    if (!points.TryGetValue(index, out var value) || string.IsNullOrWhiteSpace(value))
                        return (index + 1).ToString(CultureInfo.InvariantCulture);
                    if (isDate && double.TryParse(value, NumberStyles.Float,
                            CultureInfo.InvariantCulture, out var serial) &&
                        serial is >= 1 and <= 2958465)
                        return new DateTime(1899, 12, 30).AddDays(serial)
                            .ToString("M/d/yyyy", CultureInfo.InvariantCulture);
                    return value;
                })
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
            // A filled shape (bar, area, slice) colors via a:solidFill; a line series carries its
            // color on the stroke (a:ln/a:solidFill) instead.
            var fill = shapeProperties?.Element(A.solidFill)
                ?? shapeProperties?.Element(A.ln)?.Element(A.solidFill);
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
