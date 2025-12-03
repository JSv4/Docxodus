#nullable enable
// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;
using Xunit.Abstractions;

namespace OxPt
{
    /// <summary>
    /// Performance profiling tests for WmlToHtmlConverter.
    /// These tests identify performance bottlenecks by measuring specific operations.
    /// </summary>
    public class ConverterProfilingTests
    {
        private readonly ITestOutputHelper _output;

        public ConverterProfilingTests(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void Profile_ConvertToHtml_BreakdownByPhase()
        {
            _output.WriteLine("=== WmlToHtmlConverter Phase Breakdown Profiling ===\n");

            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");

            // Use a variety of test documents
            var testFiles = new[]
            {
                ("HC001-5DayTourPlanTemplate.docx", "Small - Tour plan"),
                ("HC006-Test-01.docx", "Small - Test document"),
                ("HC031-Complicated-Document.docx", "Medium - Complex formatting"),
                ("HC022-Table-Of-Contents.docx", "Medium - TOC with headings"),
            };

            foreach (var (fileName, description) in testFiles)
            {
                FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, fileName));
                if (!sourceDocx.Exists)
                {
                    _output.WriteLine($"Skipping {fileName} - not found");
                    continue;
                }

                var docBytes = File.ReadAllBytes(sourceDocx.FullName);

                // Warm up
                WarmUp(docBytes);

                // Measure FormattingAssembler phase separately
                var faTime = MeasureFormattingAssembler(docBytes);

                // Measure full conversion
                var (totalTime, paragraphCount, runCount, elementCount) = MeasureFullConversion(docBytes);

                var coreTime = totalTime - faTime;

                _output.WriteLine($"--- {description} ({fileName}) ---");
                _output.WriteLine($"Document Size: {sourceDocx.Length / 1024.0:F1} KB");
                _output.WriteLine($"Elements: {elementCount} | Paragraphs: {paragraphCount} | Runs: {runCount}");
                _output.WriteLine($"Total: {totalTime.TotalMilliseconds:F2} ms");
                _output.WriteLine($"  FormattingAssembler: {faTime.TotalMilliseconds:F2} ms ({faTime / totalTime * 100:F1}%)");
                _output.WriteLine($"  ConvertToHtml Core: {coreTime.TotalMilliseconds:F2} ms ({coreTime / totalTime * 100:F1}%)");
                _output.WriteLine($"  Per element: {totalTime.TotalMilliseconds / elementCount:F4} ms");
                _output.WriteLine("");
            }
        }

        [Fact]
        public void Profile_MultipleConversions_WarmupEffect()
        {
            _output.WriteLine("=== Warmup Effect Analysis (JIT compilation) ===\n");

            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, "HC006-Test-01.docx"));

            if (!sourceDocx.Exists)
            {
                _output.WriteLine("Test file not found");
                return;
            }

            var docBytes = File.ReadAllBytes(sourceDocx.FullName);
            var times = new List<long>();

            // Run 10 conversions
            for (int i = 0; i < 10; i++)
            {
                using var memStream = new MemoryStream();
                memStream.Write(docBytes, 0, docBytes.Length);
                memStream.Position = 0;
                using var wDoc = WordprocessingDocument.Open(memStream, true);

                var sw = Stopwatch.StartNew();
                var settings = new WmlToHtmlConverterSettings
                {
                    FabricateCssClasses = true,
                    CssClassPrefix = "pt-",
                };
                var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                sw.Stop();

                times.Add(sw.ElapsedMilliseconds);
                _output.WriteLine($"Run {i + 1}: {sw.ElapsedMilliseconds} ms");
            }

            _output.WriteLine("");
            _output.WriteLine($"First run (cold): {times[0]} ms");
            _output.WriteLine($"Average (excluding first): {times.Skip(1).Average():F1} ms");
            _output.WriteLine($"Min (warm): {times.Skip(1).Min()} ms");
            _output.WriteLine($"Max (warm): {times.Skip(1).Max()} ms");
            _output.WriteLine($"Warmup overhead: {times[0] - times.Skip(1).Average():F1} ms");
        }

        [Fact]
        public void Profile_CommentRenderingOverhead()
        {
            _output.WriteLine("=== Comment Rendering Overhead ===\n");

            // Find documents with comments
            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var testFiles = new[]
            {
                "HC006-Test-01.docx",
                "HC031-Complicated-Document.docx",
            };

            foreach (var fileName in testFiles)
            {
                FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, fileName));
                if (!sourceDocx.Exists) continue;

                var docBytes = File.ReadAllBytes(sourceDocx.FullName);

                // Warm up
                WarmUp(docBytes);

                // Without comments
                var timeWithoutComments = MeasureWithSettings(docBytes, new WmlToHtmlConverterSettings
                {
                    FabricateCssClasses = true,
                    RenderComments = false,
                });

                // With comments
                var timeWithComments = MeasureWithSettings(docBytes, new WmlToHtmlConverterSettings
                {
                    FabricateCssClasses = true,
                    RenderComments = true,
                });

                var overhead = timeWithComments - timeWithoutComments;
                var overheadPct = timeWithoutComments.TotalMilliseconds > 0
                    ? overhead / timeWithoutComments * 100
                    : 0;

                _output.WriteLine($"{fileName}:");
                _output.WriteLine($"  Without comments: {timeWithoutComments.TotalMilliseconds:F2} ms");
                _output.WriteLine($"  With comments: {timeWithComments.TotalMilliseconds:F2} ms");
                _output.WriteLine($"  Overhead: {overhead.TotalMilliseconds:F2} ms ({overheadPct:F1}%)");
                _output.WriteLine("");
            }
        }

        [Fact]
        public void Profile_TrackedChangesOverhead()
        {
            _output.WriteLine("=== Tracked Changes Rendering Overhead ===\n");

            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/WC/");
            var testFiles = Directory.GetFiles(sourceDir.FullName, "*.docx").Take(5).ToArray();

            foreach (var filePath in testFiles)
            {
                var fileName = Path.GetFileName(filePath);
                var docBytes = File.ReadAllBytes(filePath);

                // Warm up
                WarmUp(docBytes);

                // Without tracked changes rendering
                var timeWithoutRevisions = MeasureWithSettings(docBytes, new WmlToHtmlConverterSettings
                {
                    FabricateCssClasses = true,
                    RenderTrackedChanges = false,
                });

                // With tracked changes rendering
                var timeWithRevisions = MeasureWithSettings(docBytes, new WmlToHtmlConverterSettings
                {
                    FabricateCssClasses = true,
                    RenderTrackedChanges = true,
                });

                var overhead = timeWithRevisions - timeWithoutRevisions;
                var overheadPct = timeWithoutRevisions.TotalMilliseconds > 0
                    ? overhead / timeWithoutRevisions * 100
                    : 0;

                _output.WriteLine($"{fileName}:");
                _output.WriteLine($"  Accept revisions: {timeWithoutRevisions.TotalMilliseconds:F2} ms");
                _output.WriteLine($"  Render revisions: {timeWithRevisions.TotalMilliseconds:F2} ms");
                _output.WriteLine($"  Overhead: {overhead.TotalMilliseconds:F2} ms ({overheadPct:F1}%)");
                _output.WriteLine("");
            }
        }

        [Fact]
        public void Profile_CssClassGeneration()
        {
            _output.WriteLine("=== CSS Class Generation vs Inline Styles ===\n");

            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var testFiles = new[]
            {
                "HC006-Test-01.docx",
                "HC031-Complicated-Document.docx",
            };

            foreach (var fileName in testFiles)
            {
                FileInfo sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, fileName));
                if (!sourceDocx.Exists) continue;

                var docBytes = File.ReadAllBytes(sourceDocx.FullName);

                // Warm up
                WarmUp(docBytes);

                // With CSS classes (FabricateCssClasses = true)
                var timeWithClasses = MeasureWithSettings(docBytes, new WmlToHtmlConverterSettings
                {
                    FabricateCssClasses = true,
                    CssClassPrefix = "pt-",
                });

                // Without CSS classes (inline styles)
                var timeInline = MeasureWithSettings(docBytes, new WmlToHtmlConverterSettings
                {
                    FabricateCssClasses = false,
                });

                var diff = timeWithClasses - timeInline;
                var diffPct = timeInline.TotalMilliseconds > 0
                    ? diff / timeInline * 100
                    : 0;

                _output.WriteLine($"{fileName}:");
                _output.WriteLine($"  With CSS classes: {timeWithClasses.TotalMilliseconds:F2} ms");
                _output.WriteLine($"  Inline styles: {timeInline.TotalMilliseconds:F2} ms");
                _output.WriteLine($"  Difference: {diff.TotalMilliseconds:F2} ms ({diffPct:F1}%)");
                _output.WriteLine("");
            }
        }

        [Fact]
        public void Profile_DocumentScaling()
        {
            _output.WriteLine("=== Document Size Scaling Analysis ===\n");

            DirectoryInfo sourceDir = new DirectoryInfo("../../../../TestFiles/");

            // Get all HC test files and sort by size
            var testFiles = Directory.GetFiles(sourceDir.FullName, "HC*.docx")
                .Select(f => new FileInfo(f))
                .OrderBy(f => f.Length)
                .ToArray();

            var results = new List<(string Name, long Size, int Elements, double Time)>();

            foreach (var file in testFiles.Take(15)) // Test first 15 files
            {
                var docBytes = File.ReadAllBytes(file.FullName);

                // Warm up
                WarmUp(docBytes);

                // Measure with element count
                var (time, _, _, elements) = MeasureFullConversion(docBytes);

                results.Add((file.Name, file.Length, elements, time.TotalMilliseconds));
            }

            // Output results
            foreach (var (name, size, elements, time) in results)
            {
                var perElement = elements > 0 ? time / elements : 0;
                _output.WriteLine($"{name,-40} {size / 1024.0,6:F1} KB | {elements,5} elements | {time,7:F2} ms | {perElement:F4} ms/elem");
            }

            // Calculate correlation
            if (results.Count >= 3)
            {
                var avgTimePerElement = results.Average(r => r.Elements > 0 ? r.Time / r.Elements : 0);
                _output.WriteLine("");
                _output.WriteLine($"Average time per element: {avgTimePerElement:F4} ms");
            }
        }

        #region Helper Methods

        private void WarmUp(byte[] docBytes)
        {
            using var ms = new MemoryStream();
            ms.Write(docBytes, 0, docBytes.Length);
            ms.Position = 0;
            using var wDoc = WordprocessingDocument.Open(ms, true);
            WmlToHtmlConverter.ConvertToHtml(wDoc, new WmlToHtmlConverterSettings { FabricateCssClasses = true });
        }

        private TimeSpan MeasureFormattingAssembler(byte[] docBytes)
        {
            using var ms = new MemoryStream();
            ms.Write(docBytes, 0, docBytes.Length);
            ms.Position = 0;
            using var wDoc = WordprocessingDocument.Open(ms, true);

            var sw = Stopwatch.StartNew();
            var faSettings = new FormattingAssemblerSettings
            {
                CreateHtmlConverterAnnotationAttributes = true,
            };
            FormattingAssembler.AssembleFormatting(wDoc, faSettings);
            sw.Stop();

            return sw.Elapsed;
        }

        private (TimeSpan Time, int Paragraphs, int Runs, int Elements) MeasureFullConversion(byte[] docBytes)
        {
            using var ms = new MemoryStream();
            ms.Write(docBytes, 0, docBytes.Length);
            ms.Position = 0;
            using var wDoc = WordprocessingDocument.Open(ms, true);

            // Count elements
            var mainXDoc = wDoc.MainDocumentPart!.GetXDocument();
            var elements = mainXDoc.Descendants().Count();
            var paragraphs = mainXDoc.Descendants(W.p).Count();
            var runs = mainXDoc.Descendants(W.r).Count();

            // Time conversion
            var sw = Stopwatch.StartNew();
            var settings = new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = true,
                CssClassPrefix = "pt-",
            };
            var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
            sw.Stop();

            return (sw.Elapsed, paragraphs, runs, elements);
        }

        private TimeSpan MeasureWithSettings(byte[] docBytes, WmlToHtmlConverterSettings settings)
        {
            using var ms = new MemoryStream();
            ms.Write(docBytes, 0, docBytes.Length);
            ms.Position = 0;
            using var wDoc = WordprocessingDocument.Open(ms, true);

            var sw = Stopwatch.StartNew();
            var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
            sw.Stop();

            return sw.Elapsed;
        }

        #endregion
    }
}
