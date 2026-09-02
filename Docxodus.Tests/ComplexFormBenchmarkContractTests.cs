using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Xunit;

namespace Docxodus.Tests
{
    /// <summary>
    /// The complex-form benchmark lives outside <c>Docxodus.sln</c> and its README is the
    /// executable contract a reader relies on. Issue #669: the README kept advertising a
    /// legacy-engine stage for a week after that stage was deleted from <c>Program.cs</c>,
    /// because nothing tied the two together. These tests are that tie — they compare the
    /// stage names in the README's table against the <c>Bench(...)</c> labels the harness
    /// actually prints, in both directions.
    /// </summary>
    public class ComplexFormBenchmarkContractTests
    {
        private static readonly DirectoryInfo BenchmarkDir =
            new DirectoryInfo("../../../../benchmarks/complex-form-doc/");

        private static string ReadBenchmarkFile(string name)
        {
            var file = new FileInfo(Path.Combine(BenchmarkDir.FullName, name));
            Assert.True(file.Exists, $"expected the benchmark's {name} at {file.FullName}");
            return File.ReadAllText(file.FullName);
        }

        /// <summary>Stage labels the harness prints, i.e. every <c>Bench("…")</c> first argument.</summary>
        private static List<string> HarnessStages() =>
            Regex.Matches(ReadBenchmarkFile("Program.cs"), @"\bBench\(""([^""]+)""")
                .Select(m => m.Groups[1].Value)
                .ToList();

        /// <summary>
        /// Stage labels the README documents: the backticked first cell of every row in the
        /// "What it measures" table. Header and separator rows carry no code span, so they
        /// drop out without needing to be recognised.
        /// </summary>
        private static List<string> DocumentedStages() =>
            ReadBenchmarkFile("README.md")
                .Split('\n')
                .Select(line => Regex.Match(line.Trim(), @"^\|\s*`([^`]+)`\s*\|"))
                .Where(m => m.Success)
                .Select(m => m.Groups[1].Value)
                .ToList();

        [Fact]
        public void HarnessStagesAreDiscoverable()
        {
            // Guards the two regexes above: if either stops matching, the comparison test
            // would pass vacuously by comparing two empty sets.
            Assert.NotEmpty(HarnessStages());
            Assert.NotEmpty(DocumentedStages());
        }

        [Fact]
        public void EveryDocumentedStageExistsInTheHarness()
        {
            var undocumentedInCode = DocumentedStages().Except(HarnessStages()).ToList();
            Assert.True(
                undocumentedInCode.Count == 0,
                "README documents stages the harness no longer runs: " +
                string.Join(", ", undocumentedInCode));
        }

        [Fact]
        public void EveryHarnessStageIsDocumented()
        {
            var missingFromReadme = HarnessStages().Except(DocumentedStages()).ToList();
            Assert.True(
                missingFromReadme.Count == 0,
                "harness runs stages the README does not list: " +
                string.Join(", ", missingFromReadme));
        }

        [Fact]
        public void ReadmeDoesNotAdvertiseTheRemovedLegacyEngine()
        {
            // WmlComparer was removed from the library in v11.0.0; a README that still names
            // it promises a stage that cannot run.
            Assert.DoesNotContain("WmlComparer", ReadBenchmarkFile("README.md"), StringComparison.Ordinal);
        }
    }
}
