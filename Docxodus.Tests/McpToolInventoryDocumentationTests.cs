using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests
{
    /// <summary>
    /// Issue #673: <c>tools/mcp-server/README.md</c> is the quick-start contract an MCP host reads
    /// before it ever calls <c>tools/list</c>, and it had drifted — advertising eighteen
    /// grouped-intent tools against a catalog of twenty-two, and omitting five shipped tools
    /// entirely, which makes real capabilities undiscoverable.
    ///
    /// <para>Nothing tied the prose to <see cref="ToolCatalog.Tools"/>, so nothing could notice.
    /// These tests are that tie: the README's tool table and the advertised catalog must name
    /// exactly the same tools, and the count the README states must be the live one.</para>
    /// </summary>
    public class McpToolInventoryDocumentationTests
    {
        private static readonly FileInfo ReadmeFile =
            new FileInfo("../../../../tools/mcp-server/README.md");

        private static string Readme()
        {
            Assert.True(ReadmeFile.Exists, $"expected the MCP server README at {ReadmeFile.FullName}");
            return File.ReadAllText(ReadmeFile.FullName);
        }

        /// <summary>
        /// Tool names in the README's inventory table: the backticked first cell of every row whose
        /// name starts with the server's tool prefix. Prose mentions elsewhere in the file are not
        /// table rows and are deliberately not collected.
        /// </summary>
        private static List<string> DocumentedTools() =>
            Readme().Split('\n')
                .Select(line => Regex.Match(line.Trim(), @"^\|\s*`(docxodus_[a-z_]+)`\s*\|"))
                .Where(m => m.Success)
                .Select(m => m.Groups[1].Value)
                .ToList();

        [Fact]
        public void TheReadmeTableIsDiscoverable()
        {
            // Without this, an extraction regex that stops matching would make the comparisons
            // below pass by comparing the catalog against nothing.
            Assert.NotEmpty(DocumentedTools());
            Assert.NotEmpty(ToolCatalog.Tools);
        }

        [Fact]
        public void EveryAdvertisedToolIsDocumented()
        {
            var undocumented = ToolCatalog.Tools.Select(t => t.Name)
                .Except(DocumentedTools(), StringComparer.Ordinal)
                .ToList();

            Assert.True(
                undocumented.Count == 0,
                "tools/list advertises tools the MCP README does not document: " +
                string.Join(", ", undocumented));
        }

        [Fact]
        public void EveryDocumentedToolIsAdvertised()
        {
            var phantom = DocumentedTools()
                .Except(ToolCatalog.Tools.Select(t => t.Name), StringComparer.Ordinal)
                .ToList();

            Assert.True(
                phantom.Count == 0,
                "the MCP README documents tools the server does not advertise: " +
                string.Join(", ", phantom));
        }

        [Fact]
        public void TheReadmeTableListsEachToolExactlyOnce()
        {
            var duplicates = DocumentedTools()
                .GroupBy(name => name, StringComparer.Ordinal)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key)
                .ToList();

            Assert.True(
                duplicates.Count == 0,
                "the MCP README's tool table lists these more than once: " + string.Join(", ", duplicates));
        }

        [Fact]
        public void TheReadmeStatesTheLiveToolCount()
        {
            // The stale count ("three lifecycle tools plus eighteen grouped-intent tools") was the
            // most visible symptom, and a count written in prose cannot be derived from the table.
            Assert.Contains($"**{ToolCatalog.Tools.Count} tools.**", Readme(), StringComparison.Ordinal);
        }

        [Fact]
        public void EveryDocumentedToolCarriesAKindAndAPurpose()
        {
            var kinds = new[] { "lifecycle", "read", "grouped-intent", "sessionless" };

            var rows = Readme().Split('\n')
                .Select(line => Regex.Match(line.Trim(), @"^\|\s*`(docxodus_[a-z_]+)`\s*\|([^|]*)\|(.*)\|$"))
                .Where(m => m.Success)
                .ToList();

            Assert.Equal(ToolCatalog.Tools.Count, rows.Count);
            foreach (var row in rows)
            {
                var tool = row.Groups[1].Value;
                var kind = row.Groups[2].Value.Trim();
                var purpose = row.Groups[3].Value.Trim();

                Assert.True(kinds.Contains(kind), $"{tool} has an unrecognised kind '{kind}'");
                Assert.True(purpose.Length > 20, $"{tool} has no meaningful purpose text");
            }
        }
    }
}
