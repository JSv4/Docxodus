// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    /// <summary>
    /// The oracle for <c>WordprocessingMLUtil.CanCoalesceAdjacent</c>.
    ///
    /// Run coalescing used to group a container's children by a STRING key built per child — the
    /// child's kind, plus its run properties serialized with <c>XElement.ToString</c>. Building
    /// those strings was the single most expensive step of preparing a document for rendering,
    /// and on a paragraph whose runs all differ every string was built only to be found unequal.
    /// The rule is now a pairwise predicate that compares run properties as trees.
    ///
    /// A refactor of a grouping rule is the kind that fails silently: the output stays
    /// well-formed and merely merges a little more, or a little less, than it did. So this keeps
    /// the old key as an executable oracle and requires the two rules to agree on every adjacent
    /// pair in every run container of the committed corpus — paragraphs, hyperlinks, content
    /// control bodies, and the tracked insertions and deletions where the two could most
    /// plausibly diverge, since those keyed on author, date and id as well as properties.
    ///
    /// Pairs rather than partitions, deliberately. Both groupings chain adjacent mergeable
    /// children into one group, so equal pairwise answers give the same partition; comparing
    /// partitions directly would also trip over an artifact that changes nothing — the old key
    /// put a run of consecutive UNMERGEABLE children in one group and emitted them untouched,
    /// where the predicate leaves each on its own.
    /// </summary>
    public class CoalesceGroupingTests
    {
        private static readonly DirectoryInfo TestFilesDir = new DirectoryInfo("../../../../TestFiles/");

        private const string DontConsolidate = "DontConsolidate";

        private static string Concat(IEnumerable<string> parts) => string.Concat(parts);

        /// <summary>The pre-refactor grouping key, verbatim in behaviour.</summary>
        private static string LegacyKey(XElement ce)
        {
            if (ce.Name == W.r)
            {
                if (ce.Elements().Count(e => e.Name != W.rPr) != 1) return DontConsolidate;
                if (ce.Attribute(PtOpenXml.AbstractNumId) != null) return DontConsolidate;

                XElement? rPr = ce.Element(W.rPr);
                string rPrString = rPr != null ? rPr.ToString(SaveOptions.None) : string.Empty;

                if (ce.Element(W.t) != null) return "Wt" + rPrString;
                if (ce.Element(W.instrText) != null) return "WinstrText" + rPrString;
                return DontConsolidate;
            }

            if (ce.Name == W.ins)
            {
                if (ce.Elements(W.del).Any()) return DontConsolidate;
                if (ce.Elements().Elements().Count(e => e.Name != W.rPr) != 1
                    || !ce.Elements().Elements(W.t).Any())
                    return DontConsolidate;

                XAttribute? dateIns = ce.Attribute(W.date);
                return "Wins2"
                       + ((string?)ce.Attribute(W.author) ?? string.Empty)
                       + (dateIns != null ? ((DateTime)dateIns).ToString("s") : string.Empty)
                       + (string?)ce.Attribute(W.id)
                       + Concat(ce.Elements().Elements(W.rPr).Select(rPr => rPr.ToString(SaveOptions.None)));
            }

            if (ce.Name == W.del)
            {
                if (ce.Elements(W.r).Elements().Count(e => e.Name != W.rPr) != 1
                    || !ce.Elements().Elements(W.delText).Any())
                    return DontConsolidate;

                XAttribute? dateDel = ce.Attribute(W.date);
                return "Wdel"
                       + ((string?)ce.Attribute(W.author) ?? string.Empty)
                       + (dateDel != null ? ((DateTime)dateDel).ToString("s") : string.Empty)
                       + Concat(ce.Elements(W.r).Elements(W.rPr).Select(rPr => rPr.ToString(SaveOptions.None)));
            }

            return DontConsolidate;
        }

        private static bool LegacySaysMergeable(XElement first, XElement second)
        {
            string key = LegacyKey(first);
            return key != DontConsolidate && key == LegacyKey(second);
        }

        [Fact]
        public void CG001_TheAdjacencyPredicateAgreesWithTheKeyItReplaced()
        {
            List<FileInfo> files = TestFilesDir.GetFiles("*.docx", SearchOption.AllDirectories)
                .Where(f => !f.Name.StartsWith("~$", StringComparison.Ordinal))
                .OrderBy(f => f.FullName, StringComparer.Ordinal)
                .ToList();
            Assert.NotEmpty(files);

            int containers = 0;
            int pairs = 0;
            int mergeable = 0;

            foreach (FileInfo file in files)
            {
                byte[] bytes;
                try { bytes = File.ReadAllBytes(file.FullName); }
                catch (IOException) { continue; }

                using var stream = new MemoryStream();
                stream.Write(bytes, 0, bytes.Length);
                stream.Position = 0;

                WordprocessingDocument doc;
                try { doc = WordprocessingDocument.Open(stream, false); }
                catch (Exception) { continue; }   // not every committed fixture is a readable package

                using (doc)
                {
                    XElement? root = doc.MainDocumentPart?.GetXDocument().Root;
                    if (root == null) continue;

                    foreach (XElement container in root.DescendantsAndSelf())
                    {
                        List<XElement> children = container.Elements().ToList();
                        if (!children.Any(c => c.Name == W.r || c.Name == W.ins || c.Name == W.del)) continue;
                        containers++;

                        for (int i = 1; i < children.Count; i++)
                        {
                            XElement a = children[i - 1], b = children[i];
                            bool legacy = LegacySaysMergeable(a, b);
                            bool current = WordprocessingMLUtil.CanCoalesceAdjacent(a, b);
                            Assert.True(
                                legacy == current,
                                $"{file.Name}: <{a.Name.LocalName}> then <{b.Name.LocalName}> — "
                                + $"key rule says {legacy}, predicate says {current}");
                            pairs++;
                            if (legacy) mergeable++;
                        }
                    }
                }
            }

            // A corpus that stopped containing adjacent runs would make the loop above vacuous,
            // and one where nothing ever merges would not exercise the agreeing half of the rule.
            Assert.True(containers > 500, $"only {containers} run containers scanned");
            Assert.True(pairs > 5000, $"only {pairs} adjacent pairs compared");
            Assert.True(mergeable > 100, $"only {mergeable} mergeable pairs found");
        }
    }
}

#endif
