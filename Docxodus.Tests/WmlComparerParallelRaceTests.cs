#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Regression for issue #153 — <see cref="WmlComparer.AddFootnotesEndnotesParts"/>
/// (and the matching path in <see cref="DocumentBuilder"/>)
/// was building w:footnotes / w:endnotes / w:numbering / w:comments roots from a
/// shared static <c>XAttribute[]</c>. <see cref="System.Xml.Linq.XAttribute"/> can
/// only have one parent, so two parallel callers would race and one would throw
/// <c>"Duplicate attribute"</c> or leave the other thread's element half-built
/// (the failure mode that surfaced as flaky CI on PR #152).
/// </summary>
public class WmlComparerParallelRaceTests
{
    /// <summary>Two minimal text-only documents that differ in one paragraph — just enough
    /// for WmlComparer.Compare() to do real work, including the AddFootnotesEndnotesParts path.</summary>
    private static (WmlDocument source1, WmlDocument source2) BuildTinyComparePair(int salt)
    {
        return (Build($"Source A run #{salt}."), Build($"Source B run #{salt}."));

        static WmlDocument Build(string text)
        {
            using var ms = new MemoryStream();
            using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = wDoc.AddMainDocumentPart();
                main.Document = new Document();
                var body = new Body();
                main.Document.Body = body;
                main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
                main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
                body.Append(new Paragraph(new Run(new Text(text))));
                main.Document.Save();
            }
            return new WmlDocument("tiny.docx", ms.ToArray());
        }
    }

    [Fact]
    public async Task WC_Race001_AddFootnotesEndnotesParts_ParallelCompares_DoNotThrow()
    {
        // Run 16 concurrent WmlComparer.Compare() calls. Before the fix this would
        // throw "Duplicate attribute" or "Unexpected end of file" from at least one
        // worker within a few iterations on a multi-core box.
        const int workers = 16;
        var settings = new WmlComparerSettings();

        var tasks = Enumerable.Range(0, workers).Select(i => Task.Run(() =>
        {
            var (s1, s2) = BuildTinyComparePair(i);
            // Just exercise the path that touches AddFootnotesEndnotesParts.
            return WmlComparer.Compare(s1, s2, settings);
        })).ToArray();

        var results = await Task.WhenAll(tasks);
        Assert.All(results, r => Assert.NotNull(r));
    }
}
