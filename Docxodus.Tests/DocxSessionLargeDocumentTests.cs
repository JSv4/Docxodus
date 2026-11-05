using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

[CollectionDefinition("Large-document timing", DisableParallelization = true)]
public class LargeDocumentTimingCollection;

/// <summary>
/// Operations that look up every anchor of a document found each one by walking its whole part, so
/// sixteen times the paragraphs took about 256 times as long (issue #1389). Resolving them in one
/// pass per part takes about sixteen times as long.
/// </summary>
[Collection("Large-document timing")]
public class DocxSessionLargeDocumentTimingTests
{
    [Fact]
    public void Grep_TimeGrowsLinearlyWithParagraphCount() =>
        AssertLinear(session => Assert.Single(session.Grep("Clause 0 ")));

    [Fact]
    public void GrepCrossBlock_TimeGrowsLinearlyWithParagraphCount() =>
        AssertLinear(session => Assert.Single(session.GrepCrossBlock("Clause 0 ")));

    [Fact]
    public void GetBlockMetadatas_TimeGrowsLinearlyWithParagraphCount() =>
        AssertLinear(session => session.GetBlockMetadatas(session.Project().AnchorIndex.Keys));

    [Fact]
    public void GetAnchorInfos_TimeGrowsLinearlyWithParagraphCount() =>
        AssertLinear(session => session.GetAnchorInfos(session.Project().AnchorIndex.Keys));

    [Fact]
    public void StampedBlockRender_TimeGrowsLinearlyWithParagraphCount() =>
        AssertLinear(session => HtmlConversionOps.RenderBlockHtml(session,
            session.Project().AnchorIndex.Keys.First(), new HtmlConversionOptions { StampAnchors = true }));

    private static void AssertLinear(Action<DocxSession> operation)
    {
        double Seconds(int paragraphs)
        {
            using var session = new DocxSession(AnchorResolutionDocuments.Build(Enumerable.Range(0, paragraphs)
                .Select(i => AnchorResolutionDocuments.Paragraph($"Clause {i} sets out the terms that apply."))));
            session.Project();
            var best = double.MaxValue;
            for (int run = 0; run < 5; run++)
            {
                // A collection left over from building the document must not land in a timed run.
                GC.Collect();
                GC.WaitForPendingFinalizers();
                var stopwatch = Stopwatch.StartNew();
                operation(session);
                best = Math.Min(best, stopwatch.Elapsed.TotalSeconds);
            }
            return best;
        }

        Seconds(200); // warm the code paths both sizes share
        var small = Seconds(1000);
        var large = Seconds(16000);
        Assert.True(large < 80 * small, $"{small:F4}s over 1,000 paragraphs, {large:F4}s over 16,000");
    }
}

public class DocxSessionAnchorResolutionTests
{
    [Fact]
    public void ResolveAll_GivesTheFirstElementCarryingTheUnid_AsResolveDoes()
    {
        // Two paragraphs that carry the same Unid (a copied block keeps its attributes): both lookups
        // must return the first in document order.
        var first = AnchorResolutionDocuments.Paragraph("First copy");
        var second = AnchorResolutionDocuments.Paragraph("Second copy");
        first.SetAttributeValue(PtOpenXml.Unid, "shared");
        second.SetAttributeValue(PtOpenXml.Unid, "shared");
        using var stream = new MemoryStream();
        stream.Write(AnchorResolutionDocuments.Build(new[] { first, second }));
        using var document = WordprocessingDocument.Open(stream, true);
        var main = document.MainDocumentPart!;
        var element = main.GetXDocument().Descendants(W.p).First();
        var target = WmlToMarkdownConverter.CreateAnchorTarget(element, "p:body:shared", "p", "body", "shared",
            main.Uri.ToString(), enrich: false, document);

        var resolved = Assert.Single(AnchorTarget.ResolveAll(document, new[] { target })).Value;
        Assert.Same(target.Resolve(document), resolved);
        Assert.Equal("First copy", resolved.Value);
    }
}

internal static class AnchorResolutionDocuments
{
    internal static XElement Paragraph(string text) => new(W.p, new XElement(W.r, new XElement(W.t, text)));

    internal static byte[] Build(IEnumerable<XElement> blocks)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles)));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
            main.PutXDocument(new XDocument(new XElement(W.document,
                new XAttribute(XNamespace.Xmlns + "w", W.w), new XAttribute(XNamespace.Xmlns + "r", R.r),
                new XElement(W.body, blocks))));
        }
        return stream.ToArray();
    }
}
