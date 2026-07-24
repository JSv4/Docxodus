#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Consolidation regressions for an inline content-control envelope that differs between reviewers.  Visible
/// text alone cannot distinguish these candidates, so the assertions compare the complete inline XML tree after
/// applying the selected policy and after rejecting every revision.
/// </summary>
public class IrInlineEnvelopeCompositeTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    [Fact]
    public void Consolidate_DivergentInlineSdtEnvelopeChanges_BaseWins_UsesExactBaseEnvelope() =>
        AssertDivergentInlineSdtEnvelopePolicy(ConflictResolution.BaseWins);

    [Fact]
    public void Consolidate_DivergentInlineSdtEnvelopeChanges_FirstReviewerWins_UsesExactFirstEnvelope() =>
        AssertDivergentInlineSdtEnvelopePolicy(ConflictResolution.FirstReviewerWins);

    [Fact]
    public void Consolidate_DivergentInlineSdtEnvelopeChanges_StackAll_KeepsBothEnvelopesInReviewerOrder()
    {
        var baseDoc = Doc(Sdt("base"));
        var alice = Doc(Sdt("alice"));
        var bob = Doc(Sdt("bob"));
        var reviewers = new[]
        {
            new DocxDiffReviewer { Author = "Alice", Document = alice },
            new DocxDiffReviewer { Author = "Bob", Document = bob },
        };
        var settings = new DocxDiffConsolidateSettings { ConflictResolution = ConflictResolution.StackAll };

        var merged = DocxDiff.Consolidate(baseDoc, reviewers, settings);
        AssertSchemaValid(merged);
        Assert.Single(DocxDiff.GetConflicts(baseDoc, reviewers, settings));

        AssertBodyInlineTreeEqual(Doc(Sdt("alice") + Sdt("bob")), RevisionProcessor.AcceptRevisions(merged));
        AssertBodyInlineTreeEqual(baseDoc, RevisionProcessor.RejectRevisions(merged));
    }

    private static void AssertDivergentInlineSdtEnvelopePolicy(ConflictResolution policy)
    {
        var baseDoc = Doc(Sdt("base"));
        var alice = Doc(Sdt("alice"));
        var bob = Doc(Sdt("bob"));
        var reviewers = new[]
        {
            new DocxDiffReviewer { Author = "Alice", Document = alice },
            new DocxDiffReviewer { Author = "Bob", Document = bob },
        };
        var settings = new DocxDiffConsolidateSettings { ConflictResolution = policy };

        var merged = DocxDiff.Consolidate(baseDoc, reviewers, settings);
        AssertSchemaValid(merged);
        Assert.Single(DocxDiff.GetConflicts(baseDoc, reviewers, settings));

        var expectedAccepted = policy == ConflictResolution.BaseWins ? baseDoc : alice;
        AssertInlineTreeEqual(expectedAccepted, RevisionProcessor.AcceptRevisions(merged));
        AssertInlineTreeEqual(baseDoc, RevisionProcessor.RejectRevisions(merged));

        var revisions = DocxDiff.GetConsolidatedRevisions(baseDoc, reviewers, settings);
        if (policy == ConflictResolution.BaseWins)
            Assert.Empty(revisions);
        else
        {
            Assert.Contains(revisions, revision => revision.Type == DocxDiffRevisionType.Deleted &&
                revision.Author == "Alice" && revision.Text == "controlled" && revision.ConflictId is not null);
            Assert.Contains(revisions, revision => revision.Type == DocxDiffRevisionType.Inserted &&
                revision.Author == "Alice" && revision.Text == "controlled" && revision.ConflictId is not null);
        }

        var json = DocxDiff.GetConsolidatedEditScriptJson(baseDoc, reviewers, settings);
        Assert.Contains("\"conflicts\"", json);
        if (policy == ConflictResolution.FirstReviewerWins)
            Assert.Contains("\"requiresWholeParagraphReplace\"", json);
    }

    [Fact]
    public void Consolidate_IdenticalInlineSdtEnvelopeChanges_UsesStructuralConsensus()
    {
        var baseDoc = Doc(Sdt("base"));
        var result = Doc(Sdt("agreed"));
        var reviewers = new[]
        {
            new DocxDiffReviewer { Author = "Alice", Document = result },
            new DocxDiffReviewer { Author = "Bob", Document = result },
        };

        var merged = DocxDiff.Consolidate(baseDoc, reviewers);
        AssertSchemaValid(merged);
        Assert.Empty(DocxDiff.GetConflicts(baseDoc, reviewers));
        AssertInlineTreeEqual(result, RevisionProcessor.AcceptRevisions(merged));
        AssertInlineTreeEqual(baseDoc, RevisionProcessor.RejectRevisions(merged));
        Assert.Contains("\"requiresWholeParagraphReplace\"",
            DocxDiff.GetConsolidatedEditScriptJson(baseDoc, reviewers));
    }

    private static WmlDocument Doc(string body) => IrTestDocuments.FromBodyXml(body);

    private static string Sdt(string tag) =>
        "<w:p><w:sdt><w:sdtPr><w:tag w:val=\"" + tag + "\"/></w:sdtPr>" +
        "<w:sdtContent><w:r><w:t>controlled</w:t></w:r></w:sdtContent></w:sdt></w:p>";

    private static void AssertSchemaValid(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(wdoc)
            .Select(error => $"{error.Id}@{error.Path?.XPath}: {error.Description}")
            .ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors));
    }

    private static void AssertInlineTreeEqual(WmlDocument expected, WmlDocument actual) =>
        Assert.True(XNode.DeepEquals(InlineTree(expected), InlineTree(actual)),
            $"expected: {InlineTree(expected)}\nactual: {InlineTree(actual)}");

    private static void AssertBodyInlineTreeEqual(WmlDocument expected, WmlDocument actual) =>
        Assert.True(XNode.DeepEquals(BodyInlineTree(expected), BodyInlineTree(actual)),
            $"expected: {BodyInlineTree(expected)}\nactual: {BodyInlineTree(actual)}");

    private static XElement InlineTree(WmlDocument doc)
    {
        var paragraph = MainXml(doc).Root!.Element(W + "body")!.Elements(W + "p").Single();
        return new XElement("inline",
            paragraph.Elements().Where(element => element.Name != W + "pPr").Select(CloneWithoutNoise));
    }

    private static XElement BodyInlineTree(WmlDocument doc) =>
        new("bodyInline", MainXml(doc).Root!.Element(W + "body")!.Elements(W + "p")
            .Select(paragraph => new XElement("paragraph",
                paragraph.Elements().Where(element => element.Name != W + "pPr").Select(CloneWithoutNoise))));

    private static XElement CloneWithoutNoise(XElement source)
    {
        var clone = new XElement(source);
        foreach (var attribute in clone.DescendantsAndSelf().Attributes().Where(attribute =>
                     attribute.Name.LocalName.StartsWith("rsid", System.StringComparison.Ordinal) ||
                     attribute.Name.NamespaceName == "http://powertools.codeplex.com/2011" ||
                     attribute.Name.NamespaceName == "http://powertools.codeplex.com/documentbuilder/2011").ToList())
            attribute.Remove();
        return clone;
    }

    private static XDocument MainXml(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        using var reader = new StreamReader(wdoc.MainDocumentPart!.GetStream());
        return XDocument.Parse(reader.ReadToEnd());
    }
}
