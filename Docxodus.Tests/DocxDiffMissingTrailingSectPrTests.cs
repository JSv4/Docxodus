#nullable enable
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A body-level trailing <c>w:sectPr</c> that exists on only ONE side of the compare. Word treats
/// the absent side as the DEFAULT section: when only the right has one, its section properties are
/// adopted (a two-column right renders two-column) under a <c>w:sectPrChange</c> archiving the
/// default section; when only the left has one, the live properties become the default section and
/// the left's are archived. Accept therefore reproduces the right page setup and reject restores
/// the left's (Word cannot express "no sectPr" in a change archive, so the archived default section
/// stands in for the absent side — Word's own materialization).
/// </summary>
public class DocxDiffMissingTrailingSectPrTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string TwoColumnSectPr =
        "<w:sectPr>" +
        "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>" +
        "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>" +
        "<w:cols w:num=\"2\" w:space=\"720\"/>" +
        "</w:sectPr>";

    private static XElement? TrailingSectPr(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        var body = XDocument.Load(wd.MainDocumentPart!.GetStream()).Root?.Element(W + "body");
        return body?.Elements(W + "sectPr").LastOrDefault();
    }

    private static int? LiveColumnCount(WmlDocument doc)
    {
        var sectPr = TrailingSectPr(doc);
        var cols = sectPr?.Elements(W + "cols").FirstOrDefault();
        return (int?)cols?.Attribute(W + "num");
    }

    [Fact]
    public void Right_only_trailing_sectPr_is_adopted_with_tracked_change()
    {
        var left = IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Shared body text.</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Shared body text revised.</w:t></w:r></w:p>" + TwoColumnSectPr);

        var redline = DocxDiff.Compare(left, right);

        var sectPr = TrailingSectPr(redline);
        Assert.NotNull(sectPr);
        Assert.Equal(2, LiveColumnCount(redline));
        Assert.NotNull(sectPr!.Element(W + "sectPrChange"));

        // Accept keeps the right's page setup; the change marker is gone.
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        Assert.Equal(2, LiveColumnCount(accepted));
        Assert.Null(TrailingSectPr(accepted)?.Element(W + "sectPrChange"));

        // Reject restores the default (single-column) section.
        var rejected = RevisionProcessor.RejectRevisions(redline);
        Assert.NotEqual(2, LiveColumnCount(rejected));
    }

    [Fact]
    public void Left_only_trailing_sectPr_is_replaced_by_default_section_with_tracked_change()
    {
        var left = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Shared body text.</w:t></w:r></w:p>" + TwoColumnSectPr);
        var right = IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Shared body text revised.</w:t></w:r></w:p>");

        var redline = DocxDiff.Compare(left, right);

        var sectPr = TrailingSectPr(redline);
        Assert.NotNull(sectPr);
        // Live properties are the default section (accept ≡ right): no two-column layout.
        Assert.NotEqual(2, LiveColumnCount(redline));
        Assert.NotNull(sectPr!.Element(W + "sectPrChange"));

        var accepted = RevisionProcessor.AcceptRevisions(redline);
        Assert.NotEqual(2, LiveColumnCount(accepted));

        // Reject restores the left's two-column section.
        var rejected = RevisionProcessor.RejectRevisions(redline);
        Assert.Equal(2, LiveColumnCount(rejected));
    }

    [Fact]
    public void Right_only_trailing_sectPr_matching_the_default_section_is_adopted_untracked()
    {
        // The right's explicit properties spell out exactly the default section; Word carries the
        // sectPr but stamps no change (nothing differs from the absent-left default).
        var left = IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Shared body text.</w:t></w:r></w:p>");
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Shared body text revised.</w:t></w:r></w:p>" +
            "<w:sectPr>" +
            "<w:type w:val=\"nextPage\"/>" +
            "<w:pgSz w:w=\"12240\" w:h=\"15840\"/>" +
            "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>" +
            "<w:cols w:num=\"1\" w:space=\"720\"/>" +
            "<w:docGrid w:linePitch=\"360\"/>" +
            "</w:sectPr>");

        var redline = DocxDiff.Compare(left, right);

        var sectPr = TrailingSectPr(redline);
        Assert.NotNull(sectPr);
        Assert.Null(sectPr!.Element(W + "sectPrChange"));
    }
}
