using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

public class TrackedHeaderGroupCompatibilityTests
{
    private const string Code = "render.libreoffice_tracked_header_group";
    private const string Stamp = "w:id=\"7\" w:author=\"Fixture\" w:date=\"2026-11-04T12:00:00Z\"";

    [Theory]
    [InlineData("ins", true)]
    [InlineData("del", true)]
    [InlineData("", false)]
    [InlineData("ins", false, false)]
    [InlineData("ins", false, true, false)]
    [InlineData("ins", false, true, true, false)]
    [InlineData("ins", false, true, true, true, false)]
    public void Identifies_only_the_measured_header_shape(
        string revision, bool expected, bool grouped = true, bool textbox = true,
        bool anchored = true, bool header = true)
    {
        var input = Build(revision, grouped, textbox, anchored, header);
        Assert.Empty(PackageValidation.Errors(input.DocumentByteArray, DocumentFormat.OpenXml.FileFormatVersions.Office2019));
        var bytes = input.DocumentByteArray.ToArray();

        var report = DeliverableVerifier.VerifyDeliverable(input);

        Assert.True(report.AnalysisCompleted);
        Assert.Equal(expected, report.Findings.Any(f => f.Code == Code));
        Assert.Equal(bytes, input.DocumentByteArray);
        if (expected)
        {
            var finding = Assert.Single(report.Findings, f => f.Code == Code);
            Assert.Equal(DeliverableFindingCategory.Render, finding.Category);
            Assert.Equal(VerificationFindingSeverity.Warning, finding.Severity);
            Assert.False(finding.BlocksDelivery);
            Assert.Contains("LibreOffice", finding.Message);
            Assert.Contains("DOCX", finding.Remediation);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Public_comparison_output_reports_risk_without_changing_native_revisions(bool delete)
    {
        var empty = IrTestDocuments.FromBodyAndHeaderXml("<w:p/>", "<w:p/>");
        var drawing = Build("");
        var left = delete ? drawing : empty;
        var right = delete ? empty : drawing;
        var redline = DocxCompare.Compare(left, right);
        var bytes = redline.DocumentByteArray.ToArray();

        using var json = JsonDocument.Parse(VerificationOps.VerifyDeliverable(redline.DocumentByteArray));
        Assert.Contains(json.RootElement.GetProperty("findings").EnumerateArray(),
            finding => finding.GetProperty("code").GetString() == Code);
        Assert.Equal(bytes, redline.DocumentByteArray);
        Assert.NotEmpty(Header(redline).Descendants(delete ? W.del : W.ins));
        Assert.Empty(PackageValidation.Errors(bytes, DocumentFormat.OpenXml.FileFormatVersions.Office2019));

        foreach (var (resolved, expected) in new[]
        {
            (RevisionProcessor.AcceptRevisions(redline), right),
            (RevisionProcessor.RejectRevisions(redline), left),
        })
        {
            Assert.DoesNotContain(DeliverableVerifier.VerifyDeliverable(resolved).Findings, f => f.Code == Code);
            Assert.Equal(Drawings(expected), Drawings(resolved));
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void A_label_edit_keeps_the_group_untracked(bool editLabel)
    {
        var header = string.Concat(Header(Build("")).Elements());
        var original = IrTestDocuments.FromBodyAndHeaderXml("<w:p><w:r><w:t>Plain body.</w:t></w:r></w:p>", header);
        var revised = IrTestDocuments.FromBodyAndHeaderXml(
            "<w:p><w:r><w:t>" + (editLabel ? "Plain" : "Changed") + " body.</w:t></w:r></w:p>",
            editLabel ? header.Replace("Juniper window 0", "Cedar window 0") : header);
        var redline = DocxCompare.Compare(original, revised);
        foreach (var document in new[] { original, revised, redline })
            Assert.Empty(PackageValidation.Errors(document.DocumentByteArray, DocumentFormat.OpenXml.FileFormatVersions.Office2019));
        Assert.DoesNotContain(DeliverableVerifier.VerifyDeliverable(redline).Findings, f => f.Code == Code);
        var drawing = Assert.Single(Header(redline).Descendants(W.drawing));
        Assert.DoesNotContain(drawing.Ancestors(), e => e.Name == W.ins || e.Name == W.del);
        if (editLabel)
        {
            Assert.Single(drawing.Descendants(W.ins));
            Assert.Single(drawing.Descendants(W.del));
        }
        Assert.Equal(Drawings(revised, coalesceRuns: true), Drawings(RevisionProcessor.AcceptRevisions(redline), coalesceRuns: true));
        Assert.Equal(Drawings(original, coalesceRuns: true), Drawings(RevisionProcessor.RejectRevisions(redline), coalesceRuns: true));
    }

    [Theory]
    [InlineData("format")]
    [InlineData("insert-paragraph")]
    [InlineData("delete-paragraph")]
    [InlineData("host-and-label")]
    [InlineData("nested")]
    [InlineData("alternate")]
    [InlineData("table-format")]
    [InlineData("nested-table-format")]
    public void Textbox_edits_preserve_exact_resolved_drawing_trees(string edit)
    {
        var left = Header(Build(""));
        if (edit == "delete-paragraph")
            left.Descendants(W.txbxContent).First().Add(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Extra label."))));
        if (edit is "nested" or "nested-table-format")
        {
            var nested = Header(Build("", grouped: false, anchored: false)).Descendants(W.drawing).Single();
            nested.Descendants(WP.docPr).Single().SetAttributeValue("id", 2);
            var run = new XElement(W.r, nested);
            if (edit == "nested-table-format")
                left.Descendants(W.txbxContent).First().Add(Table(new XElement(W.p, run)));
            else
                left.Descendants(W.txbxContent).First().Element(W.p)!.Add(run);
        }
        if (edit == "table-format")
        {
            var paragraph = left.Elements(W.p).Single();
            paragraph.ReplaceWith(Table(new XElement(paragraph)));
        }
        if (edit == "alternate")
        {
            var drawing = left.Descendants(W.drawing).Single();
            var fallback = new XElement(drawing);
            fallback.Descendants(WP.docPr).Single().SetAttributeValue("id", 2);
            drawing.ReplaceWith(new XElement(MC.AlternateContent,
                new XElement(MC.Choice, new XAttribute("Requires", "wpg"), new XElement(drawing)),
                new XElement(MC.Fallback, fallback)));
        }
        if (edit == "host-and-label")
            left.Elements(W.p).Single().AddFirst(new XElement(W.r, new XElement(W.t, "Outer label.")));
        var right = new XElement(left);
        var body = right.Descendants(W.txbxContent).First();
        switch (edit)
        {
            case "format":
            case "table-format":
                body.Descendants(W.r).First().AddFirst(new XElement(W.rPr, new XElement(W.b)));
                break;
            case "nested-table-format":
                body.Descendants(W.txbxContent).Single().Descendants(W.r).Single()
                    .AddFirst(new XElement(W.rPr, new XElement(W.b)));
                break;
            case "insert-paragraph":
                body.Add(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Extra label."))));
                break;
            case "delete-paragraph":
                body.Elements(W.p).Last().Remove();
                break;
            case "nested":
                body.Descendants(W.txbxContent).Single().Descendants(W.t).Single().Value = "Cedar nested label.";
                break;
            default:
                foreach (var text in right.Descendants(W.t).Where(t => t.Value == "Juniper window 0"))
                    text.Value = "Cedar window 0";
                if (edit == "host-and-label") right.Descendants(W.t).First().Value = "Changed outer label.";
                break;
        }
        WmlDocument Document(XElement header) => IrTestDocuments.FromBodyAndHeaderXml("<w:p/>", string.Concat(header.Elements()));
        var original = Document(left);
        var revised = Document(right);
        var redline = DocxCompare.Compare(original, revised);
        Assert.DoesNotContain(DeliverableVerifier.VerifyDeliverable(redline).Findings, f => f.Code == Code);
        var ids = Header(redline).Descendants().Where(e => e.Name == W.ins || e.Name == W.del || e.Name == W.rPrChange)
            .Select(e => (string?)e.Attribute(W.id)).ToArray();
        Assert.True(ids.Length > 0, DocxDiff.GetEditScriptJson(original, revised));
        Assert.Equal(ids.Length, ids.Distinct().Count());
        foreach (var document in new[] { original, revised, redline })
            Assert.Empty(PackageValidation.Errors(document.DocumentByteArray, DocumentFormat.OpenXml.FileFormatVersions.Office2019));
        foreach (var (actual, expected) in new[]
        {
            (RevisionProcessor.AcceptRevisions(redline), revised),
            (RevisionProcessor.RejectRevisions(redline), original),
        })
        {
            Assert.Empty(PackageValidation.Errors(actual.DocumentByteArray, DocumentFormat.OpenXml.FileFormatVersions.Office2019));
            Assert.Equal(Drawings(expected, coalesceRuns: true), Drawings(actual, coalesceRuns: true));
            Assert.Equal(Header(expected).Value, Header(actual).Value);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Textbox_edits_preserve_hyperlink_relationship_ownership(bool outerLink)
    {
        WmlDocument WithLink(string text, string target)
        {
            var header = Header(Build(""));
            var wrapped = outerLink ? header.Elements(W.p).Single().Elements(W.r).Single()
                : header.Descendants(W.txbxContent).First().Descendants(W.r).Single();
            wrapped.ReplaceWith(new XElement(W.hyperlink, new XAttribute(R.id, "rIdShared"), new XElement(wrapped)));
            header.Descendants(W.t).First().Value = text;
            var input = IrTestDocuments.FromBodyAndHeaderXml("<w:p/>", string.Concat(header.Elements()));
            using var stream = new MemoryStream();
            stream.Write(input.DocumentByteArray);
            stream.Position = 0;
            using (var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, true))
                package.MainDocumentPart!.HeaderParts.Single().AddHyperlinkRelationship(new Uri(target), true, "rIdShared");
            return new WmlDocument("link.docx", stream.ToArray());
        }
        string[] Links(WmlDocument document)
        {
            using var stream = new MemoryStream(document.DocumentByteArray);
            using var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
            var part = package.MainDocumentPart!.HeaderParts.Single();
            return part.GetXDocument().Descendants(W.hyperlink).Select(link => string.Concat(link.Descendants(W.t).Select(t => t.Value)) + "=" +
                part.HyperlinkRelationships.Single(r => r.Id == (string?)link.Attribute(R.id)).Uri).ToArray();
        }
        var left = WithLink("Alice", "https://example.org/original");
        var right = WithLink("Bob", "https://example.org/revised");
        var redline = DocxCompare.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        foreach (var document in new[] { left, right, redline, accepted, rejected })
            Assert.Empty(PackageValidation.Errors(document.DocumentByteArray, DocumentFormat.OpenXml.FileFormatVersions.Office2019));
        Assert.Equal(Links(right), Links(accepted));
        Assert.Equal(Links(left), Links(rejected));
        Assert.Equal(Drawings(right, coalesceRuns: true), Drawings(accepted, coalesceRuns: true));
        Assert.Equal(Drawings(left, coalesceRuns: true), Drawings(rejected, coalesceRuns: true));
    }

    private static XElement Table(XElement paragraph) => new(W.tbl,
        new XElement(W.tblPr), new XElement(W.tblGrid, new XElement(W.gridCol, new XAttribute(W.w + "w", 3000))),
        new XElement(W.tr, new XElement(W.tc,
            new XElement(W.tcPr, new XElement(W.tcW, new XAttribute(W.w + "w", 3000), new XAttribute(W.type, "dxa"))), paragraph)));

    [Fact]
    public void Existing_strict_policy_blocks_the_warning()
    {
        var result = DeliverableVerifier.VerifyDeliverable(Build("ins"),
            new DeliverableVerificationOptions { Mode = DeliverableVerificationMode.Strict });
        Assert.True(Assert.Single(result.Findings, f => f.Code == Code).BlocksDelivery);
    }

    [Fact]
    public void Separate_compatibility_branches_do_not_combine_into_a_failing_shape()
    {
        var single = Header(Build("ins", grouped: false)).Elements().Single();
        var inlineGroup = Header(Build("ins", anchored: false)).Elements().Single();
        var carrier = new XElement(W.p,
            new XElement(MC.AlternateContent,
                new XAttribute(XNamespace.Xmlns + "wps", WPS.wps),
                new XElement(MC.Choice, new XAttribute("Requires", "wps"), single.Elements()),
                new XElement(MC.Fallback, inlineGroup.Elements())));
        var document = IrTestDocuments.FromBodyAndHeaderXml("<w:p/>", carrier.ToString());
        Assert.Empty(PackageValidation.Errors(document.DocumentByteArray, DocumentFormat.OpenXml.FileFormatVersions.Office2019));

        Assert.DoesNotContain(DeliverableVerifier.VerifyDeliverable(document).Findings, f => f.Code == Code);
    }

    [Fact]
    public void A_revision_inside_the_text_box_does_not_mark_the_group_itself()
    {
        var header = Header(Build(""));
        var run = header.Descendants(W.t).First().Parent!;
        run.ReplaceWith(new XElement(W.ins, new XAttribute(W.id, 7),
            new XAttribute(W.author, "Fixture"), new XElement(run)));
        var document = IrTestDocuments.FromBodyAndHeaderXml("<w:p/>", string.Concat(header.Elements()));

        Assert.DoesNotContain(DeliverableVerifier.VerifyDeliverable(document).Findings, f => f.Code == Code);
    }

    [Theory]
    [InlineData(true, "wgp")]
    [InlineData(true, "graphicData")]
    [InlineData(true, "graphic")]
    [InlineData(false, "wgp")]
    public void Compatibility_choices_keep_text_boxes_with_their_own_group(bool groupHasTextBox, string wrapped)
    {
        var header = Header(Build("ins", textbox: groupHasTextBox));
        var payload = header.Descendants().Single(element => element.Name.LocalName == wrapped);
        var single = Header(Build("ins", grouped: false)).Descendants(WPS.wsp).Single();
        payload.ReplaceWith(new XElement(MC.AlternateContent,
            new XElement(MC.Choice, new XAttribute("Requires", "wpg"), new XElement(payload)),
            new XElement(MC.Fallback, groupHasTextBox ? new XElement(payload) : single)));
        var document = IrTestDocuments.FromBodyAndHeaderXml("<w:p/>", string.Concat(header.Elements()));
        Assert.Empty(PackageValidation.Errors(document.DocumentByteArray, DocumentFormat.OpenXml.FileFormatVersions.Office2019));

        Assert.Equal(groupHasTextBox, DeliverableVerifier.VerifyDeliverable(document).Findings.Any(f => f.Code == Code));
    }

    [Fact]
    public void An_inline_group_nested_in_an_anchored_single_text_box_is_not_an_anchored_group()
    {
        var header = Header(Build("ins", grouped: false));
        var inline = Header(Build("", anchored: false)).Descendants(W.r).First();
        inline.Descendants(WP.docPr).Single().SetAttributeValue("id", 2);
        header.Descendants(W.txbxContent).Single().Elements(W.p).Single().Add(inline);
        var document = IrTestDocuments.FromBodyAndHeaderXml("<w:p/>", string.Concat(header.Elements()));
        Assert.Empty(PackageValidation.Errors(document.DocumentByteArray, DocumentFormat.OpenXml.FileFormatVersions.Office2019));

        Assert.DoesNotContain(DeliverableVerifier.VerifyDeliverable(document).Findings, f => f.Code == Code);
    }

    private static XElement Header(WmlDocument document)
    {
        using var package = new ZipArchive(new MemoryStream(document.DocumentByteArray));
        using var stream = package.Entries.Single(e => e.FullName.Contains("header", StringComparison.Ordinal)
            && e.FullName.EndsWith(".xml", StringComparison.Ordinal)).Open();
        return XElement.Load(stream);
    }

    private static string[] Drawings(WmlDocument document, bool coalesceRuns = false) => Header(document).Descendants(W.drawing)
        .Select(drawing =>
        {
            var copy = new XElement(drawing);
            copy.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration
                || a.Name.Namespace == PtOpenXml.pt || a.Name.LocalName == "id").Remove();
            if (coalesceRuns)
            {
                copy.Descendants(W.rPr).Where(p => !p.HasElements && !p.HasAttributes).Remove();
                copy = WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(copy);
            }
            return copy.ToString(SaveOptions.DisableFormatting);
        }).ToArray();

    private static WmlDocument Build(string revision, bool grouped = true, bool textbox = true,
        bool anchored = true, bool header = true)
    {
        const string wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        const string a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string wpg = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
        const string wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
        string Shape(int index) => $"""
            <wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>
              <a:xfrm><a:off x="{index * 1500000}" y="0"/><a:ext cx="1400000" cy="400000"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </wps:spPr>{(textbox ? $"<wps:txbx><w:txbxContent><w:p><w:r><w:t>Juniper window {index}</w:t></w:r></w:p></w:txbxContent></wps:txbx>" : "")}<wps:bodyPr/></wps:wsp>
            """;
        var shapes = grouped ? $"""
            <wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm>
              <a:off x="0" y="0"/><a:ext cx="4500000" cy="400000"/>
              <a:chOff x="0" y="0"/><a:chExt cx="4500000" cy="400000"/>
            </a:xfrm></wpg:grpSpPr>{string.Concat(Enumerable.Range(0, 3).Select(Shape))}</wpg:wgp>
            """ : Shape(0);
        var content = $"<a:graphic><a:graphicData uri=\"{(grouped ? wpg : wps)}\">{shapes}</a:graphicData></a:graphic>";
        var geometry = $"<wp:extent cx=\"4500000\" cy=\"400000\"/>{(anchored ? "<wp:wrapNone/>" : "")}<wp:docPr id=\"1\" name=\"Juniper windows\"/>";
        var placement = anchored ? $"""
            <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
              <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
              <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>{geometry}{content}
            </wp:anchor>
            """ : $"<wp:inline>{geometry}{content}</wp:inline>";
        var run = $"<w:r><w:drawing>{placement}</w:drawing></w:r>";
        var paragraph = $"<w:p xmlns:wp=\"{wp}\" xmlns:a=\"{a}\" xmlns:wpg=\"{wpg}\" xmlns:wps=\"{wps}\">"
            + (revision.Length == 0 ? run : $"<w:{revision} {Stamp}>{run}</w:{revision}>") + "</w:p>";
        return IrTestDocuments.FromBodyAndHeaderXml(header ? "<w:p/>" : paragraph,
            header ? paragraph : "<w:p/>");
    }
}
