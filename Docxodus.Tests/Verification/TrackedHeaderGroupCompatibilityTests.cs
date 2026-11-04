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

    private static string[] Drawings(WmlDocument document) => Header(document).Descendants(W.drawing)
        .Select(drawing =>
        {
            var copy = new XElement(drawing);
            copy.DescendantsAndSelf().Attributes().Where(a => a.IsNamespaceDeclaration
                || a.Name.Namespace == PtOpenXml.pt || a.Name.LocalName == "id").Remove();
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
