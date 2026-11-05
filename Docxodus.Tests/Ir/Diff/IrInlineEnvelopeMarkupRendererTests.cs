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
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Structural round trips for inline content-control/smart-tag carriers. The reader intentionally flattens
/// these wrappers for text alignment, so these tests assert the stronger package invariant: Accept retains the
/// complete right inline tree and Reject retains the complete left tree, even when visible text is unchanged.
/// </summary>
public class IrInlineEnvelopeMarkupRendererTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    private static WmlDocument Doc(string body) => IrTestDocuments.FromBodyXml(body);

    private static string Plain(string text = "controlled") =>
        $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>";

    private static string Sdt(string tag, string text = "controlled") =>
        "<w:p><w:sdt><w:sdtPr><w:tag w:val=\"" + tag + "\"/></w:sdtPr>" +
        "<w:sdtContent><w:r><w:t>" + text + "</w:t></w:r></w:sdtContent></w:sdt></w:p>";

    private static string SmartTag(string text = "controlled") =>
        "<w:p><w:smartTag w:uri=\"urn:schemas-microsoft-com:office:smarttags\" " +
        "w:element=\"country-region\"><w:r><w:t>" + text +
        "</w:t></w:r></w:smartTag></w:p>";

    private static string SdtInner(string tag, string text = "controlled") =>
        "<w:sdt><w:sdtPr><w:tag w:val=\"" + tag + "\"/></w:sdtPr>" +
        "<w:sdtContent><w:r><w:t>" + text + "</w:t></w:r></w:sdtContent></w:sdt>";

    private static IrParagraph Paragraph(WmlDocument doc) =>
        IrReader.Read(doc).Body.Blocks.OfType<IrParagraph>().Single();

    [Theory]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    public void PureTextboxCarrier_PreservesGeometryAndNestedBodyEdits(bool geometry, bool text, bool facade)
    {
        WmlDocument Input(int width, string words)
        {
            string Shape(string id, int size, string content) =>
                $"<v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" id=\"{id}\" style=\"width:{size}pt;height:40pt\">"
                + "<v:textbox><w:txbxContent>" + Plain(content) + "</w:txbxContent></v:textbox></v:shape>";
            return Doc("<w:p><w:r><w:t xml:space=\"preserve\">before </w:t></w:r><w:r><w:pict>"
                + Shape("first", width, words) + (geometry ? "" : "</w:pict></w:r><w:r><w:pict>")
                + Shape("second", 20, "second box")
                + "</w:pict></w:r><w:r><w:t xml:space=\"preserve\"> after</w:t></w:r></w:p>");
        }
        var original = Input(20, "fixed text");
        var revised = Input(geometry ? 40 : 20, text ? "changed text" : "fixed text");
        AssertSchemaValid(original);
        AssertSchemaValid(revised);
        var redline = facade ? DocxCompare.Compare(original, revised) : DocxDiff.Compare(original, revised);
        AssertSchemaValid(redline);
        foreach (var (resolved, expected) in new[]
        {
            (RevisionProcessor.AcceptRevisions(redline), revised),
            (RevisionProcessor.RejectRevisions(redline), original),
        })
        {
            AssertSchemaValid(resolved);
            AssertNoRevisionMarkup(resolved);
            var actual = MainXml(resolved);
            Assert.Equal(MainXml(expected).Root!.Value, actual.Root!.Value);
            XNamespace v = "urn:schemas-microsoft-com:vml";
            Assert.Equal(MainXml(expected).Descendants(v + "shape").Select(s => (string?)s.Attribute("style")),
                actual.Descendants(v + "shape").Select(s => (string?)s.Attribute("style")));
        }
        foreach (bool sources in new[] { false, true })
        {
            IrParagraph Read(WmlDocument doc) => Assert.IsType<IrParagraph>(
                Assert.Single(IrReader.Read(doc, new IrReaderOptions { RetainSources = sources }).Body.Blocks));
            var left = IrReader.Read(original, new IrReaderOptions { RetainSources = sources });
            var right = IrReader.Read(revised, new IrReaderOptions { RetainSources = sources });
            var leftBoxes = Read(original).Inlines.OfType<IrTextbox>().ToArray();
            var rightBoxes = Read(revised).Inlines.OfType<IrTextbox>().ToArray();
            Assert.Equal(!geometry, leftBoxes[0].CarrierDigest == rightBoxes[0].CarrierDigest);
            var op = Assert.Single(IrEditScriptBuilder.Build(left, right, new IrDiffSettings()).Operations);
            Assert.Equal(geometry, op.RequiresWholeParagraphReplace);
            if (text)
            {
                Assert.NotNull(op.TextboxDiffs);
                Assert.Equal(2, op.TextboxDiffs!.Count);
                Assert.Contains(op.TextboxDiffs[0].Ops, nested => nested.Kind != IrEditOpKind.EqualBlock);
                Assert.All(op.TextboxDiffs[1].Ops, nested => Assert.Equal(IrEditOpKind.EqualBlock, nested.Kind));
            }
        }
        var revisions = MainXml(redline).Descendants().Where(e => e.Name == W + "ins" || e.Name == W + "del").ToList();
        if (!geometry && text)
        {
            Assert.NotEmpty(revisions);
            // The markup renderer currently emits a whole-block pair for nested textbox diffs;
            // the edit script above must still retain the existing nested body operations.
        }
        if (!geometry && !text) Assert.Empty(revisions);
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(true, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, true, true)]
    public void PureTextboxCarrier_SafeParagraphSplitAndMergeRemainNative(bool merge, bool geometry, bool marker)
    {
        const string box = "<w:r><w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" "
            + "id=\"stable\" style=\"width:60pt;height:40pt\"><v:textbox><w:txbxContent>"
            + "<w:p><w:r><w:t>Stable box.</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>";
        string first = "<w:r><w:t xml:space=\"preserve\">Alder birch cedar dogwood. </w:t></w:r>"
            + (marker ? "<w:bookmarkStart w:id=\"8\" w:name=\"outside\"/><w:bookmarkEnd w:id=\"8\"/>" : "");
        const string last = "<w:r><w:t>Elm fir gingko hemlock.</w:t></w:r>";
        var joined = Doc("<w:p>" + first + box + last + "</w:p>" + Plain("Anchor remains stable across both views."));
        var split = Doc("<w:p>" + first + (geometry ? box.Replace("width:60pt", "width:100pt") : box) + "</w:p><w:p>" + last + "</w:p>" + Plain("Anchor remains stable across both views."));
        var original = merge ? split : joined;
        var revised = merge ? joined : split;
        foreach (bool sources in new[] { false, true })
        {
            var options = new IrReaderOptions { RetainSources = sources };
            var script = IrEditScriptBuilder.Build(IrReader.Read(original, options), IrReader.Read(revised, options), new IrDiffSettings());
            var kind = merge ? IrEditOpKind.MergeBlock : IrEditOpKind.SplitBlock;
            if (geometry) Assert.DoesNotContain(script.Operations, op => op.Kind == kind);
            else Assert.Single(script.Operations, op => op.Kind == kind);
        }
        var redline = DocxDiff.Compare(original, revised);
        AssertSchemaValid(redline);
        foreach (var (resolved, expected) in new[]
        {
            (RevisionProcessor.AcceptRevisions(redline), revised),
            (RevisionProcessor.RejectRevisions(redline), original),
        })
        {
            AssertSchemaValid(resolved);
            AssertNoRevisionMarkup(resolved);
            var actual = MainXml(resolved);
            Assert.Equal(MainXml(expected).Root!.Value, actual.Root!.Value);
            Assert.Single(actual.Descendants(W + "pict"));
            AssertTextboxGeometry(expected, resolved);
            Assert.Equal(TextboxMarkerOrder(expected), TextboxMarkerOrder(resolved));
            Assert.Equal(MainXml(expected).Root!.Element(W + "body")!.Elements(W + "p").Count(),
                actual.Root!.Element(W + "body")!.Elements(W + "p").Count());
        }
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public void TextboxCarrier_RegroupedMarkerGapRestoresSourceOrder(bool regroup, bool reverse, bool facade)
    {
        const string start = "<w:bookmarkStart w:id=\"5\" w:name=\"point\"/>";
        const string end = "<w:bookmarkEnd w:id=\"5\"/>";
        string Shape(string text) => "<v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" id=\"" + text
            + "\" style=\"width:60pt;height:40pt\"><v:textbox><w:txbxContent>" + Plain(text)
            + "</w:txbxContent></v:textbox></v:shape>";
        string Pict(string shapes) => "<w:r><w:pict>" + shapes + "</w:pict></w:r>";
        var separated = Doc("<w:p>" + Pict(Shape("first")) + start + end + Pict(Shape("second"))
            + "<w:r><w:t xml:space=\"preserve\"> old tail</w:t></w:r></w:p>");
        var changed = Doc("<w:p>" + (regroup ? Pict(Shape("first") + Shape("second"))
            : Pict(Shape("first")) + Pict(Shape("second")))
            + "<w:r><w:t xml:space=\"preserve\"> new tail</w:t></w:r></w:p>");
        var original = reverse ? changed : separated;
        var revised = reverse ? separated : changed;
        AssertSchemaValid(original);
        AssertSchemaValid(revised);
        foreach (var document in new[] { original, revised })
        {
            var retained = Assert.IsType<IrParagraph>(Assert.Single(IrReader.Read(document,
                new IrReaderOptions { RetainSources = true }).Body.Blocks));
            var elided = Assert.IsType<IrParagraph>(Assert.Single(IrReader.Read(document,
                new IrReaderOptions { RetainSources = false }).Body.Blocks));
            Assert.Equal(retained.Source.SharedCarriers, elided.Source.SharedCarriers);
            Assert.Equal(retained.Inlines.OfType<IrTextbox>().Select(box => box.CarrierDigest),
                elided.Inlines.OfType<IrTextbox>().Select(box => box.CarrierDigest));
        }
        foreach (bool sources in new[] { false, true })
        {
            var options = new IrReaderOptions { RetainSources = sources };
            var script = IrEditScriptBuilder.Build(IrReader.Read(original, options), IrReader.Read(revised, options), new IrDiffSettings());
            Assert.Equal(regroup, Assert.Single(script.Operations).RequiresWholeParagraphReplace);
        }
        var redline = facade ? DocxCompare.Compare(original, revised) : DocxDiff.Compare(original, revised);
        AssertSchemaValid(redline);
        foreach (var (resolved, expected) in new[]
        {
            (RevisionProcessor.AcceptRevisions(redline), revised),
            (RevisionProcessor.RejectRevisions(redline), original),
        })
        {
            AssertSchemaValid(resolved);
            AssertNoRevisionMarkup(resolved);
            Assert.Equal(TextboxMarkerOrder(expected), TextboxMarkerOrder(resolved));
            Assert.Equal(MainXml(expected).Root!.Value, MainXml(resolved).Root!.Value);
            Assert.Equal(MainXml(expected).Descendants(W + "pict").Count(), MainXml(resolved).Descendants(W + "pict").Count());
        }
    }

    private static string TextboxMarkerOrder(WmlDocument document)
    {
        var xml = MainXml(document);
        var names = xml.Descendants(W + "bookmarkStart").ToDictionary(e => (string)e.Attribute(W + "id")!,
            e => (string?)e.Attribute(W + "name"));
        return string.Concat(xml.Descendants().Where(e => !e.Ancestors(W + "txbxContent").Any())
            .Select(e => e.Name == W + "t" ? e.Value : e.Name == W + "p" ? "{p}"
                : e.Name == W + "txbxContent" ? "{textbox:" + e.Value + "}"
                : e.Name == W + "bookmarkStart" ? "{start:" + (string?)e.Attribute(W + "name") + "}"
                : e.Name == W + "bookmarkEnd" ? "{end:" + names.GetValueOrDefault((string)e.Attribute(W + "id")!, "orphan") + "}" : ""));
    }

    private static string TextboxCarrier(int width, string body, string kind = "vml")
    {
        string vml = "<w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" style=\"width:"
            + width + "pt;height:40pt\"><v:textbox><w:txbxContent>" + body
            + "</w:txbxContent></v:textbox></v:shape></w:pict>";
        string drawing = "<w:drawing xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" "
            + "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
            + "<wp:inline><wp:extent cx=\"" + width * 12700 + "\" cy=\"508000\"/><wp:docPr id=\"1\" name=\"Box\"/>"
            + "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
            + "<wps:wsp><wps:cNvSpPr txBox=\"1\"/><wps:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></wps:spPr>"
            + "<wps:txbx><w:txbxContent>" + body + "</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>"
            + "</a:graphicData></a:graphic></wp:inline></w:drawing>";
        string carrier = kind == "vml" ? vml : kind == "drawing" ? drawing
            : "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
                + "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
                + "<mc:Choice Requires=\"wps\">" + drawing + "</mc:Choice><mc:Fallback>" + vml + "</mc:Fallback></mc:AlternateContent>";
        return "<w:r>" + carrier + "</w:r>";
    }

    private static void AssertTextboxGeometry(WmlDocument expected, WmlDocument actual)
    {
        XNamespace v = "urn:schemas-microsoft-com:vml";
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        Assert.Equal(MainXml(expected).Descendants(v + "shape").Select(e => (string?)e.Attribute("style")),
            MainXml(actual).Descendants(v + "shape").Select(e => (string?)e.Attribute("style")));
        Assert.Equal(MainXml(expected).Descendants(wp + "extent").Select(e => (string?)e.Attribute("cx")),
            MainXml(actual).Descendants(wp + "extent").Select(e => (string?)e.Attribute("cx")));
    }

    [Theory]
    [InlineData("vml")]
    [InlineData("drawing")]
    [InlineData("alternate")]
    public void TextboxCarrier_GeometrySurvivesNestedScopesAndSourceElision(string kind)
    {
        string Cell(string content) => "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"6000\"/></w:tblGrid>"
            + "<w:tr><w:tc>" + content + "<w:p/></w:tc></w:tr></w:tbl>";
        foreach (string scope in new[] { "body", "cell", "nested-cell", "textbox", "sdt" })
        {
            WmlDocument Input(int width)
            {
                string body = "<w:p>" + TextboxCarrier(width, Plain("Stable body."), kind) + "</w:p>";
                body = scope switch
                {
                    "cell" => Cell(body),
                    "nested-cell" => Cell(Cell(body)),
                    "textbox" => "<w:p>" + TextboxCarrier(200, Cell(body)) + "</w:p>",
                    "sdt" => "<w:sdt><w:sdtPr/><w:sdtContent>" + body + "</w:sdtContent></w:sdt>",
                    _ => body,
                };
                return Doc(body);
            }
            var left = Input(60);
            var right = Input(100);
            AssertSchemaValid(left);
            AssertSchemaValid(right);
            string Script(bool sources) => IrEditScriptJson.Write(IrEditScriptBuilder.Build(
                IrReader.Read(left, new IrReaderOptions { RetainSources = sources }),
                IrReader.Read(right, new IrReaderOptions { RetainSources = sources }), new IrDiffSettings()));
            Assert.Equal(Script(true), Script(false));
            var redline = DocxCompare.Compare(left, right);
            AssertSchemaValid(redline);
            foreach (var (resolved, expected) in new[]
            {
                (RevisionProcessor.AcceptRevisions(redline), right),
                (RevisionProcessor.RejectRevisions(redline), left),
            })
            {
                AssertSchemaValid(resolved);
                AssertNoRevisionMarkup(resolved);
                Assert.Equal(MainXml(expected).Root!.Value, MainXml(resolved).Root!.Value);
                AssertTextboxGeometry(expected, resolved);
            }
            Assert.NotEmpty(Docxodus.Verification.SemanticDiff.Compare(left, right).Changes);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TextboxCarrier_ConsolidationDistinguishesGeometryConsensusFromConflict(bool same)
    {
        WmlDocument Input(int width) => Doc("<w:p>" + TextboxCarrier(width, Plain("Stable body.")) + "</w:p>");
        var original = Input(60);
        var first = Input(100);
        var second = Input(same ? 100 : 140);
        var script = IrCompositeMergerTests.MergeOf(original, ("Alice", first), ("Bob", second));
        Assert.Equal(same, script.Conflicts.Count == 0);
        var redline = DocxDiff.Consolidate(original, new[]
        {
            new DocxDiffReviewer { Author = "Alice", Document = first },
            new DocxDiffReviewer { Author = "Bob", Document = second },
        });
        AssertSchemaValid(redline);
        AssertTextboxGeometry(original, RevisionProcessor.RejectRevisions(redline));
        if (same) AssertTextboxGeometry(first, RevisionProcessor.AcceptRevisions(redline));
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("table")]
    [InlineData("row")]
    public void TextboxCarrier_MovedGeometryRemainsObservableAndReversible(string scope)
    {
        string Box(int width) => "<w:p><w:r><w:t>Moving boxed paragraph distinctive words.</w:t></w:r>"
            + TextboxCarrier(width, Plain("Stable textbox.")) + "</w:p>";
        var anchors = Plain("Anchor alpha beta gamma stays first.") + Plain("Anchor delta epsilon zeta stays second.")
            + Plain("Anchor eta theta iota stays third.");
        string Row(string body) => "<w:tr><w:tc>" + body + "</w:tc></w:tr>";
        string Table(string rows) => "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"6000\"/></w:tblGrid>" + rows + "</w:tbl>";
        var leftBody = Box(60) + anchors;
        var rightBody = anchors + Box(100);
        if (scope == "table")
        {
            leftBody = Table(Row(Box(60))) + anchors;
            rightBody = anchors + Table(Row(Box(100)));
        }
        else if (scope == "row")
        {
            var rows = Row(Plain("fixed alpha one")) + Row(Plain("fixed beta two")) + Row(Plain("fixed gamma three"));
            leftBody = Table(Row(Box(60)) + rows);
            rightBody = Table(rows + Row(Box(100)));
        }
        var left = Doc(leftBody);
        var right = Doc(rightBody);
        var changes = Docxodus.Verification.SemanticDiff.Compare(left, right).Changes;
        Assert.Contains(changes, change => change.Operation == Docxodus.Verification.SemanticChangeOperation.Move);
        Assert.Contains(changes, change => change.Path.Contains("textbox_carrier"));
        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        AssertSchemaValid(accepted);
        AssertSchemaValid(rejected);
        AssertTextboxGeometry(right, accepted);
        AssertTextboxGeometry(left, rejected);
        Assert.Equal(MainXml(right).Root!.Value, MainXml(accepted).Root!.Value);
        Assert.Equal(MainXml(left).Root!.Value, MainXml(rejected).Root!.Value);
    }

    [Theory]
    [InlineData("hyperlink", false)]
    [InlineData("hyperlink", true)]
    [InlineData("sdt", false)]
    [InlineData("sdt", true)]
    [InlineData("nested", false)]
    [InlineData("nested", true)]
    [InlineData("direct", false)]
    [InlineData("direct", true)]
    [InlineData("unchanged", false)]
    [InlineData("unchanged", true)]
    public void InlineCustomXmlContent_TracksTextThroughItsContainers(string shape, bool facade)
    {
        string Custom(string content) => "<w:customXml w:element=\"entry\">"
            + "<w:customXmlPr><w:attr w:name=\"key\" w:val=\"original\"/></w:customXmlPr>" + content + "</w:customXml>";
        WmlDocument Input(string text)
        {
            var inner = Custom("<w:r><w:t>Cedar " + text + "</w:t></w:r>");
            if (shape == "nested") inner = Custom(inner);
            var carrier = shape == "sdt" ? "<w:sdt><w:sdtPr/><w:sdtContent>" + inner + "</w:sdtContent></w:sdt>"
                : shape == "direct" ? inner : "<w:hyperlink w:anchor=\"cedar\">" + inner + "</w:hyperlink>";
            return Doc("<w:p><w:r><w:t>Before.</w:t></w:r>" + carrier + "<w:r><w:t>After.</w:t></w:r></w:p>");
        }
        var original = Input("opened");
        var revised = Input(shape == "unchanged" ? "opened" : "closed");
        AssertSchemaValid(original);
        AssertSchemaValid(revised);
        var redline = facade ? DocxCompare.Compare(original, revised) : DocxDiff.Compare(original, revised);
        AssertSchemaValid(redline);
        foreach (var (resolved, expected) in new[]
        {
            (RevisionProcessor.AcceptRevisions(redline), revised),
            (RevisionProcessor.RejectRevisions(redline), original),
        })
        {
            AssertSchemaValid(resolved);
            AssertNoRevisionMarkup(resolved);
            var actual = MainXml(resolved);
            Assert.Equal(MainXml(expected).Root!.Value, actual.Root!.Value);
            Assert.Equal(MainXml(expected).Descendants(W + "customXml").Count(), actual.Descendants(W + "customXml").Count());
            var text = Assert.Single(actual.Descendants(W + "t"), t => t.Value.Contains("Cedar"));
            Assert.Equal(shape == "nested" ? 2 : 1, text.Ancestors(W + "customXml").Count());
            foreach (var wrapper in text.Ancestors(W + "customXml"))
                Assert.Equal("original", (string?)wrapper.Element(W + "customXmlPr")?.Element(W + "attr")?.Attribute(W + "val"));
            foreach (var pagination in new[] { PaginationMode.None, PaginationMode.Paginated })
            {
                var html = WmlToHtmlConverter.ConvertToHtml(resolved, new WmlToHtmlConverterSettings { RenderPagination = pagination });
                Assert.Equal(MainXml(expected).Root!.Value,
                    Assert.Single(html.Descendants(Xhtml.p)).Value.Replace('\u00a0', ' '));
            }
        }
    }

    [Fact]
    public void Read_InlineEnvelopeDigest_IgnoresFormattingOnlyRunSplit()
    {
        // Run segmentation is formatting-sensitive, so a formatting-only span boundary in the
        // text BEFORE an inline carrier must not move the carrier's recorded position
        // (issue #600 — the inline-envelope sibling of the field-envelope fix in #593).
        var joined = Paragraph(Doc(
            "<w:p><w:r><w:t xml:space=\"preserve\">See </w:t></w:r>"
            + SdtInner("stable")
            + "<w:r><w:t xml:space=\"preserve\"> for details.</w:t></w:r></w:p>"));
        var split = Paragraph(Doc(
            "<w:p><w:r><w:t>S</w:t></w:r>"
            + "<w:r><w:rPr><w:i/></w:rPr><w:t xml:space=\"preserve\">ee </w:t></w:r>"
            + SdtInner("stable")
            + "<w:r><w:t xml:space=\"preserve\"> for details.</w:t></w:r></w:p>"));

        Assert.Equal(joined.ContentHash, split.ContentHash);
        Assert.Equal(joined.InlineEnvelopeDigest, split.InlineEnvelopeDigest);

        // Guard: swapping the carrier with its text neighbor is still a structural difference.
        var swapped = Paragraph(Doc(
            "<w:p>" + SdtInner("stable")
            + "<w:r><w:t xml:space=\"preserve\">See </w:t></w:r>"
            + "<w:r><w:t xml:space=\"preserve\"> for details.</w:t></w:r></w:p>"));
        Assert.NotEqual(joined.InlineEnvelopeDigest, swapped.InlineEnvelopeDigest);
    }

    [Fact]
    public void Render_FormattingOnlySpanBeforeInlineSdt_StaysFormatOnly()
    {
        // Italicizing a span that starts mid-run BEFORE the carrier and does not touch the
        // carrier itself: content is identical, only run properties differ. This must survive
        // as a FormatOnly alignment and render as w:rPrChange — not a whole-paragraph del+ins
        // pair re-typing the controlled region (issue #600). A formatting change reaching INTO
        // the sdtContent still takes the whole-carrier fallback, because the envelope entry
        // embeds the carrier XML — that conservatism is pinned by
        // Render_InlineSdtInnerTextChange_UsesWholeParagraphFallback.
        var left = Doc(
            "<w:p><w:r><w:t xml:space=\"preserve\">See </w:t></w:r>"
            + SdtInner("stable")
            + "<w:r><w:t xml:space=\"preserve\"> for details.</w:t></w:r></w:p>");
        var right = Doc(
            "<w:p><w:r><w:t>S</w:t></w:r>"
            + "<w:r><w:rPr><w:i/></w:rPr><w:t xml:space=\"preserve\">ee </w:t></w:r>"
            + SdtInner("stable")
            + "<w:r><w:rPr><w:i/></w:rPr><w:t xml:space=\"preserve\"> for</w:t></w:r>"
            + "<w:r><w:t xml:space=\"preserve\"> details.</w:t></w:r></w:p>");

        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());
        var op = Assert.Single(script.Operations);
        Assert.Equal(IrEditOpKind.FormatOnlyBlock, op.Kind);

        var redline = DocxDiff.Compare(left, right);
        AssertSchemaValid(redline);
        var root = MainXml(redline);
        Assert.Empty(root.Descendants(W + "ins"));
        Assert.Empty(root.Descendants(W + "del"));
        Assert.Single(root.Descendants(W + "sdt"));
        Assert.NotEmpty(root.Descendants(W + "rPrChange"));

        // Reject legitimately keeps the right side's run segmentation with the left formats
        // restored (modeled-semantic, not byte, equivalence — the #532-#534 convention), so
        // assert the semantics: text, carrier structure, and the formatting delta itself.
        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
        Assert.NotEmpty(MainXml(accepted).Descendants(W + "i"));
        Assert.Empty(MainXml(rejected).Descendants(W + "i"));
        Assert.Single(MainXml(accepted).Descendants(W + "sdt"));
        Assert.Single(MainXml(rejected).Descendants(W + "sdt"));
        Assert.Equal(Paragraph(left).ContentHash, Paragraph(rejected).ContentHash);
        Assert.Equal(Paragraph(left).InlineEnvelopeDigest, Paragraph(rejected).InlineEnvelopeDigest);
        Assert.Equal(Paragraph(right).InlineEnvelopeDigest, Paragraph(accepted).InlineEnvelopeDigest);
    }

    [Fact]
    public void Render_InlineSdtMovedAcrossATextStretch_StillRoundTrips()
    {
        // Moving the carrier past one of two adjacent text runs merges a text stretch, which a
        // format-neutral position encoding cannot distinguish from a formatting split in
        // reverse — so this move is caught by the CONTENT hash instead of the envelope digest.
        // Whatever path renders it, the package invariant must hold: accept keeps the right
        // inline tree, reject restores the left one.
        var left = Doc(
            "<w:p><w:r><w:t xml:space=\"preserve\">alpha </w:t></w:r>"
            + SdtInner("stable")
            + "<w:r><w:t xml:space=\"preserve\"> omega</w:t></w:r></w:p>");
        var right = Doc(
            "<w:p><w:r><w:t xml:space=\"preserve\">alpha </w:t></w:r>"
            + "<w:r><w:t xml:space=\"preserve\"> omega</w:t></w:r>"
            + SdtInner("stable") + "</w:p>");

        AssertRoundTrips(left, right);
        AssertRoundTrips(right, left);
    }

    [Fact]
    public void Render_InlineSdtAdditionAndRemoval_RoundTripsExactInlineShape()
    {
        var plain = Doc(Plain());
        var wrapped = Doc(Sdt("new"));

        AssertRoundTrips(plain, wrapped);
        AssertRoundTrips(wrapped, plain);

        var revisions = DocxDiff.GetRevisions(plain, wrapped);
        Assert.Contains(revisions, revision => revision.Type == DocxDiffRevisionType.Deleted && revision.Text == "controlled");
        Assert.Contains(revisions, revision => revision.Type == DocxDiffRevisionType.Inserted && revision.Text == "controlled");
    }

    [Fact]
    public void Render_InlineSdtMetadataOnlyChange_RoundTripsOldAndNewEnvelope()
    {
        AssertRoundTrips(Doc(Sdt("old")), Doc(Sdt("new")));
    }

    [Fact]
    public void Render_SmartTagWrapperChange_RoundTripsExactInlineShape()
    {
        // w:smartTag is legacy markup that OpenXmlValidator (Office 2019) rejects even in the source fixture;
        // assert its structural accept/reject behavior here, while the SDT cases above carry schema coverage.
        AssertRoundTrips(Doc(Plain()), Doc(SmartTag()), validateSchema: false);
    }

    [Fact]
    public void Render_InlineSdtInnerTextChange_UsesWholeParagraphFallback()
    {
        var left = Doc(Sdt("stable", "old payload"));
        var right = Doc(Sdt("stable", "new payload"));
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());
        var op = Assert.Single(script.Operations);
        Assert.True(op.RequiresWholeParagraphReplace);

        AssertRoundTrips(left, right);
    }

    [Fact]
    public void Consolidate_SingleInlineSdtEnvelopeChange_RoundTripsAndSerializesStructuralFlag()
    {
        var left = Doc(Sdt("old"));
        var right = Doc(Sdt("new"));
        var reviewers = new[] { new DocxDiffReviewer { Author = "Alice", Document = right } };

        var merged = DocxDiff.Consolidate(left, reviewers);
        AssertSchemaValid(merged);

        AssertInlineTreeEqual(right, RevisionProcessor.AcceptRevisions(merged));
        AssertInlineTreeEqual(left, RevisionProcessor.RejectRevisions(merged));
        Assert.Contains("\"requiresWholeParagraphReplace\"",
            DocxDiff.GetConsolidatedEditScriptJson(left, reviewers));

        var revisions = DocxDiff.GetConsolidatedRevisions(left, reviewers);
        Assert.Contains(revisions, revision =>
            revision.Type == DocxDiffRevisionType.Deleted &&
            revision.Text == "controlled" &&
            revision.Author == "Alice");
        Assert.Contains(revisions, revision =>
            revision.Type == DocxDiffRevisionType.Inserted &&
            revision.Text == "controlled" &&
            revision.Author == "Alice");
    }

    private static void AssertRoundTrips(WmlDocument left, WmlDocument right, bool validateSchema = true)
    {
        var redline = DocxDiff.Compare(left, right);
        if (validateSchema)
            AssertSchemaValid(redline);
        Assert.Contains(MainXml(redline).Descendants(W + "ins"), _ => true);
        Assert.Contains(MainXml(redline).Descendants(W + "del"), _ => true);

        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);

        AssertInlineTreeEqual(right, accepted);
        AssertInlineTreeEqual(left, rejected);
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
    }

    private static void AssertSchemaValid(WmlDocument doc) =>
        PackageValidation.AssertValid(doc, FileFormatVersions.Office2019);

    private static void AssertInlineTreeEqual(WmlDocument expected, WmlDocument actual) =>
        Assert.True(XNode.DeepEquals(InlineTree(expected), InlineTree(actual)),
            $"expected: {InlineTree(expected)}\nactual: {InlineTree(actual)}");

    private static XElement InlineTree(WmlDocument doc)
    {
        var paragraph = MainXml(doc).Root!.Element(W + "body")!.Elements(W + "p").Single();
        var tree = new XElement("inline",
            paragraph.Elements().Where(element => element.Name != W + "pPr").Select(CloneWithoutNoise));
        return tree;
    }

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

    private static void AssertNoRevisionMarkup(WmlDocument doc)
    {
        var revisions = new HashSet<XName>
        {
            W + "ins", W + "del", W + "moveFrom", W + "moveTo",
            W + "moveFromRangeStart", W + "moveFromRangeEnd", W + "moveToRangeStart", W + "moveToRangeEnd",
        };
        Assert.DoesNotContain(MainXml(doc).Descendants(), element => revisions.Contains(element.Name));
    }

    private static XDocument MainXml(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        using var reader = new StreamReader(wdoc.MainDocumentPart!.GetStream());
        return XDocument.Parse(reader.ReadToEnd());
    }
}
