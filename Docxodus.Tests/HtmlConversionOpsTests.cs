#nullable enable
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HtmlConversionOpsTests
{
    private static byte[] TourPlanBytes() =>
        File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "HC001-5DayTourPlanTemplate.docx"));

    [Fact]
    public void HCO001_ConvertBytes_ProducesHtmlWithPrefix()
    {
        var options = new HtmlConversionOptions { CssClassPrefix = "zz-" };

        string html = HtmlConversionOps.ConvertToHtml(TourPlanBytes(), options);

        Assert.Contains("<html", html);
        Assert.Contains("zz-", html);
    }

    [Fact]
    public void HCO002_ConvertSession_ReflectsEdit()
    {
        using var session = new DocxSession(TourPlanBytes());
        var projection = session.Project();

        // First body paragraph/heading/list-item anchor, in document order.
        // C# AnchorTarget nests the anchor: record struct Anchor(Id, Kind, Scope, Unid).
        string FirstAnchor()
        {
            string? best = null;
            int bestPos = int.MaxValue;
            foreach (var target in projection.AnchorIndex.Values)
            {
                if (target.Anchor.Scope != "body") continue;
                if (target.Anchor.Kind is not ("p" or "h" or "li")) continue;
                int pos = projection.Markdown.IndexOf("{#" + target.Anchor.Id + "}", System.StringComparison.Ordinal);
                if (pos >= 0 && pos < bestPos) { bestPos = pos; best = target.Anchor.Id; }
            }
            Assert.NotNull(best);
            return best!;
        }

        var edit = session.ReplaceText(FirstAnchor(), "HCO002UNIQUEMARKER edited body.");
        Assert.True(edit.Success, edit.Error?.Message);

        string html = HtmlConversionOps.ConvertToHtml(session, new HtmlConversionOptions());

        Assert.Contains("HCO002UNIQUEMARKER", html);
    }

    // THE FEASIBILITY GATE (spec docs/architecture/ir_editor_feasibility.md §5/§6.1):
    // The full-document render is ground truth. RenderBlockHtml(anchor) is "faithful"
    // iff its output matches the data-anchor-stamped element from the full render —
    // same tag and same visible text. Proves single-block render out of whole-doc
    // context. (List-continuation + inline-image blocks are known PoC limits, skipped.)
    [Theory]
    [InlineData("HC006-Test-01.docx")]
    [InlineData("HC001-5DayTourPlanTemplate.docx")]
    public void HCO050_RenderBlockHtml_MatchesFullRenderPerAnchor(string fileName)
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles", fileName));

        // Full render = oracle; StampAnchors assigns the same deterministic Unids.
        var full = System.Xml.Linq.XElement.Parse(
            HtmlConversionOps.ConvertToHtml(bytes,
                new HtmlConversionOptions { StampAnchors = true, FabricateCssClasses = false }));

        var fullByAnchor = full.Descendants()
            .Where(e => (string?)e.Attribute("data-anchor") != null)
            .GroupBy(e => (string)e.Attribute("data-anchor")!)
            .ToDictionary(g => g.Key, g => g.First());

        // Stamping must work at all (this is the editor's actual render path).
        Assert.NotEmpty(fullByAnchor);

        static string Norm(string s) =>
            System.Text.RegularExpressions.Regex.Replace(s, "\\s+", " ").Trim();
        static bool HasImg(System.Xml.Linq.XElement e) =>
            e.Descendants().Any(d => d.Name.LocalName == "img");

        var targets = fullByAnchor
            .Where(kv => (kv.Value.Name.LocalName is "p" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6")
                         && !HasImg(kv.Value) && Norm(kv.Value.Value).Length > 0)
            .Take(12).ToList();
        Assert.NotEmpty(targets);

        int verified = 0;
        foreach (var kv in targets)
        {
            // data-anchor carries the bare unid; RenderBlockHtml accepts a bare unid
            // OR a full kind:scope:unid (it keys on the unid tail). This is exactly
            // what the editor passes back from a DOM block's data-anchor.
            string html = HtmlConversionOps.RenderBlockHtml(bytes, kv.Key,
                new HtmlConversionOptions { FabricateCssClasses = false });
            var blockEl = System.Xml.Linq.XElement.Parse(html);
            Assert.Equal(kv.Value.Name.LocalName, blockEl.Name.LocalName);
            Assert.Equal(Norm(kv.Value.Value), Norm(blockEl.Value));
            verified++;
        }

        Assert.True(verified > 0, "no blocks verified");
    }
}
