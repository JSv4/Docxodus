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
    private readonly Xunit.Abstractions.ITestOutputHelper _output;
    public HtmlConversionOpsTests(Xunit.Abstractions.ITestOutputHelper output) => _output = output;

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
    public void HCO020_BulletListMarker_RendersUnicodeBullet()
    {
        // A bullet list item carries the Symbol-font glyph U+F0B7, which renders as a blank box in a
        // browser without the proprietary font installed. The converter should map list-marker
        // symbol glyphs to their Unicode equivalents (U+F0B7 -> U+2022 "•").
        var bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles", "Blank-wml.docx"));
        using var session = new DocxSession(bytes);
        var anchor = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind is "p" or "h" or "li").Anchor.Id;

        var edit = session.ReplaceText(anchor, "First bullet item");
        Assert.True(edit.Success, edit.Error?.Message);
        var li = session.ApplyListFormat(edit.Modified[0].Id, ListFormat.Bullet);
        Assert.True(li.Success, li.Error?.Message);

        string html = HtmlConversionOps.ConvertToHtml(session.Save(), new HtmlConversionOptions());

        Assert.Contains("•", html);       // • rendered for the bullet marker
        Assert.DoesNotContain("", html); // the raw Symbol private-use glyph is gone
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

    // Proves (a) the session-attached render resolves the SAME anchors the full render
    // stamps (one Unid scheme across convertDocxToHtml ↔ DocxSession ↔ RenderBlock) and
    // produces equivalent output, and (b) it avoids the per-call byte re-open + whole-doc
    // Unid pass, so it is no slower than the stateless path. Logs per-block latency.
    [Fact]
    public void HCO052_SessionAttachedRender_EquivalentAndNotSlower()
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "HC031-Complicated-Document.docx"));
        var opts = new HtmlConversionOptions { FabricateCssClasses = false };

        var full = System.Xml.Linq.XElement.Parse(
            HtmlConversionOps.ConvertToHtml(bytes,
                new HtmlConversionOptions { StampAnchors = true, FabricateCssClasses = false }));
        var anchors = full.Descendants()
            .Where(e => (e.Name.LocalName is "p" or "h1" or "h2" or "h3" or "h4")
                        && (string?)e.Attribute("data-anchor") != null
                        && e.Descendants().All(d => d.Name.LocalName != "img"))
            .Select(e => (string)e.Attribute("data-anchor")!)
            .Where(u => u.Length == 32)
            .Distinct().Take(20).ToList();
        Assert.NotEmpty(anchors);

        static string Text(string html) => System.Text.RegularExpressions.Regex.Replace(
            System.Xml.Linq.XElement.Parse(html).Value, "\\s+", " ").Trim();

        using var session = new DocxSession(bytes);

        // (a) Equivalence: session-attached resolves the full-render anchor (same scheme)
        // and yields the same text as the stateless path. This is the editor's invariant:
        // a DOM block's data-anchor is a valid DocxSession/RenderBlock anchor.
        foreach (var a in anchors.Take(6))
        {
            string viaBytes = HtmlConversionOps.RenderBlockHtml(bytes, a, opts);
            string viaSession = HtmlConversionOps.RenderBlockHtml(session, a, opts);
            Assert.Equal(Text(viaBytes), Text(viaSession));
        }

        // Warmup (JIT + first projection on the session path).
        HtmlConversionOps.RenderBlockHtml(bytes, anchors[0], opts);
        HtmlConversionOps.RenderBlockHtml(session, anchors[0], opts);

        var sw = System.Diagnostics.Stopwatch.StartNew();
        foreach (var a in anchors) HtmlConversionOps.RenderBlockHtml(bytes, a, opts);
        double statelessMs = sw.Elapsed.TotalMilliseconds / anchors.Count;

        sw.Restart();
        foreach (var a in anchors) HtmlConversionOps.RenderBlockHtml(session, a, opts);
        double sessionMs = sw.Elapsed.TotalMilliseconds / anchors.Count;

        _output.WriteLine($"PROFILE HC031 n={anchors.Count}: stateless={statelessMs:F2}ms/block " +
                          $"session-attached={sessionMs:F2}ms/block speedup={statelessMs / sessionMs:F2}x");

        // Session-attached must not be materially slower (it skips re-open + whole-doc
        // Unid assignment). Generous margin keeps the assertion robust to CI noise.
        Assert.True(sessionMs <= statelessMs * 1.25,
            $"session-attached slower than stateless: stateless={statelessMs:F2} session={sessionMs:F2}");
    }
}
