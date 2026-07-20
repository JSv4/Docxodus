#nullable enable
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Wp = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

public class HtmlConversionOpsTests
{
    private readonly Xunit.Abstractions.ITestOutputHelper _output;
    public HtmlConversionOpsTests(Xunit.Abstractions.ITestOutputHelper output) => _output = output;

    private static byte[] TourPlanBytes() =>
        File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "HC001-5DayTourPlanTemplate.docx"));

    private const string TransitionalMain =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string StrictMain = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    private const string TransitionalRels =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string StrictRels =
        "http://purl.oclc.org/ooxml/officeDocument/relationships";

    private static byte[] StrictDocumentOnlyDocxBytes(string text)
    {
        var bytes = DocumentOnlyDocxBytes(text);
        using var ms = new MemoryStream();
        ms.Write(bytes, 0, bytes.Length);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entry in zip.Entries.ToList())
            {
                if (!entry.FullName.EndsWith(".xml", System.StringComparison.OrdinalIgnoreCase) &&
                    !entry.FullName.EndsWith(".rels", System.StringComparison.OrdinalIgnoreCase))
                    continue;

                string xml;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                    xml = reader.ReadToEnd();
                var strict = xml
                    .Replace(TransitionalMain, StrictMain, System.StringComparison.Ordinal)
                    .Replace(TransitionalRels, StrictRels, System.StringComparison.Ordinal);
                if (strict == xml)
                    continue;
                using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                writer.BaseStream.SetLength(0);
                writer.Write(strict);
            }
        }
        return ms.ToArray();
    }

    [Fact]
    public void HCO001_ConvertBytes_ProducesHtmlWithPrefix()
    {
        var options = new HtmlConversionOptions { CssClassPrefix = "zz-" };

        string html = HtmlConversionOps.ConvertToHtml(TourPlanBytes(), options);

        Assert.Contains("<html", html);
        Assert.Contains("zz-", html);
    }

    [Fact]
    public void HCO003_PaginatedHtml_LeavesTheCaptureHostBodyFlush()
    {
        // Paginated HTML is injected into the React viewer's capture host. Its fixed-size page boxes
        // own geometry, so a converter-level body margin must not shrink/overflow that host. Standalone
        // conversion retains the readable 20px margin for existing consumers.
        string paginated = HtmlConversionOps.ConvertToHtml(TourPlanBytes(),
            new HtmlConversionOptions { PaginationMode = (int)PaginationMode.Paginated });
        string standalone = HtmlConversionOps.ConvertToHtml(TourPlanBytes(), new HtmlConversionOptions());

        Assert.Contains("body { font-family: Arial, sans-serif; margin: 0; }", paginated);
        Assert.Contains("body { font-family: Arial, sans-serif; margin: 20px; }", standalone);
    }

    [Fact]
    public void HCO076_PaginatedHtml_UsesDocumentPageSizeWithoutOuterPrintMargin()
    {
        // The paginator has already applied the Word margins within each page box. Its capture
        // path must advertise the paper size to Chromium without applying those margins again.
        string paginated = HtmlConversionOps.ConvertToHtml(
            PageSizedDocxBytes(width: 11906, height: 16838),
            new HtmlConversionOptions { PaginationMode = (int)PaginationMode.Paginated });
        string standalone = HtmlConversionOps.ConvertToHtml(
            PageSizedDocxBytes(width: 11906, height: 16838), new HtmlConversionOptions());

        Assert.Contains("@page", paginated);
        Assert.Contains("size: 8.27in 11.69in;", paginated);
        Assert.Contains("margin: 0;", paginated);
        Assert.DoesNotContain("@page docxodus-section-", paginated);
        Assert.DoesNotContain("@page", standalone);
    }

    [Fact]
    public void HCO077_MixedPaginatedSections_UseNamedPrintPages()
    {
        // The staging and final paginator page boxes retain their data-section-index. Named
        // pages let Chromium print each section at its own physical size without relying on a
        // caller to customize page.pdf options.
        string html = HtmlConversionOps.ConvertToHtml(MixedPageSizedDocxBytes(),
            new HtmlConversionOptions { PaginationMode = (int)PaginationMode.Paginated });

        Assert.Contains("@page docxodus-section-0", html);
        Assert.Contains("size: 8.27in 11.69in;", html);
        Assert.Contains("@page docxodus-section-1", html);
        Assert.Contains("size: 11.00in 8.50in;", html);
        Assert.Contains(".page-box[data-section-index=\"0\"]", html);
        Assert.Contains("page: docxodus-section-0;", html);
        Assert.Contains(".page-box[data-section-index=\"1\"]", html);
        Assert.Contains("page: docxodus-section-1;", html);
        Assert.DoesNotContain("@page {", html);
        Assert.Contains("data-section-index=\"0\"", html);
        Assert.Contains("data-page-width=\"595.3\"", html);
        Assert.Contains("data-page-height=\"841.9\"", html);
        Assert.Contains("data-section-index=\"1\"", html);
        Assert.Contains("data-page-width=\"792.0\"", html);
        Assert.Contains("data-page-height=\"612.0\"", html);
    }

    [Fact]
    public void HCO078_PaginatedHtml_WithoutSectionProperties_UsesLetterPageSize()
    {
        string html = HtmlConversionOps.ConvertToHtml(DocumentOnlyDocxBytes("No section properties"),
            new HtmlConversionOptions { PaginationMode = (int)PaginationMode.Paginated });

        Assert.Contains("size: 8.50in 11.00in;", html);
        Assert.Contains("margin: 0;", html);
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

    // The single-block render path sets SkipFormattingPartsSimplification=true to avoid re-walking
    // the (potentially huge) style gallery on every keystroke commit. That pass only strips
    // rendering-irrelevant rsids from the style parts, so it MUST be byte-for-byte rendering-neutral.
    // Prove it directly: a full-document convert with the flag on vs off produces identical HTML
    // (covers CSS classes + theme fonts + list markers, not just tag+text like HCO050).
    [Theory]
    [InlineData("HC031-Complicated-Document.docx", false)]
    [InlineData("HC001-5DayTourPlanTemplate.docx", false)]
    [InlineData("HC031-Complicated-Document.docx", true)]
    public void HCO053_SkipFormattingPartsSimplification_IsRenderingNeutral(string fileName, bool paginated)
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles", fileName));
        string Render(bool skip)
        {
            using var ms = new MemoryStream();
            ms.Write(bytes, 0, bytes.Length);
            ms.Position = 0;
            using var doc = WordprocessingDocument.Open(ms, true);
            var settings = new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = false,
                StampAnchors = true,
                RenderPagination = paginated ? PaginationMode.Paginated : PaginationMode.None,
                SkipFormattingPartsSimplification = skip,
            };
            return WmlToHtmlConverter.ConvertToHtml(doc, settings)
                .ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }

        Assert.Equal(Render(false), Render(true));
    }

    // The session-attached render path reuses a cached formatting "shell" across calls. Prove it is
    // (a) consistent across calls (cache reuse doesn't drift) and (b) byte-identical to the stateless
    // path (which HCO050 already ties to the full-render oracle).
    [Fact]
    public void HCO054_SessionShellRender_ConsistentAndMatchesStateless()
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "HC031-Complicated-Document.docx"));
        using var session = new DocxSession(bytes);
        var opts = new HtmlConversionOptions { FabricateCssClasses = false, CssClassPrefix = "pt-" };
        var anchors = session.Project().AnchorIndex.Keys
            .Where(k => k.StartsWith("p:") || k.StartsWith("h:") || k.StartsWith("li:"))
            .Take(12).ToList();
        Assert.NotEmpty(anchors);

        int verified = 0;
        foreach (var a in anchors)
        {
            string first = HtmlConversionOps.RenderBlockHtml(session, a, opts);   // builds the shell
            string second = HtmlConversionOps.RenderBlockHtml(session, a, opts);  // reuses the shell
            Assert.Equal(first, second);
            string stateless = HtmlConversionOps.RenderBlockHtml(bytes, a, opts); // independent path
            Assert.Equal(stateless, first);
            verified++;
        }
        Assert.True(verified > 0);
    }

    // A mid-session format op (ApplyListFormat) mutates the numbering part, so the cached shell MUST
    // be rebuilt (signature change) — otherwise the freshly-list-ified paragraph would render WITHOUT
    // its marker against a stale (numbering-less) shell. Also covers the no-list -> list transition.
    [Fact]
    public void HCO055_SessionShellRender_RebuildsAfterFormattingMutation()
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "HC031-Complicated-Document.docx"));
        using var session = new DocxSession(bytes);
        var opts = new HtmlConversionOptions { FabricateCssClasses = false, CssClassPrefix = "pt-" };

        var plain = session.Project().AnchorIndex
            .First(kv => kv.Key.StartsWith("p:") && kv.Value.TextPreview.Trim().Length > 3);

        // Prime the shell (no marker yet).
        string before = HtmlConversionOps.RenderBlockHtml(session, plain.Key, opts);
        Assert.DoesNotContain("data-list-marker", before);

        // Mutate the numbering part; the next render must rebuild the shell and show the marker.
        var r = session.ApplyListFormat(plain.Key, ListFormat.Bullet);
        Assert.True(r.Success, r.Error?.Message);
        string after = HtmlConversionOps.RenderBlockHtml(session, r.Modified[0].Id, opts);
        Assert.Contains("data-list-marker", after);
    }

    // A borderless layout table (w:tblBorders all w:val="none", with NO w:sz) — the standard way real
    // S-1 covers lay out multi-column rows — used to CRASH the whole conversion: both
    // FormattingAssembler.ResolveInsideBorder and WmlToHtmlConverter.ResolveCellBorder cast the
    // absent w:sz to a value type (only "nil" was special-cased; "none" fell through). It must render.
    [Fact]
    public void HCO056_BorderlessTable_DoesNotCrashConverter()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(new Wp.Body());
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            var noneBorders = new Wp.TableBorders(
                new Wp.TopBorder { Val = Wp.BorderValues.None },
                new Wp.LeftBorder { Val = Wp.BorderValues.None },
                new Wp.BottomBorder { Val = Wp.BorderValues.None },
                new Wp.RightBorder { Val = Wp.BorderValues.None },
                new Wp.InsideHorizontalBorder { Val = Wp.BorderValues.None },
                new Wp.InsideVerticalBorder { Val = Wp.BorderValues.None });
            main.Document.Body!.Append(new Wp.Table(
                new Wp.TableProperties(noneBorders),
                new Wp.TableRow(
                    new Wp.TableCell(new Wp.Paragraph(new Wp.Run(new Wp.Text("LeftCellText")))),
                    new Wp.TableCell(new Wp.Paragraph(new Wp.Run(new Wp.Text("RightCellText")))))));
            main.Document.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(), new HtmlConversionOptions());
        Assert.Contains("LeftCellText", html);
        Assert.Contains("RightCellText", html);
    }

    // Minimal OOXML packages (document.xml + styles.xml only — no word/settings.xml) are legal:
    // ECMA-376 does not require DocumentSettingsPart, and Word opens them without repair.
    // CalculateSpanWidthForTabs used to call DocumentSettingsPart.GetXDocument() unconditionally,
    // which threw ArgumentNullException("part") and aborted conversion. Default tab stop is 720 twips.
    [Fact]
    public void HCO057_MissingDocumentSettingsPart_DoesNotCrashConverter()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.Run(
                            new Wp.Text("Hello no-settings package")))));
            // Styles are required by FormattingAssembler; settings intentionally omitted.
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles(
                new Wp.DocDefaults(
                    new Wp.RunPropertiesDefault(
                        new Wp.RunPropertiesBaseStyle(
                            new Wp.RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" },
                            new Wp.FontSize { Val = "24" }))));
            main.Document.Save();
        }

        // Prove the part is absent (not just that we forgot to assert the repro shape).
        using (var reopen = WordprocessingDocument.Open(ms, false))
        {
            Assert.Null(reopen.MainDocumentPart!.DocumentSettingsPart);
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(), new HtmlConversionOptions());
        Assert.Contains("Hello no-settings package", html);
    }

    // CalculateSpanWidthForTabs (WmlToHtmlConverter.cs) computes a tab's rendered width from
    // w:defaultTabStop. This pins the actual numeric fallback (720 twips == 0.5in) that
    // HCO057 only proved didn't crash — i.e. the missing-settings path doesn't just avoid
    // throwing, it produces the SAME width Word itself defaults to for an unset tab stop.
    [Fact]
    public void HCO058_MissingDocumentSettingsPart_TabWidthDefaultsTo720Twips()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.Run(new Wp.TabChar()),
                        new Wp.Run(new Wp.Text("AfterTab")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            // DocumentSettingsPart intentionally omitted.
            main.Document.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(), new HtmlConversionOptions());

        Assert.Contains("AfterTab", html);
        // 720 twips (Word's implicit default tab stop) == 0.5in from position 0.
        Assert.Contains("margin: 0 0 0 0.50in", html);
    }

    // Same computation, but with an explicit DocumentSettingsPart that overrides
    // w:defaultTabStop — proves the "settingsPart != null" branch introduced by the same
    // refactor still reads the configured value correctly (not just the null-guard path).
    [Fact]
    public void HCO059_DocumentSettingsPartWithCustomDefaultTabStop_TabWidthUsesConfiguredValue()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.Run(new Wp.TabChar()),
                        new Wp.Run(new Wp.Text("AfterTab")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings =
                new Wp.Settings(new Wp.DefaultTabStop { Val = 1440 }); // 1 inch
            main.Document.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(), new HtmlConversionOptions());

        Assert.Contains("AfterTab", html);
        Assert.Contains("margin: 0 0 0 1.00in", html);
        Assert.DoesNotContain("margin: 0 0 0 0.50in", html);
    }

    // DocumentSettingsPart present but with no w:defaultTabStop element at all (legal — the
    // element is optional within w:settings). Must fall back to the same 720-twip default as
    // when the whole part is absent, not throw and not silently use 0.
    [Fact]
    public void HCO060_DocumentSettingsPartWithoutDefaultTabStopElement_FallsBackTo720Twips()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.Run(new Wp.TabChar()),
                        new Wp.Run(new Wp.Text("AfterTab")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings(); // no DefaultTabStop child
            main.Document.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(), new HtmlConversionOptions());

        Assert.Contains("AfterTab", html);
        Assert.Contains("margin: 0 0 0 0.50in", html);
    }

    // AddFormattingParts copies formatting parts into the RenderBlockHtml throwaway doc but no
    // longer invents a dummy DocumentSettingsPart. Regression: a source with no settings part
    // must still round-trip through RenderBlockHtml without crashing (converter defaults tab stop).
    [Fact]
    public void HCO061_RenderBlockHtml_SourceMissingDocumentSettingsPart_DoesNotCrash()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.Run(new Wp.Text("HCO061 block text")))));
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            // DocumentSettingsPart intentionally omitted from the source document.
            main.Document.Save();
        }
        byte[] bytes = ms.ToArray();

        using (var reopenStream = new MemoryStream(bytes))
        using (var reopen = WordprocessingDocument.Open(reopenStream, false))
        {
            Assert.Null(reopen.MainDocumentPart!.DocumentSettingsPart);
        }

        var opts = new HtmlConversionOptions { FabricateCssClasses = false };
        string full = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { StampAnchors = true, FabricateCssClasses = false });
        var anchorEl = System.Xml.Linq.XElement.Parse(full).Descendants()
            .First(e => (string?)e.Attribute("data-anchor") != null);
        string anchorId = (string)anchorEl.Attribute("data-anchor")!;

        string block = HtmlConversionOps.RenderBlockHtml(bytes, anchorId, opts);

        Assert.Contains("HCO061 block text", block);
    }

    // Builds a document-only package (word/document.xml only — no styles, no settings) and
    // proves the repro shape: StyleDefinitionsPart really is absent after reopen.
    private static byte[] DocumentOnlyDocxBytes(string text)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.Run(new Wp.Text(text)))));
            main.Document.Save();
        }

        using (var reopen = WordprocessingDocument.Open(ms, false))
        {
            Assert.Null(reopen.MainDocumentPart!.StyleDefinitionsPart);
        }
        return ms.ToArray();
    }

    private static byte[] PageSizedDocxBytes(uint width, uint height)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(new Wp.Run(new Wp.Text("Page-sized test content"))),
                    new Wp.SectionProperties(
                        new Wp.PageSize { Width = width, Height = height },
                        new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 })));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static byte[] MixedPageSizedDocxBytes()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.ParagraphProperties(
                            new Wp.SectionProperties(
                                new Wp.PageSize { Width = 11906, Height = 16838 },
                                new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 })),
                        new Wp.Run(new Wp.Text("A4 section"))),
                    new Wp.Paragraph(new Wp.Run(new Wp.Text("Landscape section"))),
                    new Wp.SectionProperties(
                        new Wp.PageSize { Width = 15840, Height = 12240 },
                        new Wp.PageMargin { Top = 1440, Right = 1440, Bottom = 1440, Left = 1440 })));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    // Issue #265 — sibling of the missing-settings crash fixed in #264. word/styles.xml
    // (StyleDefinitionsPart) is also optional in OOXML: Word opens a document-only package
    // without repair, but FormattingAssembler.AssembleFormatting dereferenced
    // StyleDefinitionsPart unconditionally (many sites), throwing ArgumentNullException("part")
    // at WmlToHtmlConverter.cs's AssembleFormatting call — before any HTML was produced.
    [Fact]
    public void HCO062_MissingStyleDefinitionsPart_DoesNotCrashConverter()
    {
        byte[] bytes = DocumentOnlyDocxBytes("Hello no-styles package");

        string html = HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions());

        Assert.Contains("Hello no-styles package", html);
    }

    // Same crash, real-world shape: the RPR fixtures contain ONLY [Content_Types].xml,
    // _rels/.rels, and word/document.xml (no styles, no settings) and crashed on conversion.
    [Fact]
    public void HCO063_DocumentOnlyPackage_RprFixture_ConvertsToHtml()
    {
        byte[] bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles",
            "RPR-FivePageTestDoc.docx"));

        string html = HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions());

        Assert.Contains("Page 1 paragraph 1", html);
        Assert.Contains("Page 5 paragraph 1", html);
    }

    // RenderBlockHtml's throwaway doc copies the source's formatting parts; with no styles
    // part to copy, the single-block path must survive a styles-less source end to end.
    [Fact]
    public void HCO064_RenderBlockHtml_SourceMissingStyleDefinitionsPart_DoesNotCrash()
    {
        byte[] bytes = DocumentOnlyDocxBytes("HCO064 block text");

        var opts = new HtmlConversionOptions { StampAnchors = true, FabricateCssClasses = false };
        string full = HtmlConversionOps.ConvertToHtml(bytes, opts);
        var anchorEl = System.Xml.Linq.XElement.Parse(full).Descendants()
            .First(e => (string?)e.Attribute("data-anchor") != null);
        string anchorId = (string)anchorEl.Attribute("data-anchor")!;

        string block = HtmlConversionOps.RenderBlockHtml(bytes, anchorId, opts);

        Assert.Contains("HCO064 block text", block);
    }

    // Some producer packages use a VML <v:imagedata> relationship id for a non-image part. The
    // unsupported VML image is safely omitted, but it must not cast CustomXmlPart to ImagePart and
    // take down the document conversion (the complex_style_attr benchmark fixtures have this shape).
    [Fact]
    public void HCO065_VmlImageDataReferencingCustomXmlPart_DoesNotCrashConverter()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var customXml = main.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using (var customWriter = new StreamWriter(customXml.GetStream(FileMode.Create, FileAccess.Write)))
                customWriter.Write("<payload/>");
            string relationshipId = main.GetIdOfPart(customXml);

            using (var documentWriter = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                documentWriter.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                    "xmlns:v=\"urn:schemas-microsoft-com:vml\"><w:body><w:p><w:r><w:t>" +
                    "HCO065 retained text</w:t></w:r><w:r><w:pict><v:shape style=\"width:10pt;height:10pt\">" +
                    $"<v:imagedata r:id=\"{relationshipId}\"/></v:shape></w:pict></w:r></w:p>" +
                    "<w:sectPr/></w:body></w:document>");
            }
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(), new HtmlConversionOptions());

        Assert.Contains("HCO065 retained text", html);
    }

    // Word tolerates malformed pPr payloads where lineRule is present but the required line value
    // is absent (including documents with duplicate pPr elements). Treat it as the implicit browser
    // line-height instead of casting the missing attribute and aborting the complete conversion.
    [Fact]
    public void HCO066_AutoLineRuleWithoutLineValue_DoesNotCrashConverter()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var documentWriter = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                documentWriter.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body><w:p><w:pPr><w:spacing w:lineRule=\"auto\"/></w:pPr><w:pPr/>" +
                    "<w:r><w:t>HCO066 retained text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
            }
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(), new HtmlConversionOptions());

        Assert.Contains("HCO066 retained text", html);
        Assert.DoesNotContain("line-height", html);
    }

    // Strict/compatibility producers can express paragraph spacing as fractional point measures
    // rather than raw twips. It must use the same measure parser for before and line spacing so
    // one style default cannot abort every paragraph in the document.
    [Fact]
    public void HCO073_PointSuffixedAutoLineSpacing_ConvertsToPercent()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body><w:p><w:pPr><w:spacing w:before=\"2pt\" w:line=\"12.95pt\" w:lineRule=\"auto\"/>" +
                    "</w:pPr><w:r><w:t>HCO073 retained text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
            }
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(),
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO073 retained text", html);
        Assert.Contains("line-height: 107.9%", html);
    }

    // Table indentation and preceding paragraph spacing can use point measures too. A table-cell
    // fill with no explicit shading pattern is likewise a common Word-compatible clear shading
    // form. Normalize both shapes without throwing while probing the shade mapper.
    [Fact]
    public void HCO074_CellFillWithoutShadeValue_RendersClearFill()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body><w:p><w:pPr><w:spacing w:after=\"8pt\"/></w:pPr><w:r><w:t>HCO074 preceding text</w:t></w:r></w:p>" +
                    "<w:tbl><w:tblPr><w:tblInd w:w=\"0pt\" w:type=\"dxa\"/></w:tblPr>" +
                    "<w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid><w:tr><w:tc>" +
                    "<w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/><w:shd w:fill=\"D9EAF7\"/></w:tcPr>" +
                    "<w:p><w:r><w:t>HCO074 retained text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
                    "<w:sectPr/></w:body></w:document>");
            }
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
        }

        string html = HtmlConversionOps.ConvertToHtml(ms.ToArray(),
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO074 retained text", html);
        Assert.Contains("background: #D9EAF7", html);
    }

    // The viewer's byte-based HTML bridge must open Strict OOXML packages just as DocxDiff does.
    // Exercise both full-document and anchor-addressed block rendering; before normalization the
    // converter sees no transitional w:body and throws on these packages.
    [Fact]
    public void HCO067_StrictOoxml_NormalizesBeforeFullAndBlockRender()
    {
        byte[] strict = StrictDocumentOnlyDocxBytes("HCO067 strict retained text");
        var options = new HtmlConversionOptions { StampAnchors = true, FabricateCssClasses = false };

        string full = HtmlConversionOps.ConvertToHtml(strict, options);
        Assert.Contains("HCO067 strict retained text", full);
        var anchor = System.Xml.Linq.XElement.Parse(full).Descendants()
            .First(e => (string?)e.Attribute("data-anchor") != null)
            .Attribute("data-anchor")!.Value;

        string block = HtmlConversionOps.RenderBlockHtml(strict, anchor, options);
        Assert.Contains("HCO067 strict retained text", block);
    }

    // Word writes one text box twice inside mc:AlternateContent: a modern DrawingML/wps branch
    // and a VML fallback. The renderer must select the supported modern branch exactly once,
    // retain the visible text, and not double the logical box in HTML.
    [Fact]
    public void HCO068_ModernDrawingMlTextBox_RendersChoiceWithoutVmlDuplicate()
    {
        byte[] bytes = TextBoxDocxBytes(
            "<mc:AlternateContent>" +
            "<mc:Choice Requires=\"wps\"><w:drawing><wp:inline><wp:extent cx=\"1524000\" cy=\"762000\"/>" +
            "<a:graphic><a:graphicData><wps:wsp><wps:spPr><a:solidFill><a:srgbClr val=\"FFFFFF\"/>" +
            "</a:solidFill><a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/>" +
            "</a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>" +
            "HCO068 modern text box</w:t></w:r></w:p></w:txbxContent></wps:txbx>" +
            "<wps:bodyPr lIns=\"91440\" tIns=\"45720\" rIns=\"91440\" bIns=\"45720\"><a:spAutoFit/>" +
            "</wps:bodyPr>" +
            "</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>" +
            "<mc:Fallback><w:pict><v:shape style=\"width:120pt;height:60pt\"><v:textbox>" +
            "<w:txbxContent><w:p><w:r><w:t>HCO068 fallback text box</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>" +
            "</mc:AlternateContent>");

        string html = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO068 modern text box", html);
        Assert.DoesNotContain("HCO068 fallback text box", html);
        Assert.Contains("width: 120pt", html);
        Assert.DoesNotContain("height: 60pt", html);
        Assert.Contains("margin-bottom: 0", html);
    }

    // Some legacy Word documents keep the modern DrawingML text-box body in a related XML
    // part. The synthetic package deliberately uses a distinct VML fallback so this verifies
    // that the supported choice gains its external body without rendering both copies.
    [Fact]
    public void HCO075_ExternalDrawingMlTextBox_RendersChoiceWithoutVmlDuplicate()
    {
        byte[] bytes = ExternalTextBoxDocxBytes();

        string html = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO075 external textbox text", html);
        Assert.DoesNotContain("HCO075 fallback text box", html);
        Assert.Contains("width: 120pt", html);
    }

    // Old Office 2008 wps markup is not a namespace this renderer understands. Markup
    // Compatibility requires selecting its portable VML fallback rather than dropping the
    // entire logical text box.
    [Fact]
    public void HCO069_LegacyDrawingMlTextBox_RendersVmlFallback()
    {
        byte[] bytes = TextBoxDocxBytes(
            "<mc:AlternateContent>" +
            "<mc:Choice Requires=\"legacywps\"><w:drawing><wp:inline><wp:extent cx=\"1524000\" cy=\"762000\"/>" +
            "<a:graphic><a:graphicData><legacywps:wsp/></a:graphicData></a:graphic>" +
            "</wp:inline></w:drawing></mc:Choice>" +
            "<mc:Fallback><w:pict><v:shape style=\"width:100pt;height:40pt\"><v:textbox>" +
            "<w:txbxContent><w:p><w:r><w:t>HCO069 legacy fallback text box</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>" +
            "</mc:AlternateContent>");

        string html = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO069 legacy fallback text box", html);
        Assert.Contains("width: 100pt", html);
        Assert.Contains("height: 40pt", html);
    }

    // A direct VML text box is not an AlternateContent compatibility fallback and remains a
    // supported, standalone shape. Preserve its content and size in the HTML projection.
    [Fact]
    public void HCO070_DirectVmlTextBox_RendersTextAndDimensions()
    {
        byte[] bytes = TextBoxDocxBytes(
            "<w:pict><v:shape style=\"width:100pt;height:40pt\"><v:textbox><w:txbxContent>" +
            "<w:p><w:r><w:t>HCO070 direct VML text box</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict>");

        string html = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO070 direct VML text box", html);
        Assert.Contains("width: 100pt", html);
        Assert.Contains("height: 40pt", html);
    }

    // VML theme colour values can append Word palette metadata. Keep the colour itself, but do
    // not feed the suffix (or arbitrary CSS declarations) through to the generated style string.
    [Fact]
    public void HCO071_VmlThemeColors_NormalizeAndRejectStyleInjection()
    {
        byte[] themed = TextBoxDocxBytes(
            "<w:pict><v:shape style=\"width:100pt;height:40pt\" fillcolor=\"#156082 [3204]\" " +
            "strokecolor=\"white [3212]\"><v:textbox><w:txbxContent><w:p><w:r><w:t>" +
            "HCO071 themed VML text box</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>" +
            "</w:pict>");
        byte[] unsafeColor = TextBoxDocxBytes(
            "<w:pict><v:shape style=\"width:100pt;height:40pt\" fillcolor=\"red; color: blue\"><v:textbox>" +
            "<w:txbxContent><w:p><w:r><w:t>HCO071 unsafe VML text box</w:t></w:r></w:p></w:txbxContent>" +
            "</v:textbox></v:shape></w:pict>");

        string themedHtml = HtmlConversionOps.ConvertToHtml(themed,
            new HtmlConversionOptions { FabricateCssClasses = false });
        string unsafeHtml = HtmlConversionOps.ConvertToHtml(unsafeColor,
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("background-color: #156082", themedHtml);
        Assert.Contains("border: 1pt solid white", themedHtml);
        Assert.DoesNotContain("3204", themedHtml);
        Assert.DoesNotContain("color: blue", unsafeHtml);
    }

    [Fact]
    public void HCO072_DirectAutoFitVmlTextBox_DropsStoredHeightAndTrailingSpacing()
    {
        byte[] bytes = TextBoxDocxBytes(
            "<w:pict><v:shape style=\"width:100pt;height:40pt\"><v:textbox style=\"mso-fit-shape-to-text:t\">" +
            "<w:txbxContent><w:p><w:r><w:t>HCO072 auto-fit VML text box</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict>");

        string html = HtmlConversionOps.ConvertToHtml(bytes,
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO072 auto-fit VML text box", html);
        Assert.DoesNotContain("height: 40pt", html);
        Assert.Contains("margin-bottom: 0", html);
    }

    // The clean-view fast path must recognize every revision family the accepter handles. A cell
    // deletion has no w:ins/w:del wrapper, so the former body-only detector skipped acceptance and
    // leaked its text into the supposedly accepted HTML.
    [Fact]
    public void HCO073_CleanView_AcceptsCellDeletionRevision()
    {
        string html = HtmlConversionOps.ConvertToHtml(CellDeletionTableDocxBytes(),
            new HtmlConversionOptions { FabricateCssClasses = false });

        Assert.Contains("HCO073 retained cell", html);
        Assert.DoesNotContain("HCO073 deleted cell", html);
    }

    private static byte[] CellDeletionTableDocxBytes()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2400\"/><w:gridCol w:w=\"2400\"/>" +
                    "</w:tblGrid><w:tr>" +
                    "<w:tc><w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>" +
                    "HCO073 retained cell</w:t></w:r></w:p></w:tc>" +
                    "<w:tc><w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/><w:cellDel w:id=\"1\" " +
                    "w:author=\"Test\" w:date=\"2026-01-01T00:00:00Z\"/></w:tcPr><w:p><w:r><w:t>" +
                    "HCO073 deleted cell</w:t></w:r></w:p></w:tc>" +
                    "</w:tr></w:tbl><w:sectPr/></w:body></w:document>");
            }
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
        }
        return ms.ToArray();
    }

    private static byte[] TextBoxDocxBytes(string runContent)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
                    "xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                    "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                    "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" " +
                    "xmlns:legacywps=\"http://schemas.microsoft.com/office/word/2008/6/28/wordprocessingShape\">" +
                    "<w:body><w:p><w:r>" + runContent + "</w:r></w:p><w:sectPr/></w:body></w:document>");
            }
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
        }
        return ms.ToArray();
    }

    private static byte[] ExternalTextBoxDocxBytes()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var externalTextBox = main.AddExtendedPart(
                "http://schemas.microsoft.com/office/2006/relationships/txbx",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.txbx+xml",
                ".xml",
                "rIdExternal");
            using (var writer = new StreamWriter(externalTextBox.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(
                    "<w14:txbx xmlns:w14=\"http://schemas.microsoft.com/office/word/2008/9/12/wordml\" " +
                    "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
                    "<w:p><w:r><w:t>HCO075 external textbox text</w:t></w:r></w:p></w14:txbx>");
            }

            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
            {
                writer.Write(
                    "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
                    "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
                    "xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
                    "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                    "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
                    "<w:body><w:p><w:r><mc:AlternateContent>" +
                    "<mc:Choice Requires=\"wps\"><w:drawing><wp:inline><wp:extent cx=\"1524000\" cy=\"762000\"/>" +
                    "<a:graphic><a:graphicData><wps:wsp><wps:txbx r:txbx=\"rIdExternal\"/>" +
                    "</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>" +
                    "<mc:Fallback><w:pict><v:shape style=\"width:120pt;height:60pt\"><v:textbox>" +
                    "<w:txbxContent><w:p><w:r><w:t>HCO075 fallback text box</w:t></w:r></w:p>" +
                    "</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>" +
                    "</mc:AlternateContent></w:r></w:p><w:sectPr/></w:body></w:document>");
            }
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();
            doc.Save();
        }
        return ms.ToArray();
    }
}
