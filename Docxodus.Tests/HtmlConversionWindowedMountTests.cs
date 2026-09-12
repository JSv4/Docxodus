// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Wp = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Issue #776: the engine side of the editor's windowed mount — a render plan that
/// carries section and border-box grouping, a chrome render without body units, and a range
/// render that lays a window out exactly as the full render does.</summary>
public class HtmlConversionWindowedMountTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Profile = "{\"comments\":true,\"stampAnchors\":true}";

    private static byte[] Complicated() =>
        File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles", "HC031-Complicated-Document.docx"));

    /// <summary>Two sections (a body paragraph carries the break), a table whose cell closes a
    /// third, and three bordered paragraphs of which the first two share a box.</summary>
    private static byte[] BuildSectionsAndBoxes()
    {
        Wp.Paragraph Bordered(string text, string color) => new(
            new Wp.ParagraphProperties(new Wp.ParagraphBorders(
                new Wp.TopBorder { Val = Wp.BorderValues.Single, Size = 8, Color = color },
                new Wp.BottomBorder { Val = Wp.BorderValues.Single, Size = 8, Color = color })),
            new Wp.Run(new Wp.Text(text)));
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(new Wp.Body(
                new Wp.Paragraph(new Wp.Run(new Wp.Text("first section"))),
                Bordered("boxed one", "FF0000"),
                Bordered("boxed two", "FF0000"),
                Bordered("boxed three", "0000FF"),
                new Wp.Paragraph(
                    new Wp.ParagraphProperties(new Wp.SectionProperties(
                        new Wp.PageSize { Width = 12240, Height = 15840 })),
                    new Wp.Run(new Wp.Text("closes the first section"))),
                new Wp.Paragraph(new Wp.Run(new Wp.Text("second section"))),
                new Wp.Table(
                    new Wp.TableRow(new Wp.TableCell(new Wp.Paragraph(
                        new Wp.ParagraphProperties(new Wp.SectionProperties(
                            new Wp.PageSize { Width = 15840, Height = 12240, Orient = Wp.PageOrientationValues.Landscape })),
                        new Wp.Run(new Wp.Text("cell closes the second section")))))),
                new Wp.Paragraph(new Wp.Run(new Wp.Text("third section"))),
                new Wp.SectionProperties(new Wp.PageSize { Width = 12240, Height = 15840 })));
        }
        return ms.ToArray();
    }

    [Fact]
    public void WM001_RenderPlanCarriesSectionsAndBorderGroups()
    {
        using var session = new DocxSession(BuildSectionsAndBoxes());
        var plan = session.ListBlocks();
        var texts = plan.Body.Select(unit => session.Project().AnchorIndex[unit.Id].TextPreview.Trim()).ToList();
        Assert.Equal(new[] { "first section", "boxed one", "boxed two", "boxed three",
            "closes the first section", "second section", "cell closes the second section", "third section" }, texts);
        // A block belongs to the section it closes; the next block starts the next one, whether
        // the break sits on a body paragraph or on a paragraph inside a table cell.
        Assert.Equal(new[] { 0, 0, 0, 0, 0, 1, 1, 2 }, plan.Body.Select(unit => unit.Section));
        // Adjacent paragraphs with the same visible border share one box; a different border,
        // an unbordered paragraph and a table each start a group of their own.
        var groups = plan.Body.Select(unit => unit.Group).ToList();
        Assert.Equal(groups[1], groups[2]);
        Assert.NotEqual(groups[2], groups[3]);
        Assert.Equal(groups.Count - 1, groups.Distinct().Count());

        using var json = JsonDocument.Parse(DocxSessionJson.SerializeRenderPlan(plan));
        var second = json.RootElement.GetProperty("body")[1];
        Assert.Equal(0, second.GetProperty("section").GetInt32());
        Assert.Equal(groups[1], second.GetProperty("group").GetInt32());
    }

    [Fact]
    public void WM002_ChromeRenderIsTheFullRenderWithoutItsBodyUnits()
    {
        using var session = new DocxSession(Complicated());
        var handle = SessionRegistry.OpenSession(Complicated(), new DocxSessionSettings());
        try
        {
            var full = Parse(DocxSessionOps.RenderEditorHtml(handle, Profile));
            var chrome = Parse(DocxSessionOps.RenderEditorChromeHtml(handle, Profile));

            // Same stylesheet, same section wrappers with the same page geometry.
            Assert.Equal(StyleText(full), StyleText(chrome));
            var fullSections = Sections(full);
            var chromeSections = Sections(chrome);
            string Brief(XElement e) => $"#{e.Attribute("data-section-index")?.Value} {e.Attribute("data-page-width")?.Value}x{e.Attribute("data-page-height")?.Value} units={UnitCount(e)}";
            Assert.True(fullSections.Count == chromeSections.Count,
                "full: " + string.Join(" | ", fullSections.Select(Brief))
                + "\nchrome: " + string.Join(" | ", chromeSections.Select(Brief)));
            Assert.Equal(4, fullSections.Count);
            for (int i = 0; i < fullSections.Count; i++)
                Assert.Equal(AttributeSignature(fullSections[i]), AttributeSignature(chromeSections[i]));

            // Same footnote and endnote sections, in citation order, identical markup.
            foreach (var kind in new[] { "footnotes", "endnotes" })
            {
                var fullNotes = full.Descendants().Single(e => (string?)e.Attribute("class") == kind);
                var chromeNotes = chrome.Descendants().Single(e => (string?)e.Attribute("class") == kind);
                Assert.Equal(fullNotes.ToString(SaveOptions.DisableFormatting), chromeNotes.ToString(SaveOptions.DisableFormatting));
            }

            // No body unit survived the reduction: the wrappers hold only carriers, which are
            // empty paragraphs the anchor map does not know.
            var plan = session.ListBlocks();
            var unids = plan.Body.Select(unit => unit.Id.Split(':')[^1]).ToHashSet(StringComparer.Ordinal);
            var leftovers = chromeSections.SelectMany(section => section.Elements()).ToList();
            var report = string.Join(" | ", leftovers.Select(element =>
                $"{element.Name.LocalName} anchor={(string?)element.Attribute("data-anchor")} known={unids.Contains((string?)element.Attribute("data-anchor") ?? "")} text='{element.Value.Trim()}'"));
            Assert.True(leftovers.All(element => element.Name.LocalName == "p"
                && (string?)element.Attribute("data-anchor") == HtmlConversionOps.ChromeCarrierAnchor), report);
            Assert.DoesNotContain(HtmlConversionOps.ChromeCarrierAnchor, unids);
            Assert.Equal(117, Sections(full).Sum(section => UnitCount(section)));
        }
        finally
        {
            SessionRegistry.CloseSession(handle);
        }
    }

    /// <summary>Windows of 24 as the mount cuts them by default, and windows of one unit, which
    /// cut wherever a group allows — through the table of contents' field, between list items,
    /// beside every text box — and must still lay each unit out as the full render does.</summary>
    [Theory]
    [InlineData(24, false)]
    [InlineData(1, false)]
    [InlineData(24, true)]
    public void WM003_GroupAlignedWindowsReproduceTheFullRenderSectionBySection(int windowSize, bool paginated)
    {
        // The paginated profile positions anchored objects absolutely inside their page box;
        // a window rendered for a paginated mount must draw them the same way.
        var profile = paginated ? Profile.Replace("{", "{\"paginated\":true,") : Profile;
        var handle = SessionRegistry.OpenSession(Complicated(), new DocxSessionSettings());
        try
        {
            var full = Parse(DocxSessionOps.RenderEditorHtml(handle, profile));
            var plan = JsonDocument.Parse(DocxSessionOps.ListRenderedBlocks(handle, false)).RootElement
                .GetProperty("body").EnumerateArray()
                .Select(unit => (Id: unit.GetProperty("id").GetString()!, Section: unit.GetProperty("section").GetInt32(),
                    Group: unit.GetProperty("group").GetInt32()))
                .ToList();
            Assert.Equal(117, plan.Count);

            // Windows extended to the end of the group they stop in, exactly as the mount cuts them.
            var rendered = new List<(int Section, string Html)>();
            int windows = 0;
            for (int start = 0; start < plan.Count;)
            {
                windows++;
                int end = Math.Min(start + windowSize, plan.Count);
                while (end < plan.Count && plan[end].Group == plan[end - 1].Group) end++;
                var ids = plan.Skip(start).Take(end - start).Select(unit => unit.Id).ToList();
                var nodes = JsonDocument.Parse(DocxSessionOps.RenderEditorRangeHtml(handle,
                    JsonSerializer.Serialize(ids), profile)).RootElement.EnumerateArray()
                    .Select(node => node.GetString()!).ToList();
                var bySection = plan.Skip(start).Take(end - start).ToDictionary(u => u.Id.Split(':')[^1], u => u.Section);
                foreach (var html in nodes)
                {
                    var element = XElement.Parse(html);
                    var unit = element.DescendantsAndSelf().First(e => e.Attribute("data-anchor") is not null);
                    rendered.Add((bySection[(string)unit.Attribute("data-anchor")!], element.ToString(SaveOptions.DisableFormatting)));
                }
                start = end;
            }

            Assert.True(windows >= plan.Count / windowSize, $"{windows} windows of {windowSize}");

            // Every unit landed exactly once, and each section's children are the full render's,
            // byte for byte — wrappers, grouping and source anchor ids included.
            var fullSections = Sections(full);
            for (int i = 0; i < fullSections.Count; i++)
            {
                // Page view lays the endnotes section inside the last section wrapper; that is
                // chrome the mount already holds, not a unit a window renders.
                var expected = fullSections[i].Elements().Where(e => e.Name.LocalName != "section")
                    .Select(e => e.ToString(SaveOptions.DisableFormatting)).ToList();
                var actual = rendered.Where(r => r.Section == i).Select(r => r.Html).ToList();
                string Head(string html) => XElement.Parse(html).DescendantsAndSelf()
                    .First(e => e.Attribute("data-anchor") is not null).Attribute("data-anchor")!.Value[..6];
                string Markers(string html) => string.Join("", XElement.Parse(html).Descendants()
                    .Where(e => (string?)e.Attribute("data-list-marker") == "true").Select(e => e.Value));
                var markerDiffs = expected.Zip(actual).Where(p => Markers(p.First) != Markers(p.Second))
                    .Select(p => $"{Head(p.First)} full='{Markers(p.First)}' range='{Markers(p.Second)}'").ToList();
                Assert.True(expected.Count == actual.Count && expected.Zip(actual).All(pair => pair.First == pair.Second),
                    $"section {i}: expected {expected.Count} nodes got {actual.Count}; marker diffs: {string.Join(" ", markerDiffs)}"
                    + (expected.Count == actual.Count
                        ? "\nfirst differing node:\n" + expected.Zip(actual).Where(p => p.First != p.Second).Select(p =>
                        {
                            int at = Enumerable.Range(0, Math.Min(p.First.Length, p.Second.Length)).FirstOrDefault(k => p.First[k] != p.Second[k], -1);
                            int from = Math.Max(0, at - 160);
                            return $"at {at}\nfull:  {p.First.Substring(from, Math.Min(360, p.First.Length - from))}\nrange: {p.Second.Substring(from, Math.Min(360, p.Second.Length - from))}";
                        }).FirstOrDefault()
                        : ""));
            }
        }
        finally
        {
            SessionRegistry.CloseSession(handle);
        }
    }

    /// <summary>A border a paragraph style contributes groups in the renderer, because formatting
    /// assembly folds the style chain into the paragraph before the border divs are made; the
    /// plan resolves the same chain, so two adjacent paragraphs of a bordered style share a
    /// group exactly as two with the same direct border do, and a window never lands between
    /// them. The fixture holds one box of each kind, a differently bordered neighbour and plain
    /// paragraphs around them.</summary>
    [Fact]
    public void WM004_PlanGroupsFollowStyleBordersLikeTheRenderer()
    {
        var bytes = File.ReadAllBytes(Path.Combine("..", "..", "..", "..", "TestFiles", "WM001-Bordered-Boxes.docx"));
        using var session = new DocxSession(bytes);
        var handle = SessionRegistry.OpenSession(bytes, new DocxSessionSettings());
        try
        {
            var plan = session.ListBlocks().Body;
            var groups = plan.Select(unit => unit.Group).ToList();
            Assert.Equal(8, plan.Count);
            Assert.Equal(groups[1], groups[2]);
            Assert.Equal(groups[4], groups[5]);
            Assert.Equal(groups.Count - 2, groups.Distinct().Count());

            // The renderer draws each pair inside one border div, which the range render hands
            // back as one node holding both units.
            var full = Parse(DocxSessionOps.RenderEditorHtml(handle, Profile));
            foreach (var (first, second) in new[] { (plan[1], plan[2]), (plan[4], plan[5]) })
            {
                var box = full.Descendants().Single(e => (string?)e.Attribute("data-anchor") == first.Id.Split(':')[^1]).Parent!;
                Assert.Equal("div", box.Name.LocalName);
                Assert.Contains(box.Elements(), e => (string?)e.Attribute("data-anchor") == second.Id.Split(':')[^1]);
                var nodes = JsonDocument.Parse(DocxSessionOps.RenderEditorRangeHtml(handle,
                    JsonSerializer.Serialize(new[] { first.Id, second.Id }), Profile)).RootElement.EnumerateArray().ToList();
                Assert.Single(nodes);
                Assert.Equal(box.ToString(SaveOptions.DisableFormatting), XElement.Parse(nodes[0].GetString()!).ToString(SaveOptions.DisableFormatting));
            }
        }
        finally
        {
            SessionRegistry.CloseSession(handle);
        }
    }

    private static XElement Parse(string html)
    {
        if (html.TrimStart().StartsWith('{')) throw new Xunit.Sdk.XunitException(html);
        return XElement.Parse(html);
    }

    private static string StyleText(XElement html) =>
        string.Concat(html.Descendants().Where(e => e.Name.LocalName == "style").Select(e => e.Value));

    private static List<XElement> Sections(XElement html) =>
        html.Descendants().Where(e => e.Attribute("data-section-index") is not null).ToList();

    private static string AttributeSignature(XElement element) =>
        string.Join(";", element.Attributes().OrderBy(a => a.Name.ToString(), StringComparer.Ordinal)
            .Select(a => $"{a.Name}={a.Value}"));

    private static int UnitCount(XElement section) =>
        section.Descendants().Count(e => e.Attribute("data-anchor") is not null
            && !e.Ancestors().TakeWhile(a => !ReferenceEquals(a, section)).Any(a => a.Attribute("data-anchor") is not null));
}
