using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class LegacyFormFieldRenderingTests
{
    [Theory]
    [InlineData("0", null, "false")]
    [InlineData("1", null, "true")]
    [InlineData("1", "0", "false")]
    [InlineData("0", "1", "true")]
    [InlineData("false", "on", "true")]
    [InlineData("true", "off", "false")]
    [InlineData("0", "", "true")]
    public void CheckboxShowsTheExplicitStateOrDefault(string defaultValue, string? explicitValue, string expected)
    {
        var properties = new XElement(W.checkBox, new XElement(W.sizeAuto),
            new XElement(W._default, new XAttribute(W.val, defaultValue)),
            explicitValue == null ? null : new XElement(W._checked,
                explicitValue == "" ? null : new XAttribute(W.val, explicitValue)));
        var html = Render("FORMCHECKBOX", properties, placeholders: false);
        var checkbox = Assert.Single(html.Descendants().Where(e => (string?)e.Attribute("role") == "checkbox"));
        Assert.Equal(expected, (string?)checkbox.Attribute("aria-checked"));
        Assert.True(checkbox.Value.Length > 0 || checkbox.Descendants().Any(e => e.Name.LocalName == "svg"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CheckboxSizeUsesTheFieldCharactersUnlessExplicitlyFixed(bool fixedSize)
    {
        var properties = new XElement(W.checkBox,
            fixedSize ? new XElement(W.size, new XAttribute(W.val, "20")) : new XElement(W.sizeAuto),
            new XElement(W._checked));
        var html = Render("FORMCHECKBOX", properties, "stale", mutate: p =>
            p.Descendants(W.ffData).Single().Parent!.Parent!.AddFirst(new XElement(W.rPr,
                new XElement(W.sz, new XAttribute(W.val, "40")))));
        var checkbox = Assert.Single(html.Descendants().Where(e => (string?)e.Attribute("role") == "checkbox"));
        var sizeOwner = fixedSize ? checkbox : checkbox.Ancestors(Xhtml.span).First();
        Assert.Matches(fixedSize ? @"font-size:\s*10(\.0+)?pt" : @"font-size:\s*20(\.0+)?pt",
            (string?)sizeOwner.Attribute("style") ?? "");
        Assert.DoesNotContain("stale", html.Descendants(Xhtml.body).Single().Value);
    }

    [Theory]
    [InlineData("FORMCHECKBOX")]
    [InlineData("FORMDROPDOWN")]
    public void MismatchedFormMetadataIsReportedWithoutLosingItsCache(string instruction)
    {
        var html = Render(instruction, new XElement(W.textInput), "Retained text");
        var marker = Assert.Single(html.Descendants().Where(e => (string?)e.Attribute("data-content-type") == "FormField"));
        marker.Remove();
        Assert.Equal("Before Retained text After", html.Descendants(Xhtml.body).Single().Value);
    }

    [Theory]
    [InlineData(null, "1", "Copper bin")]
    [InlineData("0", "1", "Amber rack")]
    [InlineData("1", "0", "Copper bin")]
    public void DropdownShowsTheSelectionOnceEvenWithACachedResult(string? selection, string defaultValue, string expected)
    {
        var properties = new XElement(W.ddList,
            selection == null ? null : new XElement(W.result, new XAttribute(W.val, selection)),
            new XElement(W._default, new XAttribute(W.val, defaultValue)),
            new XElement(W.listEntry, new XAttribute(W.val, "Amber rack")),
            new XElement(W.listEntry, new XAttribute(W.val, "Copper bin")));
        foreach (var cached in new[] { "", expected })
        {
            var body = Render("FORMDROPDOWN", properties, cached).Descendants(Xhtml.body).Single();
            Assert.Equal("Before " + expected + " After", body.Value);
        }
    }

    [Theory]
    [InlineData("header")]
    [InlineData("footer")]
    public void DropdownMetadataRendersInRunningStories(string story)
    {
        var properties = new XElement(W.ddList,
            new XElement(W.result, new XAttribute(W.val, "1")),
            new XElement(W.listEntry, new XAttribute(W.val, "Amber rack")),
            new XElement(W.listEntry, new XAttribute(W.val, "Copper bin")));
        var html = Render("FORMDROPDOWN", properties, story: story);
        var runningStory = Assert.Single(html.Descendants().Where(e => e.Name.LocalName == story));
        Assert.Equal("Before Copper bin After", runningStory.Value);
    }

    [Fact]
    public void DropdownMetadataReplacesAStaleCacheWithoutDuplicatingTheValue()
    {
        var properties = new XElement(W.ddList,
            new XElement(W.result, new XAttribute(W.val, "1")),
            new XElement(W.listEntry, new XAttribute(W.val, "Amber rack")),
            new XElement(W.listEntry, new XAttribute(W.val, "Copper bin")));
        var body = Render("FORMDROPDOWN", properties, "Amber rack").Descendants(Xhtml.body).Single();
        Assert.Equal("Before Copper bin After", body.Value);
    }

    [Theory]
    [InlineData("Retained text")]
    [InlineData("")]
    public void UnsupportedFormTypesAreReportedAndRetainTheirCachedText(string cached)
    {
        var html = Render("FORMUNKNOWN", new XElement(W.textInput), cached);
        var marker = Assert.Single(html.Descendants().Where(e =>
            (string?)e.Attribute("data-content-type") == "FormField"));
        Assert.NotEmpty(marker.Value);
        marker.Remove();
        Assert.Equal("Before " + cached + " After", html.Descendants(Xhtml.body).Single().Value);
    }

    [Fact]
    public void ACompleteTextFormCacheIsRetainedWithoutAnUnsupportedWarning()
    {
        var html = Render("FORMTEXT", new XElement(W.textInput), "Retained text");
        Assert.Equal("Before Retained text After", html.Descendants(Xhtml.body).Single().Value);
        Assert.DoesNotContain(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "FormField");
    }

    [Theory]
    [InlineData("1", "Copper bin")]
    [InlineData("7", "Amber rack")]
    [InlineData(null, "Amber rack")]
    public void AnOutOfRangeSelectionFallsBackToTheDefaultThenTheFirstEntry(string? defaultValue, string expected)
    {
        var properties = new XElement(W.ddList,
            new XElement(W.result, new XAttribute(W.val, "7")),
            defaultValue == null ? null : new XElement(W._default, new XAttribute(W.val, defaultValue)),
            new XElement(W.listEntry, new XAttribute(W.val, "Amber rack")),
            new XElement(W.listEntry, new XAttribute(W.val, "Copper bin")));
        var body = Render("FORMDROPDOWN", properties).Descendants(Xhtml.body).Single();
        Assert.Equal("Before " + expected + " After", body.Value);
    }

    [Fact]
    public void ADropdownWithNoEntriesOrCachedValueIsReported()
    {
        var html = Render("FORMDROPDOWN", new XElement(W.ddList));
        Assert.Single(html.Descendants().Where(e =>
            (string?)e.Attribute("data-content-type") == "FormField"));
    }

    private static XElement Dropdown() => new XElement(W.ddList,
        new XElement(W.result, new XAttribute(W.val, "1")),
        new XElement(W.listEntry, new XAttribute(W.val, "Amber rack")),
        new XElement(W.listEntry, new XAttribute(W.val, "Copper bin")));

    [Fact]
    public void AReplacementKeepsItsCachedRunFormatting()
    {
        var html = Render("FORMDROPDOWN", Dropdown(), "stale", mutate: p =>
            p.Elements(W.r).Single(r => r.Element(W.t)?.Value == "stale").AddFirst(
                new XElement(W.rPr, new XElement(W.b), new XElement(W.color, new XAttribute(W.val, "B40000")),
                    new XElement(W.sz, new XAttribute(W.val, "28")))));
        var result = html.Descendants(Xhtml.span).First(e => e.Value == "Copper bin");
        var style = (string?)result.Attribute("style") ?? "";
        Assert.Matches(@"font-weight:\s*bold", style);
        Assert.Matches(@"font-size:\s*14(\.0+)?pt", style);
        Assert.Matches(@"color:\s*#B40000", style);
    }

    [Fact]
    public void FieldCharactersInOneRunKeepTheTextBeforeAndAfterTheField()
    {
        var html = Render("FORMDROPDOWN", Dropdown(), "stale", mutate: p =>
            p.ReplaceNodes(new XElement(W.r, p.Elements(W.r).Elements().ToList())));
        Assert.Equal("Before Copper bin After", html.Descendants(Xhtml.body).Single().Value);
    }

    [Fact]
    public void AdjacentFieldsUseTheirOwnMetadata()
    {
        var html = Render("FORMDROPDOWN", Dropdown(), mutate: p =>
        {
            var second = p.Elements(W.r).Skip(1).Take(5).Select(r => new XElement(r)).ToList();
            second.SelectMany(r => r.Descendants(W.result)).Single().SetAttributeValue(W.val, "0");
            p.Elements().Last().AddBeforeSelf(second);
        });
        Assert.Equal("Before Copper binAmber rack After", html.Descendants(Xhtml.body).Single().Value);
    }

    [Fact]
    public void AFormInsideAHyperlinkFieldKeepsItsValueAndLink()
    {
        var html = Render("FORMDROPDOWN", Dropdown(), mutate: p =>
        {
            var runs = p.Elements().ToList();
            p.ReplaceNodes(runs[0],
                new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"))),
                new XElement(W.r, new XElement(W.instrText, "HYPERLINK \"https://example.test/rack\"")),
                new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate"))),
                runs.Skip(1).SkipLast(1),
                new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))), runs[^1]);
        });
        var link = Assert.Single(html.Descendants(Xhtml.a).Where(e =>
            (string?)e.Attribute("href") == "https://example.test/rack"));
        Assert.Equal("Copper bin", link.Value);
        Assert.Equal("Before Copper bin After", html.Descendants(Xhtml.body).Single().Value);
    }

    [Theory]
    [InlineData("ins", true)]
    [InlineData("del", true)]
    [InlineData("del", false)]
    public void ResolvedValuesRespectTrackedRevisionVisibility(string kind, bool showDeleted)
    {
        var html = Render("FORMDROPDOWN", Dropdown(), showDeleted: showDeleted, mutate: p =>
        {
            var runs = p.Elements().Skip(1).SkipLast(1).ToList();
            foreach (var run in runs) run.Remove();
            if (kind == "del")
                foreach (var text in runs.SelectMany(r => r.Elements()).Where(e => e.Name == W.t || e.Name == W.instrText))
                    text.Name = text.Name == W.t ? W.delText : W.delInstrText;
            p.Elements().Last().AddBeforeSelf(new XElement(W.w + kind,
                new XAttribute(W.id, "1"), new XAttribute(W.author, "Cedar Review"), runs));
        });
        var visible = kind == "ins" || showDeleted;
        Assert.Equal(visible ? "Before Copper bin After" : "Before  After", html.Descendants(Xhtml.body).Single().Value);
        if (visible)
            Assert.Equal("Copper bin", Assert.Single(html.Descendants(Xhtml.xhtml + kind)).Value);
    }

    [Theory]
    [InlineData("")]
    [InlineData("stale")]
    public void ATrailingTabMeasuresTheResolvedValueEvenInTheFirstParagraph(string cached)
    {
        var selected = string.Join(" ", Enumerable.Repeat("Copper bin", 5));
        XElement WithTab(string cache)
        {
            var dropdown = Dropdown();
            dropdown.Elements(W.listEntry).Last().SetAttributeValue(W.val, selected);
            return Render("FORMDROPDOWN", dropdown, cache, mutate: p =>
            {
                p.AddFirst(new XElement(W.pPr, new XElement(W.tabs, new XElement(W.tab,
                    new XAttribute(W.val, "left"), new XAttribute(W.pos, "2880")))));
                p.Elements().Last().AddBeforeSelf(new XElement(W.r, new XElement(W.tab)));
            });
        }
        var actual = WithTab(cached);
        var control = WithTab(selected);
        Assert.Equal(control.Descendants(Xhtml.body).Single().Value, actual.Descendants(Xhtml.body).Single().Value);
        Assert.Equal(control.Descendants().Where(e => e.Attribute("data-docx-tab") != null).Select(e => e.ToString()),
            actual.Descendants().Where(e => e.Attribute("data-docx-tab") != null).Select(e => e.ToString()));
        Assert.NotEmpty(actual.Descendants().Where(e => e.Attribute("data-docx-tab") != null));
    }

    [Theory]
    [InlineData(true, "Before Old value After")]
    [InlineData(false, "Before  After")]
    public void ADeletedResultIsNotRegeneratedOutsideItsRevision(bool showDeleted, string expected)
    {
        var html = Render("FORMDROPDOWN", Dropdown(), "Old value", showDeleted: showDeleted, mutate: p =>
        {
            var run = p.Elements(W.r).Single(r => r.Element(W.t)?.Value == "Old value");
            run.Element(W.t)!.Name = W.delText;
            run.ReplaceWith(new XElement(W.del, new XAttribute(W.id, "1"),
                new XAttribute(W.author, "Cedar Review"), new XElement(run)));
        });
        Assert.Equal(expected, html.Descendants(Xhtml.body).Single().Value);
    }

    [Fact]
    public void AFormInsideAnOuterInstructionDoesNotLeakIntoTheResult()
    {
        var html = Render("FORMDROPDOWN", Dropdown(), mutate: p =>
        {
            var runs = p.Elements().ToList();
            p.ReplaceNodes(runs[0],
                new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"))),
                new XElement(W.r, new XElement(W.instrText, "IF ")),
                runs.Skip(1).SkipLast(1),
                new XElement(W.r, new XElement(W.instrText, " = \"Copper bin\" \"Yes\" \"No\"")),
                new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate"))),
                new XElement(W.r, new XElement(W.t, "Yes")),
                new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))), runs[^1]);
        });
        Assert.Equal("Before Yes After", html.Descendants(Xhtml.body).Single().Value);
    }

    [Fact]
    public void AFieldOnlyParagraphNeedsNoSyntheticBlankText()
    {
        var html = Render("FORMDROPDOWN", Dropdown(), mutate: p =>
        {
            p.Elements().First().Remove();
            p.Elements().Last().Remove();
        });
        Assert.Equal("Copper bin", html.Descendants(Xhtml.body).Single().Value);
    }

    [Fact]
    public void BookmarksAndCommentReferencesInsideAResultSurviveReplacement()
    {
        var html = Render("FORMDROPDOWN", Dropdown(), "stale", mutate: p =>
        {
            var run = p.Elements(W.r).Single(r => r.Element(W.t)?.Value == "stale");
            run.AddBeforeSelf(new XElement(W.bookmarkStart, new XAttribute(W.id, "4"), new XAttribute(W.name, "RackTarget")),
                new XElement(W.commentRangeStart, new XAttribute(W.id, "7")));
            run.AddAfterSelf(new XElement(W.commentRangeEnd, new XAttribute(W.id, "7")),
                new XElement(W.bookmarkEnd, new XAttribute(W.id, "4")),
                new XElement(W.r, new XElement(W.commentReference, new XAttribute(W.id, "7"))));
        });
        Assert.Single(html.Descendants().Where(e => (string?)e.Attribute("id") == "RackTarget"));
        Assert.Single(html.Descendants().Where(e => (string?)e.Attribute("id") == "comment-ref-7"));
        Assert.Contains(html.Descendants(), e => e.Value == "Copper bin" && (string?)e.Attribute("data-comment-id") == "7");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnUnrelatedHeaderFieldDoesNotEnterFormNormalization(bool headers)
    {
        var html = Render("FORMDROPDOWN", Dropdown(), story: "header", headers: headers, mutate: p =>
            p.ReplaceNodes(new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))),
                new XElement(W.r, new XElement(W.t, "Inactive header"))));
        Assert.Equal(headers ? "Inactive headerBody control" : "Body control", html.Descendants(Xhtml.body).Single().Value);
    }

    private static XElement Render(string instruction, XElement properties, string cached = "", string? story = null,
        bool placeholders = true, bool showDeleted = true, bool headers = true, Action<XElement>? mutate = null)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var paragraph = new XElement(W.p,
                        new XElement(W.r, new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), "Before ")),
                        new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"),
                            new XElement(W.ffData, new XElement(W.name, new XAttribute(W.val, "CedarField")), properties))),
                        new XElement(W.r, new XElement(W.instrText, instruction)),
                        new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate"))),
                        new XElement(W.r, new XElement(W.t, cached)),
                        new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))),
                        new XElement(W.r, new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), " After")));
            mutate?.Invoke(paragraph);
            var body = new XElement(W.body);
            if (story == null)
                body.Add(paragraph);
            else
            {
                OpenXmlPart part = story == "header" ? main.AddNewPart<HeaderPart>() : main.AddNewPart<FooterPart>();
                part.PutXDocument(new XDocument(new XElement(story == "header" ? W.hdr : W.ftr,
                    new XAttribute(XNamespace.Xmlns + "w", W.w), paragraph)));
                body.Add(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Body control"))),
                    new XElement(W.sectPr, new XElement(story == "header" ? W.headerReference : W.footerReference,
                        new XAttribute(W.type, "default"), new XAttribute(R.id, main.GetIdOfPart(part)))));
            }
            main.PutXDocument(new XDocument(new XElement(W.document,
                new XAttribute(XNamespace.Xmlns + "w", W.w), body)));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles,
                new XAttribute(XNamespace.Xmlns + "w", W.w))));
            if (paragraph.Descendants(W.commentReference).Any())
                main.AddNewPart<WordprocessingCommentsPart>().PutXDocument(new XDocument(new XElement(W.comments,
                    new XElement(W.comment, new XAttribute(W.id, "7"), new XAttribute(W.author, "Cedar Review"),
                        new XElement(W.p, new XElement(W.r, new XElement(W.t, "Review rack")))))));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(document));
        }
        return XElement.Parse(HtmlConversionOps.ConvertToHtml(stream.ToArray(), new HtmlConversionOptions
        {
            RenderUnsupportedContentPlaceholders = placeholders,
            RenderHeadersAndFooters = headers,
            RenderTrackedChanges = true,
            ShowDeletedContent = showDeleted,
            FabricateCssClasses = false,
            CommentRenderMode = 1,
        }));
    }
}
