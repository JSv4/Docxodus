using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class DocumentFieldRenderingTests
{
    private const string Title = "Juniper equipment register";

    public static IEnumerable<object[]> PropertyCases()
    {
        foreach (var instruction in new[] { "TITLE", "DOCPROPERTY Title" })
        foreach (var simple in new[] { false, true })
        foreach (var story in new[] { "body", "header", "footer" })
        foreach (var cache in new[] { "", "Old title", Title })
            yield return new object[] { instruction, simple, story, cache };
    }

    [Theory]
    [MemberData(nameof(PropertyCases))]
    public void UnlockedTitleFieldsResolveThePropertyInEveryStory(string instruction, bool simple, string story, string cache)
    {
        var html = Render(instruction, simple, story, cache);
        Assert.Equal("Before " + Title + " After", Story(html, story).Value);
    }

    [Theory]
    [InlineData(false, "true", "Locked title")]
    [InlineData(false, "1", "Locked title")]
    [InlineData(false, "on", "")]
    [InlineData(true, "true", "Locked title")]
    [InlineData(true, "1", "Locked title")]
    [InlineData(true, "on", "")]
    public void LockedFieldsKeepTheirCacheEvenWhenThePropertyDiffers(bool simple, string locked, string cache)
    {
        var html = Render("TITLE", simple, "body", cache, locked);
        Assert.Equal("Before " + cache + " After", Story(html, "body").Value);
    }

    [Theory]
    [InlineData(false, "false")]
    [InlineData(false, "0")]
    [InlineData(false, "off")]
    [InlineData(true, "false")]
    [InlineData(true, "0")]
    [InlineData(true, "off")]
    public void FalseLockFlagsAllowPropertyResolution(bool simple, string locked)
    {
        Assert.Equal("Before " + Title + " After", Story(Render("TITLE", simple, "body", "Old title", locked), "body").Value);
    }

    [Theory]
    [InlineData(false, "body")]
    [InlineData(false, "header")]
    [InlineData(false, "footer")]
    [InlineData(true, "body")]
    [InlineData(true, "header")]
    [InlineData(true, "footer")]
    public void UnrelatedDynamicFieldsKeepTheirStoredResults(bool simple, string story)
    {
        Assert.Equal("Before Stored date After", Story(Render("DATE", simple, story, "Stored date"), story).Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void APropertyReplacementRetainsItsCachedRunFormatting(bool simple)
    {
        var html = Render("DOCPROPERTY \"Title\" \\* MERGEFORMAT", simple, "body", "Old title", mutate: paragraph =>
        {
            var run = paragraph.Descendants(W.r).Single(r => r.Element(W.t)?.Value == "Old title");
            run.AddFirst(new XElement(W.rPr, new XElement(W.b), new XElement(W.color, new XAttribute(W.val, "146532")),
                new XElement(W.sz, new XAttribute(W.val, "28"))));
        });
        var value = Assert.Single(html.Descendants(Xhtml.span), e => e.Value == Title);
        var style = (string?)value.Attribute("style") ?? "";
        Assert.Contains("font-weight:bold", style.Replace(" ", ""));
        Assert.Contains("#146532", style);
        Assert.Matches(@"font-size:\s*14(\.0+)?pt", style);
    }

    [Theory]
    [InlineData(false, "")]
    [InlineData(false, "Locked title")]
    [InlineData(true, "")]
    [InlineData(true, "Locked title")]
    public void DirtyDoesNotOverrideALock(bool simple, string cache)
    {
        var html = Render("TITLE", simple, "body", cache, "true", mutate: paragraph =>
            paragraph.Descendants(simple ? W.fldSimple : W.fldChar).First().SetAttributeValue(W.dirty, "true"));
        Assert.Equal("Before " + cache + " After", Story(html, "body").Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnExplicitEmptyPropertyClearsAStaleCache(bool simple)
    {
        Assert.Equal("Before  After", Story(Render("TITLE", simple, "body", "Old title", title: ""), "body").Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnAbsentPropertyPreservesTheCacheAndReportsIt(bool simple)
    {
        var html = Render("TITLE", simple, "body", "Old title", title: null, placeholders: true);
        RemoveDiagnostic(html);
        Assert.Equal("Before Old title After", Story(html, "body").Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnExistingCorePropertiesPartWithoutATitlePreservesTheCacheAndReportsIt(bool simple)
    {
        var html = Render("TITLE", simple, "body", "Old title", title: null,
            creator: "Juniper Author", placeholders: true);
        RemoveDiagnostic(html);
        Assert.Equal("Before Old title After", Story(html, "body").Value);
    }

    [Theory]
    [InlineData(false, "TITLE \"Changed title\"")]
    [InlineData(true, "TITLE \"Changed title\"")]
    [InlineData(false, "DOCPROPERTY MissingProperty")]
    [InlineData(true, "DOCPROPERTY MissingProperty")]
    [InlineData(false, "DOCPROPERTY Title \\Unknown")]
    [InlineData(true, "DOCPROPERTY Title \\Unknown")]
    public void UnsupportedPropertyInstructionsPreserveTheirCacheAndReportIt(bool simple, string instruction)
    {
        var html = Render(instruction, simple, "body", "Stored result", placeholders: true);
        RemoveDiagnostic(html);
        Assert.Equal("Before Stored result After", Story(html, "body").Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void APropertyFieldInsideAnotherFieldsInstructionDoesNotLeak(bool simple)
    {
        var html = Render("TITLE", simple, "body", "Old title", mutate: paragraph =>
        {
            var field = paragraph.Elements().Skip(1).SkipLast(1).ToArray();
            field[0].AddBeforeSelf(FieldChar("begin"), new XElement(W.r,
                new XElement(W.instrText, " IF \"")));
            field[^1].AddAfterSelf(new XElement(W.r, new XElement(W.instrText,
                "\" = \"" + Title + "\" \"Outer result\" \"Other result\" ")), FieldChar("separate"), new XElement(W.r,
                new XElement(W.t, "Outer result")), FieldChar("end"));
        });
        Assert.Equal("Before Outer result After", Story(html, "body").Value);
    }

    [Theory]
    [InlineData("TITLE")]
    [InlineData("DOCPROPERTY Title")]
    public void SharedRunFieldBoundariesPreserveSurroundingText(string instruction)
    {
        var html = Render(instruction, false, "body", "Old title", mutate: paragraph =>
        {
            var leaves = paragraph.Elements(W.r).SelectMany(run => run.Elements()).ToArray();
            paragraph.ReplaceNodes(new XElement(W.r, leaves));
        });
        Assert.Equal("Before " + Title + " After", Story(html, "body").Value);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void ADeletedCacheIsNotResurrectedOutsideItsRevision(bool simple, bool showDeleted)
    {
        var html = Render("TITLE", simple, "body", "Old title", showDeleted: showDeleted, mutate: paragraph =>
        {
            var text = paragraph.Descendants(W.t).Single(e => e.Value == "Old title");
            text.Name = W.delText;
            var run = text.Parent!;
            run.ReplaceWith(new XElement(W.del, new XAttribute(W.id, "1"),
                new XAttribute(W.author, "Juniper Review"), new XElement(run)));
        });
        Assert.Equal(showDeleted ? "Before Old title After" : "Before  After", Story(html, "body").Value);
    }

    [Theory]
    [InlineData(false, "DATE")]
    [InlineData(true, "DATE")]
    [InlineData(false, "HYPERLINK \"https://example.test/juniper\"")]
    [InlineData(true, "HYPERLINK \"https://example.test/juniper\"")]
    public void PropertyResolutionLeavesOtherFieldsInTheSameStoryIntact(bool simple, string instruction)
    {
        var html = Render("TITLE", simple, "body", "Old title", placeholders: true, mutate: paragraph =>
            paragraph.Elements().Last().AddBeforeSelf(
                new XElement(W.r, new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), " / ")),
                Field(instruction, simple, "Stored result")));
        Assert.DoesNotContain(html.Descendants(), e => e.Attribute("data-content-type") != null);
        if (!simple && instruction.StartsWith("HYPERLINK", StringComparison.Ordinal))
            Assert.Equal("Stored result", Assert.Single(html.Descendants(Xhtml.a),
                e => (string?)e.Attribute("href") == "https://example.test/juniper").Value);
        Assert.Equal("Before " + Title + " / Stored result After", Story(html, "body").Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnAuthoritativeOuterResultReplacesItsNestedCachedField(bool simple)
    {
        var html = Render("TITLE", simple, "body", "Old title", mutate: paragraph =>
            paragraph.Descendants(W.t).Single(e => e.Value == "Old title").Parent!
                .ReplaceWith(Field("DATE", !simple, "Nested stale result")));
        Assert.Equal("Before " + Title + " After", Story(html, "body").Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void APropertyFieldInsideAnotherFieldsResultRemainsVisible(bool simple)
    {
        var html = Render("TITLE", simple, "body", "Old title", mutate: paragraph =>
        {
            var field = paragraph.Elements().Skip(1).SkipLast(1).ToArray();
            field[0].AddBeforeSelf(FieldChar("begin"), new XElement(W.r,
                new XElement(W.instrText, " IF 1 = 1 \"Stored result\" \"Other result\" ")), FieldChar("separate"));
            field[^1].AddAfterSelf(FieldChar("end"));
        });
        Assert.Equal("Before " + Title + " After", Story(html, "body").Value);
    }

    private static XElement[] Field(string instruction, bool simple, string cache, string? locked = null)
    {
        var result = new XElement(W.r, new XElement(W.t, cache));
        return simple ? new[] { new XElement(W.fldSimple, new XAttribute(W.instr, instruction),
            locked == null ? null : new XAttribute(W.fldLock, locked), result) } : new[] {
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"),
                locked == null ? null : new XAttribute(W.fldLock, locked))),
            new XElement(W.r, new XElement(W.instrText, instruction)), FieldChar("separate"), result, FieldChar("end") };
    }

    [Fact]
    public void HiddenInstructionTextDoesNotRemoveAnEmptyParagraphsLineBox()
    {
        var html = Render("TITLE", false, "body", "Old title", mutate: paragraph =>
            paragraph.ReplaceNodes(ConditionalField("")));
        Assert.Equal("\u00a0", Story(html, "body").Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AReferenceInsideAFieldInstructionDoesNotRenderOrConsumeANumber(bool visibleNote)
    {
        var html = Render("TITLE", false, "body", "Old title", notes: true, mutate: paragraph =>
        {
            var citation = new XElement(W.r, new XElement(W.footnoteReference, new XAttribute(W.id, "1")));
            paragraph.ReplaceNodes(ConditionalField("Visible result", citation));
            if (visibleNote) paragraph.Add(new XElement(W.r, new XElement(W.footnoteReference, new XAttribute(W.id, "2"))));
        }, setup: (main, _) =>
        {
            XElement Note(int id) => new XElement(W.footnote, new XAttribute(W.id, id), new XElement(W.p,
                new XElement(W.r, new XElement(W.footnoteRef)), new XElement(W.r, new XElement(W.t, "Note " + id))));
            main.AddNewPart<FootnotesPart>().PutXDocument(new XDocument(new XElement(W.footnotes,
                Note(1), visibleNote ? Note(2) : null)));
        });
        var citations = Story(html, "body").Descendants(Xhtml.a).Where(e => e.Attribute("data-footnote-id") != null).ToList();
        if (visibleNote)
        {
            var citation = Assert.Single(citations);
            Assert.Equal("2", (string?)citation.Attribute("data-footnote-id"));
            Assert.Equal("1", citation.Value);
        }
        else Assert.Empty(citations);
        Assert.DoesNotContain(Story(html, "body").Descendants(Xhtml.a), e =>
            (string?)e.Attribute("href") == "#fn-ref-1");
    }

    [Fact]
    public void MergeFormatUsesTheVisibleCachedResultRatherThanAnInstructionCache()
    {
        var html = Render("TITLE \\* MERGEFORMAT", false, "body", "Old title", mutate: paragraph =>
        {
            paragraph.Descendants(W.t).Single(e => e.Value == "Old title").Parent!
                .ReplaceWith(ConditionalField("Visible cache"));
            foreach (var text in paragraph.Descendants(W.t).Where(e => e.Value is "Hidden title" or "Visible cache"))
                text.Parent!.AddFirst(new XElement(W.rPr, new XElement(W.color,
                    new XAttribute(W.val, text.Value == "Hidden title" ? "FF0000" : "008800"))));
        });
        var resolved = Assert.Single(html.Descendants(Xhtml.span), e => e.Value == Title);
        Assert.Contains("#008800", (string?)resolved.Attribute("style"));
    }

    [Fact]
    public void AnUnrelatedSplitKeywordDoesNotActivateFieldNormalization()
    {
        var html = Render("TITLECASE", false, "header", "Unrelated header", mutate: paragraph =>
            paragraph.Descendants(W.instrText).Single().Parent!.ReplaceWith(
                new XElement(W.r, new XElement(W.instrText, "TITLE")),
                new XElement(W.r, new XElement(W.rPr, new XElement(W.b)), new XElement(W.instrText, "CASE "))),
            setup: (main, _) =>
            {
                var header = main.HeaderParts.Single();
                header.GetXDocument().Root!.AddFirst(new XElement(W.p, FieldChar("end")));
                header.PutXDocument();
            });
        Assert.Contains("Before Unrelated header After", Story(html, "header").Value);
    }

    private static IEnumerable<XElement> ConditionalField(string cache, params XElement[] instructionExtra) =>
        new[] { FieldChar("begin"), new XElement(W.r, new XElement(W.instrText, " IF \"")) }
            .Concat(Field("TITLE", false, "Hidden title").SkipLast(1)).Concat(instructionExtra)
            .Concat(new[] { FieldChar("end"), new XElement(W.r, new XElement(W.instrText,
                "\" = \"x\" \"Visible result\" \"Other result\" ")), FieldChar("separate"),
                new XElement(W.r, new XElement(W.t, cache)), FieldChar("end") });

    private static XElement FieldChar(string type) =>
        new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, type)));

    private static void RemoveDiagnostic(XElement html) =>
        Assert.Single(html.Descendants(), e => e.Attribute("data-content-type") != null).Remove();

    private static XElement Story(XElement html, string story) =>
        Assert.Single(html.Descendants(), e => e.Name.LocalName == story);

    internal static XElement Render(string instruction, bool simple, string story, string cache,
        string? locked = null, Action<XElement>? mutate = null, string? title = Title,
        bool placeholders = false, bool showDeleted = true, string? creator = null,
        Action<MainDocumentPart, XElement>? setup = null, bool notes = false, bool paginate = false)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            document.PackageProperties.Title = title;
            document.PackageProperties.Creator = creator;
            var main = document.AddMainDocumentPart();
            var paragraph = new XElement(W.p,
                new XElement(W.r, new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), "Before ")),
                Field(instruction, simple, cache, locked),
                new XElement(W.r, new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), " After")));
            mutate?.Invoke(paragraph);
            var body = new XElement(W.body);
            if (story == "body") body.Add(paragraph);
            else
            {
                OpenXmlPart part = story == "header" ? main.AddNewPart<HeaderPart>() : main.AddNewPart<FooterPart>();
                part.PutXDocument(new XDocument(new XElement(story == "header" ? W.hdr : W.ftr,
                    new XAttribute(XNamespace.Xmlns + "w", W.w), paragraph)));
                body.Add(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Body control"))),
                    new XElement(W.sectPr, new XElement(story == "header" ? W.headerReference : W.footerReference,
                        new XAttribute(W.type, "default"), new XAttribute(R.id, main.GetIdOfPart(part)))));
            }
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles,
                new XAttribute(XNamespace.Xmlns + "w", W.w))));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
            setup?.Invoke(main, body);
            main.PutXDocument(new XDocument(new XElement(W.document, new XAttribute(XNamespace.Xmlns + "w", W.w), body)));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(document));
        }
        return XElement.Parse(HtmlConversionOps.ConvertToHtml(stream.ToArray(), new HtmlConversionOptions
        {
            RenderHeadersAndFooters = true, RenderTrackedChanges = true, FabricateCssClasses = false,
            RenderUnsupportedContentPlaceholders = placeholders,
            ShowDeletedContent = showDeleted,
            RenderFootnotesAndEndnotes = notes,
            PaginationMode = paginate ? 1 : 0,
        }));
    }
}
