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
        var value = Assert.Single(html.Descendants(Xhtml.span).Where(e => e.Value == Title));
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

    private static XElement FieldChar(string type) =>
        new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, type)));

    private static void RemoveDiagnostic(XElement html) =>
        Assert.Single(html.Descendants().Where(e => e.Attribute("data-content-type") != null)).Remove();

    private static XElement Story(XElement html, string story) =>
        Assert.Single(html.Descendants().Where(e => e.Name.LocalName == story));

    private static XElement Render(string instruction, bool simple, string story, string cache,
        string? locked = null, Action<XElement>? mutate = null, string? title = Title,
        bool placeholders = false, bool showDeleted = true)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            document.PackageProperties.Title = title;
            var main = document.AddMainDocumentPart();
            var result = new XElement(W.r, new XElement(W.t, cache));
            var field = simple ? new object[] { new XElement(W.fldSimple, new XAttribute(W.instr, instruction),
                locked == null ? null : new XAttribute(W.fldLock, locked), result) } : new object[] {
                new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"),
                    locked == null ? null : new XAttribute(W.fldLock, locked))),
                new XElement(W.r, new XElement(W.instrText, instruction)),
                new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate"))), result,
                new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))) };
            var paragraph = new XElement(W.p,
                new XElement(W.r, new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), "Before ")),
                field, new XElement(W.r, new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), " After")));
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
            main.PutXDocument(new XDocument(new XElement(W.document, new XAttribute(XNamespace.Xmlns + "w", W.w), body)));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles,
                new XAttribute(XNamespace.Xmlns + "w", W.w))));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
            Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(document));
        }
        return XElement.Parse(HtmlConversionOps.ConvertToHtml(stream.ToArray(), new HtmlConversionOptions
        {
            RenderHeadersAndFooters = true, RenderTrackedChanges = true, FabricateCssClasses = false,
            RenderUnsupportedContentPlaceholders = placeholders,
            ShowDeletedContent = showDeleted,
        }));
    }
}
