// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Docxodus;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1299: with tracked changes on, the running-story and section-property operations must
/// leave a reviewable revision that rejecting undoes and accepting keeps, or refuse without
/// changing anything when OOXML has no revision form for the change.
/// </summary>
public class DocxSessionTrackedHeaderFooterTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>A session over the Letter-section fixture, its first body anchor, and (when
    /// <paramref name="withStories"/>) a default footer "Page " and header "Original header"
    /// authored before tracking is switched on.</summary>
    private static (DocxSession Session, string Body, string? FooterParagraph) Open(bool withStories = false)
    {
        var s = new DocxSession(DocxSessionTests.BuildDocWithTrailingSection(),
            new DocxSessionSettings { RevisionAuthor = "Bob" });
        var body = s.Project().AnchorIndex.Values.Select(v => v.Anchor)
            .First(a => a.Kind == "p" && a.Scope == "body").Id;
        string? footerParagraph = null;
        if (withStories)
        {
            var footer = s.SetFooterText(body, HeaderFooterKind.Default, "Page ");
            Assert.True(footer.Success, footer.Error?.Message);
            footerParagraph = footer.Created.First().Id;
            Assert.True(s.SetHeaderText(body, HeaderFooterKind.Default, "Original header").Success);
        }
        s.SetTrackedChanges(TrackedChangeMode.RenderInline);
        return (s, body, footerParagraph);
    }

    private static XElement Root(byte[] bytes, Func<MainDocumentPart, OpenXmlPart?> select)
    {
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        using var stream = select(doc.MainDocumentPart!)!.GetStream();
        return XDocument.Load(stream).Root!;
    }

    /// <summary>The saved package has no schema error — revision payloads included.</summary>
    private static void AssertSchemaValid(DocxSession s) =>
        Assert.Empty(PackageValidation.Errors(s.Save(), FileFormatVersions.Office2019)
            .Where(PackageValidation.IsSchemaError).Select(PackageValidation.Keys.Description));

    private static XElement BodySectPr(DocxSession s) =>
        Root(s.Save(), m => m).Element(W + "body")!.Element(W + "sectPr")!;

    private static XElement Settings(DocxSession s) => Root(s.Save(), m => m.DocumentSettingsPart);

    private static string HeaderText(DocxSession s) =>
        string.Concat(Root(s.Save(), m => m.HeaderParts.Single()).Descendants(W + "t").Select(t => t.Value));

    /// <summary>The section's property children, ignoring the change marker itself — what a
    /// reject must restore and an accept must keep.</summary>
    private static string SectionProperties(XElement sectPr) =>
        string.Join("|", sectPr.Elements()
            .Where(e => e.Name != W + "sectPrChange")
            .Select(e => new XElement(e.Name, e.Attributes()).ToString(SaveOptions.DisableFormatting)));

    [Fact]
    public void SectPrChange_FromWord_IsListedAndRejectRestoresTheOldPageSize()
    {
        // The registry and resolver already own w:sectPrChange; this pins that before the session
        // starts writing it.
        var bytes = File.ReadAllBytes(
            Path.Combine("../../../../TestFiles/", "FA/RevTracking/020-SectionPropertiesChange.docx"));
        using var s = new DocxSession(bytes);
        Assert.Single(s.ListRevisions());
        Assert.True(s.RejectAllRevisions().Success);
        var pgSz = BodySectPr(s).Element(W + "pgSz")!;
        Assert.Equal("12240", (string?)pgSz.Attribute(W + "w"));
        Assert.Null(BodySectPr(s).Element(W + "sectPrChange"));
    }

    [Fact]
    public void TrackedPageSetup_RecordsSectionChange_RejectRestoresAcceptKeeps()
    {
        var (s, body, _) = Open();
        using var _s = s;
        var before = SectionProperties(BodySectPr(s));

        Assert.True(s.SetPageSetup(body, new PageSetupOp { MarginLeftTwips = 2160 }).Success);

        var revision = Assert.Single(s.ListRevisions());
        Assert.Equal("Bob", revision.Author);
        AssertSchemaValid(s);
        Assert.NotNull(BodySectPr(s).Element(W + "sectPrChange"));
        var after = SectionProperties(BodySectPr(s));
        Assert.Equal("2160", (string?)BodySectPr(s).Element(W + "pgMar")!.Attribute(W + "left"));

        Assert.True(s.RejectAllRevisions().Success);
        Assert.Equal(before, SectionProperties(BodySectPr(s)));
        Assert.Null(BodySectPr(s).Element(W + "sectPrChange"));

        Assert.True(s.SetPageSetup(body, new PageSetupOp { MarginLeftTwips = 2160 }).Success);
        Assert.True(s.AcceptAllRevisions().Success);
        Assert.Equal(after, SectionProperties(BodySectPr(s)));
        Assert.Empty(s.ListRevisions());
    }

    [Fact]
    public void TrackedSectionEdits_FoldIntoOneChange_AndRejectRestoresTheOriginalSection()
    {
        var (s, body, _) = Open(withStories: true);
        using var _s = s;
        var before = SectionProperties(BodySectPr(s));

        Assert.True(s.SetPageSetup(body, new PageSetupOp { MarginLeftTwips = 2160 }).Success);
        Assert.True(s.SetPageNumbering(body, new PageNumberingOp { Start = 5 }).Success);
        Assert.True(s.SetHeaderFooterKindEnabled(body, HeaderFooterKind.First, true).Success);

        var sectPr = BodySectPr(s);
        Assert.Single(sectPr.Elements(W + "sectPrChange"));
        Assert.NotNull(sectPr.Element(W + "pgNumType"));
        Assert.NotNull(sectPr.Element(W + "titlePg"));
        Assert.Single(s.ListRevisions());
        AssertSchemaValid(s);

        Assert.True(s.RejectAllRevisions().Success);
        sectPr = BodySectPr(s);
        Assert.Equal(before, SectionProperties(sectPr));
        // The header/footer references sit outside the property change and survive the reject.
        Assert.NotNull(sectPr.Element(W + "headerReference"));
        Assert.NotNull(sectPr.Element(W + "footerReference"));
    }

    [Fact]
    public void TrackedSectionEdit_UndoneByTheNextEdit_LeavesNoRevision()
    {
        var (s, body, _) = Open();
        using var _s = s;
        var before = SectionProperties(BodySectPr(s));

        Assert.True(s.SetPageNumbering(body, new PageNumberingOp { Start = 5 }).Success);
        Assert.Single(s.ListRevisions());
        Assert.True(s.ClearPageNumbering(body).Success);

        Assert.Empty(s.ListRevisions());
        Assert.Null(BodySectPr(s).Element(W + "sectPrChange"));
        Assert.Equal(before, SectionProperties(BodySectPr(s)));
    }

    [Fact]
    public void TrackedClearPageNumbering_OnAnExistingSetup_RejectBringsItBack()
    {
        var s = new DocxSession(DocxSessionTests.BuildDocWithTrailingSection(),
            new DocxSessionSettings { RevisionAuthor = "Bob" });
        using var _s = s;
        var body = s.Project().AnchorIndex.Values.Select(v => v.Anchor).First(a => a.Kind == "p").Id;
        Assert.True(s.SetPageNumbering(body, new PageNumberingOp { Start = 3 }).Success);
        s.SetTrackedChanges(TrackedChangeMode.RenderInline);

        Assert.True(s.ClearPageNumbering(body).Success);
        Assert.Null(BodySectPr(s).Element(W + "pgNumType"));
        Assert.Single(s.ListRevisions());

        Assert.True(s.RejectAllRevisions().Success);
        Assert.Equal("3", (string?)BodySectPr(s).Element(W + "pgNumType")?.Attribute(W + "start"));
    }

    [Fact]
    public void TrackedFirstPageFlag_InAFreshTrackedSession_IsASectionChange()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDocWithTrailingSection(),
            new DocxSessionSettings { RevisionAuthor = "Bob", TrackedChanges = TrackedChangeMode.RenderInline });
        var body = s.Project().AnchorIndex.Values.Select(v => v.Anchor).First(a => a.Kind == "p").Id;

        Assert.True(s.EnsureHeaderFooterVisible(body, HeaderFooterKind.First).Success);
        Assert.NotNull(BodySectPr(s).Element(W + "titlePg"));
        Assert.Single(s.ListRevisions());

        Assert.True(s.RejectAllRevisions().Success);
        Assert.Null(BodySectPr(s).Element(W + "titlePg"));
    }

    [Fact]
    public void TrackedFirstPageFlag_SwitchedOff_RejectSwitchesItBackOn()
    {
        var s = new DocxSession(DocxSessionTests.BuildDocWithTrailingSection(),
            new DocxSessionSettings { RevisionAuthor = "Bob" });
        using var _s = s;
        var body = s.Project().AnchorIndex.Values.Select(v => v.Anchor).First(a => a.Kind == "p").Id;
        Assert.True(s.EnsureHeaderFooterVisible(body, HeaderFooterKind.First).Success);
        s.SetTrackedChanges(TrackedChangeMode.RenderInline);

        Assert.True(s.SetHeaderFooterKindEnabled(body, HeaderFooterKind.First, false).Success);
        Assert.Null(BodySectPr(s).Element(W + "titlePg"));
        Assert.Single(s.ListRevisions());

        Assert.True(s.RejectAllRevisions().Success);
        Assert.NotNull(BodySectPr(s).Element(W + "titlePg"));

        // Off and straight back on puts the section back as it was: no revision is left behind.
        Assert.True(s.SetHeaderFooterKindEnabled(body, HeaderFooterKind.First, false).Success);
        Assert.True(s.SetHeaderFooterKindEnabled(body, HeaderFooterKind.First, true).Success);
        Assert.Empty(s.ListRevisions());
        Assert.Null(BodySectPr(s).Element(W + "sectPrChange"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void TrackedEvenAndOddSwitch_IsRefusedAndChangesNothing(bool enable)
    {
        var s = new DocxSession(DocxSessionTests.BuildDocWithTrailingSection(),
            new DocxSessionSettings { RevisionAuthor = "Bob" });
        using var _s = s;
        var body = s.Project().AnchorIndex.Values.Select(v => v.Anchor).First(a => a.Kind == "p").Id;
        if (!enable) Assert.True(s.EnsureHeaderFooterVisible(body, HeaderFooterKind.Even).Success);
        s.SetTrackedChanges(TrackedChangeMode.RenderInline);
        var before = Settings(s).ToString();
        var undoBefore = s.UndoCount;

        var result = s.SetHeaderFooterKindEnabled(body, HeaderFooterKind.Even, enable);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, result.Error!.Code);
        Assert.Equal(before, Settings(s).ToString());
        Assert.Equal(undoBefore, s.UndoCount);
    }

    [Fact]
    public void TrackedEvenStorySeed_WithoutTheEvenFlag_IsRefusedAndChangesNothing()
    {
        var (s, body, _) = Open();
        using var _s = s;
        var before = s.Save();

        var result = s.SetHeaderText(body, HeaderFooterKind.Even, "Even header");

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, result.Error!.Code);
        Assert.Equal(before, s.Save());
    }

    [Fact]
    public void TrackedPageNumberField_IsAnInsertion_RejectRemovesItAcceptKeepsIt()
    {
        var (s, body, footerParagraph) = Open(withStories: true);
        using var _s = s;

        Assert.True(s.InsertPageNumberField(footerParagraph!, PageNumberField.CurrentPage).Success);

        var footer = Root(s.Save(), m => m.FooterParts.Single());
        Assert.All(footer.Descendants(W + "fldChar"), f => Assert.NotNull(f.Ancestors(W + "ins").FirstOrDefault()));
        Assert.Single(s.ListRevisions());
        AssertSchemaValid(s);

        Assert.True(s.RejectAllRevisions().Success);
        footer = Root(s.Save(), m => m.FooterParts.Single());
        Assert.Empty(footer.Descendants(W + "fldChar"));
        Assert.Equal("Page ", string.Concat(footer.Descendants(W + "t").Select(t => t.Value)));

        Assert.True(s.InsertPageNumberField(footerParagraph!, PageNumberField.PageOfTotal).Success);
        Assert.Single(s.ListRevisions());
        Assert.True(s.AcceptAllRevisions().Success);
        footer = Root(s.Save(), m => m.FooterParts.Single());
        Assert.Equal(4, footer.Descendants(W + "fldChar").Count(f => (string?)f.Attribute(W + "fldCharType") != "separate"));
        Assert.Empty(footer.Descendants(W + "ins"));
    }

    [Theory]
    [InlineData("Original header", "Replacement header")]
    [InlineData("First old line\n\nSecond old line", "New one\n\nNew two\n\nNew three")]
    public void TrackedHeaderReplacement_IsReviewable_RejectRestoresAcceptKeeps(string original, string replacement)
    {
        var (s, body, _) = Open();
        using var _s = s;
        s.SetTrackedChanges(TrackedChangeMode.Accept);
        Assert.True(s.SetHeaderText(body, HeaderFooterKind.Default, original).Success);
        var originalHeader = Root(s.Save(), m => m.HeaderParts.Single());
        s.SetTrackedChanges(TrackedChangeMode.RenderInline);

        var result = s.SetHeaderText(body, HeaderFooterKind.Default, replacement);
        Assert.True(result.Success, result.Error?.Message);

        var header = Root(s.Save(), m => m.HeaderParts.Single());
        Assert.NotEmpty(header.Descendants(W + "del"));
        Assert.NotEmpty(header.Descendants(W + "ins"));
        Assert.NotEmpty(s.ListRevisions());
        AssertSchemaValid(s);

        Assert.True(s.RejectAllRevisions().Success);
        var rejected = Root(s.Save(), m => m.HeaderParts.Single());
        Assert.Equal(ParagraphTexts(originalHeader), ParagraphTexts(rejected));
        Assert.Equal(ParagraphStyles(originalHeader), ParagraphStyles(rejected));
        AssertSchemaValid(s);

        Assert.True(s.SetHeaderText(body, HeaderFooterKind.Default, replacement).Success);
        Assert.True(s.AcceptAllRevisions().Success);
        var accepted = Root(s.Save(), m => m.HeaderParts.Single());
        Assert.Equal(replacement.Split("\n\n"), ParagraphTexts(accepted));
        Assert.All(ParagraphStyles(accepted), style => Assert.Equal("Header", style));
        AssertSchemaValid(s);
    }

    [Fact]
    public void TrackedNewHeaderStory_IsAnInsertion_RejectEmptiesIt()
    {
        var (s, body, _) = Open();
        using var _s = s;

        Assert.True(s.SetHeaderText(body, HeaderFooterKind.Default, "Brand new header").Success);
        Assert.NotEmpty(s.ListRevisions());

        Assert.True(s.RejectAllRevisions().Success);
        Assert.Equal("", HeaderText(s));
        AssertSchemaValid(s);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void TrackedNewHeaderStory_RewrittenByTheSameAuthor_ResolvesToTheLastText(bool accept)
    {
        var (s, body, _) = Open();
        using var _s = s;

        Assert.True(s.SetHeaderText(body, HeaderFooterKind.Default, "First draft\n\nSecond line").Success);
        Assert.True(s.SetHeaderText(body, HeaderFooterKind.Default, "Final header").Success);

        Assert.True(accept ? s.AcceptAllRevisions().Success : s.RejectAllRevisions().Success);
        Assert.Equal(accept ? new[] { "Final header" } : new[] { "" },
            ParagraphTexts(Root(s.Save(), m => m.HeaderParts.Single())));
        AssertSchemaValid(s);
    }

    [Fact]
    public void TrackedHeaderReplacement_OverABookmarkedLastParagraph_IsRefusedAndChangesNothing()
    {
        var (s, body, _) = Open();
        using var _s = s;
        s.SetTrackedChanges(TrackedChangeMode.Accept);
        var header = s.SetHeaderText(body, HeaderFooterKind.Default, "Bookmarked header");
        Assert.True(s.AddBookmark("HeaderMark", DocumentRange.In(header.Created.Single().Id, new CharSpan(0, 10))).Success);
        s.SetTrackedChanges(TrackedChangeMode.RenderInline);
        var before = s.Save();

        var result = s.SetHeaderText(body, HeaderFooterKind.Default, "Replacement");

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.TrackedOperationUnsupported, result.Error!.Code);
        Assert.Equal(before, s.Save());
    }

    [Fact]
    public void Issue1299_Reproduction_EveryCallLeavesARevisionOrRefuses()
    {
        var (s, body, footerParagraph) = Open(withStories: true);
        using var _s = s;

        Assert.True(s.InsertPageNumberField(footerParagraph!, PageNumberField.CurrentPage).Success);
        Assert.NotEmpty(s.ListRevisions());
        Assert.True(s.SetPageSetup(body, new PageSetupOp { MarginLeftTwips = 2160 }).Success);
        Assert.True(s.SetPageNumbering(body, new PageNumberingOp { Start = 5 }).Success);
        Assert.True(s.SetHeaderFooterKindEnabled(body, HeaderFooterKind.First, true).Success);
        Assert.False(s.SetHeaderFooterKindEnabled(body, HeaderFooterKind.Even, true).Success);
        Assert.True(s.ClearPageNumbering(body).Success);
        Assert.True(s.SetHeaderText(body, HeaderFooterKind.Default, "Replacement header").Success);

        Assert.Null(Settings(s).Element(W + "evenAndOddHeaders"));
        Assert.NotNull(BodySectPr(s).Element(W + "sectPrChange"));
        AssertSchemaValid(s);

        Assert.True(s.RejectAllRevisions().Success);
        Assert.Equal("Original header", HeaderText(s));
        AssertSchemaValid(s);
        var sectPr = BodySectPr(s);
        Assert.Null(sectPr.Element(W + "titlePg"));
        Assert.Null(sectPr.Element(W + "pgNumType"));
        Assert.Equal("1440", (string?)sectPr.Element(W + "pgMar")!.Attribute(W + "left"));
        Assert.Empty(Root(s.Save(), m => m.FooterParts.Single()).Descendants(W + "fldChar"));
    }

    private static string[] ParagraphTexts(XElement story) =>
        story.Elements(W + "p").Select(p => string.Concat(p.Descendants(W + "t").Select(t => t.Value))).ToArray();

    private static string?[] ParagraphStyles(XElement story) =>
        story.Elements(W + "p").Select(p => (string?)p.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val")).ToArray();
}
