#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tests for the engine additions behind the browser editor's Word-parity redesign: run
/// highlight / caps / small caps on <see cref="FormatOp"/>, the numeric id on
/// <see cref="CommentListEntry"/>, <see cref="DocxSession.SetHeaderFooterKindEnabled"/>,
/// <see cref="DocxSession.SetPageSetup"/> (with the new <see cref="SectionInfo"/> fields), and
/// trailing-space preservation through <see cref="DocxSession.ReplaceText"/>. Test IDs use the
/// DEP range.
/// </summary>
public class DocxSessionEditorParityTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string FirstBodyParagraph(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;

    private static XElement DocumentXml(byte[] docxBytes)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.GetXDocument().Root!;
    }

    private static XElement? SettingsXml(byte[] docxBytes)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.DocumentSettingsPart?.GetXDocument().Root;
    }

    private static XElement? CommentsXml(byte[] docxBytes)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.WordprocessingCommentsPart?.GetXDocument().Root;
    }

    private static XElement FirstTextRunProps(byte[] docxBytes) =>
        DocumentXml(docxBytes).Descendants(W + "r").First(r => r.Value.Length > 0).Element(W + "rPr")!;

    private static XElement GoverningSectPr(byte[] docxBytes) =>
        DocumentXml(docxBytes).Element(W + "body")!.Element(W + "sectPr")!;

    /// <summary>Two body paragraphs under a Letter / one-inch-margin <c>w:sectPr</c> — the shape
    /// Word writes, and what <c>GetSectionInfo</c> needs to resolve (a body with no sectPr at all
    /// reports null). Only the main part, so the settings part is created on demand.</summary>
    private static byte[] BuildTwoParagraphsWithSection()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body(
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("First paragraph"))),
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text("Second paragraph"))),
                    new DocumentFormat.OpenXml.Wordprocessing.SectionProperties(
                        new DocumentFormat.OpenXml.Wordprocessing.PageSize { Width = 12240, Height = 15840 },
                        new DocumentFormat.OpenXml.Wordprocessing.PageMargin
                        {
                            Top = 1440, Right = 1440, Bottom = 1440, Left = 1440, Header = 720, Footer = 720, Gutter = 0,
                        })));
        }
        return ms.ToArray();
    }

    // ─── FormatOp: highlight, caps, small caps ──────────────────────────

    [Fact]
    public void DEP001_Highlight_WritesHighlightElementInCanonicalCase()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);

        var r = session.ApplyFormat(anchor, null, new FormatOp { Highlight = "yellow" });
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal("yellow", (string?)FirstTextRunProps(session.Save()).Element(W + "highlight")?.Attribute(W + "val"));

        // Case-insensitive on input, canonical ST_HighlightColor casing on output.
        r = session.ApplyFormat(anchor, null, new FormatOp { Highlight = "DARKBLUE" });
        Assert.True(r.Success, r.Error?.Message);
        var rPr = FirstTextRunProps(session.Save());
        Assert.Single(rPr.Elements(W + "highlight"));
        Assert.Equal("darkBlue", (string?)rPr.Element(W + "highlight")?.Attribute(W + "val"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("none")]
    public void DEP002_Highlight_EmptyOrNoneClears(string clearToken)
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);
        Assert.True(session.ApplyFormat(anchor, null, new FormatOp { Highlight = "green" }).Success);
        Assert.NotNull(FirstTextRunProps(session.Save()).Element(W + "highlight"));

        var r = session.ApplyFormat(anchor, null, new FormatOp { Highlight = clearToken });
        Assert.True(r.Success, r.Error?.Message);
        Assert.Null(FirstTextRunProps(session.Save()).Element(W + "highlight"));
    }

    [Fact]
    public void DEP003_Highlight_InvalidNameFailsAndRollsBack()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);
        var before = DocumentXml(session.Save()).ToString(SaveOptions.DisableFormatting);
        int undoBefore = session.UndoCount;

        // Same surfacing as an invalid vertAlign: the writer throws ArgumentException, ApplyFormat
        // maps it to InternalError and rolls the pre-op snapshot back.
        var r = session.ApplyFormat(anchor, null, new FormatOp { Highlight = "pink" });
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.InternalError, r.Error!.Code);
        Assert.Contains("invalid highlight", r.Error.Message);
        Assert.Equal(before, DocumentXml(session.Save()).ToString(SaveOptions.DisableFormatting));
        Assert.Equal(undoBefore, session.UndoCount);
    }

    [Fact]
    public void DEP004_Caps_And_SmallCaps_AreMutuallyExclusive()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);

        Assert.True(session.ApplyFormat(anchor, null, new FormatOp { Caps = true }).Success);
        var rPr = FirstTextRunProps(session.Save());
        Assert.NotNull(rPr.Element(W + "caps"));
        Assert.Null(rPr.Element(W + "smallCaps"));

        // Word's rule: turning small caps on turns caps off.
        Assert.True(session.ApplyFormat(anchor, null, new FormatOp { SmallCaps = true }).Success);
        rPr = FirstTextRunProps(session.Save());
        Assert.Null(rPr.Element(W + "caps"));
        Assert.NotNull(rPr.Element(W + "smallCaps"));

        // And back.
        Assert.True(session.ApplyFormat(anchor, null, new FormatOp { Caps = true }).Success);
        rPr = FirstTextRunProps(session.Save());
        Assert.NotNull(rPr.Element(W + "caps"));
        Assert.Null(rPr.Element(W + "smallCaps"));

        // false removes without touching the other slot.
        Assert.True(session.ApplyFormat(anchor, null, new FormatOp { Caps = false }).Success);
        rPr = FirstTextRunProps(session.Save());
        Assert.Null(rPr.Element(W + "caps"));
        Assert.Null(rPr.Element(W + "smallCaps"));
    }

    [Fact]
    public void DEP005_FormatOp_JsonParsesHighlightCapsSmallCaps()
    {
        var op = DocxSessionJson.ParseFormatOp("{\"highlight\":\"green\",\"caps\":true,\"smallCaps\":false,\"bold\":true}");
        Assert.Equal("green", op.Highlight);
        Assert.True(op.Caps);
        Assert.False(op.SmallCaps);
        Assert.True(op.Bold);

        var untouched = DocxSessionJson.ParseFormatOp("{\"bold\":true}");
        Assert.Null(untouched.Highlight);
        Assert.Null(untouched.Caps);
        Assert.Null(untouched.SmallCaps);

        // And the parsed op drives the writer the same way the typed one does.
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);
        Assert.True(session.ApplyFormat(anchor, null, op).Success);
        var rPr = FirstTextRunProps(session.Save());
        Assert.Equal("green", (string?)rPr.Element(W + "highlight")?.Attribute(W + "val"));
        Assert.NotNull(rPr.Element(W + "caps"));
    }

    [Fact]
    public void DEP006_Highlight_TrackedModeRecordsRPrChange()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection(),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchor = FirstBodyParagraph(session);

        var r = session.ApplyFormat(anchor, null, new FormatOp { Highlight = "cyan", SmallCaps = true });
        Assert.True(r.Success, r.Error?.Message);
        var rPr = FirstTextRunProps(session.Save());
        Assert.Equal("cyan", (string?)rPr.Element(W + "highlight")?.Attribute(W + "val"));
        Assert.NotNull(rPr.Element(W + "smallCaps"));
        var change = rPr.Element(W + "rPrChange");
        Assert.NotNull(change);
        Assert.Null(change!.Element(W + "rPr")?.Element(W + "highlight"));
    }

    // ─── CommentListEntry.Id ────────────────────────────────────────────

    [Fact]
    public void DEP010_ListComments_ReportsNumericIdMatchingCommentsPart()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);
        Assert.True(session.AddComment(anchor, null, "Reviewer", "Check this").Success);

        var entries = session.ListComments();
        var entry = Assert.Single(entries);

        var comment = CommentsXml(session.Save())!.Elements(W + "comment").Single();
        Assert.Equal(int.Parse((string)comment.Attribute(W + "id")!), entry.Id);

        var json = DocxSessionJson.SerializeCommentList(entries);
        Assert.Contains($"\"id\":{entry.Id}", json);
        Assert.Contains("\"anchorId\":\"cmt:cmt:", json);
    }

    // ─── SetHeaderFooterKindEnabled ─────────────────────────────────────

    [Fact]
    public void DEP020_FirstPage_DisableRemovesTitlePgAndKeepsThePart()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);
        Assert.True(session.SetHeaderText(anchor, HeaderFooterKind.First, "First-page header").Success);

        var info = session.GetSectionInfo(anchor)!;
        Assert.True(info.TitlePage);
        Assert.NotNull(GoverningSectPr(session.Save()).Element(W + "titlePg"));

        int undoBefore = session.UndoCount;
        var r = session.SetHeaderFooterKindEnabled(anchor, HeaderFooterKind.First, false);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal(undoBefore + 1, session.UndoCount);

        var sectPr = GoverningSectPr(session.Save());
        Assert.Null(sectPr.Element(W + "titlePg"));
        // Word leaves the story behind when the checkbox is cleared: the reference and part survive.
        Assert.Contains(sectPr.Elements(W + "headerReference"), h => (string?)h.Attribute(W + "type") == "first");
        info = session.GetSectionInfo(anchor)!;
        Assert.False(info.TitlePage);
        Assert.Contains(info.HeaderRefs, h => h.Kind == HeaderFooterKind.First);

        // Re-enabling brings the flag straight back.
        Assert.True(session.SetHeaderFooterKindEnabled(anchor, HeaderFooterKind.First, true).Success);
        Assert.True(session.GetSectionInfo(anchor)!.TitlePage);
        Assert.NotNull(GoverningSectPr(session.Save()).Element(W + "titlePg"));

        // Undo of the disable restores the flag.
        Assert.True(session.Undo());
        Assert.True(session.Undo());
        Assert.True(session.GetSectionInfo(anchor)!.TitlePage);
    }

    [Fact]
    public void DEP021_EvenPages_RoundTripsThroughSettingsPart()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);
        Assert.False(session.GetSectionInfo(anchor)!.EvenAndOddHeaders);

        var r = session.SetHeaderFooterKindEnabled(anchor, HeaderFooterKind.Even, true);
        Assert.True(r.Success, r.Error?.Message);
        Assert.NotNull(SettingsXml(session.Save())?.Element(W + "evenAndOddHeaders"));
        Assert.True(session.GetSectionInfo(anchor)!.EvenAndOddHeaders);

        r = session.SetHeaderFooterKindEnabled(anchor, HeaderFooterKind.Even, false);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Null(SettingsXml(session.Save())?.Element(W + "evenAndOddHeaders"));
        Assert.False(session.GetSectionInfo(anchor)!.EvenAndOddHeaders);
    }

    [Fact]
    public void DEP022_DisableWhenAlreadyOff_RecordsNoUndoStep()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);
        int undoBefore = session.UndoCount;
        long versionBefore = session.Version;

        Assert.True(session.SetHeaderFooterKindEnabled(anchor, HeaderFooterKind.First, false).Success);
        Assert.True(session.SetHeaderFooterKindEnabled(anchor, HeaderFooterKind.Even, false).Success);
        Assert.Equal(undoBefore, session.UndoCount);
        Assert.Equal(versionBefore, session.Version);
        Assert.False(session.Undo());
    }

    [Fact]
    public void DEP023_DisableDefault_IsInvalidPageSetup_EnableDefaultIsNoOp()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);

        var r = session.SetHeaderFooterKindEnabled(anchor, HeaderFooterKind.Default, false);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.InvalidPageSetup, r.Error!.Code);

        // Enabling delegates to EnsureHeaderFooterVisible, whose Default is a successful no-op.
        Assert.True(session.SetHeaderFooterKindEnabled(anchor, HeaderFooterKind.Default, true).Success);
        Assert.Equal(0, session.UndoCount);
    }

    [Fact]
    public void DEP024_SetHeaderFooterKindEnabled_WireRipple()
    {
        int handle = DocxSessionOps.OpenSession(BuildTwoParagraphsWithSection(), null);
        try
        {
            var session = SessionRegistry.Get(handle);
            var anchor = FirstBodyParagraph(session);
            var json = DocxSessionOps.SetHeaderFooterKindEnabled(handle, anchor, "first", true);
            Assert.Contains("\"success\":true", json);
            Assert.True(session.GetSectionInfo(anchor)!.TitlePage);

            var info = DocxSessionOps.GetSectionInfo(handle, anchor);
            Assert.Contains("\"titlePage\":true", info);
            Assert.Contains("\"evenAndOddHeaders\":false", info);
            Assert.Contains("\"headerDistanceTwips\":720", info);
            Assert.Contains("\"footerDistanceTwips\":720", info);

            json = DocxSessionOps.SetHeaderFooterKindEnabled(handle, anchor, "default", false);
            Assert.Contains("invalid_page_setup", json);
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    // ─── SetPageSetup ───────────────────────────────────────────────────

    [Fact]
    public void DEP030_Margins_WrittenAndReportedBySectionInfo()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);
        int undoBefore = session.UndoCount;

        var r = session.SetPageSetup(anchor, new PageSetupOp
        {
            MarginTopTwips = 720, MarginBottomTwips = 720, MarginLeftTwips = 1080, MarginRightTwips = 1080,
        });
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal(undoBefore + 1, session.UndoCount);

        var pgMar = GoverningSectPr(session.Save()).Element(W + "pgMar");
        Assert.NotNull(pgMar);
        Assert.Equal("720", (string?)pgMar!.Attribute(W + "top"));
        Assert.Equal("720", (string?)pgMar.Attribute(W + "bottom"));
        Assert.Equal("1080", (string?)pgMar.Attribute(W + "left"));
        Assert.Equal("1080", (string?)pgMar.Attribute(W + "right"));

        var info = session.GetSectionInfo(anchor)!;
        Assert.Equal(720, info.MarginTopTwips);
        Assert.Equal(720, info.MarginBottomTwips);
        Assert.Equal(1080, info.MarginLeftTwips);
        Assert.Equal(1080, info.MarginRightTwips);

        // A partial op leaves the other attributes alone.
        Assert.True(session.SetPageSetup(anchor, new PageSetupOp { MarginTopTwips = 1440 }).Success);
        info = session.GetSectionInfo(anchor)!;
        Assert.Equal(1440, info.MarginTopTwips);
        Assert.Equal(1080, info.MarginLeftTwips);
    }

    [Fact]
    public void DEP031_Landscape_SwapsDimensionsAndBack()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);
        var before = session.GetSectionInfo(anchor)!;
        Assert.False(before.Landscape);
        Assert.True(before.PageWidthTwips < before.PageHeightTwips);

        var r = session.SetPageSetup(anchor, new PageSetupOp { Landscape = true });
        Assert.True(r.Success, r.Error?.Message);
        var pgSz = GoverningSectPr(session.Save()).Element(W + "pgSz")!;
        Assert.Equal("landscape", (string?)pgSz.Attribute(W + "orient"));
        Assert.Equal(before.PageHeightTwips.ToString(), (string?)pgSz.Attribute(W + "w"));
        Assert.Equal(before.PageWidthTwips.ToString(), (string?)pgSz.Attribute(W + "h"));
        var info = session.GetSectionInfo(anchor)!;
        Assert.True(info.Landscape);
        Assert.Equal(before.PageHeightTwips, info.PageWidthTwips);
        Assert.Equal(before.PageWidthTwips, info.PageHeightTwips);

        r = session.SetPageSetup(anchor, new PageSetupOp { Landscape = false });
        Assert.True(r.Success, r.Error?.Message);
        pgSz = GoverningSectPr(session.Save()).Element(W + "pgSz")!;
        Assert.Null(pgSz.Attribute(W + "orient"));
        info = session.GetSectionInfo(anchor)!;
        Assert.False(info.Landscape);
        Assert.Equal(before.PageWidthTwips, info.PageWidthTwips);
        Assert.Equal(before.PageHeightTwips, info.PageHeightTwips);
    }

    [Fact]
    public void DEP032_ExplicitSizeWithLandscape_WritesAsGiven()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);

        // A4 landscape, stated explicitly: no swap second-guesses the caller.
        var r = session.SetPageSetup(anchor, new PageSetupOp
        {
            PageWidthTwips = 16838, PageHeightTwips = 11906, Landscape = true,
        });
        Assert.True(r.Success, r.Error?.Message);
        var pgSz = GoverningSectPr(session.Save()).Element(W + "pgSz")!;
        Assert.Equal("16838", (string?)pgSz.Attribute(W + "w"));
        Assert.Equal("11906", (string?)pgSz.Attribute(W + "h"));
        Assert.Equal("landscape", (string?)pgSz.Attribute(W + "orient"));

        var info = session.GetSectionInfo(anchor)!;
        Assert.Equal(16838, info.PageWidthTwips);
        Assert.Equal(11906, info.PageHeightTwips);
        Assert.True(info.Landscape);
    }

    [Fact]
    public void DEP033_Validation_LeavesXmlUntouchedAndRecordsNoUndo()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);
        var before = DocumentXml(session.Save()).ToString(SaveOptions.DisableFormatting);
        int undoBefore = session.UndoCount;

        // Opposing margins that leave no room on a Letter page.
        var r = session.SetPageSetup(anchor, new PageSetupOp { MarginLeftTwips = 7000, MarginRightTwips = 7000 });
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.InvalidPageSetup, r.Error!.Code);

        r = session.SetPageSetup(anchor, new PageSetupOp { MarginTopTwips = -1 });
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.InvalidPageSetup, r.Error!.Code);

        r = session.SetPageSetup(anchor, new PageSetupOp { PageWidthTwips = 0 });
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.InvalidPageSetup, r.Error!.Code);

        // Validation is against the merged section: a new narrower page makes existing margins invalid.
        r = session.SetPageSetup(anchor, new PageSetupOp { PageWidthTwips = 2000 });
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.InvalidPageSetup, r.Error!.Code);

        Assert.Equal(before, DocumentXml(session.Save()).ToString(SaveOptions.DisableFormatting));
        Assert.Equal(undoBefore, session.UndoCount);
    }

    [Fact]
    public void DEP034_HeaderFooterDistance_RoundTrip()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);
        var before = session.GetSectionInfo(anchor)!;
        Assert.Equal(720, before.HeaderDistanceTwips);
        Assert.Equal(720, before.FooterDistanceTwips);

        var r = session.SetPageSetup(anchor, new PageSetupOp { HeaderDistanceTwips = 360, FooterDistanceTwips = 480 });
        Assert.True(r.Success, r.Error?.Message);
        var pgMar = GoverningSectPr(session.Save()).Element(W + "pgMar")!;
        Assert.Equal("360", (string?)pgMar.Attribute(W + "header"));
        Assert.Equal("480", (string?)pgMar.Attribute(W + "footer"));
        // The margins the op did not name are written with their effective values, not dropped.
        Assert.Equal("1440", (string?)pgMar.Attribute(W + "top"));

        var info = session.GetSectionInfo(anchor)!;
        Assert.Equal(360, info.HeaderDistanceTwips);
        Assert.Equal(480, info.FooterDistanceTwips);
    }

    [Fact]
    public void DEP035_NoOp_RecordsNoUndoStep_AndEmptyOpIsNoOp()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);

        Assert.True(session.SetPageSetup(anchor, new PageSetupOp()).Success);
        Assert.Equal(0, session.UndoCount);

        var op = new PageSetupOp { MarginTopTwips = 900, Landscape = true };
        Assert.True(session.SetPageSetup(anchor, op).Success);
        Assert.Equal(1, session.UndoCount);
        Assert.True(session.SetPageSetup(anchor, op).Success);
        Assert.Equal(1, session.UndoCount);
    }

    [Fact]
    public void DEP036_PageSetupOp_WireRipple()
    {
        var op = DocxSessionJson.ParsePageSetupOp(
            "{\"pageWidthTwips\":15840,\"pageHeightTwips\":12240,\"landscape\":true,\"marginTopTwips\":720," +
            "\"marginBottomTwips\":720,\"marginLeftTwips\":720,\"marginRightTwips\":720," +
            "\"headerDistanceTwips\":360,\"footerDistanceTwips\":360}");
        Assert.Equal(15840, op.PageWidthTwips);
        Assert.Equal(12240, op.PageHeightTwips);
        Assert.True(op.Landscape);
        Assert.Equal(720, op.MarginTopTwips);
        Assert.Equal(360, op.HeaderDistanceTwips);
        Assert.Equal(360, op.FooterDistanceTwips);
        Assert.Null(DocxSessionJson.ParsePageSetupOp("{}").PageWidthTwips);
        Assert.Null(DocxSessionJson.ParsePageSetupOp("").Landscape);

        int handle = DocxSessionOps.OpenSession(BuildTwoParagraphsWithSection(), null);
        try
        {
            var session = SessionRegistry.Get(handle);
            var anchor = FirstBodyParagraph(session);
            var json = DocxSessionOps.SetPageSetup(handle, anchor, "{\"landscape\":true,\"headerDistanceTwips\":500}");
            Assert.Contains("\"success\":true", json);
            var info = DocxSessionOps.GetSectionInfo(handle, anchor);
            Assert.Contains("\"landscape\":true", info);
            Assert.Contains("\"pageWidthTwips\":15840", info);
            Assert.Contains("\"headerDistanceTwips\":500", info);

            json = DocxSessionOps.SetPageSetup(handle, anchor, "{\"marginLeftTwips\":-5}");
            Assert.Contains("invalid_page_setup", json);
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    // ─── ReplaceText keeps the trailing space ───────────────────────────

    [Fact]
    public void DEP040_ReplaceText_PreservesTrailingSpace()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var anchor = FirstBodyParagraph(session);

        var r = session.ReplaceText(anchor, "Page ");
        Assert.True(r.Success, r.Error?.Message);

        var xml = DocumentXml(session.Save()).ToString(SaveOptions.DisableFormatting);
        Assert.Contains("<w:t xml:space=\"preserve\">Page </w:t>", xml);
        Assert.Equal("Page ", DocumentXml(session.Save()).Descendants(W + "p").First().Value);

        // The projection keeps it too: a "Page " followed by an inserted field must not read "Page1".
        Assert.Contains("Page ", session.Project().Markdown);
        Assert.Equal("Page ", session.Grep("Page ").Single().Text);

        // Leading space and the tracked-change path preserve it as well.
        Assert.True(session.ReplaceText(anchor, " Page ").Success);
        Assert.Equal(" Page ", DocumentXml(session.Save()).Descendants(W + "p").First().Value);

        using var tracked = new DocxSession(BuildTwoParagraphsWithSection(),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var trackedAnchor = FirstBodyParagraph(tracked);
        Assert.True(tracked.ReplaceText(trackedAnchor, "Page ").Success);
        var insRuns = DocumentXml(tracked.Save()).Descendants(W + "ins").Descendants(W + "t").ToList();
        Assert.Contains(insRuns, t => t.Value == "Page ");
    }

    // ─── InsertPageNumberField: "Page X of Y" ───────────────────────────

    [Fact]
    public void DEP041_InsertPageNumberField_PageOfTotal_EmitsTextAndBothFields()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var body = FirstBodyParagraph(session);
        // A bold footer, so the composite has run properties to inherit.
        var footer = session.SetFooterText(body, HeaderFooterKind.Default, "**Draft**");
        Assert.True(footer.Success, footer.Error?.Message);
        var footerPara = footer.Created[0].Id;

        var r = session.InsertPageNumberField(footerPara, PageNumberField.PageOfTotal);
        Assert.True(r.Success, r.Error?.Message);

        byte[] saved = session.Save();
        XElement ftr;
        using (var ms = new MemoryStream(saved))
        using (var doc = WordprocessingDocument.Open(ms, false))
        {
            ftr = doc.MainDocumentPart!.FooterParts.Single().GetXDocument().Root!;
        }
        var para = ftr.Descendants(W + "p").First();
        var runs = para.Elements(W + "r").ToList();

        // Word's gallery shape: "Page " + PAGE field (5 runs) + " of " + NUMPAGES field (5 runs),
        // after the existing "Draft" run.
        string Sig(XElement run) =>
            run.Element(W + "fldChar") is { } fc ? "fld:" + (string?)fc.Attribute(W + "fldCharType")
            : run.Element(W + "instrText") is { } it ? "instr:" + it.Value.Trim()
            : "t:" + (run.Element(W + "t")?.Value ?? "");
        var expected = new[]
        {
            "t:Draft",
            "t:Page ", "fld:begin", "instr:PAGE", "fld:separate", "t:1", "fld:end",
            "t: of ", "fld:begin", "instr:NUMPAGES", "fld:separate", "t:1", "fld:end",
        };
        Assert.Equal(expected, runs.Select(Sig).ToArray());

        // The literal runs preserve their spaces and every appended run inherits the bold rPr.
        var pageText = runs[1].Element(W + "t")!;
        Assert.Equal("preserve", (string?)pageText.Attribute(XNamespace.Xml + "space"));
        Assert.All(runs.Skip(1), run => Assert.NotNull(run.Element(W + "rPr")?.Element(W + "b")));

        // The projection reads the cached results: "Page 1 of 1", spaces intact.
        var projection = session.Project();
        Assert.Contains("Page 1 of 1", projection.Markdown);
        Assert.Equal("DraftPage 1 of 1", string.Concat(para.Descendants(W + "t").Select(t => t.Value)));

        // The wire token round-trips through the parser used by every transport.
        Assert.Equal(PageNumberField.PageOfTotal, DocxSessionJson.ParsePageNumberField("pageOfTotal"));
        Assert.Equal(PageNumberField.PageOfTotal, DocxSessionJson.ParsePageNumberField("page_of_total"));
    }

    /// <summary>
    /// A zero-length span inserts a NEW run at a run boundary and steps outside a complex field's
    /// chrome: the browser editor types after "Page X of Y" and the keystrokes must land after the
    /// NUMPAGES end run, never inside the result run Word's next field update discards.
    /// </summary>
    [Fact]
    public void DEP042_ReplaceTextAtSpan_ZeroLength_InsertsOutsideFields()
    {
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var body = FirstBodyParagraph(session);
        var footer = session.SetFooterText(body, HeaderFooterKind.Default, "**Draft**");
        Assert.True(footer.Success, footer.Error?.Message);
        var footerPara = footer.Created[0].Id;
        Assert.True(session.InsertPageNumberField(footerPara, PageNumberField.PageOfTotal).Success);
        var length = session.Project().AnchorIndex[footerPara].TextPreview!.Length; // "DraftPage 1 of 1"

        // After the last field: the new run follows the NUMPAGES end run.
        var tail = session.ReplaceTextAtSpan(footerPara, length, 0, " (rev A)");
        Assert.True(tail.Success, tail.Error?.Message);
        // Before the first run: a new run at the very start.
        var head = session.ReplaceTextAtSpan(footerPara, 0, 0, "> ");
        Assert.True(head.Success, head.Error?.Message);
        // Strictly inside a run's text is not a boundary.
        var mid = session.ReplaceTextAtSpan(footerPara, 3, 0, "x");
        Assert.False(mid.Success);
        Assert.Equal(EditErrorCode.OffsetOutOfRange, mid.Error!.Code);

        var xml = XElement.Parse(session.Raw.GetXml(footerPara));
        var runs = xml.Elements(W + "r").ToList();
        string Sig(XElement run) =>
            run.Element(W + "fldChar") is { } fc ? "fld:" + (string?)fc.Attribute(W + "fldCharType")
            : run.Element(W + "instrText") is { } it ? "instr:" + it.Value.Trim()
            : "t:" + (run.Element(W + "t")?.Value ?? "");
        // The head insert sat against the plain "Draft" run, so it extends that run rather than
        // adding a sibling — a separate run's leading "> " would force the converter to emit the
        // space as &#160;. The tail insert sat against the NUMPAGES end field, so it stays its own
        // run outside the field. Both keep the bold of the run they joined / followed.
        Assert.Equal("t:> Draft", Sig(runs[0]));
        Assert.Equal("fld:end", Sig(runs[^2]));
        Assert.Equal("t: (rev A)", Sig(runs[^1]));
        Assert.NotNull(runs[0].Element(W + "rPr")?.Element(W + "b"));
        Assert.NotNull(runs[^1].Element(W + "rPr")?.Element(W + "b"));
        Assert.Equal("> DraftPage 1 of 1 (rev A)", session.Project().AnchorIndex[footerPara].TextPreview);

        // Tracked mode records the insertion as w:ins around the new run.
        session.SetTrackedChanges(TrackedChangeMode.RenderInline);
        var tracked = session.ReplaceTextAtSpan(footerPara, session.Project().AnchorIndex[footerPara].TextPreview!.Length, 0, "!");
        Assert.True(tracked.Success, tracked.Error?.Message);
        var last = XElement.Parse(session.Raw.GetXml(footerPara)).Elements().Last();
        Assert.Equal(W + "ins", last.Name);
        Assert.Equal("!", last.Element(W + "r")?.Element(W + "t")?.Value);
    }

    [Fact]
    public void DEP043_ZeroLengthInsert_AtPlainRunBoundary_ExtendsTheRun()
    {
        // Appending text after an ordinary run extends that run rather than dropping a sibling
        // beside it. A separate run whose text begins with a space forces the converter to render
        // that space as a non-breaking one (a run boundary is where HTML would collapse it), so a
        // plainly typed " world" came back with a &#160;. Coalescing keeps one run, one space.
        using var session = new DocxSession(BuildTwoParagraphsWithSection());
        var para = FirstBodyParagraph(session);
        Assert.True(session.ReplaceText(para, "Hello").Success);
        var end = session.Project().AnchorIndex[para].TextPreview!.Length;

        var appended = session.ReplaceTextAtSpan(para, end, 0, " world");
        Assert.True(appended.Success, appended.Error?.Message);

        var runs = XElement.Parse(session.Raw.GetXml(para)).Elements(W + "r").ToList();
        Assert.Single(runs);
        var text = runs[0].Element(W + "t");
        Assert.Equal("Hello world", text?.Value);
        Assert.Equal("preserve", (string?)text?.Attribute(XNamespace.Xml + "space"));
        Assert.Equal("Hello world", session.Project().AnchorIndex[para].TextPreview);
    }
}
