#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Reference-field authoring on <see cref="DocxSession"/> (issue #607): the table of contents,
/// table of figures and table of authorities that narrowing the library to the DOCX toolchain took
/// away with <c>ReferenceAdder</c>.
/// </summary>
/// <remarks>
/// The point of these ops is that the caller never writes a switch string, so the tests assert the
/// switch string: a malformed one renders as <em>nothing</em> in Word, silently, and no schema
/// check catches it. Test IDs use the DS44x range.
/// </remarks>
public class DocxSessionReferenceFieldTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static DocxSession Session() =>
        new(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { PersistAnchorIds = false });

    private static string FirstBody(DocxSession session) =>
        session.Project().AnchorIndex.Keys.First(k => k.StartsWith("p:body:", StringComparison.Ordinal));

    private static XElement Body(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.GetXDocument().Root!;
    }

    /// <summary>The field instruction of the one field in the document.</summary>
    private static string Instruction(byte[] bytes) =>
        string.Concat(Body(bytes).Descendants(W + "instrText").Select(t => (string)t)).Trim();

    [Fact]
    public void DS440_TableOfContents_IsADirtyTocFieldInWordsContentControl()
    {
        using var session = Session();
        var anchor = FirstBody(session);

        var result = session.InsertTableOfContents(anchor, Position.Before);

        Assert.True(result.Success, result.Error?.Message);
        var saved = session.Save(persistAnchorIds: false);
        var body = Body(saved);

        // Word's own wrapper — this is what puts an "Update Table" control on it.
        var sdt = Assert.Single(body.Descendants(W + "sdt"));
        Assert.Equal("Table of Contents",
            (string?)sdt.Descendants(W + "docPartGallery").Single().Attribute(W + "val"));

        // The switches the typed options mean, in Word's order.
        Assert.Equal("TOC \\o \"1-3\" \\h \\z \\u", Instruction(saved));

        // Dirty, and with no cached result between separate and end: Word fills it, we do not.
        var fldChars = body.Descendants(W + "fldChar").ToList();
        Assert.Equal(new[] { "begin", "separate", "end" },
            fldChars.Select(f => (string?)f.Attribute(W + "fldCharType")));
        Assert.Equal("true", (string?)fldChars[0].Attribute(W + "dirty"));

        // …so the document has to ask Word to update fields on open, or the reader sees nothing.
        using var ms = new MemoryStream(saved);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Equal("true", (string?)doc.MainDocumentPart!.DocumentSettingsPart!
            .GetXDocument().Root!.Element(W + "updateFields")!.Attribute(W + "val"));
    }

    [Theory]
    [InlineData("1-3", "TOC \\o \"1-3\" \\h \\z \\u")]
    [InlineData("2", "TOC \\o \"2-2\" \\h \\z \\u")]
    [InlineData(" 1 - 9 ", "TOC \\o \"1-9\" \\h \\z \\u")]
    public void DS441_TocLevels_AreNormalizedIntoTheSwitch(string levels, string expected)
    {
        using var session = Session();
        Assert.True(session.InsertTableOfContents(
            FirstBody(session), Position.Before, new TableOfContentsOptions { Levels = levels }).Success);
        Assert.Equal(expected, Instruction(session.Save(persistAnchorIds: false)));
    }

    [Fact]
    public void DS442_TocSwitchesFollowTheirOptions()
    {
        using var session = Session();
        Assert.True(session.InsertTableOfContents(FirstBody(session), Position.Before,
            new TableOfContentsOptions
            {
                Levels = "1-2",
                Hyperlinks = false,
                HideTabAndPageNumbersInWeb = false,
                UseOutlineLevels = false,
                Title = null,
            }).Success);

        var saved = session.Save(persistAnchorIds: false);
        Assert.Equal("TOC \\o \"1-2\"", Instruction(saved));
        // Title = null means no heading paragraph at all.
        Assert.DoesNotContain(Body(saved).Descendants(W + "pStyle"),
            p => (string?)p.Attribute(W + "val") == "TOCHeading");
    }

    [Theory]
    [InlineData("0-3")]
    [InlineData("3-1")]
    [InlineData("1-10")]
    [InlineData("one")]
    [InlineData("")]
    public void DS443_MalformedLevels_AreRefusedWithoutTouchingTheDocument(string levels)
    {
        using var session = Session();
        var anchor = FirstBody(session);
        var before = session.GetPackageContentHash();
        var undoCount = session.UndoCount;

        var result = session.InsertTableOfContents(
            anchor, Position.Before, new TableOfContentsOptions { Levels = levels });

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.InvalidReferenceField, result.Error!.Code);
        // Nothing written, and no undo entry burned on a rejected call.
        Assert.Equal(before, session.GetPackageContentHash());
        Assert.Equal(undoCount, session.UndoCount);
    }

    [Fact]
    public void DS444_TableOfFiguresAndAuthorities_CarryTheirOwnSwitchesAndStyles()
    {
        using var session = Session();
        var anchor = FirstBody(session);

        Assert.True(session.InsertTableOfFigures(anchor, Position.Before,
            new TableOfFiguresOptions { CaptionLabel = "Exhibit" }).Success);
        var tof = session.Save(persistAnchorIds: false);
        Assert.Equal("TOC \\c \"Exhibit\" \\h", Instruction(tof));
        Assert.Contains(Body(tof).Descendants(W + "pStyle"),
            p => (string?)p.Attribute(W + "val") == "TableofFigures");
        // Word writes a table of figures as a bare paragraph, not inside a content control.
        Assert.Empty(Body(tof).Descendants(W + "sdt"));

        Assert.True(session.Undo());

        Assert.True(session.InsertTableOfAuthorities(anchor, Position.Before,
            new TableOfAuthoritiesOptions
            {
                Category = AuthorityCategory.Statutes,
                EntryPageSeparator = ", ",
            }).Success);
        var toa = session.Save(persistAnchorIds: false);
        Assert.Equal("TOA \\c \"2\" \\h \\e \", \"", Instruction(toa));
        Assert.Contains(Body(toa).Descendants(W + "pStyle"),
            p => (string?)p.Attribute(W + "val") == "TableofAuthorities");
    }

    [Fact]
    public void DS445_InsertedTables_AreSchemaValidAndUndoable()
    {
        using var session = Session();
        var anchor = FirstBody(session);
        var before = session.Save(persistAnchorIds: false);

        Assert.True(session.InsertTableOfContents(anchor, Position.Before).Success);
        Assert.True(session.InsertTableOfFigures(anchor, Position.After).Success);
        Assert.True(session.InsertTableOfAuthorities(anchor, Position.After).Success);

        using (var ms = new MemoryStream(session.Save(persistAnchorIds: false)))
        using (var doc = WordprocessingDocument.Open(ms, false))
        {
            var errors = new OpenXmlValidator().Validate(doc).ToList();
            Assert.True(errors.Count == 0,
                string.Join(" | ", errors.Take(3).Select(e => e.Description)));
        }

        // Each op is exactly one undo step, and three undos restore the original package byte for
        // byte. The comparison is on the SAVED package rather than the live checkpoint hash: an op
        // that flushes a non-projected part (here styles and settings) leaves the flushed stream
        // behind in the live package even after undo restores the cache, so the checkpoint hash
        // does not return — InsertFootnote, which also ensures styles, behaves identically. What a
        // caller sees is the save, and the save returns.
        Assert.Equal(3, session.UndoCount);
        Assert.True(session.Undo());
        Assert.True(session.Undo());
        Assert.True(session.Undo());
        Assert.Equal(before, session.Save(persistAnchorIds: false));
    }

    /// <summary>
    /// A generated table is regenerated wholesale by Word on every field update, so there is no
    /// reversible way to redline it. Refuse under recording rather than write a mark rejection
    /// cannot take back — the shape #614 established for note insertion.
    /// </summary>
    [Fact]
    public void DS446_UnderTrackedChangeRecording_TheOpsRefuseWithoutMutating()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings
            {
                PersistAnchorIds = false,
                TrackedChanges = TrackedChangeMode.RenderInline,
            });
        var anchor = FirstBody(session);
        var before = session.GetPackageContentHash();

        foreach (var result in new[]
        {
            session.InsertTableOfContents(anchor, Position.Before),
            session.InsertTableOfFigures(anchor, Position.Before),
            session.InsertTableOfAuthorities(anchor, Position.Before),
        })
        {
            Assert.False(result.Success);
            Assert.Equal(EditErrorCode.TrackedOperationUnsupported, result.Error!.Code);
        }

        Assert.Equal(before, session.GetPackageContentHash());
    }

    /// <summary>Word does not generate a reference table inside a running story or a note, so a
    /// non-body anchor is refused rather than silently written somewhere Word will not fill.</summary>
    [Fact]
    public void DS447_ANonBodyAnchor_IsRefused()
    {
        using var session = Session();
        var body = FirstBody(session);
        var header = session.SetHeaderText(body, HeaderFooterKind.Default, "Running head.");
        Assert.True(header.Success, header.Error?.Message);
        var headerAnchor = header.Created.First(a => a.Scope.StartsWith("hdr", StringComparison.Ordinal));

        var result = session.InsertTableOfContents(headerAnchor.Id, Position.Before);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, result.Error!.Code);
    }

    /// <summary>A document that already defines the styles keeps its own — a firm's house TOC
    /// formatting must survive inserting a table of contents.</summary>
    [Fact]
    public void DS448_ExistingReferenceStyles_AreLeftAlone()
    {
        using var session = Session();
        var anchor = FirstBody(session);
        Assert.True(session.InsertTableOfContents(anchor, Position.Before).Success);

        var firstPass = StylesXml(session.Save(persistAnchorIds: false));
        Assert.True(session.InsertTableOfContents(anchor, Position.Before).Success);
        var secondPass = StylesXml(session.Save(persistAnchorIds: false));

        // The second insert finds the styles already there and adds nothing.
        Assert.Equal(
            firstPass.Elements(W + "style").Count(),
            secondPass.Elements(W + "style").Count());
    }

    private static XElement StylesXml(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.StyleDefinitionsPart!.GetXDocument().Root!;
    }
}
