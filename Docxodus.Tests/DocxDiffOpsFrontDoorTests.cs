// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #961: the browser's typed comparison exports (<c>DocumentComparer</c>) route through
/// <see cref="DocxDiffOps.CompareFrontDoor"/> and <see cref="DocxDiffOps.CompareToHtml"/>. These pin
/// what that buys: one author default, deterministic dates, and the same bytes as the .NET front door.
/// </summary>
public class DocxDiffOpsFrontDoorTests
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");
    private static byte[] Wc(string name) => File.ReadAllBytes(Path.Combine(TestFilesDir.FullName, "WC", name));

    private static readonly byte[] Left = Wc("WC001-Digits.docx");
    private static readonly byte[] Right = Wc("WC001-Digits-Mod.docx");

    private static string[] RevisionAuthors(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var w = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        return doc.MainDocumentPart!.GetXDocument().Descendants()
            .Where(e => e.Name == w + "ins" || e.Name == w + "del")
            .Select(e => (string?)e.Attribute(w + "author") ?? "")
            .Distinct()
            .ToArray();
    }

    [Fact]
    public void Default_author_is_Docxodus_on_the_public_settings_and_the_engine()
    {
        Assert.Equal("Docxodus", new DocxDiffSettings().AuthorForRevisions);
        Assert.Equal(new DocxDiffSettings().AuthorForRevisions, new Docxodus.Ir.Diff.IrDiffSettings().AuthorForRevisions);
    }

    [Fact]
    public void A_compare_naming_no_author_stamps_Docxodus()
    {
        var redline = DocxCompare.Compare(new WmlDocument("l.docx", Left), new WmlDocument("r.docx", Right));

        Assert.Equal(new[] { "Docxodus" }, RevisionAuthors(redline.DocumentByteArray));
    }

    [Theory]
    [InlineData(null, "Docxodus")]
    [InlineData("", "Docxodus")]
    [InlineData("Reviewer", "Reviewer")]
    public void Front_door_settings_take_the_core_author_unless_one_is_named(string? author, string expected)
    {
        var redline = DocxDiffOps.CompareFrontDoor(Left, Right, DocxDiffOps.FrontDoorSettings(author, caseInsensitive: false));

        Assert.Equal(new[] { expected }, RevisionAuthors(redline));
    }

    [Fact]
    public void Front_door_compare_is_byte_identical_to_DocxCompare_and_repeatable()
    {
        var settings = DocxDiffOps.FrontDoorSettings("Reviewer", caseInsensitive: true);

        var first = DocxDiffOps.CompareFrontDoor(Left, Right, settings);
        var second = DocxDiffOps.CompareFrontDoor(Left, Right, settings);
        var dotnet = DocxCompare.Compare(
            new WmlDocument("l.docx", Left),
            new WmlDocument("r.docx", Right),
            new DocxDiffSettings { AuthorForRevisions = "Reviewer", CaseInsensitive = true });

        // The browser used to stamp DateTime.UtcNow, so two runs never matched.
        Assert.Equal(first, second);
        Assert.Equal(dotnet.DocumentByteArray, first);
    }

    [Fact]
    public void Front_door_compare_matches_the_raw_facade_with_the_front_door_policy()
    {
        // The front door is DocxDiff plus pre-accepting the inputs' revisions; on these
        // revision-free inputs the raw facade with that policy produces the same package.
        var frontDoor = DocxDiffOps.CompareFrontDoor(Left, Right, DocxDiffOps.FrontDoorSettings(null, caseInsensitive: false));
        var raw = DocxDiffOps.Compare(Left, Right, "{\"preAcceptInputRevisions\":true}");

        Assert.Equal(raw, frontDoor);
    }

    [Fact]
    public void Compare_to_html_renders_the_front_door_redline_through_HtmlConversionOps()
    {
        var settings = DocxDiffOps.FrontDoorSettings("Reviewer", caseInsensitive: false);

        var html = DocxDiffOps.CompareToHtml(Left, Right, settings, renderTrackedChanges: true);
        var expected = HtmlConversionOps.ConvertToHtml(
            DocxDiffOps.CompareFrontDoor(Left, Right, settings),
            new HtmlConversionOptions
            {
                PageTitle = "Document Comparison",
                CssClassPrefix = "redline-",
                RenderTrackedChanges = true,
                AuthorColors = new System.Collections.Generic.Dictionary<string, string> { ["Reviewer"] = "#007bff" },
            });

        Assert.Equal(expected, html);
        Assert.Contains("[data-author=\"Reviewer\"]", html);
        Assert.Contains("<ins", html);
    }

    [Fact]
    public void Compare_to_html_without_tracked_changes_shows_the_accepted_text()
    {
        var html = DocxDiffOps.CompareToHtml(Left, Right, DocxDiffOps.FrontDoorSettings(null, false), renderTrackedChanges: false);

        Assert.DoesNotContain("<ins", html);
        Assert.DoesNotContain("<del", html);
        Assert.DoesNotContain("data-author=", html);
    }

    [Fact]
    public void Missing_input_is_an_argument_error()
    {
        Assert.Throws<System.ArgumentException>(() =>
            DocxDiffOps.CompareFrontDoor(System.Array.Empty<byte>(), Right, new DocxDiffSettings()));
    }
}
