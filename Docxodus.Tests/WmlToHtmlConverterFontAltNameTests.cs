#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #670: standalone export reported `font_unavailable` for families that cannot exist —
/// `(normal text)` (Word's UI label for the theme font, written into `word/numbering.xml` as if it
/// were a family) and `TimesNewRomanPSMT` (a PostScript spelling).
///
/// <para>The document already answers for both. ECMA-376 §17.8.3.1's `w:altName` is exactly "use
/// this family when the primary one is unavailable", and `TestFiles/NVCA-Model-COI.docx` declares
/// `altName="Times New Roman"` for each of them in its font table. The converter emitted only the
/// primary name, so the renderer was asked for a family the package itself had already said to
/// substitute.</para>
/// </summary>
public class WmlToHtmlConverterFontAltNameTests
{
    private static readonly XNamespace Xh = "http://www.w3.org/1999/xhtml";

    private static readonly string NvcaPath =
        Path.Combine("../../../../TestFiles/", "NVCA-Model-COI.docx");

    private static XElement Convert(bool fabricateCssClasses = true) =>
        WmlToHtmlConverter.ConvertToHtml(
            new WmlDocument(Path.GetFullPath(NvcaPath)),
            new WmlToHtmlConverterSettings
            {
                FabricateCssClasses = fabricateCssClasses,
                RenderFootnotesAndEndnotes = true,
                RenderHeadersAndFooters = true,
            });

    /// <summary>Every `font-family` declaration in the rendered document, inline or in CSS.</summary>
    private static string AllFontFamilyDeclarations(XElement html)
    {
        var inline = html.Descendants()
            .Select(e => (string?)e.Attribute("style"))
            .OfType<string>();
        var css = html.Descendants(Xh + "style").Select(e => e.Value);
        return string.Join("\n", inline.Concat(css));
    }

    [Fact]
    public void TheFixtureDeclaresAltNamesForTheAliasesUnderTest()
    {
        // Guards the premise: if the fixture's font table ever stops carrying these, the
        // assertions below would be checking a document that never had the problem.
        using var ms = new MemoryStream(File.ReadAllBytes(Path.GetFullPath(NvcaPath)));
        using var zip = new System.IO.Compression.ZipArchive(ms);
        var fontTable = XDocument.Load(zip.GetEntry("word/fontTable.xml")!.Open());

        string? AltNameOf(string family) => fontTable.Root!.Elements(W.font)
            .Where(f => (string?)f.Attribute(W.name) == family)
            .Select(f => (string?)f.Element(W.altName)?.Attribute(W.val))
            .FirstOrDefault();

        Assert.Equal("Times New Roman", AltNameOf("(normal text)"));
        Assert.Equal("Times New Roman", AltNameOf("TimesNewRomanPSMT"));
        // Declared with no alternate — it stays an honest unresolved request rather than being
        // guessed at from its name.
        Assert.Null(AltNameOf("Times New Roman Bold"));
    }

    [Fact]
    public void AnAliasedFamilyCarriesTheDocumentsDeclaredFallbackInItsStack()
    {
        var declarations = AllFontFamilyDeclarations(Convert());

        // Wherever the alias is requested, the family the package nominates follows it — so the
        // renderer has something real to land on instead of only a name that cannot resolve.
        Assert.Contains("'TimesNewRomanPSMT', 'Times New Roman'", declarations, StringComparison.Ordinal);
    }

    [Fact]
    public void NoAliasedFamilyIsRequestedWithoutItsDeclaredFallback()
    {
        var declarations = AllFontFamilyDeclarations(Convert());

        foreach (var alias in new[] { "TimesNewRomanPSMT", "(normal text)" })
        {
            var orphaned = declarations
                .Split('\n')
                .Where(line => line.Contains($"'{alias}'", StringComparison.Ordinal)
                    && !line.Contains($"'{alias}', 'Times New Roman'", StringComparison.Ordinal))
                .ToList();

            Assert.True(
                orphaned.Count == 0,
                $"{alias} was requested without its declared alternate in: "
                    + string.Join(" | ", orphaned.Take(3)));
        }
    }

    [Fact]
    public void AFamilyWithNoDeclaredAlternateIsLeftAlone()
    {
        var declarations = AllFontFamilyDeclarations(Convert());

        // Calibri has no w:altName. It must keep its plain stack — the generic fallback only —
        // rather than acquire an invented alternate.
        Assert.Contains("'Calibri', sans-serif", declarations, StringComparison.Ordinal);
    }

    [Fact]
    public void RunningStoriesResolveAltNamesToo()
    {
        // A note's content resolves converter annotations through its own part root, so a fix
        // annotated onto the body alone would silently skip every non-body story. Rendered with
        // class fabrication off so each element's stack is readable on the element itself.
        var html = Convert(fabricateCssClasses: false);

        var noteDeclarations = html.Descendants()
            .Where(e => e.AncestorsAndSelf().Any(a =>
                (string?)a.Attribute("class") is "footnotes" or "endnotes"))
            .Select(e => (string?)e.Attribute("style"))
            .OfType<string>()
            .ToList();

        Assert.NotEmpty(noteDeclarations);
        var aliased = noteDeclarations
            .Where(d => d.Contains("'TimesNewRomanPSMT'", StringComparison.Ordinal))
            .ToList();
        Assert.NotEmpty(aliased);
        Assert.All(aliased, d =>
            Assert.Contains("'TimesNewRomanPSMT', 'Times New Roman'", d, StringComparison.Ordinal));
    }
}
